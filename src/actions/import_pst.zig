const std = @import("std");
const tui = @import("zigtui");
const Session = @import("../session.zig").Session;
const input = @import("../cli/input.zig");
const menu = @import("../cli/menu.zig");
const file_browser = @import("../cli/file_browser.zig");
const pst_tree_selector = @import("../cli/pst_tree_selector.zig");
const scripts = @import("../scripts/scripts.zig");
const import_progress = @import("import_progress.zig");

const SCAN_TEXT_MAX: usize = 320;

const ScanViewState = struct {
    pst_path: []const u8,
    phase: [SCAN_TEXT_MAX]u8 = undefined,
    phase_len: usize = 0,
    folder: [SCAN_TEXT_MAX]u8 = undefined,
    folder_len: usize = 0,
    log: [SCAN_TEXT_MAX]u8 = undefined,
    log_len: usize = 0,
    percent: u8 = 0,
    total_folders: usize = 0,
    scanned_folders: usize = 0,
    current_item_count: usize = 0,
    accumulated_items: usize = 0,
    pst_size_bytes: u64 = 0,
    elapsed_ms: u64 = 0,
    spinner_idx: usize = 0,
};

const PstFolder = struct {
    path: []u8,
    item_count: usize,
    year_summary: ?[]u8 = null,
};

/// Ejecuta el wizard de importacion de PST. Devuelve exit code del script
/// (0 = exito) o un codigo de error del wizard.
pub fn run(session: Session, allocator: std.mem.Allocator) !u8 {
    std.debug.print("\n\x1b[1;36m== Importar correos desde PST ==\x1b[0m\n", .{});
    session.print();
    std.debug.print("\n", .{});

    // 1) Ruta del PST
    const pst_path = try askPstPath(allocator);
    defer allocator.free(pst_path);

    // 2) Seleccion de carpetas del PST (multi-seleccion)
    const selected_folders = try askPstFolders(allocator, pst_path);
    defer {
        for (selected_folders) |p| allocator.free(p);
        allocator.free(selected_folders);
    }

    // 3) Accion Copy / Move
    const action = askAction() orelse {
        std.debug.print("\n\x1b[33m[!] Importacion cancelada.\x1b[0m\n", .{});
        return 2;
    };

    // 4) Filtro por anio
    const filter_year = try askFilterYear(allocator);
    defer if (filter_year) |fy| allocator.free(fy);

    // 5) Filtro por meses (solo si hay anio)
    var filter_months: ?[]u8 = null;
    defer if (filter_months) |fm| allocator.free(fm);
    if (filter_year != null) {
        filter_months = try askFilterMonths(allocator);
    }

    // 6) Saltar duplicados
    const skip_dupes = input.confirm("\xc2\xbfSaltar duplicados (Message-Id/Search-Key + Asunto + Tamano en bytes)?", false);

    // 7) Resumen y confirmacion
    printSummary(session, pst_path, selected_folders, action, filter_year, filter_months, skip_dupes);
    if (!input.confirm("\xc2\xbfEjecutar la importacion con estos parametros?", true)) {
        std.debug.print("\n\x1b[33m[!] Importacion cancelada.\x1b[0m\n", .{});
        return 2;
    }

    // 8) Ejecutar PowerShell heredando stdio para ver progreso
    return try executePowerShell(allocator, session, pst_path, selected_folders, action, filter_year, filter_months, skip_dupes);
}

fn askPstFolders(allocator: std.mem.Allocator, pst_path: []const u8) ![][]u8 {
    const folders = try listPstFolders(allocator, pst_path);
    defer {
        for (folders) |f| {
            allocator.free(f.path);
            if (f.year_summary) |ys| allocator.free(ys);
        }
        allocator.free(folders);
    }

    if (folders.len == 0) {
        std.debug.print("\x1b[31m[X]\x1b[0m  El PST no contiene carpetas importables.\n", .{});
        return error.NoFoldersFound;
    }

    var entries = try allocator.alloc(pst_tree_selector.FolderEntry, folders.len);
    defer allocator.free(entries);

    for (folders, 0..) |f, i| {
        entries[i] = .{
            .path = f.path,
            .item_count = f.item_count,
            .year_summary = f.year_summary,
        };
    }

    const idxs = try pst_tree_selector.select(allocator, entries) orelse {
        return error.UserCancelled;
    };
    defer allocator.free(idxs);

    var out = try allocator.alloc([]u8, idxs.len);
    errdefer {
        for (out[0..]) |p| allocator.free(p);
        allocator.free(out);
    }
    for (idxs, 0..) |idx, j| {
        out[j] = try allocator.dupe(u8, folders[idx].path);
    }
    return out;
}

fn listPstFolders(allocator: std.mem.Allocator, pst_path: []const u8) ![]PstFolder {
    const script_path = try scripts.getScriptPath(allocator, .import_pst);
    defer allocator.free(script_path);

    var list = std.ArrayList(PstFolder){};
    defer list.deinit(allocator);

    var child = std.process.Child.init(&.{
        "powershell",
        "-NoProfile",
        "-ExecutionPolicy",
        "Bypass",
        "-File",
        script_path,
        "-PstPath",
        pst_path,
        "-ListFolders",
        "-Json",
    }, allocator);
    child.stdin_behavior = .Ignore;
    child.stdout_behavior = .Pipe;
    child.stderr_behavior = .Pipe;

    child.spawn() catch return error.ListFoldersFailed;

    var scan_state = ScanViewState{ .pst_path = pst_path };
    scan_state.phase_len = copyText(&scan_state.phase, "Preparando escaneo...");
    scan_state.log_len = copyText(&scan_state.log, "Leyendo estructura del PST");

    var backend = try tui.init(allocator);
    defer backend.deinit();

    var terminal = try tui.Terminal.init(allocator, backend.interface());
    defer terminal.deinit();

    try terminal.hideCursor();
    defer terminal.showCursor() catch {};

    try terminal.draw(&scan_state, renderScanView);

    const stdout_file = child.stdout orelse return error.ListFoldersFailed;
    var pending_stdout = std.ArrayList(u8){};
    defer pending_stdout.deinit(allocator);

    var read_chunk: [4096]u8 = undefined;
    var child_error_msg: ?[]u8 = null;
    defer if (child_error_msg) |msg| allocator.free(msg);

    while (true) {
        const read_n = try stdout_file.read(&read_chunk);
        if (read_n == 0) break;
        try pending_stdout.appendSlice(allocator, read_chunk[0..read_n]);

        while (std.mem.indexOfScalar(u8, pending_stdout.items, '\n')) |line_end| {
            const line = pending_stdout.items[0..line_end];
            const trimmed = std.mem.trim(u8, line, " \r\n\t");
            if (trimmed.len > 0 and trimmed[0] == '{') {
                var parsed = std.json.parseFromSlice(std.json.Value, allocator, trimmed, .{}) catch {
                    const consumed = line_end + 1;
                    const rem = pending_stdout.items[consumed..];
                    std.mem.copyForwards(u8, pending_stdout.items[0..rem.len], rem);
                    pending_stdout.items.len = rem.len;
                    continue;
                };
                defer parsed.deinit();

                if (parsed.value == .object) {
                    const obj = parsed.value.object;
                    const type_val = obj.get("type");
                    if (type_val != null and type_val.? == .string) {
                        if (std.mem.eql(u8, type_val.?.string, "scanMeta")) {
                            scan_state.total_folders = parsePositiveUsize(obj.get("totalFolders"));
                            scan_state.pst_size_bytes = parsePositiveU64(obj.get("pstSizeBytes"));
                            scan_state.log_len = copyText(&scan_state.log, "Escaneando carpetas y conteos por anio...");
                            scan_state.spinner_idx +%= 1;
                            try terminal.draw(&scan_state, renderScanView);
                        } else if (std.mem.eql(u8, type_val.?.string, "scanProgress")) {
                            const phase_val = obj.get("phase");
                            if (phase_val != null and phase_val.? == .string) {
                                scan_state.phase_len = copyText(&scan_state.phase, phase_val.?.string);
                            }
                            const folder_val = obj.get("folderPath");
                            if (folder_val != null and folder_val.? == .string and folder_val.?.string.len > 0) {
                                scan_state.folder_len = copyText(&scan_state.folder, folder_val.?.string);
                            }
                            scan_state.percent = parsePercent(obj.get("percent"));
                            scan_state.scanned_folders = parsePositiveUsize(obj.get("scannedFolders"));
                            scan_state.total_folders = parsePositiveUsizeOr(scan_state.total_folders, obj.get("totalFolders"));
                            scan_state.current_item_count = parsePositiveUsize(obj.get("currentItemCount"));
                            scan_state.accumulated_items = parsePositiveUsize(obj.get("accumulatedItems"));
                            scan_state.elapsed_ms = parsePositiveU64(obj.get("elapsedMs"));
                            scan_state.pst_size_bytes = parsePositiveU64Or(scan_state.pst_size_bytes, obj.get("pstSizeBytes"));
                            scan_state.spinner_idx +%= 1;
                            try terminal.draw(&scan_state, renderScanView);
                        } else if (std.mem.eql(u8, type_val.?.string, "log")) {
                            const msg_val = obj.get("message");
                            if (msg_val != null and msg_val.? == .string and msg_val.?.string.len > 0) {
                                scan_state.log_len = copyText(&scan_state.log, msg_val.?.string);
                                scan_state.spinner_idx +%= 1;
                                try terminal.draw(&scan_state, renderScanView);
                            }
                        } else if (std.mem.eql(u8, type_val.?.string, "error")) {
                            const msg_val = obj.get("message");
                            if (msg_val != null and msg_val.? == .string and msg_val.?.string.len > 0) {
                                if (child_error_msg) |old_msg| allocator.free(old_msg);
                                child_error_msg = try allocator.dupe(u8, msg_val.?.string);
                            }
                        } else if (std.mem.eql(u8, type_val.?.string, "folder")) {
                            const path_val = obj.get("path");
                            if (path_val != null and path_val.? == .string) {
                                const count: usize = parsePositiveUsize(obj.get("itemCount"));
                                const undated_count = parseUndatedCount(obj.get("undatedCount"));
                                const year_summary = blk: {
                                    const base = try parseYearSummary(allocator, obj.get("yearBreakdown"));
                                    if (undated_count == 0) break :blk base;
                                    if (base) |b| {
                                        const with_undated = try std.fmt.allocPrint(allocator, "{s} · sin_fecha({d})", .{ b, undated_count });
                                        allocator.free(b);
                                        break :blk with_undated;
                                    }
                                    break :blk try std.fmt.allocPrint(allocator, "Años: sin_fecha({d})", .{undated_count});
                                };

                                if (count == 0) {
                                    if (year_summary) |ys| allocator.free(ys);
                                } else {
                                    try list.append(allocator, .{
                                        .path = try allocator.dupe(u8, path_val.?.string),
                                        .item_count = count,
                                        .year_summary = year_summary,
                                    });
                                }
                            }
                        }
                    }
                }
            }

            const consumed = line_end + 1;
            const rem = pending_stdout.items[consumed..];
            std.mem.copyForwards(u8, pending_stdout.items[0..rem.len], rem);
            pending_stdout.items.len = rem.len;
        }
    }

    if (pending_stdout.items.len > 0) {
        const trimmed = std.mem.trim(u8, pending_stdout.items, " \r\n\t");
        if (trimmed.len > 0 and trimmed[0] == '{') {
            var parsed = std.json.parseFromSlice(std.json.Value, allocator, trimmed, .{}) catch null;
            if (parsed) |*p| {
                defer p.deinit();
                if (p.value == .object) {
                    const obj = p.value.object;
                    const type_val = obj.get("type");
                    if (type_val != null and type_val.? == .string and std.mem.eql(u8, type_val.?.string, "folder")) {
                        const path_val = obj.get("path");
                        if (path_val != null and path_val.? == .string) {
                            const count: usize = parsePositiveUsize(obj.get("itemCount"));
                            const undated_count = parseUndatedCount(obj.get("undatedCount"));
                            const year_summary = blk: {
                                const base = try parseYearSummary(allocator, obj.get("yearBreakdown"));
                                if (undated_count == 0) break :blk base;
                                if (base) |b| {
                                    const with_undated = try std.fmt.allocPrint(allocator, "{s} · sin_fecha({d})", .{ b, undated_count });
                                    allocator.free(b);
                                    break :blk with_undated;
                                }
                                break :blk try std.fmt.allocPrint(allocator, "Años: sin_fecha({d})", .{undated_count});
                            };

                            if (count == 0) {
                                if (year_summary) |ys| allocator.free(ys);
                            } else {
                                try list.append(allocator, .{
                                    .path = try allocator.dupe(u8, path_val.?.string),
                                    .item_count = count,
                                    .year_summary = year_summary,
                                });
                            }
                        }
                    }
                }
            }
        }
    }

    const stderr_text = blk: {
        if (child.stderr) |stderr_file| {
            break :blk try stderr_file.readToEndAlloc(allocator, 128 * 1024);
        }
        break :blk try allocator.dupe(u8, "");
    };
    defer allocator.free(stderr_text);

    const term = try child.wait();
    switch (term) {
        .Exited => |code| {
            if (code != 0) {
                std.debug.print("\x1b[2J\x1b[H", .{});
                if (child_error_msg) |msg| {
                    std.debug.print("\n\x1b[31mError listando carpetas PST:\x1b[0m\n{s}\n", .{msg});
                } else if (stderr_text.len > 0) {
                    std.debug.print("\n\x1b[31mError listando carpetas PST:\x1b[0m\n{s}\n", .{stderr_text});
                }
                return error.ListFoldersFailed;
            }
        },
        else => return error.ListFoldersFailed,
    }

    std.debug.print("\x1b[2J\x1b[H", .{});

    return try list.toOwnedSlice(allocator);
}

fn renderScanView(state: *ScanViewState, buf: *tui.Buffer) anyerror!void {
    const spinner = [_][]const u8{ "|", "/", "-", "\\" };
    const spin = spinner[state.spinner_idx % spinner.len];

    var size_buf: [32]u8 = undefined;
    const size_txt = formatBytes(&size_buf, state.pst_size_bytes);

    var elapsed_buf: [24]u8 = undefined;
    const elapsed_txt = formatElapsed(&elapsed_buf, state.elapsed_ms);

    const area = buf.getArea();
    if (area.width < 48 or area.height < 14) return;

    const popup_area = tui.centeredRectPct(area, 88, 72);
    const popup = tui.widgets.Popup{
        .title = " Escaneo de PST ",
        .border_style = .{ .fg = .cyan },
        .backdrop_style = .{ .fg = .dark_gray },
        .show_backdrop = true,
    };
    popup.render(popup_area, buf);

    const inner = tui.widgets.Popup.innerArea(popup_area);
    if (inner.width < 32 or inner.height < 10) return;

    var title_buf: [96]u8 = undefined;
    const title = std.fmt.bufPrint(&title_buf, "Escaneando estructura y metricas {s}", .{spin}) catch "Escaneando";
    buf.setStringTruncated(inner.x, inner.y, title, inner.width, .{ .fg = .light_cyan, .modifier = .{ .bold = true } });
    buf.setStringTruncated(inner.x, inner.y + 1, state.pst_path, inner.width, .{ .fg = .gray });

    var stats1_buf: [128]u8 = undefined;
    const stats1 = std.fmt.bufPrint(&stats1_buf, "Tamano: {s} | Tiempo: {s}", .{ size_txt, elapsed_txt }) catch "";
    buf.setStringTruncated(inner.x, inner.y + 2, stats1, inner.width, .{ .fg = .white });

    var stats2_buf: [128]u8 = undefined;
    const stats2 = std.fmt.bufPrint(
        &stats2_buf,
        "Carpetas: {d}/{d} | Item actual: {d} | Acumulados: {d}",
        .{ state.scanned_folders, state.total_folders, state.current_item_count, state.accumulated_items },
    ) catch "";
    buf.setStringTruncated(inner.x, inner.y + 3, stats2, inner.width, .{ .fg = .white });

    const graph_y = inner.y + 4;
    const graph_h: u16 = if (inner.height > 9) 4 else 2;
    const graph_w = inner.width;
    if (graph_h >= 2 and graph_w >= 10) {
        const graph_area = tui.Rect{ .x = inner.x, .y = graph_y, .width = graph_w, .height = graph_h };
        var canvas = tui.widgets.Canvas.init(graph_area, buf, .{ .fg = .white });
        canvas.drawBox(.{ .x = 0, .y = 0, .width = graph_w, .height = graph_h }, .{ .fg = .dark_gray });

        const usable_w: u16 = graph_w -| 2;
        const fill: u16 = @intCast((@as(usize, state.percent) * usable_w) / 100);
        const bar_y: u16 = if (graph_h > 2) graph_h / 2 else 1;
        if (fill > 0) {
            canvas.drawHLine(1, fill, bar_y, '█', .{ .fg = .green });
        }
        if (fill < usable_w) {
            canvas.drawHLine(1 + fill, usable_w - fill, bar_y, '░', .{ .fg = .dark_gray });
        }

        var pct_buf: [16]u8 = undefined;
        const pct = std.fmt.bufPrint(&pct_buf, "{d}%", .{state.percent}) catch "";
        const pct_x: u16 = if (graph_w > pct.len) (graph_w - @as(u16, @intCast(pct.len))) / 2 else 1;
        canvas.drawText(pct_x, 0, pct, .{ .fg = .yellow, .modifier = .{ .bold = true } });
    }

    var line_y = graph_y + graph_h;
    if (line_y < inner.y + inner.height) {
        if (state.phase_len > 0) {
            var phase_buf: [SCAN_TEXT_MAX + 16]u8 = undefined;
            const phase = std.fmt.bufPrint(&phase_buf, "Fase: {s}", .{state.phase[0..state.phase_len]}) catch "";
            buf.setStringTruncated(inner.x, line_y, phase, inner.width, .{ .fg = .light_white });
            line_y += 1;
        }
    }
    if (line_y < inner.y + inner.height) {
        if (state.folder_len > 0) {
            var folder_buf: [SCAN_TEXT_MAX + 24]u8 = undefined;
            const folder = std.fmt.bufPrint(&folder_buf, "Carpeta: {s}", .{state.folder[0..state.folder_len]}) catch "";
            buf.setStringTruncated(inner.x, line_y, folder, inner.width, .{ .fg = .gray });
            line_y += 1;
        }
    }
    if (line_y < inner.y + inner.height and state.log_len > 0) {
        var log_buf: [SCAN_TEXT_MAX + 24]u8 = undefined;
        const log_line = std.fmt.bufPrint(&log_buf, "Detalle: {s}", .{state.log[0..state.log_len]}) catch "";
        buf.setStringTruncated(inner.x, line_y, log_line, inner.width, .{ .fg = .gray });
    }
}

fn copyText(dest: *[SCAN_TEXT_MAX]u8, src: []const u8) usize {
    const n = @min(dest.len, src.len);
    @memcpy(dest[0..n], src[0..n]);
    return n;
}

fn parsePositiveUsize(value: ?std.json.Value) usize {
    const v = value orelse return 0;
    return switch (v) {
        .integer => |i| if (i > 0) @intCast(i) else 0,
        .float => |f| if (f > 0) @intFromFloat(f) else 0,
        else => 0,
    };
}

fn parsePositiveUsizeOr(current: usize, value: ?std.json.Value) usize {
    const parsed = parsePositiveUsize(value);
    return if (parsed > 0) parsed else current;
}

fn parsePositiveU64(value: ?std.json.Value) u64 {
    const v = value orelse return 0;
    return switch (v) {
        .integer => |i| if (i > 0) @intCast(i) else 0,
        .float => |f| if (f > 0) @intFromFloat(f) else 0,
        else => 0,
    };
}

fn parsePositiveU64Or(current: u64, value: ?std.json.Value) u64 {
    const parsed = parsePositiveU64(value);
    return if (parsed > 0) parsed else current;
}

fn parsePercent(value: ?std.json.Value) u8 {
    const p = parsePositiveUsize(value);
    if (p >= 100) return 100;
    return @intCast(p);
}

fn formatBytes(buf: []u8, bytes: u64) []const u8 {
    if (bytes >= 1024 * 1024 * 1024) {
        return std.fmt.bufPrint(buf, "{d:.2} GB", .{@as(f64, @floatFromInt(bytes)) / 1073741824.0}) catch "0 B";
    }
    if (bytes >= 1024 * 1024) {
        return std.fmt.bufPrint(buf, "{d:.2} MB", .{@as(f64, @floatFromInt(bytes)) / 1048576.0}) catch "0 B";
    }
    if (bytes >= 1024) {
        return std.fmt.bufPrint(buf, "{d:.2} KB", .{@as(f64, @floatFromInt(bytes)) / 1024.0}) catch "0 B";
    }
    return std.fmt.bufPrint(buf, "{d} B", .{bytes}) catch "0 B";
}

fn formatElapsed(buf: []u8, elapsed_ms: u64) []const u8 {
    const total_seconds: u64 = elapsed_ms / 1000;
    const minutes: u64 = total_seconds / 60;
    const seconds: u64 = total_seconds % 60;
    return std.fmt.bufPrint(buf, "{d}m {d}s", .{ minutes, seconds }) catch "0m 0s";
}

fn parseYearSummary(allocator: std.mem.Allocator, year_breakdown_val: ?std.json.Value) !?[]u8 {
    const yb = year_breakdown_val orelse return null;
    if (yb != .array) return null;
    if (yb.array.items.len == 0) return null;

    var parts = std.ArrayList([]const u8){};
    defer {
        for (parts.items) |p| allocator.free(p);
        parts.deinit(allocator);
    }

    for (yb.array.items) |entry| {
        if (entry != .object) continue;
        const eobj = entry.object;
        const yv = eobj.get("year") orelse continue;
        const cv = eobj.get("count") orelse continue;

        const year: i64 = switch (yv) {
            .integer => |v| v,
            .float => |v| @intFromFloat(v),
            else => continue,
        };

        const cnt_raw: i64 = switch (cv) {
            .integer => |v| v,
            .float => |v| @intFromFloat(v),
            else => continue,
        };

        const cnt: usize = if (cnt_raw > 0) @intCast(cnt_raw) else 0;

        const part = try std.fmt.allocPrint(allocator, "{d}({d})", .{ year, cnt });
        try parts.append(allocator, part);
    }

    if (parts.items.len == 0) return null;

    var buf = std.ArrayList(u8){};
    defer buf.deinit(allocator);

    try buf.appendSlice(allocator, "Años: ");

    for (parts.items, 0..) |part, i| {
        if (i > 0) try buf.appendSlice(allocator, " · ");
        try buf.appendSlice(allocator, part);
    }

    return try buf.toOwnedSlice(allocator);
}

fn parseUndatedCount(undated_count_val: ?std.json.Value) usize {
    const uv = undated_count_val orelse return 0;
    return switch (uv) {
        .integer => |v| if (v > 0) @intCast(v) else 0,
        .float => |v| if (v > 0) @intFromFloat(v) else 0,
        else => 0,
    };
}

fn askPstPath(allocator: std.mem.Allocator) ![]u8 {
    while (true) {
        const method_items = [_]menu.MenuItem{
            .{
                .label = "Buscar archivo (explorador de terminal)",
                .description = "Navega con flechas, filtra .pst automaticamente",
            },
            .{
                .label = "Escribir la ruta manualmente",
                .description = "Pega o escribe una ruta completa al archivo .pst",
            },
        };
        const idx = menu.select("\xc2\xbfComo indicaras la ruta del PST?", &method_items) orelse {
            return error.UserCancelled;
        };

        const path_opt: ?[]u8 = switch (idx) {
            0 => try file_browser.browseForFile(allocator, null, ".pst"),
            else => try askPstPathManual(allocator),
        };

        const path = path_opt orelse {
            std.debug.print("\x1b[33m[!]\x1b[0m  Seleccion cancelada. Intenta de nuevo.\n", .{});
            continue;
        };

        // Validar que el archivo exista (por si la ruta manual se escribio mal)
        if (std.fs.cwd().access(path, .{})) |_| {
            return path;
        } else |_| {
            std.debug.print("\x1b[31m[X]\x1b[0m  Archivo no encontrado: {s}\n", .{path});
            allocator.free(path);
            continue;
        }
    }
}

fn askPstPathManual(allocator: std.mem.Allocator) !?[]u8 {
    const raw = try input.prompt(allocator, "Ruta del archivo PST (.pst)");
    defer allocator.free(raw);

    const trimmed = std.mem.trim(u8, raw, " \t\"'");
    if (trimmed.len == 0) return null;

    if (!std.ascii.endsWithIgnoreCase(trimmed, ".pst")) {
        std.debug.print("\x1b[33m[!]\x1b[0m  La ruta no termina en .pst. Se intentara igual.\n", .{});
    }

    return try allocator.dupe(u8, trimmed);
}

fn askAction() ?[]const u8 {
    const items = [_]menu.MenuItem{
        .{ .label = "Copy", .description = "Copiar correos al buzon (no modifica el PST)" },
        .{ .label = "Move", .description = "Mover correos (los elimina del PST)" },
    };
    const idx = menu.select("\xc2\xbfQue accion realizar con los correos?", &items) orelse return null;
    return if (idx == 0) "Copy" else "Move";
}

fn askFilterYear(allocator: std.mem.Allocator) !?[]u8 {
    if (!input.confirm("\xc2\xbfFiltrar por un anio especifico?", false)) return null;

    while (true) {
        const raw = try input.prompt(allocator, "Anio (ej. 2026)");
        defer allocator.free(raw);

        const trimmed = std.mem.trim(u8, raw, " \t");
        if (trimmed.len == 0) {
            std.debug.print("\x1b[31m[X]\x1b[0m  Debes ingresar un anio.\n", .{});
            continue;
        }

        const year = std.fmt.parseInt(i32, trimmed, 10) catch {
            std.debug.print("\x1b[31m[X]\x1b[0m  Anio invalido: {s}\n", .{trimmed});
            continue;
        };

        if (year < 1900 or year > 9999) {
            std.debug.print("\x1b[31m[X]\x1b[0m  Anio fuera de rango (1900-9999).\n", .{});
            continue;
        }

        return try allocator.dupe(u8, trimmed);
    }
}

fn askFilterMonths(allocator: std.mem.Allocator) !?[]u8 {
    if (!input.confirm("\xc2\xbfFiltrar tambien por meses especificos?", false)) return null;

    while (true) {
        const raw = try input.prompt(
            allocator,
            "Meses separados por coma (ej. 1,2,3 o enero,febrero,marzo)",
        );
        defer allocator.free(raw);

        const trimmed = std.mem.trim(u8, raw, " \t");
        if (trimmed.len == 0) {
            std.debug.print("\x1b[31m[X]\x1b[0m  Debes ingresar al menos un mes.\n", .{});
            continue;
        }

        return try allocator.dupe(u8, trimmed);
    }
}

fn printSummary(
    session: Session,
    pst_path: []const u8,
    selected_folders: []const []const u8,
    action: []const u8,
    filter_year: ?[]const u8,
    filter_months: ?[]const u8,
    skip_dupes: bool,
) void {
    std.debug.print("\n\x1b[1;36m== Resumen de importacion ==\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mCuenta destino:\x1b[0m    \x1b[1m{s}\x1b[0m\n", .{session.email});
    std.debug.print("  \x1b[90mStoreId:\x1b[0m           {s}\n", .{session.store_id});
    std.debug.print("  \x1b[90mArchivo PST:\x1b[0m       {s}\n", .{pst_path});
    std.debug.print("  \x1b[90mCarpetas PST:\x1b[0m      {d} seleccionada(s)\n", .{selected_folders.len});
    std.debug.print("  \x1b[90mAccion:\x1b[0m            {s}\n", .{action});
    if (filter_year) |fy| {
        std.debug.print("  \x1b[90mFiltro anio:\x1b[0m       {s}\n", .{fy});
    } else {
        std.debug.print("  \x1b[90mFiltro anio:\x1b[0m       (ninguno)\n", .{});
    }
    if (filter_months) |fm| {
        std.debug.print("  \x1b[90mFiltro meses:\x1b[0m      {s}\n", .{fm});
    } else {
        std.debug.print("  \x1b[90mFiltro meses:\x1b[0m      (ninguno)\n", .{});
    }
    std.debug.print("  \x1b[90mSaltar duplicados:\x1b[0m {s}\n", .{if (skip_dupes) "Si" else "No"});
    std.debug.print("\n", .{});
}

fn executePowerShell(
    allocator: std.mem.Allocator,
    session: Session,
    pst_path: []const u8,
    selected_folders: []const []const u8,
    action: []const u8,
    filter_year: ?[]const u8,
    filter_months: ?[]const u8,
    skip_dupes: bool,
) !u8 {
    const script_path = try scripts.getScriptPath(allocator, .import_pst);
    defer allocator.free(script_path);

    const selected_folders_json = try buildJsonArrayString(allocator, selected_folders);
    defer allocator.free(selected_folders_json);

    var argv = std.ArrayList([]const u8){};
    defer argv.deinit(allocator);

    try argv.append(allocator, "powershell");
    try argv.append(allocator, "-NoProfile");
    try argv.append(allocator, "-ExecutionPolicy");
    try argv.append(allocator, "Bypass");
    try argv.append(allocator, "-File");
    try argv.append(allocator, script_path);

    try argv.append(allocator, "-PstPath");
    try argv.append(allocator, pst_path);

    try argv.append(allocator, "-TargetStoreId");
    try argv.append(allocator, session.store_id);

    try argv.append(allocator, "-Action");
    try argv.append(allocator, action);

    try argv.append(allocator, "-IncludeFoldersJson");
    try argv.append(allocator, selected_folders_json);

    if (filter_year) |fy| {
        try argv.append(allocator, "-FilterOnlyYear");
        try argv.append(allocator, fy);
    }
    if (filter_months) |fm| {
        try argv.append(allocator, "-FilterOnlyMonths");
        try argv.append(allocator, fm);
    }
    if (skip_dupes) {
        try argv.append(allocator, "-SkipDuplicates");
    }

    try argv.append(allocator, "-Json");

    var child = std.process.Child.init(argv.items, allocator);
    child.stdin_behavior = .Inherit;
    child.stdout_behavior = .Pipe;
    child.stderr_behavior = .Pipe;

    child.spawn() catch |err| {
        std.debug.print("\x1b[31m[X]\x1b[0m  No se pudo iniciar PowerShell: {s}\n", .{@errorName(err)});
        return 1;
    };

    const header = import_progress.Header{
        .cuenta = session.email,
        .pst_path = pst_path,
        .action = action,
        .filter_year = filter_year,
        .filter_months = filter_months,
    };

    const code = try import_progress.run(allocator, &child, header);

    std.debug.print("\n", .{});
    if (code == 0) {
        std.debug.print("\x1b[1;32m[OK]\x1b[0m Importacion finalizada correctamente.\n", .{});
    } else {
        std.debug.print("\x1b[1;31m[X]\x1b[0m Importacion finalizo con errores (exit code {d}).\n", .{code});
    }
    return code;
}

fn buildJsonArrayString(allocator: std.mem.Allocator, values: []const []const u8) ![]u8 {
    var buf = std.ArrayList(u8){};
    errdefer buf.deinit(allocator);

    try buf.append(allocator, '[');
    for (values, 0..) |value, i| {
        if (i > 0) try buf.appendSlice(allocator, ",");
        try appendJsonString(&buf, allocator, value);
    }
    try buf.append(allocator, ']');

    return try buf.toOwnedSlice(allocator);
}

fn appendJsonString(buf: *std.ArrayList(u8), allocator: std.mem.Allocator, value: []const u8) !void {
    try buf.append(allocator, '"');
    for (value) |c| {
        switch (c) {
            '"' => try buf.appendSlice(allocator, "\\\""),
            '\\' => try buf.appendSlice(allocator, "\\\\"),
            '\n' => try buf.appendSlice(allocator, "\\n"),
            '\r' => try buf.appendSlice(allocator, "\\r"),
            '\t' => try buf.appendSlice(allocator, "\\t"),
            else => try buf.append(allocator, c),
        }
    }
    try buf.append(allocator, '"');
}
