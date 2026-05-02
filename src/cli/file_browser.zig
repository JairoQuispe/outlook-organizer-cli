const std = @import("std");
const tui = @import("zigtui");

extern "kernel32" fn GetLogicalDrives() callconv(.winapi) u32;

const Entry = struct {
    name: []u8,
    is_dir: bool,
    is_parent: bool, // ".."
    is_drive: bool, // "C:\"
    size: u64,
};

const VIEWPORT_MAX: usize = 64;
const NAME_CELL_MAX: usize = 512;
const LINE_CELL_MAX: usize = 640;

fn sanitizeUtf8ForDisplay(input: []const u8, out: []u8) []const u8 {
    if (out.len == 0) return "";

    var in_i: usize = 0;
    var out_i: usize = 0;

    while (in_i < input.len and out_i < out.len) {
        const first = input[in_i];
        const seq_len = std.unicode.utf8ByteSequenceLength(first) catch {
            out[out_i] = '?';
            out_i += 1;
            in_i += 1;
            continue;
        };

        if (seq_len == 1) {
            out[out_i] = first;
            out_i += 1;
            in_i += 1;
            continue;
        }

        if (in_i + seq_len > input.len or out_i + seq_len > out.len) {
            break;
        }

        _ = std.unicode.utf8Decode(input[in_i .. in_i + seq_len]) catch {
            out[out_i] = '?';
            out_i += 1;
            in_i += 1;
            continue;
        };

        @memcpy(out[out_i .. out_i + seq_len], input[in_i .. in_i + seq_len]);
        out_i += seq_len;
        in_i += seq_len;
    }

    return out[0..out_i];
}

/// Navegador de archivos interactivo. Devuelve la ruta absoluta del archivo
/// seleccionado, o null si el usuario cancela con ESC.
/// `extension_filter` opcional (ej. ".pst") filtra archivos. Los directorios
/// siempre se muestran.
pub fn browseForFile(
    allocator: std.mem.Allocator,
    start_dir: ?[]const u8,
    extension_filter: ?[]const u8,
) !?[]u8 {
    // Directorio inicial
    var current_dir: []u8 = undefined;
    if (start_dir) |sd| {
        current_dir = try allocator.dupe(u8, sd);
    } else {
        var buf: [std.fs.max_path_bytes]u8 = undefined;
        const cwd = try std.process.getCwd(&buf);
        current_dir = try allocator.dupe(u8, cwd);
    }
    defer allocator.free(current_dir);

    var backend = try tui.init(allocator);
    defer backend.deinit();

    var terminal = try tui.Terminal.init(allocator, backend.interface());
    defer terminal.deinit();

    try terminal.hideCursor();
    defer terminal.showCursor() catch {};

    var cursor: usize = 0;
    var scroll_top: usize = 0;
    var in_drives_view: bool = !std.unicode.utf8ValidateSlice(current_dir);

    while (true) {
        var entries = std.ArrayList(Entry){};
        defer {
            for (entries.items) |e| allocator.free(e.name);
            entries.deinit(allocator);
        }

        if (in_drives_view) {
            try loadDrives(allocator, &entries);
        } else {
            loadDirectory(allocator, current_dir, extension_filter, &entries) catch |err| {
                const msg = @errorName(err);
                const error_state = RenderState{
                    .current_dir = current_dir,
                    .in_drives_view = in_drives_view,
                    .entries = entries.items,
                    .cursor = cursor,
                    .scroll_top = scroll_top,
                    .extension_filter = extension_filter,
                    .error_message = msg,
                };
                try terminal.draw(
                    error_state,
                    render,
                );

                while (true) {
                    const ev = backend.interface().pollEvent(100) catch tui.Event.none;
                    if (ev != .key) continue;
                    if (ev.key.code == .esc) return null;

                    in_drives_view = true;
                    cursor = 0;
                    scroll_top = 0;
                    break;
                }
                continue;
            };
        }

        if (cursor >= entries.items.len and entries.items.len > 0) {
            cursor = entries.items.len - 1;
        }
        // ajustar scroll
        if (cursor < scroll_top) scroll_top = cursor;
        if (cursor >= scroll_top + VIEWPORT_MAX) scroll_top = cursor - VIEWPORT_MAX + 1;

        const render_state = RenderState{
            .current_dir = current_dir,
            .in_drives_view = in_drives_view,
            .entries = entries.items,
            .cursor = cursor,
            .scroll_top = scroll_top,
            .extension_filter = extension_filter,
            .error_message = null,
        };
        try terminal.draw(render_state, render);

        const ev = backend.interface().pollEvent(100) catch tui.Event.none;
        if (ev != .key) continue;

        switch (ev.key.code) {
            .up => {
                if (entries.items.len == 0) continue;
                cursor = if (cursor == 0) entries.items.len - 1 else cursor - 1;
            },
            .down => {
                if (entries.items.len == 0) continue;
                cursor = (cursor + 1) % entries.items.len;
            },
            .left => {
                if (!in_drives_view) {
                    const res = try goUp(allocator, current_dir);
                    if (res) |new_path| {
                        allocator.free(current_dir);
                        current_dir = new_path;
                        cursor = 0;
                        scroll_top = 0;
                    } else {
                        in_drives_view = true;
                        cursor = 0;
                        scroll_top = 0;
                    }
                }
            },
            .enter => {
                if (entries.items.len == 0) continue;
                const selected = entries.items[cursor];

                if (selected.is_drive) {
                    allocator.free(current_dir);
                    current_dir = try allocator.dupe(u8, selected.name);
                    in_drives_view = false;
                    cursor = 0;
                    scroll_top = 0;
                    continue;
                }

                if (selected.is_parent) {
                    const res = try goUp(allocator, current_dir);
                    if (res) |new_path| {
                        allocator.free(current_dir);
                        current_dir = new_path;
                    } else {
                        in_drives_view = true;
                    }
                    cursor = 0;
                    scroll_top = 0;
                    continue;
                }

                if (selected.is_dir) {
                    const new_path = try std.fs.path.join(allocator, &.{ current_dir, selected.name });
                    allocator.free(current_dir);
                    current_dir = new_path;
                    cursor = 0;
                    scroll_top = 0;
                    continue;
                }

                // Archivo: devolver path absoluto
                const result = try std.fs.path.join(allocator, &.{ current_dir, selected.name });
                return result;
            },
            .esc => return null,
            else => {},
        }
    }
}

fn goUp(allocator: std.mem.Allocator, current_dir: []const u8) !?[]u8 {
    const parent = std.fs.path.dirname(current_dir) orelse return null;
    // Si ya estamos en "C:\" dirname devuelve "C:" -> tratarlo como drive root
    if (std.mem.eql(u8, parent, current_dir)) return null;
    if (parent.len == 2 and parent[1] == ':') {
        // "C:" -> "C:\"
        var buf = try allocator.alloc(u8, 3);
        buf[0] = parent[0];
        buf[1] = ':';
        buf[2] = '\\';
        return buf;
    }
    return try allocator.dupe(u8, parent);
}

fn loadDrives(allocator: std.mem.Allocator, entries: *std.ArrayList(Entry)) !void {
    const mask = GetLogicalDrives();
    var i: u5 = 0;
    while (i < 26) : (i += 1) {
        const bit = @as(u32, 1) << i;
        if ((mask & bit) != 0) {
            const letter: u8 = 'A' + @as(u8, i);
            var name = try allocator.alloc(u8, 3);
            name[0] = letter;
            name[1] = ':';
            name[2] = '\\';
            try entries.append(allocator, .{
                .name = name,
                .is_dir = true,
                .is_parent = false,
                .is_drive = true,
                .size = 0,
            });
        }
    }
}

fn loadDirectory(
    allocator: std.mem.Allocator,
    path: []const u8,
    extension_filter: ?[]const u8,
    entries: *std.ArrayList(Entry),
) !void {
    if (!std.unicode.utf8ValidateSlice(path)) return error.InvalidUtf8Path;

    var dir = try std.fs.openDirAbsolute(path, .{ .iterate = true });
    defer dir.close();

    // Agregar ".." siempre
    const parent_name = try allocator.dupe(u8, "..");
    try entries.append(allocator, .{
        .name = parent_name,
        .is_dir = true,
        .is_parent = true,
        .is_drive = false,
        .size = 0,
    });

    var it = dir.iterate();
    while (try it.next()) |entry| {
        if (!std.unicode.utf8ValidateSlice(entry.name)) continue;

        const is_dir = entry.kind == .directory;

        if (!is_dir) {
            if (extension_filter) |ext| {
                if (!std.ascii.endsWithIgnoreCase(entry.name, ext)) continue;
            }
        }

        // omitir archivos/carpetas ocultos del sistema (empiezan con '.')
        if (entry.name.len > 0 and entry.name[0] == '.' and !is_dir) continue;

        var size: u64 = 0;
        if (!is_dir) {
            const stat_result = dir.statFile(entry.name) catch null;
            if (stat_result) |st| size = st.size;
        }

        const name_copy = try allocator.dupe(u8, entry.name);
        try entries.append(allocator, .{
            .name = name_copy,
            .is_dir = is_dir,
            .is_parent = false,
            .is_drive = false,
            .size = size,
        });
    }

    // Ordenar: carpetas primero (excepto ..), luego archivos; alfabetico
    std.mem.sort(Entry, entries.items[1..], {}, entryLessThan);
}

fn entryLessThan(_: void, a: Entry, b: Entry) bool {
    if (a.is_dir and !b.is_dir) return true;
    if (!a.is_dir and b.is_dir) return false;
    return std.ascii.lessThanIgnoreCase(a.name, b.name);
}

const RenderState = struct {
    current_dir: []const u8,
    in_drives_view: bool,
    entries: []const Entry,
    cursor: usize,
    scroll_top: usize,
    extension_filter: ?[]const u8,
    error_message: ?[]const u8,
};

fn render(state: RenderState, buf: *tui.Buffer) anyerror!void {
    const area = buf.getArea();
    clearRect(buf, area);
    if (area.width < 52 or area.height < 12) {
        if (area.width > 0 and area.height > 0) {
            const message = "Ventana muy pequena: amplia la terminal para usar el explorador.";
            const y = area.y + area.height / 2;
            buf.setStringTruncated(area.x, y, message, area.width, .{ .fg = .yellow, .modifier = .{ .bold = true } });
        }
        return;
    }

    const root = tui.Block{
        .title = " Explorador de archivos ",
        .borders = tui.Borders.ALL,
        .border_style = .{ .fg = .cyan },
        .title_style = .{ .modifier = .{ .bold = true } },
        .border_symbols = tui.BorderSymbols.rounded(),
    };
    root.render(area, buf);
    const inner = root.inner(area);
    if (inner.height < 7 or inner.width < 20) return;

    buf.setStringTruncated(
        inner.x,
        inner.y,
        "Arriba/Abajo: navegar | Enter: abrir/seleccionar | Izquierda: subir | ESC: cancelar",
        inner.width,
        .{ .fg = .gray },
    );

    var location_buf: [std.fs.max_path_bytes + 32]u8 = undefined;
    var dir_display_buf: [std.fs.max_path_bytes]u8 = undefined;
    const dir_display = sanitizeUtf8ForDisplay(state.current_dir, dir_display_buf[0..]);
    const location = if (state.in_drives_view)
        "Ubicacion: [Unidades del sistema]"
    else
        std.fmt.bufPrint(&location_buf, "Ubicacion: {s}", .{dir_display}) catch "Ubicacion";
    buf.setStringTruncated(inner.x, inner.y + 1, location, inner.width, .{ .fg = .white });

    if (state.extension_filter) |ext| {
        var filter_buf: [64]u8 = undefined;
        const filter = std.fmt.bufPrint(&filter_buf, "Filtro activo: {s}", .{ext}) catch "";
        buf.setStringTruncated(inner.x, inner.y + 2, filter, inner.width, .{ .fg = .gray });
    }

    const table_y = inner.y + 3;
    const footer_h: u16 = 2;
    if (inner.y + inner.height <= table_y + footer_h) return;
    const table_h = inner.y + inner.height - table_y - footer_h;
    if (table_h <= 1) return;

    const max_rows: usize = @max(1, @min(@as(usize, table_h - 1), VIEWPORT_MAX));
    var render_scroll_top = state.scroll_top;
    if (state.entries.len > 0) {
        if (state.cursor < render_scroll_top) {
            render_scroll_top = state.cursor;
        }
        if (state.cursor >= render_scroll_top + max_rows) {
            render_scroll_top = state.cursor - max_rows + 1;
        }
    }

    buf.setStringTruncated(inner.x, table_y, "Tipo   Nombre                              Tamano", inner.width, .{ .fg = .light_cyan, .modifier = .{ .bold = true } });

    var name_cells: [VIEWPORT_MAX][NAME_CELL_MAX]u8 = undefined;
    var size_cells: [VIEWPORT_MAX][32]u8 = undefined;
    var line_cells: [VIEWPORT_MAX][LINE_CELL_MAX]u8 = undefined;
    var roots: [VIEWPORT_MAX]tui.widgets.TreeNode = undefined;
    var row_count: usize = 0;

    if (state.entries.len > 0) {
        const end = @min(render_scroll_top + max_rows, state.entries.len);
        var i: usize = render_scroll_top;
        while (i < end and row_count < max_rows) : (i += 1) {
            const e = state.entries[i];
            const tipo: []const u8 = if (e.is_drive)
                "UND"
            else if (e.is_parent)
                ".."
            else if (e.is_dir)
                "DIR"
            else
                "FILE";

            const size_txt: []const u8 = if (e.is_dir) "-" else formatSize(e.size, size_cells[row_count][0..]);
            const display_name = sanitizeUtf8ForDisplay(e.name, name_cells[row_count][0..]);
            const line = std.fmt.bufPrint(
                line_cells[row_count][0..],
                "{s:<5} {s}  {s}",
                .{ tipo, display_name, size_txt },
            ) catch display_name;

            roots[row_count] = .{ .label = line };
            row_count += 1;
        }
    }

    const selected_in_view: ?usize = if (state.entries.len == 0 or state.cursor < render_scroll_top or state.cursor >= render_scroll_top + row_count)
        null
    else
        state.cursor - render_scroll_top;

    if (row_count > 0) {
        const selected_idx = selected_in_view orelse 0;
        const tree = tui.widgets.Tree{
            .roots = roots[0..row_count],
            .selected = selected_idx,
            .highlight_style = .{ .fg = .black, .bg = .cyan, .modifier = .{ .bold = true } },
            .indent = 0,
            .expanded_symbol = "",
            .collapsed_symbol = "",
            .leaf_symbol = "",
        };
        tree.render(.{ .x = inner.x, .y = table_y + 1, .width = inner.width, .height = table_h - 1 }, buf);
    }

    const footer_y = inner.y + inner.height - 2;
    if (state.error_message) |err| {
        var err_buf: [128]u8 = undefined;
        const err_line = std.fmt.bufPrint(&err_buf, "Error leyendo directorio: {s}. ESC cancela; otra tecla va a unidades.", .{err}) catch "Error";
        buf.setStringTruncated(inner.x, footer_y, err_line, inner.width, .{ .fg = .light_red, .modifier = .{ .bold = true } });
    } else if (state.entries.len == 0) {
        buf.setStringTruncated(inner.x, footer_y, "Directorio vacio o sin archivos que coincidan con el filtro.", inner.width, .{ .fg = .gray });
    }

    var pos_buf: [64]u8 = undefined;
    const pos = if (state.entries.len == 0)
        "0/0"
    else
        std.fmt.bufPrint(&pos_buf, "{d}/{d}", .{ state.cursor + 1, state.entries.len }) catch "";
    buf.setStringTruncated(inner.x, inner.y + inner.height - 1, pos, inner.width, .{ .fg = .yellow });
}

fn formatSize(bytes: u64, buf: []u8) []const u8 {
    const GB: u64 = 1024 * 1024 * 1024;
    const MB: u64 = 1024 * 1024;
    const KB: u64 = 1024;
    if (bytes >= GB) {
        return std.fmt.bufPrint(buf, "{d:.2} GB", .{@as(f64, @floatFromInt(bytes)) / @as(f64, GB)}) catch "";
    } else if (bytes >= MB) {
        return std.fmt.bufPrint(buf, "{d:.2} MB", .{@as(f64, @floatFromInt(bytes)) / @as(f64, MB)}) catch "";
    } else if (bytes >= KB) {
        return std.fmt.bufPrint(buf, "{d:.2} KB", .{@as(f64, @floatFromInt(bytes)) / @as(f64, KB)}) catch "";
    } else {
        return std.fmt.bufPrint(buf, "{d} B", .{bytes}) catch "";
    }
}

fn clearRect(buf: *tui.Buffer, rect: tui.Rect) void {
    if (rect.width == 0 or rect.height == 0) return;

    var row: u16 = 0;
    while (row < rect.height) : (row += 1) {
        clearLine(buf, rect.x, rect.y + row, rect.width);
    }
}

fn clearLine(buf: *tui.Buffer, x: u16, y: u16, width: u16) void {
    if (width == 0) return;

    const spaces = "                                                                ";
    var written: u16 = 0;
    while (written < width) {
        const remaining: u16 = width - written;
        const chunk_len: usize = @min(@as(usize, remaining), spaces.len);
        buf.setString(x + written, y, spaces[0..chunk_len], .{});
        written += @as(u16, @intCast(chunk_len));
    }
}
