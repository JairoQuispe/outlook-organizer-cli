const std = @import("std");
const tui = @import("zigtui");

const MAX_LOG_LINES: usize = 40;
const LINE_BUFFER: usize = 256;

const LogLevel = enum(u8) { info = 0, warn = 1, err = 2 };

const LogEntry = struct {
    level: LogLevel = .info,
    text_len: usize = 0,
    text: [LINE_BUFFER]u8 = undefined,
};

pub const Header = struct {
    cuenta: []const u8,
    pst_path: []const u8,
    action: []const u8,
    filter_year: ?[]const u8,
    filter_months: ?[]const u8,
};

const State = struct {
    mutex: std.Thread.Mutex = .{},

    header: Header,

    activity_buf: [LINE_BUFFER]u8 = undefined,
    activity_len: usize = 0,
    status_buf: [LINE_BUFFER]u8 = undefined,
    status_len: usize = 0,
    percent: u8 = 0,

    copied: u64 = 0,
    moved: u64 = 0,
    skipped: u64 = 0,
    failed: u64 = 0,

    throttle_rate: f64 = 0,
    throttle_waited_ms: u64 = 0,
    throttle_errors: u32 = 0,

    logs: [MAX_LOG_LINES]LogEntry = [_]LogEntry{.{}} ** MAX_LOG_LINES,
    log_head: usize = 0,
    log_count: usize = 0,

    stdout_eof: bool = false,
    stderr_eof: bool = false,
    cancelled: bool = false,
    finished: bool = false,
    exit_code: u8 = 0,
    final_message_buf: [LINE_BUFFER]u8 = undefined,
    final_message_len: usize = 0,
    started_ms: i64 = 0,
    elapsed_ms: u64 = 0,

    fn setActivity(self: *State, text: []const u8) void {
        const n = @min(text.len, self.activity_buf.len);
        @memcpy(self.activity_buf[0..n], text[0..n]);
        self.activity_len = n;
    }
    fn setStatus(self: *State, text: []const u8) void {
        const n = @min(text.len, self.status_buf.len);
        @memcpy(self.status_buf[0..n], text[0..n]);
        self.status_len = n;
    }
    fn setFinalMessage(self: *State, text: []const u8) void {
        const n = @min(text.len, self.final_message_buf.len);
        @memcpy(self.final_message_buf[0..n], text[0..n]);
        self.final_message_len = n;
    }
    fn pushLog(self: *State, level: LogLevel, text: []const u8) void {
        const idx = self.log_head;
        self.logs[idx].level = level;
        const n = @min(text.len, LINE_BUFFER);
        @memcpy(self.logs[idx].text[0..n], text[0..n]);
        self.logs[idx].text_len = n;
        self.log_head = (self.log_head + 1) % MAX_LOG_LINES;
        if (self.log_count < MAX_LOG_LINES) self.log_count += 1;
    }
};

pub fn run(allocator: std.mem.Allocator, child: *std.process.Child, header: Header) !u8 {
    var state = State{
        .header = header,
        .started_ms = std.time.milliTimestamp(),
    };
    state.setActivity("Iniciando...");

    var backend = try tui.init(allocator);
    defer backend.deinit();

    var terminal = try tui.Terminal.init(allocator, backend.interface());
    defer terminal.deinit();

    try terminal.hideCursor();
    defer terminal.showCursor() catch {};

    const stdout_file = child.stdout orelse return error.NoStdoutPipe;
    const stderr_file = child.stderr orelse return error.NoStderrPipe;

    const stdout_thread = try std.Thread.spawn(.{}, readerLoop, .{ stdout_file, &state, allocator, true });
    const stderr_thread = try std.Thread.spawn(.{}, readerLoop, .{ stderr_file, &state, allocator, false });

    var waited = false;

    while (true) {
        const ev = backend.interface().pollEvent(50) catch tui.Event.none;
        switch (ev) {
            .key => |k| {
                switch (k.code) {
                    .esc => {
                        state.mutex.lock();
                        const already = state.cancelled;
                        state.cancelled = true;
                        state.mutex.unlock();
                        if (!already) {
                            _ = child.kill() catch {};
                        }
                    },
                    .enter => {
                        state.mutex.lock();
                        const finished = state.finished;
                        state.mutex.unlock();
                        if (finished) break;
                    },
                    .char => |c| {
                        if (c == 'q' or c == 'Q') {
                            state.mutex.lock();
                            const finished = state.finished;
                            state.mutex.unlock();
                            if (finished) break;
                        }
                    },
                    else => {},
                }
            },
            else => {},
        }

        state.mutex.lock();
        const both_eof = state.stdout_eof and state.stderr_eof;
        state.mutex.unlock();

        if (both_eof and !waited) {
            waited = true;
            const term = child.wait() catch std.process.Child.Term{ .Unknown = 1 };
            const code: u8 = switch (term) {
                .Exited => |c| @truncate(c),
                else => 1,
            };
            state.mutex.lock();
            state.finished = true;
            state.exit_code = code;
            state.elapsed_ms = @intCast(std.time.milliTimestamp() - state.started_ms);
            if (state.cancelled) {
                state.setFinalMessage("Cancelado por el usuario. Enter para volver.");
            } else if (code == 0) {
                state.setFinalMessage("Importacion completada. Enter para volver.");
            } else {
                state.setFinalMessage("Importacion con errores. Enter para volver.");
            }
            state.mutex.unlock();
        }

        state.mutex.lock();
        if (!state.finished) {
            state.elapsed_ms = @intCast(std.time.milliTimestamp() - state.started_ms);
        }
        state.mutex.unlock();

        try terminal.draw(&state, renderState);
    }

    stdout_thread.join();
    stderr_thread.join();

    return state.exit_code;
}

// ── Reader thread ─────────────────────────────────────────────────

fn readerLoop(file: std.fs.File, state: *State, allocator: std.mem.Allocator, is_stdout: bool) void {
    var line_buf = std.ArrayList(u8){};
    defer line_buf.deinit(allocator);

    var read_buf: [4096]u8 = undefined;

    while (true) {
        const n = file.read(&read_buf) catch break;
        if (n == 0) break;

        for (read_buf[0..n]) |byte| {
            if (byte == '\n') {
                processLine(line_buf.items, state, allocator, is_stdout);
                line_buf.clearRetainingCapacity();
            } else if (byte != '\r') {
                line_buf.append(allocator, byte) catch {
                    line_buf.clearRetainingCapacity();
                };
            }
        }
    }

    if (line_buf.items.len > 0) {
        processLine(line_buf.items, state, allocator, is_stdout);
    }

    state.mutex.lock();
    if (is_stdout) {
        state.stdout_eof = true;
    } else {
        state.stderr_eof = true;
    }
    state.mutex.unlock();
}

fn processLine(line: []const u8, state: *State, allocator: std.mem.Allocator, is_stdout: bool) void {
    const trimmed = std.mem.trim(u8, line, " \t\r");
    if (trimmed.len == 0) return;

    if (is_stdout and trimmed[0] == '{') {
        if (parseAndApply(trimmed, state, allocator)) {
            return;
        }
    }

    state.mutex.lock();
    defer state.mutex.unlock();
    state.pushLog(if (is_stdout) .info else .warn, trimmed);
}

fn parseAndApply(line: []const u8, state: *State, allocator: std.mem.Allocator) bool {
    var parsed = std.json.parseFromSlice(std.json.Value, allocator, line, .{}) catch return false;
    defer parsed.deinit();

    const root = switch (parsed.value) {
        .object => |o| o,
        else => return false,
    };

    const type_val = root.get("type") orelse return false;
    const type_str = switch (type_val) {
        .string => |s| s,
        else => return false,
    };

    state.mutex.lock();
    defer state.mutex.unlock();

    if (std.mem.eql(u8, type_str, "log")) {
        const level_str = strField(root, "level") orelse "info";
        const message = strField(root, "message") orelse "";
        const lvl: LogLevel = if (std.mem.eql(u8, level_str, "warn"))
            .warn
        else if (std.mem.eql(u8, level_str, "error"))
            .err
        else
            .info;
        state.pushLog(lvl, message);
    } else if (std.mem.eql(u8, type_str, "progress")) {
        if (strField(root, "activity")) |a| state.setActivity(a);
        if (strField(root, "status")) |s| state.setStatus(s);
        if (intField(root, "percent")) |p| state.percent = @intCast(@max(0, @min(100, p)));
    } else if (std.mem.eql(u8, type_str, "throttleStats")) {
        if (floatField(root, "effectiveRate")) |r| state.throttle_rate = r;
        if (intField(root, "totalWaitedMs")) |w| state.throttle_waited_ms = @intCast(@max(0, w));
        if (intField(root, "throttleErrors")) |e| state.throttle_errors = @intCast(@max(0, e));
    } else if (std.mem.eql(u8, type_str, "restoreResult")) {
        if (intField(root, "copied")) |v| state.copied = @intCast(@max(0, v));
        if (intField(root, "moved")) |v| state.moved = @intCast(@max(0, v));
        if (intField(root, "skipped")) |v| state.skipped = @intCast(@max(0, v));
        if (intField(root, "failed")) |v| state.failed = @intCast(@max(0, v));
        state.percent = 100;
        state.pushLog(.info, "Resultado final recibido.");
    } else if (std.mem.eql(u8, type_str, "error")) {
        const message = strField(root, "message") orelse "(sin mensaje)";
        state.pushLog(.err, message);
    }

    return true;
}

fn strField(obj: std.json.ObjectMap, key: []const u8) ?[]const u8 {
    const v = obj.get(key) orelse return null;
    return switch (v) {
        .string => |s| s,
        else => null,
    };
}

fn intField(obj: std.json.ObjectMap, key: []const u8) ?i64 {
    const v = obj.get(key) orelse return null;
    return switch (v) {
        .integer => |i| i,
        .float => |f| @intFromFloat(f),
        else => null,
    };
}

fn floatField(obj: std.json.ObjectMap, key: []const u8) ?f64 {
    const v = obj.get(key) orelse return null;
    return switch (v) {
        .float => |f| f,
        .integer => |i| @floatFromInt(i),
        else => null,
    };
}

// ── Render ────────────────────────────────────────────────────────

fn renderState(state: *State, buf: *tui.Buffer) anyerror!void {
    state.mutex.lock();
    defer state.mutex.unlock();

    const area = buf.getArea();
    if (area.width < 30 or area.height < 10) return;

    const root_block = tui.Block{
        .title = " Importar correos desde PST ",
        .borders = tui.Borders.ALL,
        .border_style = .{ .fg = .cyan },
        .title_style = .{ .modifier = .{ .bold = true } },
        .border_symbols = tui.BorderSymbols.rounded(),
    };
    root_block.render(area, buf);
    const inner = root_block.inner(area);
    if (inner.width < 10 or inner.height < 8) return;

    // Layout: header(4) | progress(3) | stats(2) | throttle(2) | log(rest) | footer(1)
    var y = inner.y;
    const x = inner.x;
    const w = inner.width;

    // ── Header ────────────────────────────────────
    const dim: tui.Style = .{ .fg = .gray };
    const accent: tui.Style = .{ .fg = .light_cyan, .modifier = .{ .bold = true } };

    drawLabeled(buf, x, y, w, "Cuenta:    ", state.header.cuenta, dim, accent);
    y += 1;
    drawLabeled(buf, x, y, w, "PST:       ", state.header.pst_path, dim, .{});
    y += 1;
    drawLabeled(buf, x, y, w, "Accion:    ", state.header.action, dim, .{});
    y += 1;

    var filter_buf: [256]u8 = undefined;
    const filter_str = formatFilters(&filter_buf, state.header.filter_year, state.header.filter_months);
    drawLabeled(buf, x, y, w, "Filtros:   ", filter_str, dim, .{});
    y += 2;

    if (y >= inner.y + inner.height) return;

    // ── Progress ──────────────────────────────────
    var prog_label_buf: [LINE_BUFFER]u8 = undefined;
    const activity = state.activity_buf[0..state.activity_len];
    const status = state.status_buf[0..state.status_len];
    const prog_label = std.fmt.bufPrint(&prog_label_buf, "{s} - {s}", .{ activity, status }) catch activity;

    buf.setStringTruncated(x, y, prog_label, w, accent);
    y += 1;

    const gauge_rect = tui.Rect{ .x = x, .y = y, .width = w, .height = 1 };
    const gauge = tui.Gauge{
        .ratio = @as(f64, @floatFromInt(state.percent)) / 100.0,
        .style = .{ .fg = .dark_gray },
        .gauge_style = .{ .fg = .green, .bg = .green },
        .use_unicode = true,
    };
    gauge.render(gauge_rect, buf);
    var pct_buf: [16]u8 = undefined;
    const pct_str = std.fmt.bufPrint(&pct_buf, " {d}% ", .{state.percent}) catch " ";
    if (w >= pct_str.len) {
        const pct_x: u16 = @intCast(x + (w -| @as(u16, @intCast(pct_str.len))) / 2);
        buf.setString(pct_x, y, pct_str, .{ .fg = .black, .bg = .green, .modifier = .{ .bold = true } });
    }
    y += 2;

    if (y >= inner.y + inner.height) return;

    // ── Stats ─────────────────────────────────────
    var stats_buf: [256]u8 = undefined;
    const stats_str = std.fmt.bufPrint(
        &stats_buf,
        "Copied: {d}  Moved: {d}  Skipped: {d}  Failed: {d}",
        .{ state.copied, state.moved, state.skipped, state.failed },
    ) catch "";
    buf.setStringTruncated(x, y, stats_str, w, .{ .fg = .light_white });
    y += 1;

    var throttle_buf: [256]u8 = undefined;
    const throttle_str = std.fmt.bufPrint(
        &throttle_buf,
        "Throttle: {d:.1} items/min  esperado: {d:.1}s  errores: {d}  elapsed: {d:.1}s",
        .{
            state.throttle_rate,
            @as(f64, @floatFromInt(state.throttle_waited_ms)) / 1000.0,
            state.throttle_errors,
            @as(f64, @floatFromInt(state.elapsed_ms)) / 1000.0,
        },
    ) catch "";
    buf.setStringTruncated(x, y, throttle_str, w, dim);
    y += 2;

    if (y >= inner.y + inner.height) return;

    // ── Log ───────────────────────────────────────
    const footer_h: u16 = 2;
    const log_top_h: u16 = 1;
    if (inner.y + inner.height < y + footer_h + log_top_h) return;

    const log_height: u16 = inner.y + inner.height - y - footer_h;
    if (log_height >= 2) {
        buf.setString(x, y, "Log:", .{ .fg = .yellow, .modifier = .{ .bold = true } });
        y += 1;

        const visible = @min(@as(usize, log_height - 1), state.log_count);
        // Empezar por las mas viejas visibles
        const start_offset: usize = state.log_count - visible;
        var i: usize = 0;
        while (i < visible) : (i += 1) {
            const logical = start_offset + i;
            const idx = (state.log_head + MAX_LOG_LINES - state.log_count + logical) % MAX_LOG_LINES;
            const entry = state.logs[idx];
            const text = entry.text[0..entry.text_len];

            const style: tui.Style = switch (entry.level) {
                .info => .{ .fg = .gray },
                .warn => .{ .fg = .yellow },
                .err => .{ .fg = .light_red, .modifier = .{ .bold = true } },
            };
            const prefix: []const u8 = switch (entry.level) {
                .info => " . ",
                .warn => " ! ",
                .err => " X ",
            };
            buf.setString(x, y, prefix, style);
            buf.setStringTruncated(x + 3, y, text, w -| 3, style);
            y += 1;
        }
        // Si quedan filas vacias del log, no se rellenan (ya estan en blanco)
        y = inner.y + inner.height - footer_h;
    }

    // ── Footer ────────────────────────────────────
    const footer_y = inner.y + inner.height - 1;
    if (state.finished) {
        const final = state.final_message_buf[0..state.final_message_len];
        const final_style: tui.Style = if (state.cancelled or state.exit_code != 0)
            .{ .fg = .light_yellow, .modifier = .{ .bold = true } }
        else
            .{ .fg = .light_green, .modifier = .{ .bold = true } };
        buf.setStringTruncated(x, footer_y, final, w, final_style);
    } else {
        buf.setStringTruncated(x, footer_y, "ESC: cancelar | Q/Enter al finalizar para volver", w, dim);
    }
}

fn drawLabeled(
    buf: *tui.Buffer,
    x: u16,
    y: u16,
    w: u16,
    label: []const u8,
    value: []const u8,
    label_style: tui.Style,
    value_style: tui.Style,
) void {
    buf.setString(x, y, label, label_style);
    const lx: u16 = @intCast(x + label.len);
    if (lx < x + w) {
        buf.setStringTruncated(lx, y, value, w -| @as(u16, @intCast(label.len)), value_style);
    }
}

fn formatFilters(buf: []u8, year: ?[]const u8, months: ?[]const u8) []const u8 {
    if (year == null and months == null) return "(sin filtros)";
    if (year != null and months != null) {
        return std.fmt.bufPrint(buf, "anio={s}  meses={s}", .{ year.?, months.? }) catch "(filtros)";
    }
    if (year) |y| {
        return std.fmt.bufPrint(buf, "anio={s}", .{y}) catch "(filtro anio)";
    }
    if (months) |m| {
        return std.fmt.bufPrint(buf, "meses={s}", .{m}) catch "(filtro meses)";
    }
    return "(sin filtros)";
}
