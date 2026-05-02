const std = @import("std");
const menu = @import("menu.zig");
const tui = @import("zigtui");

const STD_INPUT_HANDLE: i32 = -10;
extern "kernel32" fn GetStdHandle(nStdHandle: i32) callconv(.winapi) ?*anyopaque;
extern "kernel32" fn ReadFile(
    hFile: ?*anyopaque,
    lpBuffer: *anyopaque,
    nNumberOfBytesToRead: u32,
    lpNumberOfBytesRead: *u32,
    lpOverlapped: ?*anyopaque,
) callconv(.winapi) i32;

/// Lee una linea del stdin (modo linea, termina en Enter).
/// Devuelve string sin `\r` ni `\n`. Caller debe liberar.
pub fn readLine(allocator: std.mem.Allocator) ![]u8 {
    const handle = GetStdHandle(STD_INPUT_HANDLE) orelse return error.NoStdin;
    var buf = std.ArrayList(u8){};
    defer buf.deinit(allocator);

    while (true) {
        var byte: u8 = 0;
        var read: u32 = 0;
        if (ReadFile(handle, &byte, 1, &read, null) == 0) return error.ReadFailed;
        if (read == 0) break; // EOF
        if (byte == '\r') continue;
        if (byte == '\n') break;
        try buf.append(allocator, byte);
    }

    return try buf.toOwnedSlice(allocator);
}

/// Pregunta al usuario y devuelve la linea tipeada (puede estar vacia).
pub fn prompt(allocator: std.mem.Allocator, question: []const u8) ![]u8 {
    std.debug.print("\x1b[1;36m?\x1b[0m {s} \x1b[90m>\x1b[0m ", .{question});
    return try readLine(allocator);
}

/// Pregunta al usuario con un valor por defecto si deja vacio.
pub fn promptWithDefault(allocator: std.mem.Allocator, question: []const u8, default_value: []const u8) ![]u8 {
    std.debug.print("\x1b[1;36m?\x1b[0m {s} \x1b[90m[{s}]\x1b[0m \x1b[90m>\x1b[0m ", .{ question, default_value });
    const line = try readLine(allocator);
    if (line.len == 0) {
        allocator.free(line);
        return try allocator.dupe(u8, default_value);
    }
    return line;
}

/// Pregunta si/no usando el menu interactivo. true = si.
pub fn confirm(question: []const u8, default_yes: bool) bool {
    if (confirmWithDialog(question, default_yes)) |value| {
        return value;
    }

    const items = [_]menu.MenuItem{
        .{ .label = "Si", .description = null },
        .{ .label = "No", .description = null },
    };
    const idx = menu.select(question, &items) orelse return false;
    return idx == 0;
}

const ConfirmView = struct {
    question: []const u8,
    selected_button: usize,
};

fn confirmWithDialog(question: []const u8, default_yes: bool) ?bool {
    var backend = tui.init(std.heap.page_allocator) catch return null;
    defer backend.deinit();

    var terminal = tui.Terminal.init(std.heap.page_allocator, backend.interface()) catch return null;
    defer terminal.deinit();

    terminal.hideCursor() catch return null;
    defer terminal.showCursor() catch {};

    var view = ConfirmView{
        .question = question,
        .selected_button = if (default_yes) 0 else 1,
    };

    while (true) {
        terminal.draw(&view, renderConfirmView) catch return null;

        const ev = backend.interface().pollEvent(100) catch tui.Event.none;
        if (ev != .key) continue;

        switch (ev.key.code) {
            .left, .back_tab => {
                view.selected_button = if (view.selected_button == 0) 1 else 0;
            },
            .right, .tab => {
                view.selected_button = (view.selected_button + 1) % 2;
            },
            .enter => return view.selected_button == 0,
            .esc => return false,
            .char => |c| {
                if (c == 'y' or c == 'Y' or c == 's' or c == 'S') return true;
                if (c == 'n' or c == 'N') return false;
            },
            else => {},
        }
    }
}

fn renderConfirmView(view: *ConfirmView, buf: *tui.Buffer) anyerror!void {
    const area = buf.getArea();
    clearRect(buf, area);
    if (area.width < 40 or area.height < 12) {
        if (area.width > 0 and area.height > 0) {
            const message = "Ventana muy pequena: amplia la terminal para confirmar.";
            const y = area.y + area.height / 2;
            buf.setStringTruncated(area.x, y, message, area.width, .{ .fg = .yellow, .modifier = .{ .bold = true } });
        }
        return;
    }

    const popup_w_pct: u8 = if (area.width >= 120) 84 else 92;
    const popup_h_pct: u8 = if (area.height >= 32) 46 else 62;
    const popup_area = tui.centeredRectPct(area, popup_w_pct, popup_h_pct);
    const popup = tui.widgets.Popup{
        .title = " Confirmacion de filtro ",
        .border_style = .{ .fg = .cyan },
        .backdrop_style = .{ .fg = .dark_gray },
        .show_backdrop = true,
    };
    popup.render(popup_area, buf);

    const inner = tui.widgets.Popup.innerArea(popup_area);
    if (inner.height < 8) return;

    buf.setStringTruncated(inner.x, inner.y, view.question, inner.width, .{ .fg = .white, .modifier = .{ .bold = true } });
    buf.setStringTruncated(inner.x, inner.y + 2, "Usa Izq/Der o Tab para elegir opcion", inner.width, .{ .fg = .gray });
    buf.setStringTruncated(inner.x, inner.y + 3, "Enter confirma | ESC cancela", inner.width, .{ .fg = .gray });

    const dialog = tui.widgets.Dialog{
        .title = " Elige una opcion ",
        .message = "Selecciona Si o No",
        .buttons = &.{ "Si", "No" },
        .selected_button = view.selected_button,
        .border_style = .{ .fg = .cyan },
        .message_style = .{ .fg = .white },
        .button_style = .{ .fg = .white },
        .selected_button_style = .{ .fg = .black, .bg = .cyan, .modifier = .{ .bold = true } },
    };
    const dialog_area = tui.widgets.Dialog.dialogArea(
        .{ .x = inner.x, .y = inner.y + 4, .width = inner.width, .height = inner.height - 4 },
        @min(inner.width, 62),
        9,
    );
    dialog.render(dialog_area, buf);
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
