const std = @import("std");
const keyinput = @import("../windows/keyinput.zig");

pub const MenuItem = struct {
    label: []const u8,
    description: ?[]const u8 = null,
};

/// Muestra un menu interactivo y devuelve el indice elegido o null si se cancela con ESC.
/// Soporta flechas Arriba/Abajo, numeros 1..N y Enter.
pub fn select(title: []const u8, items: []const MenuItem) ?usize {
    if (items.len == 0) return null;
    if (items.len == 1) {
        printHeader(title);
        renderItems(items, 0);
        std.debug.print("\n\x1b[90mUna sola opcion disponible, seleccionada automaticamente.\x1b[0m\n", .{});
        return 0;
    }

    var raw = keyinput.RawMode.enter();
    defer raw.leave();

    std.debug.print("\x1b[?25l", .{}); // ocultar cursor
    defer std.debug.print("\x1b[?25h", .{});

    var current: usize = 0;
    printHeader(title);
    renderItems(items, current);

    while (true) {
        const key = keyinput.readKey();
        switch (key) {
            .arrow_up => {
                current = if (current == 0) items.len - 1 else current - 1;
                redrawItems(items, current);
            },
            .arrow_down => {
                current = (current + 1) % items.len;
                redrawItems(items, current);
            },
            .digit => |d| {
                const n: usize = d - '0';
                if (n >= 1 and n <= items.len) {
                    current = n - 1;
                    redrawItems(items, current);
                }
            },
            .enter => return current,
            .escape => return null,
            else => {},
        }
    }
}

fn printHeader(title: []const u8) void {
    std.debug.print("\n\x1b[1;36m{s}\x1b[0m\n", .{title});
    std.debug.print("\x1b[90m(usa flechas Arriba/Abajo o numero, Enter para confirmar, ESC para cancelar)\x1b[0m\n\n", .{});
}

fn renderItems(items: []const MenuItem, selected: usize) void {
    for (items, 0..) |item, i| {
        renderRow(item, i, i == selected);
    }
}

fn redrawItems(items: []const MenuItem, selected: usize) void {
    var idx: usize = 0;
    while (idx < items.len) : (idx += 1) {
        std.debug.print("\x1b[1A\x1b[2K", .{});
    }
    renderItems(items, selected);
}

fn renderRow(item: MenuItem, index: usize, selected: bool) void {
    const num = index + 1;
    if (selected) {
        if (item.description) |d| {
            std.debug.print("\x1b[1;42;30m > [{d}] {s}\x1b[0m  \x1b[90m{s}\x1b[0m\n", .{ num, item.label, d });
        } else {
            std.debug.print("\x1b[1;42;30m > [{d}] {s}\x1b[0m\n", .{ num, item.label });
        }
    } else {
        if (item.description) |d| {
            std.debug.print("   [{d}] {s}  \x1b[90m{s}\x1b[0m\n", .{ num, item.label, d });
        } else {
            std.debug.print("   [{d}] {s}\n", .{ num, item.label });
        }
    }
}
