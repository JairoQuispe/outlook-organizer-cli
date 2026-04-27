const std = @import("std");

pub const MenuItem = struct {
    label: []const u8,
    description: ?[]const u8 = null,
};

/// Muestra un menu interactivo y devuelve el indice elegido o null si se cancela con ESC.
/// Soporta flechas Arriba/Abajo, numeros 1..N y Enter.
pub fn select(title: []const u8, items: []const MenuItem) ?usize {
    if (items.len == 0) return null;
    if (items.len == 1) {
        std.debug.print("\n\x1b[1;36m{s}\x1b[0m\n", .{title});
        std.debug.print("   [1] {s}", .{items[0].label});
        if (items[0].description) |d| {
            std.debug.print("  \x1b[90m{s}\x1b[0m", .{d});
        }
        std.debug.print("\n\n\x1b[90mUna sola opcion disponible, seleccionada automaticamente.\x1b[0m\n", .{});
        return 0;
    }

    // Usar ANSI + raw mode para compatibilidad con Windows 10/11
    return selectWithAnsi(title, items);
}

fn selectWithAnsi(title: []const u8, items: []const MenuItem) ?usize {
    const keyinput = @import("../windows/keyinput.zig");

    var raw = keyinput.RawMode.enter();
    defer raw.leave();

    std.debug.print("\x1b[?25l", .{});
    defer std.debug.print("\x1b[?25h", .{});

    var current: usize = 0;
    std.debug.print("\n\x1b[1;36m{s}\x1b[0m\n", .{title});
    std.debug.print("\x1b[90m(usa flechas Arriba/Abajo o numero, Enter para confirmar, ESC para cancelar)\x1b[0m\n\n", .{});
    std.debug.print("\x1b[s", .{}); // guardar posicion del cursor para redibujar
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

fn renderItems(items: []const MenuItem, selected: usize) void {
    for (items, 0..) |item, i| {
        renderRow(item, i, i == selected);
    }
}

fn redrawItems(items: []const MenuItem, selected: usize) void {
    // Restaura cursor a la posicion guardada (justo antes de las filas) y
    // limpia desde alli hasta el final de pantalla. Robusto frente a wrap
    // de lineas largas en terminales angostos.
    std.debug.print("\x1b[u\x1b[J", .{});
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
