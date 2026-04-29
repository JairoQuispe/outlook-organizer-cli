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

/// Muestra un menu interactivo multi-seleccion y devuelve los indices elegidos.
/// Teclas: Arriba/Abajo mover, Espacio toggle, A seleccionar todos,
/// N limpiar seleccion, Enter confirmar, ESC cancelar.
pub fn selectMultiple(allocator: std.mem.Allocator, title: []const u8, items: []const MenuItem) !?[]usize {
    if (items.len == 0) return null;

    const keyinput = @import("../windows/keyinput.zig");
    var selected = try allocator.alloc(bool, items.len);
    errdefer allocator.free(selected);
    @memset(selected, false);

    var raw = keyinput.RawMode.enter();
    defer raw.leave();

    std.debug.print("\x1b[?25l", .{});
    defer std.debug.print("\x1b[?25h", .{});

    var current: usize = 0;
    std.debug.print("\n\x1b[1;36m{s}\x1b[0m\n", .{title});
    std.debug.print("\x1b[90m(Arriba/Abajo, Espacio=marcar, A=todos, N=ninguno, Enter=confirmar, ESC=cancelar)\x1b[0m\n\n", .{});
    renderMultiItems(items, selected, current);

    while (true) {
        const key = keyinput.readKey();
        switch (key) {
            .arrow_up => {
                current = if (current == 0) items.len - 1 else current - 1;
                redrawMultiItems(items, selected, current);
            },
            .arrow_down => {
                current = (current + 1) % items.len;
                redrawMultiItems(items, selected, current);
            },
            .digit => |d| {
                const n: usize = d - '0';
                if (n >= 1 and n <= items.len) {
                    current = n - 1;
                    selected[current] = !selected[current];
                    redrawMultiItems(items, selected, current);
                }
            },
            .char => |c| {
                switch (c) {
                    ' ' => {
                        selected[current] = !selected[current];
                        redrawMultiItems(items, selected, current);
                    },
                    'a', 'A' => {
                        @memset(selected, true);
                        redrawMultiItems(items, selected, current);
                    },
                    'n', 'N' => {
                        @memset(selected, false);
                        redrawMultiItems(items, selected, current);
                    },
                    else => {},
                }
            },
            .enter => {
                var count: usize = 0;
                for (selected) |s| {
                    if (s) count += 1;
                }
                if (count == 0) {
                    continue;
                }
                var out = try allocator.alloc(usize, count);
                var j: usize = 0;
                for (selected, 0..) |s, i| {
                    if (s) {
                        out[j] = i;
                        j += 1;
                    }
                }
                allocator.free(selected);
                return out;
            },
            .escape => {
                allocator.free(selected);
                return null;
            },
            else => {},
        }
    }
}

fn renderMultiRow(item: MenuItem, index: usize, checked: bool, current: bool) void {
    const num = index + 1;
    const mark: []const u8 = if (checked) "X" else " ";
    if (current) {
        if (item.description) |d| {
            std.debug.print("\x1b[1;42;30m > [{d}] [{s}] {s}\x1b[0m  \x1b[90m{s}\x1b[0m\n", .{ num, mark, item.label, d });
        } else {
            std.debug.print("\x1b[1;42;30m > [{d}] [{s}] {s}\x1b[0m\n", .{ num, mark, item.label });
        }
    } else {
        if (item.description) |d| {
            std.debug.print("   [{d}] [{s}] {s}  \x1b[90m{s}\x1b[0m\n", .{ num, mark, item.label, d });
        } else {
            std.debug.print("   [{d}] [{s}] {s}\n", .{ num, mark, item.label });
        }
    }
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

fn renderMultiItems(items: []const MenuItem, selected_map: []const bool, current: usize) void {
    for (items, 0..) |item, i| {
        renderMultiRow(item, i, selected_map[i], i == current);
    }
}

fn redrawItems(items: []const MenuItem, selected: usize) void {
    // Mueve el cursor hacia arriba N lineas y limpia desde alli
    std.debug.print("\x1b[{d}A\x1b[J", .{items.len});
    renderItems(items, selected);
}

fn redrawMultiItems(items: []const MenuItem, selected_map: []const bool, current: usize) void {
    std.debug.print("\x1b[{d}A\x1b[J", .{items.len});
    renderMultiItems(items, selected_map, current);
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
