const std = @import("std");
const list_stores = @import("../stores/list_stores.zig");
const keyinput = @import("../windows/keyinput.zig");

const Store = list_stores.Store;

/// Renderiza un selector de stores con flechas + numero + Enter.
/// Devuelve el indice seleccionado o null si el usuario cancela (ESC).
pub fn selectStore(stores: []const Store) ?usize {
    if (stores.len == 0) {
        std.debug.print("\x1b[33m[!] No hay buzones conectados a Outlook.\x1b[0m\n", .{});
        return null;
    }

    if (stores.len == 1) {
        renderStores(stores, 0);
        std.debug.print("\n\x1b[90mSolo hay un buzon conectado, seleccionado automaticamente.\x1b[0m\n", .{});
        return 0;
    }

    var raw = keyinput.RawMode.enter();
    defer raw.leave();

    var current: usize = 0;

    // Oculta cursor para mejor presentacion
    std.debug.print("\x1b[?25l", .{});
    defer std.debug.print("\x1b[?25h", .{});

    renderStores(stores, current);

    while (true) {
        const key = keyinput.readKey();
        switch (key) {
            .arrow_up => {
                if (current == 0) {
                    current = stores.len - 1;
                } else {
                    current -= 1;
                }
                redrawStores(stores, current);
            },
            .arrow_down => {
                current = (current + 1) % stores.len;
                redrawStores(stores, current);
            },
            .digit => |d| {
                // Permite 1..9 directos (los index visibles empiezan en 1)
                const n: usize = d - '0';
                if (n >= 1 and n <= stores.len) {
                    current = n - 1;
                    redrawStores(stores, current);
                }
            },
            .enter => return current,
            .escape => return null,
            else => {},
        }
    }
}

fn renderStores(stores: []const Store, selected: usize) void {
    std.debug.print("\n\x1b[1;36mBuzones conectados a Outlook:\x1b[0m\n", .{});
    std.debug.print("\x1b[90m(usa flechas Arriba/Abajo o numero, Enter para confirmar, ESC para cancelar)\x1b[0m\n\n", .{});

    for (stores, 0..) |s, i| {
        renderRow(s, i, i == selected);
    }
}

fn redrawStores(stores: []const Store, selected: usize) void {
    // Sube N lineas (una por store) y reescribe cada una
    const lines = stores.len;
    // Mover cursor al inicio del bloque y limpiar cada linea
    var idx: usize = 0;
    while (idx < lines) : (idx += 1) {
        std.debug.print("\x1b[1A\x1b[2K", .{});
    }
    for (stores, 0..) |s, i| {
        renderRow(s, i, i == selected);
    }
}

fn renderRow(s: Store, index: usize, selected: bool) void {
    const num = index + 1;
    const path = s.file_path orelse "Modo Online / Sin ruta local";
    const exch = list_stores.describeExchangeType(s.exchange_store_type);

    if (selected) {
        std.debug.print("\x1b[1;42;30m > [{d}] {s}{s}\x1b[0m \x1b[90m({s})\x1b[0m\n", .{ num, s.display_name, exch, path });
    } else {
        std.debug.print("   [{d}] {s}{s} \x1b[90m({s})\x1b[0m\n", .{ num, s.display_name, exch, path });
    }
}
