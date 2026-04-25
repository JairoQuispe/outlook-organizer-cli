const std = @import("std");
const tui = @import("zigtui");

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

    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer {
        const leaked = gpa.deinit();
        if (leaked == .leak) {
            // Silenciar warnings de leak - el buffer del TUI maneja su propia memoria
        }
    }
    const allocator = gpa.allocator();

    var backend = tui.init(allocator) catch {
        return selectFallback(title, items);
    };
    defer backend.deinit();

    var terminal = tui.Terminal.init(allocator, backend.interface()) catch {
        return selectFallback(title, items);
    };
    defer terminal.deinit();

    terminal.hideCursor() catch {};
    defer terminal.showCursor() catch {};

    var state = MenuState{
        .title = title,
        .items = items,
        .selected = 0,
    };

    while (true) {
        const ev = backend.interface().pollEvent(50) catch tui.Event.none;
        switch (ev) {
            .key => |k| {
                switch (k.code) {
                    .up => {
                        state.selected = if (state.selected == 0) items.len - 1 else state.selected - 1;
                    },
                    .down => {
                        state.selected = (state.selected + 1) % items.len;
                    },
                    .char => |c| {
                        if (c >= '1' and c <= '9') {
                            const n: usize = c - '0';
                            if (n >= 1 and n <= items.len) {
                                state.selected = n - 1;
                            }
                        }
                    },
                    .enter => return state.selected,
                    .esc => return null,
                    else => {},
                }
            },
            else => {},
        }

        terminal.draw(&state, renderMenu) catch {};
    }
}

const MenuState = struct {
    title: []const u8,
    items: []const MenuItem,
    selected: usize,
};

fn renderMenu(state: *const MenuState, buf: *tui.Buffer) anyerror!void {
    const area = tui.Rect{
        .x = 0,
        .y = 0,
        .width = buf.width,
        .height = buf.height,
    };

    var block = tui.widgets.Block{
        .title = state.title,
        .borders = tui.widgets.Borders.ALL,
        .border_style = .{ .fg = .cyan },
    };

    const inner = block.inner(area);
    block.render(area, buf);

    var list_items = std.ArrayList(tui.widgets.ListItem){};
    defer {
        for (list_items.items) |item| {
            buf.allocator.free(item.content);
        }
        list_items.deinit(buf.allocator);
    }

    for (state.items, 0..) |item, i| {
        const num = i + 1;
        var text_buf: [512]u8 = undefined;
        const text = if (item.description) |d|
            std.fmt.bufPrint(&text_buf, "[{d}] {s}  {s}", .{ num, item.label, d }) catch item.label
        else
            std.fmt.bufPrint(&text_buf, "[{d}] {s}", .{ num, item.label }) catch item.label;

        const owned = buf.allocator.dupe(u8, text) catch continue;
        list_items.append(buf.allocator, .{ .content = owned }) catch continue;
    }

    const list = tui.widgets.List{
        .items = list_items.items,
        .highlight_style = .{
            .fg = .black,
            .bg = .green,
            .modifier = tui.Modifier.BOLD,
        },
        .highlight_symbol = " > ",
        .selected = state.selected,
    };

    list.render(inner, buf);

    const help_y = if (area.height > 2) area.height - 2 else 0;
    if (help_y > 0) {
        const help_area = tui.Rect{
            .x = area.x,
            .y = area.y + help_y,
            .width = area.width,
            .height = 1,
        };
        const help = tui.widgets.Paragraph{
            .text = "Flechas/Numeros: navegar | Enter: confirmar | ESC: cancelar",
            .style = .{ .fg = .gray },
        };
        help.render(help_area, buf);
    }
}

fn selectFallback(title: []const u8, items: []const MenuItem) ?usize {
    const keyinput = @import("../windows/keyinput.zig");

    var raw = keyinput.RawMode.enter();
    defer raw.leave();

    std.debug.print("\x1b[?25l", .{});
    defer std.debug.print("\x1b[?25h", .{});

    var current: usize = 0;
    std.debug.print("\n\x1b[1;36m{s}\x1b[0m\n", .{title});
    std.debug.print("\x1b[90m(usa flechas Arriba/Abajo o numero, Enter para confirmar, ESC para cancelar)\x1b[0m\n\n", .{});
    renderItemsFallback(items, current);

    while (true) {
        const key = keyinput.readKey();
        switch (key) {
            .arrow_up => {
                current = if (current == 0) items.len - 1 else current - 1;
                redrawItemsFallback(items, current);
            },
            .arrow_down => {
                current = (current + 1) % items.len;
                redrawItemsFallback(items, current);
            },
            .digit => |d| {
                const n: usize = d - '0';
                if (n >= 1 and n <= items.len) {
                    current = n - 1;
                    redrawItemsFallback(items, current);
                }
            },
            .enter => return current,
            .escape => return null,
            else => {},
        }
    }
}

fn renderItemsFallback(items: []const MenuItem, selected: usize) void {
    for (items, 0..) |item, i| {
        renderRowFallback(item, i, i == selected);
    }
}

fn redrawItemsFallback(items: []const MenuItem, selected: usize) void {
    var idx: usize = 0;
    while (idx < items.len) : (idx += 1) {
        std.debug.print("\x1b[1A\x1b[2K", .{});
    }
    renderItemsFallback(items, selected);
}

fn renderRowFallback(item: MenuItem, index: usize, selected: bool) void {
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
