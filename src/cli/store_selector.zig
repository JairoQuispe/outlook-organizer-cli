const std = @import("std");
const list_stores = @import("../stores/list_stores.zig");
const tui = @import("zigtui");

const Store = list_stores.Store;

/// Renderiza un selector de stores con flechas + numero + Enter.
/// Devuelve el indice seleccionado o null si el usuario cancela (ESC).
pub fn selectStore(allocator: std.mem.Allocator, stores: []const Store) ?usize {
    if (stores.len == 0) {
        std.debug.print("\x1b[33m[!] No hay buzones conectados a Outlook.\x1b[0m\n", .{});
        return null;
    }

    if (stores.len == 1) {
        const exch = normalizeExchangeLabel(stores[0].exchange_store_type);
        const path = formatPath(stores[0].file_path);
        const type_label = if (isMissingStoreId(stores[0])) "Sin StoreId" else exch;
        std.debug.print("\n\x1b[1;36mBuzones conectados a Outlook:\x1b[0m\n", .{});
        std.debug.print("   [1] {s} {s} \x1b[90m({s})\x1b[0m\n", .{ stores[0].display_name, type_label, path });
        std.debug.print("\n\x1b[90mSolo hay un buzon conectado, seleccionado automaticamente.\x1b[0m\n", .{});
        return 0;
    }

    var view = StoreTableView.init(allocator, stores) catch return null;
    defer view.deinit();

    var backend = tui.init(allocator) catch return null;
    defer backend.deinit();

    var terminal = tui.Terminal.init(allocator, backend.interface()) catch return null;
    defer terminal.deinit();

    terminal.hideCursor() catch return null;
    defer terminal.showCursor() catch {};

    while (true) {
        terminal.draw(&view, renderStoreTableView) catch return null;

        const ev = backend.interface().pollEvent(100) catch tui.Event.none;
        if (ev != .key) continue;

        switch (ev.key.code) {
            .up => {
                if (view.selected == 0) {
                    view.selected = stores.len - 1;
                } else {
                    view.selected -= 1;
                }
            },
            .down => {
                view.selected = (view.selected + 1) % stores.len;
            },
            .enter => return view.selected,
            .esc => return null,
            .char => |c| {
                if (c >= '1' and c <= '9') {
                    const n: usize = c - '0';
                    if (n >= 1 and n <= stores.len) {
                        view.selected = n - 1;
                    }
                }
            },
            else => {},
        }
    }
}

const StoreTableView = struct {
    allocator: std.mem.Allocator,
    stores: []const Store,
    selected: usize = 0,
    number_labels: [][]u8,

    fn init(allocator: std.mem.Allocator, stores: []const Store) !StoreTableView {
        var labels = try allocator.alloc([]u8, stores.len);
        errdefer allocator.free(labels);

        for (stores, 0..) |_, i| {
            labels[i] = try std.fmt.allocPrint(allocator, "{d}", .{i + 1});
        }

        return .{
            .allocator = allocator,
            .stores = stores,
            .selected = 0,
            .number_labels = labels,
        };
    }

    fn deinit(self: *StoreTableView) void {
        for (self.number_labels) |txt| {
            self.allocator.free(txt);
        }
        self.allocator.free(self.number_labels);
    }
};

fn renderStoreTableView(view: *StoreTableView, buf: *tui.Buffer) anyerror!void {
    const area = buf.getArea();
    if (area.width < 40 or area.height < 8) return;

    const root = tui.Block{
        .title = " Buzones conectados a Outlook ",
        .borders = tui.Borders.ALL,
        .border_style = .{ .fg = .cyan },
        .title_style = .{ .modifier = .{ .bold = true } },
        .border_symbols = tui.BorderSymbols.rounded(),
    };
    root.render(area, buf);
    const inner = root.inner(area);
    if (inner.height <= 2 or inner.width <= 4) return;

    buf.setStringTruncated(
        inner.x,
        inner.y,
        "Flechas: navegar | 1..9: salto rapido | Enter: confirmar | ESC: cancelar",
        inner.width,
        .{ .fg = .gray },
    );

    const table_y = inner.y + 1;
    const table_h: u16 = inner.height - 1;
    if (table_h == 0) return;

    const cols = [_]tui.widgets.Column{
        .{ .header = "#", .width = 4 },
        .{ .header = "Nombre" },
        .{ .header = "Tipo", .width = 12 },
        .{ .header = "Ruta" },
    };

    var cell_rows = std.ArrayList([4][]const u8){};
    defer cell_rows.deinit(view.allocator);
    cell_rows.ensureTotalCapacity(view.allocator, view.stores.len) catch return;

    for (view.stores, 0..) |s, i| {
        const path = formatPath(s.file_path);
        const exch = normalizeExchangeLabel(s.exchange_store_type);
        const type_label = if (isMissingStoreId(s)) "Sin StoreId" else exch;
        cell_rows.appendAssumeCapacity(.{ view.number_labels[i], s.display_name, type_label, path });
    }

    var rows_buf = std.ArrayList(tui.widgets.Row){};
    defer rows_buf.deinit(view.allocator);
    rows_buf.ensureTotalCapacity(view.allocator, view.stores.len) catch return;

    for (cell_rows.items) |*row_cells| {
        rows_buf.appendAssumeCapacity(.{ .cells = row_cells[0..] });
    }

    const table = tui.widgets.Table{
        .columns = &cols,
        .rows = rows_buf.items,
        .header_style = .{ .fg = .light_cyan, .modifier = .{ .bold = true } },
        .selected_style = .{ .fg = .black, .bg = .cyan, .modifier = .{ .bold = true } },
        .selected = view.selected,
    };
    table.render(.{ .x = inner.x, .y = table_y, .width = inner.width, .height = table_h }, buf);
}

fn formatPath(path_opt: ?[]const u8) []const u8 {
    if (path_opt) |path| {
        if (path.len > 0) return path;
    }
    return "Online / Sin ruta local";
}

fn normalizeExchangeLabel(t: ?i64) []const u8 {
    const raw = list_stores.describeExchangeType(t);
    if (raw.len == 0) return "";
    if (raw[0] == ' ' and raw.len > 1) return raw[1..];
    return raw;
}

fn isMissingStoreId(store: Store) bool {
    if (store.store_id) |id| {
        return id.len == 0;
    }
    return true;
}
