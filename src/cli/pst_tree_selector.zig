const std = @import("std");
const tui = @import("zigtui");

pub const FolderEntry = struct {
    path: []const u8,
    item_count: usize,
    year_summary: ?[]const u8 = null,
};

const Node = struct {
    segment: []u8,
    parent: ?*Node,
    children: std.ArrayList(*Node),
    folder_index: ?usize = null,
    item_count: usize = 0,
    year_summary: ?[]const u8 = null,
    expanded: bool = false,
    selected: bool = false,

    fn init(allocator: std.mem.Allocator, segment: []const u8, parent: ?*Node) !*Node {
        const ptr = try allocator.create(Node);
        ptr.* = .{
            .segment = try allocator.dupe(u8, segment),
            .parent = parent,
            .children = std.ArrayList(*Node){},
        };
        return ptr;
    }

    fn deinit(self: *Node, allocator: std.mem.Allocator) void {
        for (self.children.items) |child| {
            child.deinit(allocator);
            allocator.destroy(child);
        }
        self.children.deinit(allocator);
        allocator.free(self.segment);
    }
};

const View = struct {
    allocator: std.mem.Allocator,
    folders: []const FolderEntry,
    root: *Node,
    visible: std.ArrayList(*Node),
    selected_visible: usize = 0,
    frame_arena: std.heap.ArenaAllocator,

    fn init(allocator: std.mem.Allocator, folders: []const FolderEntry) !View {
        const root = try Node.init(allocator, "", null);
        root.expanded = true;

        var view = View{
            .allocator = allocator,
            .folders = folders,
            .root = root,
            .visible = std.ArrayList(*Node){},
            .selected_visible = 0,
            .frame_arena = std.heap.ArenaAllocator.init(allocator),
        };

        for (folders, 0..) |f, idx| {
            try insertFolderPath(allocator, root, f, idx);
        }

        for (root.children.items) |child| {
            child.expanded = true;
        }

        view.rebuildVisible();
        return view;
    }

    fn deinit(self: *View) void {
        self.visible.deinit(self.allocator);
        self.frame_arena.deinit();
        self.root.deinit(self.allocator);
        self.allocator.destroy(self.root);
    }

    fn rebuildVisible(self: *View) void {
        self.visible.clearRetainingCapacity();
        for (self.root.children.items) |child| {
            self.appendVisibleRecursive(child);
        }
        if (self.visible.items.len == 0) {
            self.selected_visible = 0;
        } else if (self.selected_visible >= self.visible.items.len) {
            self.selected_visible = self.visible.items.len - 1;
        }
    }

    fn appendVisibleRecursive(self: *View, node: *Node) void {
        self.visible.append(self.allocator, node) catch return;
        if (!node.expanded) return;
        for (node.children.items) |child| {
            self.appendVisibleRecursive(child);
        }
    }

    fn selectedNode(self: *View) ?*Node {
        if (self.visible.items.len == 0) return null;
        return self.visible.items[self.selected_visible];
    }

    fn toggleCurrentExpanded(self: *View) void {
        const node = self.selectedNode() orelse return;
        if (node.children.items.len == 0) return;
        node.expanded = !node.expanded;
        self.rebuildVisible();
    }

    fn collapseOrGoParent(self: *View) void {
        const node = self.selectedNode() orelse return;
        if (node.expanded and node.children.items.len > 0) {
            node.expanded = false;
            self.rebuildVisible();
            return;
        }

        const p = node.parent orelse return;
        if (p == self.root) return;

        for (self.visible.items, 0..) |v, i| {
            if (v == p) {
                self.selected_visible = i;
                break;
            }
        }
    }

    fn expandOrGoChild(self: *View) void {
        const node = self.selectedNode() orelse return;
        if (node.children.items.len == 0) return;
        if (!node.expanded) {
            node.expanded = true;
            self.rebuildVisible();
            return;
        }
        if (self.selected_visible + 1 < self.visible.items.len) {
            self.selected_visible += 1;
        }
    }

    fn toggleCurrentSelected(self: *View) void {
        const node = self.selectedNode() orelse return;
        if (node.folder_index == null) return;
        node.selected = !node.selected;
    }

    fn clearAllSelections(self: *View) void {
        clearRecursive(self.root);
    }

    fn selectAll(self: *View) void {
        selectRecursive(self.root);
    }

    fn hasAnySelected(self: *View) bool {
        return anySelectedRecursive(self.root);
    }

    fn collectSelectedIndexes(self: *View) ![]usize {
        const picked = try self.allocator.alloc(bool, self.folders.len);
        defer self.allocator.free(picked);
        @memset(picked, false);

        markSelectedRecursive(self.root, picked);

        var count: usize = 0;
        for (picked) |p| {
            if (p) count += 1;
        }

        var out = try self.allocator.alloc(usize, count);
        var j: usize = 0;
        for (picked, 0..) |p, i| {
            if (p) {
                out[j] = i;
                j += 1;
            }
        }
        return out;
    }
};

pub fn select(allocator: std.mem.Allocator, folders: []const FolderEntry) !?[]usize {
    if (folders.len == 0) return null;

    var view = try View.init(allocator, folders);
    defer view.deinit();

    var backend = try tui.init(allocator);
    defer backend.deinit();

    var terminal = try tui.Terminal.init(allocator, backend.interface());
    defer terminal.deinit();

    try terminal.hideCursor();
    defer terminal.showCursor() catch {};

    while (true) {
        try terminal.draw(&view, render);

        const ev = backend.interface().pollEvent(100) catch tui.Event.none;
        if (ev != .key) continue;

        switch (ev.key.code) {
            .up => {
                if (view.visible.items.len == 0) continue;
                if (view.selected_visible == 0) {
                    view.selected_visible = view.visible.items.len - 1;
                } else {
                    view.selected_visible -= 1;
                }
            },
            .down => {
                if (view.visible.items.len == 0) continue;
                view.selected_visible = (view.selected_visible + 1) % view.visible.items.len;
            },
            .left => view.collapseOrGoParent(),
            .right => view.expandOrGoChild(),
            .enter => {
                if (!view.hasAnySelected()) continue;
                return try view.collectSelectedIndexes();
            },
            .esc => return null,
            .char => |c| {
                switch (c) {
                    ' ' => view.toggleCurrentSelected(),
                    'a', 'A' => view.selectAll(),
                    'n', 'N' => view.clearAllSelections(),
                    else => {
                        if (c >= '1' and c <= '9') {
                            const n: usize = c - '0';
                            if (n >= 1 and n <= view.visible.items.len) {
                                view.selected_visible = n - 1;
                            }
                        }
                    },
                }
            },
            else => {},
        }
    }
}

fn insertFolderPath(allocator: std.mem.Allocator, root: *Node, folder: FolderEntry, idx: usize) !void {
    var current = root;
    var start: usize = 0;
    while (start < folder.path.len) {
        var end = start;
        while (end < folder.path.len and folder.path[end] != '\\' and folder.path[end] != '/') : (end += 1) {}

        const part = std.mem.trim(u8, folder.path[start..end], " \t");
        if (part.len > 0) {
            current = try findOrCreateChild(allocator, current, part);
        }

        start = end + 1;
    }

    current.folder_index = idx;
    current.item_count = folder.item_count;
    current.year_summary = folder.year_summary;
}

fn findOrCreateChild(allocator: std.mem.Allocator, parent: *Node, segment: []const u8) !*Node {
    for (parent.children.items) |child| {
        if (std.mem.eql(u8, child.segment, segment)) return child;
    }

    const child = try Node.init(allocator, segment, parent);
    try parent.children.append(allocator, child);
    return child;
}

fn clearRecursive(node: *Node) void {
    node.selected = false;
    for (node.children.items) |c| {
        clearRecursive(c);
    }
}

fn selectRecursive(node: *Node) void {
    if (node.folder_index != null) {
        node.selected = true;
    }
    for (node.children.items) |c| {
        selectRecursive(c);
    }
}

fn anySelectedRecursive(node: *Node) bool {
    if (node.folder_index != null and node.selected) return true;
    for (node.children.items) |c| {
        if (anySelectedRecursive(c)) return true;
    }
    return false;
}

fn markSelectedRecursive(node: *Node, picked: []bool) void {
    if (node.folder_index) |idx| {
        if (node.selected and idx < picked.len) picked[idx] = true;
    }
    for (node.children.items) |c| {
        markSelectedRecursive(c, picked);
    }
}

fn render(view: *View, buf: *tui.Buffer) anyerror!void {
    _ = view.frame_arena.reset(.retain_capacity);
    const arena = view.frame_arena.allocator();

    const area = buf.getArea();
    if (area.width < 44 or area.height < 10) return;

    const root = tui.Block{
        .title = " Seleccion de carpetas PST ",
        .borders = tui.Borders.ALL,
        .border_style = .{ .fg = .cyan },
        .title_style = .{ .modifier = .{ .bold = true } },
        .border_symbols = tui.BorderSymbols.rounded(),
    };
    root.render(area, buf);

    const inner = root.inner(area);
    if (inner.height <= 3) return;

    buf.setStringTruncated(
        inner.x,
        inner.y,
        "Arriba/Abajo: mover | Izq/Der: contraer/expandir | Espacio: marcar | A: todas | N: ninguna | Enter: confirmar",
        inner.width,
        .{ .fg = .gray },
    );

    var roots = try arena.alloc(tui.widgets.TreeNode, view.root.children.items.len);
    for (view.root.children.items, 0..) |child, i| {
        roots[i] = try buildTreeNode(arena, child);
    }

    const tree = tui.widgets.Tree{
        .roots = roots,
        .selected = view.selected_visible,
        .highlight_style = .{ .fg = .black, .bg = .cyan, .modifier = .{ .bold = true } },
        .indent = 2,
        .expanded_symbol = "▼ ",
        .collapsed_symbol = "▶ ",
        .leaf_symbol = "  ",
    };

    tree.render(
        .{ .x = inner.x, .y = inner.y + 1, .width = inner.width, .height = inner.height - 2 },
        buf,
    );

    var info_buf: [96]u8 = undefined;
    const info = std.fmt.bufPrint(
        &info_buf,
        "Seleccionadas: {d} / {d} (ESC cancela)",
        .{ countSelected(view.root), view.folders.len },
    ) catch "";
    buf.setStringTruncated(inner.x, inner.y + inner.height - 1, info, inner.width, .{ .fg = .yellow });
}

fn buildTreeNode(allocator: std.mem.Allocator, node: *Node) !tui.widgets.TreeNode {
    const label = try buildLabel(allocator, node);

    var out = tui.widgets.TreeNode{
        .label = label,
        .expanded = node.expanded,
    };

    if (node.children.items.len > 0) {
        var kids = try allocator.alloc(tui.widgets.TreeNode, node.children.items.len);
        for (node.children.items, 0..) |c, i| {
            kids[i] = try buildTreeNode(allocator, c);
        }
        out.children = kids;
    }

    return out;
}

fn buildLabel(allocator: std.mem.Allocator, node: *Node) ![]const u8 {
    if (node.folder_index != null) {
        const mark = if (node.selected) "[x]" else "[ ]";
        if (node.year_summary) |ys| {
            return try std.fmt.allocPrint(allocator, "{s} {s} ({d}) - {s}", .{ mark, node.segment, node.item_count, ys });
        }
        return try std.fmt.allocPrint(allocator, "{s} {s} ({d})", .{ mark, node.segment, node.item_count });
    }

    return try std.fmt.allocPrint(allocator, "    {s}", .{node.segment});
}

fn countSelected(node: *Node) usize {
    var total: usize = 0;
    if (node.folder_index != null and node.selected) total += 1;
    for (node.children.items) |c| {
        total += countSelected(c);
    }
    return total;
}
