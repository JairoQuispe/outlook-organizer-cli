const std = @import("std");

pub const Store = struct {
    display_name: []const u8,
    store_id: ?[]const u8,
    file_path: ?[]const u8,
    file_size: ?[]const u8,
    exchange_store_type: ?i64,
};

pub const StoresResult = struct {
    arena: std.heap.ArenaAllocator,
    stores: []Store,

    pub fn deinit(self: *StoresResult) void {
        self.arena.deinit();
    }
};

/// Ejecuta el script PowerShell outlook-list-stores.ps1 con -Json
/// y parsea el resultado a un slice de Store.
pub fn fetchStores(parent_allocator: std.mem.Allocator, script_path: []const u8) !StoresResult {
    var arena = std.heap.ArenaAllocator.init(parent_allocator);
    errdefer arena.deinit();
    const allocator = arena.allocator();

    const result = try std.process.Child.run(.{
        .allocator = parent_allocator,
        .argv = &.{
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            script_path,
            "-Json",
        },
        .max_output_bytes = 4 * 1024 * 1024,
    });
    defer parent_allocator.free(result.stdout);
    defer parent_allocator.free(result.stderr);

    switch (result.term) {
        .Exited => |code| if (code != 0) {
            if (result.stderr.len > 0) {
                std.debug.print("\n\x1b[31mError del script PowerShell:\x1b[0m\n{s}\n", .{result.stderr});
            }
            return error.ScriptFailed;
        },
        else => return error.ScriptFailed,
    }

    // Buscar la primera linea que parezca JSON valido (empieza con '{')
    var json_text: []const u8 = "";
    var it = std.mem.splitScalar(u8, result.stdout, '\n');
    while (it.next()) |line| {
        const trimmed = std.mem.trim(u8, line, " \r\n\t");
        if (trimmed.len == 0) continue;
        if (trimmed[0] == '{') {
            json_text = trimmed;
            break;
        }
    }
    if (json_text.len == 0) return error.NoJsonInOutput;

    const parsed = try std.json.parseFromSlice(std.json.Value, allocator, json_text, .{});
    const root = parsed.value;
    if (root != .object) return error.InvalidPayload;

    const stores_value = root.object.get("stores") orelse return error.NoStoresKey;

    var list = std.ArrayList(Store){};
    if (stores_value == .array) {
        for (stores_value.array.items) |item| {
            const store = try parseStore(allocator, item);
            try list.append(allocator, store);
        }
    } else if (stores_value == .object) {
        try list.append(allocator, try parseStore(allocator, stores_value));
    } else {
        return error.InvalidStoresShape;
    }

    return StoresResult{
        .arena = arena,
        .stores = try list.toOwnedSlice(allocator),
    };
}

fn parseStore(allocator: std.mem.Allocator, value: std.json.Value) !Store {
    if (value != .object) return error.InvalidStore;
    const obj = value.object;

    return Store{
        .display_name = try dupString(allocator, obj.get("displayName")) orelse "(sin nombre)",
        .store_id = try dupString(allocator, obj.get("storeId")),
        .file_path = try dupString(allocator, obj.get("filePath")),
        .file_size = try dupString(allocator, obj.get("fileSize")),
        .exchange_store_type = readInt(obj.get("exchangeStoreType")),
    };
}

fn dupString(allocator: std.mem.Allocator, maybe: ?std.json.Value) !?[]const u8 {
    const v = maybe orelse return null;
    return switch (v) {
        .string => |s| blk: {
            const trimmed = std.mem.trim(u8, s, " \t\r\n");
            if (trimmed.len == 0) break :blk null;

            // Validate UTF-8 and sanitize if invalid
            if (!std.unicode.utf8ValidateSlice(trimmed)) {
                // Replace invalid bytes with '?'
                var sanitized = try allocator.alloc(u8, trimmed.len);
                for (trimmed, 0..) |byte, i| {
                    sanitized[i] = if (byte < 128) byte else '?';
                }
                break :blk sanitized;
            }

            break :blk try allocator.dupe(u8, trimmed);
        },
        .null => null,
        else => null,
    };
}

fn readInt(maybe: ?std.json.Value) ?i64 {
    const v = maybe orelse return null;
    return switch (v) {
        .integer => |i| i,
        .float => |f| @intFromFloat(f),
        .null => null,
        else => null,
    };
}

pub fn describeExchangeType(t: ?i64) []const u8 {
    if (t == null) return "";
    return switch (t.?) {
        0 => " [Exchange]",
        1 => " [Exchange Principal]",
        2 => " [Exchange Delegado]",
        3 => " [PST/Publicas]",
        4 => " [OST Cache]",
        else => "",
    };
}
