const std = @import("std");
const menu = @import("menu.zig");

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
    const items = [_]menu.MenuItem{
        .{ .label = "Si", .description = null },
        .{ .label = "No", .description = null },
    };
    _ = default_yes; // reservado para futuro default prerseleccionado
    const idx = menu.select(question, &items) orelse return false;
    return idx == 0;
}
