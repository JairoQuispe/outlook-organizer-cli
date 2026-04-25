const std = @import("std");
const builtin = @import("builtin");

pub const Key = union(enum) {
    arrow_up,
    arrow_down,
    arrow_left,
    arrow_right,
    enter,
    escape,
    digit: u8, // '0'..'9'
    char: u8,
    other,
};

const STD_INPUT_HANDLE: i32 = -10;
const ENABLE_LINE_INPUT: u32 = 0x0002;
const ENABLE_ECHO_INPUT: u32 = 0x0004;
const ENABLE_PROCESSED_INPUT: u32 = 0x0001;
const ENABLE_VIRTUAL_TERMINAL_INPUT: u32 = 0x0200;

extern "kernel32" fn GetStdHandle(nStdHandle: i32) callconv(.winapi) ?*anyopaque;
extern "kernel32" fn GetConsoleMode(hConsoleHandle: ?*anyopaque, lpMode: *u32) callconv(.winapi) i32;
extern "kernel32" fn SetConsoleMode(hConsoleHandle: ?*anyopaque, dwMode: u32) callconv(.winapi) i32;
extern "kernel32" fn ReadFile(
    hFile: ?*anyopaque,
    lpBuffer: *anyopaque,
    nNumberOfBytesToRead: u32,
    lpNumberOfBytesRead: *u32,
    lpOverlapped: ?*anyopaque,
) callconv(.winapi) i32;

pub const RawMode = struct {
    handle: ?*anyopaque,
    original_mode: u32,
    active: bool,

    pub fn enter() RawMode {
        if (builtin.os.tag != .windows) {
            return .{ .handle = null, .original_mode = 0, .active = false };
        }
        const handle = GetStdHandle(STD_INPUT_HANDLE) orelse return .{ .handle = null, .original_mode = 0, .active = false };
        var mode: u32 = 0;
        if (GetConsoleMode(handle, &mode) == 0) return .{ .handle = handle, .original_mode = 0, .active = false };

        var new_mode: u32 = mode;
        new_mode &= ~ENABLE_LINE_INPUT;
        new_mode &= ~ENABLE_ECHO_INPUT;
        new_mode |= ENABLE_PROCESSED_INPUT;
        new_mode |= ENABLE_VIRTUAL_TERMINAL_INPUT;

        _ = SetConsoleMode(handle, new_mode);
        return .{ .handle = handle, .original_mode = mode, .active = true };
    }

    pub fn leave(self: *RawMode) void {
        if (!self.active) return;
        _ = SetConsoleMode(self.handle, self.original_mode);
        self.active = false;
    }
};

pub fn readKey() Key {
    if (builtin.os.tag != .windows) return .other;
    const handle = GetStdHandle(STD_INPUT_HANDLE) orelse return .other;

    var byte: u8 = 0;
    var read: u32 = 0;
    if (ReadFile(handle, &byte, 1, &read, null) == 0 or read == 0) return .other;

    switch (byte) {
        '\r', '\n' => return .enter,
        0x1b => {
            // Posible secuencia ANSI: ESC [ X
            var next: u8 = 0;
            var n: u32 = 0;
            if (ReadFile(handle, &next, 1, &n, null) == 0 or n == 0) return .escape;
            if (next != '[') return .escape;
            var third: u8 = 0;
            if (ReadFile(handle, &third, 1, &n, null) == 0 or n == 0) return .escape;
            return switch (third) {
                'A' => .arrow_up,
                'B' => .arrow_down,
                'C' => .arrow_right,
                'D' => .arrow_left,
                else => .other,
            };
        },
        '0'...'9' => return .{ .digit = byte },
        else => return .{ .char = byte },
    }
}
