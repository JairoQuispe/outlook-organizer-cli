const std = @import("std");
const builtin = @import("builtin");

const ENABLE_VIRTUAL_TERMINAL_PROCESSING: u32 = 0x0004;
const STD_OUTPUT_HANDLE: i32 = -11;
const STD_ERROR_HANDLE: i32 = -12;

extern "kernel32" fn GetStdHandle(nStdHandle: i32) callconv(.winapi) ?*anyopaque;
extern "kernel32" fn GetConsoleMode(hConsoleHandle: ?*anyopaque, lpMode: *u32) callconv(.winapi) i32;
extern "kernel32" fn SetConsoleMode(hConsoleHandle: ?*anyopaque, dwMode: u32) callconv(.winapi) i32;
extern "kernel32" fn SetConsoleOutputCP(wCodePageID: u32) callconv(.winapi) i32;

/// Habilita procesamiento de secuencias ANSI/VT y UTF-8 en la consola de Windows
/// para que spinner, colores y emojis se rendericen correctamente.
pub fn enableModernConsole() void {
    if (builtin.os.tag != .windows) return;

    // UTF-8 para que los caracteres Braille y emojis se vean bien
    _ = SetConsoleOutputCP(65001);

    enableVtForHandle(STD_OUTPUT_HANDLE);
    enableVtForHandle(STD_ERROR_HANDLE);
}

fn enableVtForHandle(handle_id: i32) void {
    const handle = GetStdHandle(handle_id) orelse return;
    var mode: u32 = 0;
    if (GetConsoleMode(handle, &mode) == 0) return;
    _ = SetConsoleMode(handle, mode | ENABLE_VIRTUAL_TERMINAL_PROCESSING);
}
