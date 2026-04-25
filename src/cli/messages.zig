const std = @import("std");

pub fn printHelp() void {
    std.debug.print(
        \\Outlook Organizer CLI (Zig)
        \\
        \\Uso:
        \\  outlook-organizer [comando]
        \\
        \\Comandos:
        \\  preflight    Ejecuta validaciones iniciales (por defecto)
        \\  help         Muestra esta ayuda
        \\
        \\Opciones:
        \\  --json       Imprime salida estructurada JSON
        \\
    , .{});
}

pub fn printPreflightSuccess() void {
    std.debug.print("[OK] Outlook clasico detectado. Validacion de instalacion completada.\n", .{});
}

pub fn printPreflightFailure() void {
    std.debug.print("[ERROR] No se detecto Outlook clasico instalado en este Windows.\n", .{});
}
