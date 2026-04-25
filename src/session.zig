const std = @import("std");

/// Contexto de la sesion del usuario tras pasar el preflight y seleccionar
/// un buzon. Se propaga a las acciones (importar / respaldar).
pub const Session = struct {
    store_id: []const u8,
    email: []const u8, // usualmente el DisplayName del store en Outlook (suele ser el correo)
    display_name: []const u8,
    file_path: ?[]const u8,
    exchange_store_type: ?i64,

    pub fn print(self: Session) void {
        std.debug.print("\x1b[90m   Cuenta:\x1b[0m  \x1b[1m{s}\x1b[0m\n", .{self.email});
        std.debug.print("\x1b[90m   StoreId:\x1b[0m {s}\n", .{self.store_id});
        if (self.file_path) |fp| {
            std.debug.print("\x1b[90m   Ruta:\x1b[0m    {s}\n", .{fp});
        }
    }
};
