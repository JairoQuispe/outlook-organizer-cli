const std = @import("std");
const Session = @import("../session.zig").Session;
const menu = @import("../cli/menu.zig");
const import_pst = @import("import_pst.zig");

pub const Action = enum {
    import_pst,
    backup_pst,
};

pub const ActionResult = enum(u8) {
    success = 0,
    failure = 1,
    cancelled = 2,
};

pub fn promptAction() ?Action {
    const items = [_]menu.MenuItem{
        .{
            .label = "Importar correos desde un PST a la cuenta",
            .description = "Carga un archivo .pst y mueve/copia sus correos al buzon seleccionado",
        },
        .{
            .label = "Respaldar correos hacia un PST",
            .description = "Genera un archivo .pst con los correos del buzon seleccionado",
        },
    };

    const idx = menu.select("\xc2\xbfQue accion deseas realizar?", &items) orelse return null;
    return switch (idx) {
        0 => .import_pst,
        1 => .backup_pst,
        else => null,
    };
}

pub fn runAction(action: Action, session: Session, allocator: std.mem.Allocator) ActionResult {
    return switch (action) {
        .import_pst => runImport(session, allocator),
        .backup_pst => runBackup(session),
    };
}

fn runImport(session: Session, allocator: std.mem.Allocator) ActionResult {
    const code = import_pst.run(session, allocator) catch |err| {
        std.debug.print("\x1b[31m[X]\x1b[0m  Error en el wizard de importacion: {s}\n", .{@errorName(err)});
        return .failure;
    };
    return switch (code) {
        0 => .success,
        2 => .cancelled,
        else => .failure,
    };
}

fn runBackup(session: Session) ActionResult {
    std.debug.print("\n\x1b[1;36m== Respaldar correos hacia PST ==\x1b[0m\n", .{});
    session.print();
    std.debug.print("\n\x1b[33m[!] Accion pendiente de implementar (se enganchara a backup_outlook_yearly.ps1).\x1b[0m\n", .{});
    return .success;
}
