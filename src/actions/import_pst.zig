const std = @import("std");
const Session = @import("../session.zig").Session;
const input = @import("../cli/input.zig");
const menu = @import("../cli/menu.zig");
const file_browser = @import("../cli/file_browser.zig");
const scripts = @import("../scripts/scripts.zig");
const import_progress = @import("import_progress.zig");

/// Ejecuta el wizard de importacion de PST. Devuelve exit code del script
/// (0 = exito) o un codigo de error del wizard.
pub fn run(session: Session, allocator: std.mem.Allocator) !u8 {
    std.debug.print("\n\x1b[1;36m== Importar correos desde PST ==\x1b[0m\n", .{});
    session.print();
    std.debug.print("\n", .{});

    // 1) Ruta del PST
    const pst_path = try askPstPath(allocator);
    defer allocator.free(pst_path);

    // 2) Accion Copy / Move
    const action = askAction() orelse {
        std.debug.print("\n\x1b[33m[!] Importacion cancelada.\x1b[0m\n", .{});
        return 2;
    };

    // 3) Filtro por anio
    const filter_year = try askFilterYear(allocator);
    defer if (filter_year) |fy| allocator.free(fy);

    // 4) Filtro por meses (solo si hay anio)
    var filter_months: ?[]u8 = null;
    defer if (filter_months) |fm| allocator.free(fm);
    if (filter_year != null) {
        filter_months = try askFilterMonths(allocator);
    }

    // 5) Saltar duplicados
    const skip_dupes = input.confirm("\xc2\xbfSaltar duplicados (por Message-Id)?", true);

    // 6) Resumen y confirmacion
    printSummary(session, pst_path, action, filter_year, filter_months, skip_dupes);
    if (!input.confirm("\xc2\xbfEjecutar la importacion con estos parametros?", true)) {
        std.debug.print("\n\x1b[33m[!] Importacion cancelada.\x1b[0m\n", .{});
        return 2;
    }

    // 7) Ejecutar PowerShell heredando stdio para ver progreso
    return try executePowerShell(allocator, session, pst_path, action, filter_year, filter_months, skip_dupes);
}

fn askPstPath(allocator: std.mem.Allocator) ![]u8 {
    while (true) {
        const method_items = [_]menu.MenuItem{
            .{
                .label = "Buscar archivo (explorador de terminal)",
                .description = "Navega con flechas, filtra .pst automaticamente",
            },
            .{
                .label = "Escribir la ruta manualmente",
                .description = "Pega o escribe una ruta completa al archivo .pst",
            },
        };
        const idx = menu.select("\xc2\xbfComo indicaras la ruta del PST?", &method_items) orelse {
            return error.UserCancelled;
        };

        const path_opt: ?[]u8 = switch (idx) {
            0 => try file_browser.browseForFile(allocator, null, ".pst"),
            else => try askPstPathManual(allocator),
        };

        const path = path_opt orelse {
            std.debug.print("\x1b[33m[!]\x1b[0m  Seleccion cancelada. Intenta de nuevo.\n", .{});
            continue;
        };

        // Validar que el archivo exista (por si la ruta manual se escribio mal)
        if (std.fs.cwd().access(path, .{})) |_| {
            return path;
        } else |_| {
            std.debug.print("\x1b[31m[X]\x1b[0m  Archivo no encontrado: {s}\n", .{path});
            allocator.free(path);
            continue;
        }
    }
}

fn askPstPathManual(allocator: std.mem.Allocator) !?[]u8 {
    const raw = try input.prompt(allocator, "Ruta del archivo PST (.pst)");
    defer allocator.free(raw);

    const trimmed = std.mem.trim(u8, raw, " \t\"'");
    if (trimmed.len == 0) return null;

    if (!std.ascii.endsWithIgnoreCase(trimmed, ".pst")) {
        std.debug.print("\x1b[33m[!]\x1b[0m  La ruta no termina en .pst. Se intentara igual.\n", .{});
    }

    return try allocator.dupe(u8, trimmed);
}

fn askAction() ?[]const u8 {
    const items = [_]menu.MenuItem{
        .{ .label = "Copy", .description = "Copiar correos al buzon (no modifica el PST)" },
        .{ .label = "Move", .description = "Mover correos (los elimina del PST)" },
    };
    const idx = menu.select("\xc2\xbfQue accion realizar con los correos?", &items) orelse return null;
    return if (idx == 0) "Copy" else "Move";
}

fn askFilterYear(allocator: std.mem.Allocator) !?[]u8 {
    if (!input.confirm("\xc2\xbfFiltrar por un anio especifico?", false)) return null;

    while (true) {
        const raw = try input.prompt(allocator, "Anio (ej. 2026)");
        defer allocator.free(raw);

        const trimmed = std.mem.trim(u8, raw, " \t");
        if (trimmed.len == 0) {
            std.debug.print("\x1b[31m[X]\x1b[0m  Debes ingresar un anio.\n", .{});
            continue;
        }

        const year = std.fmt.parseInt(i32, trimmed, 10) catch {
            std.debug.print("\x1b[31m[X]\x1b[0m  Anio invalido: {s}\n", .{trimmed});
            continue;
        };

        if (year < 1900 or year > 9999) {
            std.debug.print("\x1b[31m[X]\x1b[0m  Anio fuera de rango (1900-9999).\n", .{});
            continue;
        }

        return try allocator.dupe(u8, trimmed);
    }
}

fn askFilterMonths(allocator: std.mem.Allocator) !?[]u8 {
    if (!input.confirm("\xc2\xbfFiltrar tambien por meses especificos?", false)) return null;

    while (true) {
        const raw = try input.prompt(
            allocator,
            "Meses separados por coma (ej. 1,2,3 o enero,febrero,marzo)",
        );
        defer allocator.free(raw);

        const trimmed = std.mem.trim(u8, raw, " \t");
        if (trimmed.len == 0) {
            std.debug.print("\x1b[31m[X]\x1b[0m  Debes ingresar al menos un mes.\n", .{});
            continue;
        }

        return try allocator.dupe(u8, trimmed);
    }
}

fn printSummary(
    session: Session,
    pst_path: []const u8,
    action: []const u8,
    filter_year: ?[]const u8,
    filter_months: ?[]const u8,
    skip_dupes: bool,
) void {
    std.debug.print("\n\x1b[1;36m== Resumen de importacion ==\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mCuenta destino:\x1b[0m    \x1b[1m{s}\x1b[0m\n", .{session.email});
    std.debug.print("  \x1b[90mStoreId:\x1b[0m           {s}\n", .{session.store_id});
    std.debug.print("  \x1b[90mArchivo PST:\x1b[0m       {s}\n", .{pst_path});
    std.debug.print("  \x1b[90mAccion:\x1b[0m            {s}\n", .{action});
    if (filter_year) |fy| {
        std.debug.print("  \x1b[90mFiltro anio:\x1b[0m       {s}\n", .{fy});
    } else {
        std.debug.print("  \x1b[90mFiltro anio:\x1b[0m       (ninguno)\n", .{});
    }
    if (filter_months) |fm| {
        std.debug.print("  \x1b[90mFiltro meses:\x1b[0m      {s}\n", .{fm});
    } else {
        std.debug.print("  \x1b[90mFiltro meses:\x1b[0m      (ninguno)\n", .{});
    }
    std.debug.print("  \x1b[90mSaltar duplicados:\x1b[0m {s}\n", .{if (skip_dupes) "Si" else "No"});
    std.debug.print("\n", .{});
}

fn executePowerShell(
    allocator: std.mem.Allocator,
    session: Session,
    pst_path: []const u8,
    action: []const u8,
    filter_year: ?[]const u8,
    filter_months: ?[]const u8,
    skip_dupes: bool,
) !u8 {
    const script_path = try scripts.getScriptPath(allocator, .import_pst);
    defer allocator.free(script_path);

    var argv = std.ArrayList([]const u8){};
    defer argv.deinit(allocator);

    try argv.append(allocator, "powershell");
    try argv.append(allocator, "-NoProfile");
    try argv.append(allocator, "-ExecutionPolicy");
    try argv.append(allocator, "Bypass");
    try argv.append(allocator, "-File");
    try argv.append(allocator, script_path);

    try argv.append(allocator, "-PstPath");
    try argv.append(allocator, pst_path);

    try argv.append(allocator, "-TargetStoreId");
    try argv.append(allocator, session.store_id);

    try argv.append(allocator, "-Action");
    try argv.append(allocator, action);

    if (filter_year) |fy| {
        try argv.append(allocator, "-FilterOnlyYear");
        try argv.append(allocator, fy);
    }
    if (filter_months) |fm| {
        try argv.append(allocator, "-FilterOnlyMonths");
        try argv.append(allocator, fm);
    }
    if (skip_dupes) {
        try argv.append(allocator, "-SkipDuplicates");
    }

    try argv.append(allocator, "-Json");

    var child = std.process.Child.init(argv.items, allocator);
    child.stdin_behavior = .Inherit;
    child.stdout_behavior = .Pipe;
    child.stderr_behavior = .Pipe;

    child.spawn() catch |err| {
        std.debug.print("\x1b[31m[X]\x1b[0m  No se pudo iniciar PowerShell: {s}\n", .{@errorName(err)});
        return 1;
    };

    const header = import_progress.Header{
        .cuenta = session.email,
        .pst_path = pst_path,
        .action = action,
        .filter_year = filter_year,
        .filter_months = filter_months,
    };

    const code = try import_progress.run(allocator, &child, header);

    std.debug.print("\n", .{});
    if (code == 0) {
        std.debug.print("\x1b[1;32m[OK]\x1b[0m Importacion finalizada correctamente.\n", .{});
    } else {
        std.debug.print("\x1b[1;31m[X]\x1b[0m Importacion finalizo con errores (exit code {d}).\n", .{code});
    }
    return code;
}
