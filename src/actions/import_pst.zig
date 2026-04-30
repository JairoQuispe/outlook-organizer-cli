const std = @import("std");
const Session = @import("../session.zig").Session;
const input = @import("../cli/input.zig");
const menu = @import("../cli/menu.zig");
const file_browser = @import("../cli/file_browser.zig");
const scripts = @import("../scripts/scripts.zig");
const import_progress = @import("import_progress.zig");

const PstFolder = struct {
    path: []u8,
    item_count: usize,
    year_summary: ?[]u8 = null,
};

/// Ejecuta el wizard de importacion de PST. Devuelve exit code del script
/// (0 = exito) o un codigo de error del wizard.
pub fn run(session: Session, allocator: std.mem.Allocator) !u8 {
    std.debug.print("\n\x1b[1;36m== Importar correos desde PST ==\x1b[0m\n", .{});
    session.print();
    std.debug.print("\n", .{});

    // 1) Ruta del PST
    const pst_path = try askPstPath(allocator);
    defer allocator.free(pst_path);

    // 2) Seleccion de carpetas del PST (multi-seleccion)
    const selected_folders = try askPstFolders(allocator, pst_path);
    defer {
        for (selected_folders) |p| allocator.free(p);
        allocator.free(selected_folders);
    }

    // 3) Accion Copy / Move
    const action = askAction() orelse {
        std.debug.print("\n\x1b[33m[!] Importacion cancelada.\x1b[0m\n", .{});
        return 2;
    };

    // 4) Filtro por anio
    const filter_year = try askFilterYear(allocator);
    defer if (filter_year) |fy| allocator.free(fy);

    // 5) Filtro por meses (solo si hay anio)
    var filter_months: ?[]u8 = null;
    defer if (filter_months) |fm| allocator.free(fm);
    if (filter_year != null) {
        filter_months = try askFilterMonths(allocator);
    }

    // 6) Saltar duplicados
    const skip_dupes = input.confirm("\xc2\xbfSaltar duplicados (por Message-Id)?", false);

    // 7) Resumen y confirmacion
    printSummary(session, pst_path, selected_folders, action, filter_year, filter_months, skip_dupes);
    if (!input.confirm("\xc2\xbfEjecutar la importacion con estos parametros?", true)) {
        std.debug.print("\n\x1b[33m[!] Importacion cancelada.\x1b[0m\n", .{});
        return 2;
    }

    // 8) Ejecutar PowerShell heredando stdio para ver progreso
    return try executePowerShell(allocator, session, pst_path, selected_folders, action, filter_year, filter_months, skip_dupes);
}

fn askPstFolders(allocator: std.mem.Allocator, pst_path: []const u8) ![][]u8 {
    const folders = try listPstFolders(allocator, pst_path);
    defer {
        for (folders) |f| {
            allocator.free(f.path);
            if (f.year_summary) |ys| allocator.free(ys);
        }
        allocator.free(folders);
    }

    if (folders.len == 0) {
        std.debug.print("\x1b[31m[X]\x1b[0m  El PST no contiene carpetas importables.\n", .{});
        return error.NoFoldersFound;
    }

    var items = try allocator.alloc(menu.MenuItem, folders.len);
    defer allocator.free(items);

    var descriptions = try allocator.alloc(?[]u8, folders.len);
    defer {
        for (descriptions) |d_opt| {
            if (d_opt) |d| allocator.free(d);
        }
        allocator.free(descriptions);
    }

    for (folders, 0..) |f, i| {
        const desc = if (f.year_summary) |ys|
            try std.fmt.allocPrint(allocator, "{d} correo(s) | {s}", .{ f.item_count, ys })
        else
            try std.fmt.allocPrint(allocator, "{d} correo(s)", .{f.item_count});
        descriptions[i] = desc;
        items[i] = .{ .label = f.path, .description = desc };
    }

    const idxs = try menu.selectMultiple(allocator, "Selecciona carpetas del PST a importar", items) orelse {
        return error.UserCancelled;
    };
    defer allocator.free(idxs);

    var out = try allocator.alloc([]u8, idxs.len);
    errdefer {
        for (out[0..]) |p| allocator.free(p);
        allocator.free(out);
    }
    for (idxs, 0..) |idx, j| {
        out[j] = try allocator.dupe(u8, folders[idx].path);
    }
    return out;
}

fn listPstFolders(allocator: std.mem.Allocator, pst_path: []const u8) ![]PstFolder {
    const script_path = try scripts.getScriptPath(allocator, .import_pst);
    defer allocator.free(script_path);

    const result = try std.process.Child.run(.{
        .allocator = allocator,
        .argv = &.{
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-File",
            script_path,
            "-PstPath",
            pst_path,
            "-ListFolders",
            "-Json",
        },
        .max_output_bytes = 8 * 1024 * 1024,
    });
    defer allocator.free(result.stdout);
    defer allocator.free(result.stderr);

    switch (result.term) {
        .Exited => |code| if (code != 0) {
            if (result.stderr.len > 0) {
                std.debug.print("\n\x1b[31mError listando carpetas PST:\x1b[0m\n{s}\n", .{result.stderr});
            }
            return error.ListFoldersFailed;
        },
        else => return error.ListFoldersFailed,
    }

    var list = std.ArrayList(PstFolder){};
    defer list.deinit(allocator);

    var it = std.mem.splitScalar(u8, result.stdout, '\n');
    while (it.next()) |line| {
        const trimmed = std.mem.trim(u8, line, " \r\n\t");
        if (trimmed.len == 0 or trimmed[0] != '{') continue;

        var parsed = std.json.parseFromSlice(std.json.Value, allocator, trimmed, .{}) catch continue;
        defer parsed.deinit();
        if (parsed.value != .object) continue;
        const obj = parsed.value.object;

        const type_val = obj.get("type") orelse continue;
        if (type_val != .string or !std.mem.eql(u8, type_val.string, "folder")) continue;

        const path_val = obj.get("path") orelse continue;
        if (path_val != .string) continue;
        const count_val = obj.get("itemCount");
        const count: usize = if (count_val) |v| switch (v) {
            .integer => |i| if (i > 0) @intCast(i) else 0,
            .float => |f| if (f > 0) @intFromFloat(f) else 0,
            else => 0,
        } else 0;

        const undated_count = parseUndatedCount(obj.get("undatedCount"));
        const year_summary = blk: {
            const base = try parseYearSummary(allocator, obj.get("yearBreakdown"));
            if (undated_count == 0) break :blk base;

            if (base) |b| {
                const with_undated = try std.fmt.allocPrint(allocator, "{s} · sin_fecha({d})", .{ b, undated_count });
                allocator.free(b);
                break :blk with_undated;
            }

            break :blk try std.fmt.allocPrint(allocator, "Años: sin_fecha({d})", .{undated_count});
        };

        try list.append(allocator, .{
            .path = try allocator.dupe(u8, path_val.string),
            .item_count = count,
            .year_summary = year_summary,
        });
    }

    return try list.toOwnedSlice(allocator);
}

fn parseYearSummary(allocator: std.mem.Allocator, year_breakdown_val: ?std.json.Value) !?[]u8 {
    const yb = year_breakdown_val orelse return null;
    if (yb != .array) return null;
    if (yb.array.items.len == 0) return null;

    var parts = std.ArrayList([]const u8){};
    defer {
        for (parts.items) |p| allocator.free(p);
        parts.deinit(allocator);
    }

    for (yb.array.items) |entry| {
        if (entry != .object) continue;
        const eobj = entry.object;
        const yv = eobj.get("year") orelse continue;
        const cv = eobj.get("count") orelse continue;

        const year: i64 = switch (yv) {
            .integer => |v| v,
            .float => |v| @intFromFloat(v),
            else => continue,
        };

        const cnt_raw: i64 = switch (cv) {
            .integer => |v| v,
            .float => |v| @intFromFloat(v),
            else => continue,
        };

        const cnt: usize = if (cnt_raw > 0) @intCast(cnt_raw) else 0;

        const part = try std.fmt.allocPrint(allocator, "{d}({d})", .{ year, cnt });
        try parts.append(allocator, part);
    }

    if (parts.items.len == 0) return null;

    var buf = std.ArrayList(u8){};
    defer buf.deinit(allocator);

    try buf.appendSlice(allocator, "Años: ");

    for (parts.items, 0..) |part, i| {
        if (i > 0) try buf.appendSlice(allocator, " · ");
        try buf.appendSlice(allocator, part);
    }

    return try buf.toOwnedSlice(allocator);
}

fn parseUndatedCount(undated_count_val: ?std.json.Value) usize {
    const uv = undated_count_val orelse return 0;
    return switch (uv) {
        .integer => |v| if (v > 0) @intCast(v) else 0,
        .float => |v| if (v > 0) @intFromFloat(v) else 0,
        else => 0,
    };
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
    selected_folders: []const []const u8,
    action: []const u8,
    filter_year: ?[]const u8,
    filter_months: ?[]const u8,
    skip_dupes: bool,
) void {
    std.debug.print("\n\x1b[1;36m== Resumen de importacion ==\x1b[0m\n", .{});
    std.debug.print("  \x1b[90mCuenta destino:\x1b[0m    \x1b[1m{s}\x1b[0m\n", .{session.email});
    std.debug.print("  \x1b[90mStoreId:\x1b[0m           {s}\n", .{session.store_id});
    std.debug.print("  \x1b[90mArchivo PST:\x1b[0m       {s}\n", .{pst_path});
    std.debug.print("  \x1b[90mCarpetas PST:\x1b[0m      {d} seleccionada(s)\n", .{selected_folders.len});
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
    selected_folders: []const []const u8,
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

    for (selected_folders) |folder_path| {
        try argv.append(allocator, "-IncludeFolders");
        try argv.append(allocator, folder_path);
    }

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
