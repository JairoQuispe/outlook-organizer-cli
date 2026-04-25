const std = @import("std");
const build_options = @import("build_options");

pub const embed_enabled: bool = build_options.embed_scripts;

pub const ScriptKind = enum {
    list_stores,
    import_pst,
};

pub fn fileName(kind: ScriptKind) []const u8 {
    return switch (kind) {
        .list_stores => "outlook-list-stores.ps1",
        .import_pst => "outlook-import-pst.ps1",
    };
}

fn embeddedContent(comptime kind: ScriptKind) []const u8 {
    if (!embed_enabled) @compileError("embeddedContent llamado con embed_scripts deshabilitado");
    return switch (kind) {
        .list_stores => @embedFile("script_list_stores"),
        .import_pst => @embedFile("script_import_pst"),
    };
}

/// Cache de la carpeta extraida en %TEMP% (solo cuando embed_enabled).
var extracted_dir: ?[]u8 = null;
var extracted_mutex: std.Thread.Mutex = .{};

/// Obtiene la ruta absoluta del script solicitado. Si el binario tiene
/// los scripts incrustados los extrae una unica vez a %TEMP%\outlook-organizer-<pid>\
/// y devuelve esa ruta. En caso contrario busca en disco (scripts-powershell/
/// junto al exe o en cwd).
pub fn getScriptPath(allocator: std.mem.Allocator, kind: ScriptKind) ![]u8 {
    if (comptime embed_enabled) {
        const dir = try ensureExtracted(allocator);
        return try std.fs.path.join(allocator, &.{ dir, fileName(kind) });
    } else {
        return try resolveFromDisk(allocator, kind);
    }
}

/// Limpia la carpeta temporal si se extrajeron scripts.
pub fn cleanup(allocator: std.mem.Allocator) void {
    extracted_mutex.lock();
    defer extracted_mutex.unlock();

    if (extracted_dir) |dir| {
        std.fs.deleteTreeAbsolute(dir) catch {};
        allocator.free(dir);
        extracted_dir = null;
    }
}

fn ensureExtracted(allocator: std.mem.Allocator) ![]const u8 {
    extracted_mutex.lock();
    defer extracted_mutex.unlock();

    if (extracted_dir) |dir| return dir;

    // Construir ruta: %TEMP%\outlook-organizer-<pid>\
    const temp_base = std.process.getEnvVarOwned(allocator, "TEMP") catch
        try allocator.dupe(u8, "C:\\Windows\\Temp");
    defer allocator.free(temp_base);

    const pid = std.os.windows.GetCurrentProcessId();
    const dir_path = try std.fmt.allocPrint(allocator, "{s}\\outlook-organizer-{d}", .{ temp_base, pid });
    errdefer allocator.free(dir_path);

    std.fs.makeDirAbsolute(dir_path) catch |err| switch (err) {
        error.PathAlreadyExists => {},
        else => return err,
    };

    // Extraer cada script embebido (kind debe ser comptime para @embedFile)
    inline for (.{ ScriptKind.list_stores, ScriptKind.import_pst }) |kind| {
        try writeEmbeddedFile(allocator, dir_path, kind);
    }

    extracted_dir = dir_path;
    return dir_path;
}

fn writeEmbeddedFile(allocator: std.mem.Allocator, dir_path: []const u8, comptime kind: ScriptKind) !void {
    const full = try std.fs.path.join(allocator, &.{ dir_path, fileName(kind) });
    defer allocator.free(full);

    var f = try std.fs.createFileAbsolute(full, .{ .truncate = true });
    defer f.close();

    try f.writeAll(embeddedContent(kind));
}

fn resolveFromDisk(allocator: std.mem.Allocator, kind: ScriptKind) ![]u8 {
    const name = fileName(kind);

    // 1) cwd: scripts-powershell\<name>
    const cwd_path = try std.fs.path.join(allocator, &.{ "scripts-powershell", name });
    if (std.fs.cwd().access(cwd_path, .{})) |_| {
        return cwd_path;
    } else |_| {
        allocator.free(cwd_path);
    }

    // 2) Relativo al binario
    var exe_dir_buf: [std.fs.max_path_bytes]u8 = undefined;
    const exe_dir = try std.fs.selfExeDirPath(&exe_dir_buf);

    const relatives = [_][]const u8{
        "scripts-powershell",
        "..\\scripts-powershell",
        "..\\..\\scripts-powershell",
    };

    for (relatives) |rel| {
        const candidate = try std.fs.path.join(allocator, &.{ exe_dir, rel, name });
        if (std.fs.cwd().access(candidate, .{})) |_| {
            return candidate;
        } else |_| {
            allocator.free(candidate);
        }
    }

    return error.ScriptNotFound;
}
