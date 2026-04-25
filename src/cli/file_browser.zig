const std = @import("std");
const keyinput = @import("../windows/keyinput.zig");

extern "kernel32" fn GetLogicalDrives() callconv(.winapi) u32;

const Entry = struct {
    name: []u8,
    is_dir: bool,
    is_parent: bool, // ".."
    is_drive: bool, // "C:\"
    size: u64,
};

const VIEWPORT: usize = 15;

/// Navegador de archivos interactivo. Devuelve la ruta absoluta del archivo
/// seleccionado, o null si el usuario cancela con ESC.
/// `extension_filter` opcional (ej. ".pst") filtra archivos. Los directorios
/// siempre se muestran.
pub fn browseForFile(
    allocator: std.mem.Allocator,
    start_dir: ?[]const u8,
    extension_filter: ?[]const u8,
) !?[]u8 {
    // Directorio inicial
    var current_dir: []u8 = undefined;
    if (start_dir) |sd| {
        current_dir = try allocator.dupe(u8, sd);
    } else {
        var buf: [std.fs.max_path_bytes]u8 = undefined;
        const cwd = try std.process.getCwd(&buf);
        current_dir = try allocator.dupe(u8, cwd);
    }
    defer allocator.free(current_dir);

    var raw = keyinput.RawMode.enter();
    defer raw.leave();

    std.debug.print("\x1b[?25l", .{}); // ocultar cursor
    defer std.debug.print("\x1b[?25h", .{});

    var cursor: usize = 0;
    var scroll_top: usize = 0;
    var in_drives_view: bool = false;

    while (true) {
        var entries = std.ArrayList(Entry){};
        defer {
            for (entries.items) |e| allocator.free(e.name);
            entries.deinit(allocator);
        }

        if (in_drives_view) {
            try loadDrives(allocator, &entries);
        } else {
            loadDirectory(allocator, current_dir, extension_filter, &entries) catch |err| {
                std.debug.print("\x1b[2J\x1b[H", .{});
                std.debug.print("\x1b[31m[X]\x1b[0m  Error leyendo directorio: {s}\n", .{@errorName(err)});
                std.debug.print("Presiona ESC para salir o cualquier tecla para continuar...\n", .{});
                const k = keyinput.readKey();
                if (k == .escape) return null;
                // Ir a drives
                in_drives_view = true;
                cursor = 0;
                scroll_top = 0;
                continue;
            };
        }

        if (cursor >= entries.items.len and entries.items.len > 0) {
            cursor = entries.items.len - 1;
        }
        // ajustar scroll
        if (cursor < scroll_top) scroll_top = cursor;
        if (cursor >= scroll_top + VIEWPORT) scroll_top = cursor - VIEWPORT + 1;

        render(current_dir, in_drives_view, entries.items, cursor, scroll_top, extension_filter);

        const key = keyinput.readKey();
        switch (key) {
            .arrow_up => {
                if (entries.items.len == 0) continue;
                cursor = if (cursor == 0) entries.items.len - 1 else cursor - 1;
            },
            .arrow_down => {
                if (entries.items.len == 0) continue;
                cursor = (cursor + 1) % entries.items.len;
            },
            .arrow_left => {
                if (!in_drives_view) {
                    const res = try goUp(allocator, current_dir);
                    if (res) |new_path| {
                        allocator.free(current_dir);
                        current_dir = new_path;
                        cursor = 0;
                        scroll_top = 0;
                    } else {
                        in_drives_view = true;
                        cursor = 0;
                        scroll_top = 0;
                    }
                }
            },
            .enter => {
                if (entries.items.len == 0) continue;
                const selected = entries.items[cursor];

                if (selected.is_drive) {
                    allocator.free(current_dir);
                    current_dir = try allocator.dupe(u8, selected.name);
                    in_drives_view = false;
                    cursor = 0;
                    scroll_top = 0;
                    continue;
                }

                if (selected.is_parent) {
                    const res = try goUp(allocator, current_dir);
                    if (res) |new_path| {
                        allocator.free(current_dir);
                        current_dir = new_path;
                    } else {
                        in_drives_view = true;
                    }
                    cursor = 0;
                    scroll_top = 0;
                    continue;
                }

                if (selected.is_dir) {
                    const new_path = try std.fs.path.join(allocator, &.{ current_dir, selected.name });
                    allocator.free(current_dir);
                    current_dir = new_path;
                    cursor = 0;
                    scroll_top = 0;
                    continue;
                }

                // Archivo: devolver path absoluto
                const result = try std.fs.path.join(allocator, &.{ current_dir, selected.name });
                std.debug.print("\x1b[2J\x1b[H", .{}); // limpia pantalla antes de volver
                return result;
            },
            .escape => {
                std.debug.print("\x1b[2J\x1b[H", .{});
                return null;
            },
            else => {},
        }
    }
}

fn goUp(allocator: std.mem.Allocator, current_dir: []const u8) !?[]u8 {
    const parent = std.fs.path.dirname(current_dir) orelse return null;
    // Si ya estamos en "C:\" dirname devuelve "C:" -> tratarlo como drive root
    if (std.mem.eql(u8, parent, current_dir)) return null;
    if (parent.len == 2 and parent[1] == ':') {
        // "C:" -> "C:\"
        var buf = try allocator.alloc(u8, 3);
        buf[0] = parent[0];
        buf[1] = ':';
        buf[2] = '\\';
        return buf;
    }
    return try allocator.dupe(u8, parent);
}

fn loadDrives(allocator: std.mem.Allocator, entries: *std.ArrayList(Entry)) !void {
    const mask = GetLogicalDrives();
    var i: u5 = 0;
    while (i < 26) : (i += 1) {
        const bit = @as(u32, 1) << i;
        if ((mask & bit) != 0) {
            const letter: u8 = 'A' + @as(u8, i);
            var name = try allocator.alloc(u8, 3);
            name[0] = letter;
            name[1] = ':';
            name[2] = '\\';
            try entries.append(allocator, .{
                .name = name,
                .is_dir = true,
                .is_parent = false,
                .is_drive = true,
                .size = 0,
            });
        }
    }
}

fn loadDirectory(
    allocator: std.mem.Allocator,
    path: []const u8,
    extension_filter: ?[]const u8,
    entries: *std.ArrayList(Entry),
) !void {
    var dir = try std.fs.openDirAbsolute(path, .{ .iterate = true });
    defer dir.close();

    // Agregar ".." siempre
    const parent_name = try allocator.dupe(u8, "..");
    try entries.append(allocator, .{
        .name = parent_name,
        .is_dir = true,
        .is_parent = true,
        .is_drive = false,
        .size = 0,
    });

    var it = dir.iterate();
    while (try it.next()) |entry| {
        const is_dir = entry.kind == .directory;

        if (!is_dir) {
            if (extension_filter) |ext| {
                if (!std.ascii.endsWithIgnoreCase(entry.name, ext)) continue;
            }
        }

        // omitir archivos/carpetas ocultos del sistema (empiezan con '.')
        if (entry.name.len > 0 and entry.name[0] == '.' and !is_dir) continue;

        var size: u64 = 0;
        if (!is_dir) {
            const stat_result = dir.statFile(entry.name) catch null;
            if (stat_result) |st| size = st.size;
        }

        const name_copy = try allocator.dupe(u8, entry.name);
        try entries.append(allocator, .{
            .name = name_copy,
            .is_dir = is_dir,
            .is_parent = false,
            .is_drive = false,
            .size = size,
        });
    }

    // Ordenar: carpetas primero (excepto ..), luego archivos; alfabetico
    std.mem.sort(Entry, entries.items[1..], {}, entryLessThan);
}

fn entryLessThan(_: void, a: Entry, b: Entry) bool {
    if (a.is_dir and !b.is_dir) return true;
    if (!a.is_dir and b.is_dir) return false;
    return std.ascii.lessThanIgnoreCase(a.name, b.name);
}

fn render(
    current_dir: []const u8,
    in_drives_view: bool,
    entries: []const Entry,
    cursor: usize,
    scroll_top: usize,
    extension_filter: ?[]const u8,
) void {
    std.debug.print("\x1b[2J\x1b[H", .{}); // limpia pantalla + home
    std.debug.print("\x1b[1;36m== Explorador de archivos ==\x1b[0m\n", .{});
    std.debug.print("\x1b[90m(Flechas Arriba/Abajo: navegar | Enter: abrir/seleccionar | Flecha Izquierda: subir | ESC: cancelar)\x1b[0m\n", .{});

    if (extension_filter) |ext| {
        std.debug.print("\x1b[90mFiltro: mostrando solo archivos {s}\x1b[0m\n", .{ext});
    }

    std.debug.print("\n\x1b[1mUbicacion:\x1b[0m ", .{});
    if (in_drives_view) {
        std.debug.print("\x1b[33m[Unidades del sistema]\x1b[0m\n", .{});
    } else {
        std.debug.print("{s}\n", .{current_dir});
    }
    std.debug.print("\n", .{});

    if (entries.len == 0) {
        std.debug.print("\x1b[90m   (directorio vacio o sin archivos que coincidan con el filtro)\x1b[0m\n", .{});
        return;
    }

    const end = @min(scroll_top + VIEWPORT, entries.len);
    if (scroll_top > 0) {
        std.debug.print("\x1b[90m   ...\x1b[0m\n", .{});
    }

    var i: usize = scroll_top;
    while (i < end) : (i += 1) {
        renderRow(entries[i], i == cursor);
    }

    if (end < entries.len) {
        std.debug.print("\x1b[90m   ... ({d} mas)\x1b[0m\n", .{entries.len - end});
    }

    std.debug.print("\n\x1b[90m   {d}/{d}\x1b[0m\n", .{ cursor + 1, entries.len });
}

fn renderRow(e: Entry, selected: bool) void {
    const prefix: []const u8 = if (e.is_drive)
        "[UND]"
    else if (e.is_parent)
        " .. "
    else if (e.is_dir)
        "[DIR]"
    else
        "     ";

    if (selected) {
        if (e.is_dir) {
            std.debug.print("\x1b[1;42;30m > {s} {s}\\\x1b[0m\n", .{ prefix, e.name });
        } else {
            var size_buf: [32]u8 = undefined;
            const size_str = formatSize(e.size, &size_buf);
            std.debug.print("\x1b[1;42;30m > {s} {s}  {s}\x1b[0m\n", .{ prefix, e.name, size_str });
        }
    } else {
        if (e.is_dir) {
            std.debug.print("   \x1b[1;34m{s}\x1b[0m \x1b[34m{s}\\\x1b[0m\n", .{ prefix, e.name });
        } else {
            var size_buf: [32]u8 = undefined;
            const size_str = formatSize(e.size, &size_buf);
            std.debug.print("   \x1b[90m{s}\x1b[0m {s}  \x1b[90m{s}\x1b[0m\n", .{ prefix, e.name, size_str });
        }
    }
}

fn formatSize(bytes: u64, buf: []u8) []const u8 {
    const GB: u64 = 1024 * 1024 * 1024;
    const MB: u64 = 1024 * 1024;
    const KB: u64 = 1024;
    if (bytes >= GB) {
        return std.fmt.bufPrint(buf, "{d:.2} GB", .{@as(f64, @floatFromInt(bytes)) / @as(f64, GB)}) catch "";
    } else if (bytes >= MB) {
        return std.fmt.bufPrint(buf, "{d:.2} MB", .{@as(f64, @floatFromInt(bytes)) / @as(f64, MB)}) catch "";
    } else if (bytes >= KB) {
        return std.fmt.bufPrint(buf, "{d:.2} KB", .{@as(f64, @floatFromInt(bytes)) / @as(f64, KB)}) catch "";
    } else {
        return std.fmt.bufPrint(buf, "{d} B", .{bytes}) catch "";
    }
}
