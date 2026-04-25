const std = @import("std");
const args_mod = @import("cli/args.zig");
const messages = @import("cli/messages.zig");
const preflight = @import("preflight/preflight.zig");
const console = @import("windows/console.zig");
const list_stores = @import("stores/list_stores.zig");
const store_selector = @import("cli/store_selector.zig");
const Session = @import("session.zig").Session;
const dispatcher = @import("actions/dispatcher.zig");
const scripts = @import("scripts/scripts.zig");

pub fn main() !void {
    var gpa = std.heap.GeneralPurposeAllocator(.{}){};
    defer _ = gpa.deinit();
    const allocator = gpa.allocator();

    console.enableModernConsole();
    defer scripts.cleanup(allocator);

    const raw_args = try std.process.argsAlloc(allocator);
    defer std.process.argsFree(allocator, raw_args);

    const parsed = args_mod.parseArgs(raw_args);

    const exit_code: preflight.ExitCode = switch (parsed.command) {
        .preflight => runPreflightAndSelect(allocator, parsed.json),
        .help => blk: {
            messages.printHelp();
            break :blk .success;
        },
    };

    std.process.exit(@intFromEnum(exit_code));
}

fn runPreflightAndSelect(allocator: std.mem.Allocator, as_json: bool) preflight.ExitCode {
    const pre = preflight.run(allocator, as_json);
    if (pre != .success) return pre;

    const script_path = scripts.getScriptPath(allocator, .list_stores) catch |err| {
        std.debug.print("\x1b[31m[X]\x1b[0m  No se encontro outlook-list-stores.ps1: {s}\n", .{@errorName(err)});
        return .failure;
    };
    defer allocator.free(script_path);

    var stores_result = list_stores.fetchStores(allocator, script_path) catch |err| {
        std.debug.print("\x1b[31m[X]\x1b[0m  Error al listar buzones: {s}\n", .{@errorName(err)});
        return .failure;
    };
    defer stores_result.deinit();

    // Filtrar: solo buzones con StoreId valido (requerido por las acciones siguientes).
    var filtered_buf = std.ArrayList(list_stores.Store){};
    defer filtered_buf.deinit(allocator);
    var skipped: usize = 0;
    for (stores_result.stores) |s| {
        if (s.store_id) |id| {
            if (id.len > 0) {
                filtered_buf.append(allocator, s) catch return .failure;
                continue;
            }
        }
        skipped += 1;
    }
    const filtered = filtered_buf.items;

    if (skipped > 0 and !as_json) {
        std.debug.print("\n\x1b[90m[i] Se omitieron {d} buzon(es) sin StoreId valido.\x1b[0m\n", .{skipped});
    }

    if (filtered.len == 0) {
        std.debug.print("\n\x1b[31m[X]\x1b[0m  No hay buzones con StoreId valido para continuar.\n", .{});
        return .failure;
    }

    if (as_json) {
        // En modo JSON solo emitimos los stores filtrados y salimos.
        std.debug.print("{{\"type\":\"stores\",\"count\":{d},\"skipped\":{d}}}\n", .{ filtered.len, skipped });
        for (filtered, 0..) |s, i| {
            std.debug.print(
                "{{\"type\":\"store\",\"index\":{d},\"displayName\":\"{s}\",\"storeId\":\"{s}\",\"filePath\":\"{s}\"}}\n",
                .{
                    i,
                    s.display_name,
                    s.store_id orelse "",
                    s.file_path orelse "",
                },
            );
        }
        return .success;
    }

    const selected_idx = store_selector.selectStore(filtered) orelse {
        std.debug.print("\n\x1b[33m[!] Seleccion cancelada por el usuario.\x1b[0m\n", .{});
        return .failure;
    };

    const selected = filtered[selected_idx];
    const session = Session{
        .store_id = selected.store_id orelse "",
        .email = selected.display_name,
        .display_name = selected.display_name,
        .file_path = selected.file_path,
        .exchange_store_type = selected.exchange_store_type,
    };

    std.debug.print("\n\x1b[1;32m[OK]\x1b[0m Buzon seleccionado: \x1b[1m{s}\x1b[0m\n", .{session.display_name});
    if (session.store_id.len > 0) {
        std.debug.print("     StoreId: \x1b[90m{s}\x1b[0m\n", .{session.store_id});
    }
    if (session.file_path) |fp| {
        std.debug.print("     Ruta:    \x1b[90m{s}\x1b[0m\n", .{fp});
    }

    if (session.store_id.len == 0) {
        std.debug.print("\n\x1b[31m[X]\x1b[0m  El buzon seleccionado no expone un StoreId. No se puede continuar.\n", .{});
        return .failure;
    }

    const action = dispatcher.promptAction() orelse {
        std.debug.print("\n\x1b[33m[!] Operacion cancelada por el usuario.\x1b[0m\n", .{});
        return .failure;
    };

    const action_result = dispatcher.runAction(action, session, allocator);
    return switch (action_result) {
        .success => .success,
        else => .failure,
    };
}
