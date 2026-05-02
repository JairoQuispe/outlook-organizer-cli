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

    const all_stores = stores_result.stores;
    var valid_buf = std.ArrayList(list_stores.Store){};
    defer valid_buf.deinit(allocator);
    var skipped: usize = 0;
    for (all_stores) |s| {
        if (s.store_id) |id| {
            if (id.len > 0) {
                valid_buf.append(allocator, s) catch return .failure;
                continue;
            }
        }
        skipped += 1;
    }
    const valid = valid_buf.items;

    if (skipped > 0 and !as_json) {
        std.debug.print("\n\x1b[90m[i] Se detectaron {d} buzon(es) sin StoreId; se muestran pero no se pueden seleccionar.\x1b[0m\n", .{skipped});
    }

    if (valid.len == 0) {
        std.debug.print("\n\x1b[31m[X]\x1b[0m  No hay buzones con StoreId valido para continuar.\n", .{});
        return .failure;
    }

    if (as_json) {
        // En modo JSON solo emitimos los stores filtrados y salimos.
        std.debug.print("{{\"type\":\"stores\",\"count\":{d},\"skipped\":{d}}}\n", .{ valid.len, skipped });
        for (valid, 0..) |s, i| {
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

    var selected: list_stores.Store = undefined;
    while (true) {
        const selected_idx = store_selector.selectStore(allocator, all_stores) orelse {
            std.debug.print("\n\x1b[33m[!] Seleccion cancelada por el usuario.\x1b[0m\n", .{});
            return .failure;
        };

        selected = all_stores[selected_idx];
        if (selected.store_id) |id| {
            if (id.len > 0) break;
        }
        std.debug.print("\n\x1b[31m[X]\x1b[0m  El buzon seleccionado no expone un StoreId. Selecciona otro.\n", .{});
    }
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
