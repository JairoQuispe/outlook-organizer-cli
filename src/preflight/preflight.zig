const std = @import("std");
const messages = @import("../cli/messages.zig");
const Spinner = @import("../cli/spinner.zig").Spinner;
const install_check = @import("../windows/outlook_install_check.zig");
const window_check = @import("../windows/outlook_window_check.zig");

pub const ExitCode = enum(u8) {
    success = 0,
    failure = 1,
};

pub fn run(allocator: std.mem.Allocator, as_json: bool) ExitCode {
    const os_tag = @import("builtin").os.tag;
    if (os_tag != .windows) {
        if (as_json) {
            std.debug.print("{{\"type\":\"preflight\",\"ok\":false,\"message\":\"Esta CLI solo esta soportada en Windows.\"}}\n", .{});
        } else {
            std.debug.print("\x1b[31m[X]\x1b[0m  Esta CLI solo esta soportada en Windows.\n", .{});
        }
        return .failure;
    }

    if (!as_json) {
        // Encabezado limpio en ASCII (cmd-friendly)
        std.debug.print("\n\x1b[1;36m== Outlook Organizer - Preflight ==\x1b[0m\n", .{});
        std.debug.print("\x1b[90m------------------------------------\x1b[0m\n", .{});
    }

    var spinner = Spinner.init(!as_json);

    // Etapa 1: instalacion
    spinner.start("Verificando instalacion de Outlook clasico...") catch {};
    const install = install_check.runInstallChecks(allocator);

    if (!install.isInstalled()) {
        spinner.stopWithFailure("Outlook clasico no esta instalado en este equipo.");
        if (as_json) {
            emitInstallJson(install);
            std.debug.print("{{\"type\":\"preflight_error\",\"stage\":\"install\",\"message\":\"Outlook clasico no esta instalado en este equipo.\"}}\n", .{});
        }
        return .failure;
    }
    spinner.stopWithSuccess("Outlook clasico instalado.");

    // Etapa 2: proceso en ejecucion + ventana
    spinner.start("Verificando que Outlook este abierto...") catch {};
    const runtime = window_check.runRuntimeChecks(allocator);

    if (!runtime.process_running) {
        spinner.stopWithFailure("Outlook no esta abierto. Por favor abre Outlook clasico antes de continuar.");
        if (as_json) {
            emitInstallJson(install);
            emitRuntimeJson(runtime);
            std.debug.print("{{\"type\":\"preflight_error\",\"stage\":\"process\",\"message\":\"Outlook no esta abierto.\"}}\n", .{});
        }
        return .failure;
    }
    spinner.stopWithSuccess("Outlook esta en ejecucion.");

    spinner.start("Verificando ventana principal de Outlook...") catch {};
    if (!runtime.main_window_found) {
        spinner.stopWithFailure("Outlook esta en ejecucion pero sin ventana principal.");
        if (as_json) {
            emitInstallJson(install);
            emitRuntimeJson(runtime);
            std.debug.print("{{\"type\":\"preflight_error\",\"stage\":\"window\",\"message\":\"Outlook sin ventana principal.\"}}\n", .{});
        }
        return .failure;
    }

    if (!runtime.window_visible or !runtime.window_not_minimized) {
        spinner.stopWithFailure("Outlook esta minimizado u oculto. Restaura su ventana antes de continuar.");
        if (as_json) {
            emitInstallJson(install);
            emitRuntimeJson(runtime);
            std.debug.print("{{\"type\":\"preflight_error\",\"stage\":\"window_state\",\"message\":\"Outlook minimizado u oculto.\"}}\n", .{});
        }
        return .failure;
    }
    spinner.stopWithSuccess("Ventana principal visible y restaurada.");

    if (as_json) {
        emitInstallJson(install);
        emitRuntimeJson(runtime);
        std.debug.print("{{\"type\":\"preflight_ok\",\"message\":\"Preflight completado correctamente.\"}}\n", .{});
    } else {
        std.debug.print("\n\x1b[1;32m=> Preflight completado correctamente.\x1b[0m\n\n", .{});
        messages.printPreflightSuccess();
    }

    return .success;
}

fn emitInstallJson(install: install_check.InstallCheckResult) void {
    std.debug.print(
        "{{\"type\":\"check\",\"stage\":\"install\",\"clsid\":{},\"app_paths\":{},\"click_to_run\":{},\"known_paths\":{},\"ok\":{}}}\n",
        .{
            install.clsid_ok,
            install.app_paths_ok,
            install.click_to_run_ok,
            install.known_path_ok,
            install.isInstalled(),
        },
    );
}

fn emitRuntimeJson(runtime: window_check.RuntimeCheckResult) void {
    std.debug.print(
        "{{\"type\":\"check\",\"stage\":\"runtime\",\"process_running\":{},\"main_window_found\":{},\"window_visible\":{},\"window_not_minimized\":{},\"ok\":{}}}\n",
        .{
            runtime.process_running,
            runtime.main_window_found,
            runtime.window_visible,
            runtime.window_not_minimized,
            runtime.isReady(),
        },
    );
}
