const std = @import("std");

pub fn build(b: *std.Build) void {
    const target = b.standardTargetOptions(.{});
    const optimize = b.standardOptimizeOption(.{});

    // Si es true, los .ps1 se incrustan en el binario y se extraen a %TEMP%
    // en tiempo de ejecucion. Si es false (default), se resuelven desde disco
    // (carpeta scripts-powershell/ junto al exe o en cwd).
    const embed_scripts = b.option(
        bool,
        "embed-scripts",
        "Embed PowerShell scripts into the binary (recommended for release)",
    ) orelse false;

    const build_options = b.addOptions();
    build_options.addOption(bool, "embed_scripts", embed_scripts);

    const zigtui_dep = b.dependency("zigtui", .{
        .target = target,
        .optimize = optimize,
    });

    const root_mod = b.createModule(.{
        .root_source_file = b.path("src/main.zig"),
        .target = target,
        .optimize = optimize,
    });
    root_mod.addOptions("build_options", build_options);
    root_mod.addImport("zigtui", zigtui_dep.module("zigtui"));

    if (embed_scripts) {
        root_mod.addAnonymousImport("script_list_stores", .{
            .root_source_file = b.path("scripts-powershell/outlook-list-stores.ps1"),
        });
        root_mod.addAnonymousImport("script_import_pst", .{
            .root_source_file = b.path("scripts-powershell/outlook-import-pst.ps1"),
        });
    }

    const exe = b.addExecutable(.{
        .name = "outlook-organizer",
        .root_module = root_mod,
    });

    b.installArtifact(exe);

    const run_cmd = b.addRunArtifact(exe);
    if (b.args) |args| {
        run_cmd.addArgs(args);
    }

    const run_step = b.step("run", "Run Outlook Organizer CLI");
    run_step.dependOn(&run_cmd.step);
}
