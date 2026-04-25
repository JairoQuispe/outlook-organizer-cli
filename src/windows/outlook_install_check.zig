const std = @import("std");

pub const InstallCheckResult = struct {
    clsid_ok: bool,
    app_paths_ok: bool,
    click_to_run_ok: bool,
    known_path_ok: bool,

    pub fn isInstalled(self: InstallCheckResult) bool {
        return self.clsid_ok or self.app_paths_ok or self.click_to_run_ok;
    }
};

pub fn runInstallChecks(allocator: std.mem.Allocator) InstallCheckResult {
    const clsid_ok = checkClsidAndLocalServer(allocator);
    const app_paths_ok = checkAppPaths(allocator);
    const click_to_run_ok = checkClickToRun(allocator);
    const known_path_ok = checkKnownPaths(allocator);

    return .{
        .clsid_ok = clsid_ok,
        .app_paths_ok = app_paths_ok,
        .click_to_run_ok = click_to_run_ok,
        .known_path_ok = known_path_ok,
    };
}

fn checkClsidAndLocalServer(allocator: std.mem.Allocator) bool {
    const clsid_term = runRegQuery(allocator, &.{
        "reg",
        "query",
        "HKCR\\Outlook.Application\\CLSID",
        "/ve",
    });
    if (!clsid_term) return false;

    return runRegQuery(allocator, &.{
        "reg",
        "query",
        "HKCR\\CLSID",
        "/s",
        "/f",
        "OUTLOOK.EXE",
    });
}

fn checkAppPaths(allocator: std.mem.Allocator) bool {
    return runRegQuery(allocator, &.{
        "reg",
        "query",
        "HKLM\\SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\App Paths\\OUTLOOK.EXE",
        "/ve",
    });
}

fn checkClickToRun(allocator: std.mem.Allocator) bool {
    return runRegQuery(allocator, &.{
        "reg",
        "query",
        "HKLM\\SOFTWARE\\Microsoft\\Office\\ClickToRun\\Configuration",
        "/v",
        "ClientVersionToReport",
    });
}

fn checkKnownPaths(allocator: std.mem.Allocator) bool {
    const maybe_program_files = std.process.getEnvVarOwned(allocator, "ProgramFiles") catch null;
    defer if (maybe_program_files) |value| allocator.free(value);

    const maybe_program_files_x86 = std.process.getEnvVarOwned(allocator, "ProgramFiles(x86)") catch null;
    defer if (maybe_program_files_x86) |value| allocator.free(value);

    if (maybe_program_files) |root| {
        if (existsUnderRoot(allocator, root)) return true;
    }

    if (maybe_program_files_x86) |root| {
        if (existsUnderRoot(allocator, root)) return true;
    }

    return false;
}

fn existsUnderRoot(allocator: std.mem.Allocator, root: []const u8) bool {
    const candidates = [_][]const u8{
        "\\Microsoft Office\\root\\Office16\\OUTLOOK.EXE",
        "\\Microsoft Office\\Office16\\OUTLOOK.EXE",
        "\\Microsoft Office\\Office15\\OUTLOOK.EXE",
    };

    for (candidates) |suffix| {
        const full_path = std.mem.concat(allocator, u8, &.{ root, suffix }) catch continue;
        defer allocator.free(full_path);

        std.fs.accessAbsolute(full_path, .{}) catch continue;
        return true;
    }

    return false;
}

fn runRegQuery(allocator: std.mem.Allocator, argv: []const []const u8) bool {
    const result = std.process.Child.run(.{
        .allocator = allocator,
        .argv = argv,
    }) catch return false;

    switch (result.term) {
        .Exited => |code| return code == 0,
        else => return false,
    }
}
