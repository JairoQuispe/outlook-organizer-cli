const std = @import("std");

pub const RuntimeCheckResult = struct {
    process_running: bool,
    main_window_found: bool,
    window_visible: bool,
    window_not_minimized: bool,

    pub fn isReady(self: RuntimeCheckResult) bool {
        return self.process_running and self.main_window_found and self.window_visible and self.window_not_minimized;
    }
};

pub fn runRuntimeChecks(allocator: std.mem.Allocator) RuntimeCheckResult {
    const process_running = checkOutlookProcessRunning(allocator);
    if (!process_running) {
        return .{
            .process_running = false,
            .main_window_found = false,
            .window_visible = false,
            .window_not_minimized = false,
        };
    }

    const maybe_hwnd = getOutlookMainWindowHandle(allocator);
    if (maybe_hwnd == null) {
        return .{
            .process_running = true,
            .main_window_found = false,
            .window_visible = false,
            .window_not_minimized = false,
        };
    }

    const hwnd = maybe_hwnd.?;
    const visible = isWindowVisible(hwnd, allocator);
    const not_minimized = isWindowNotMinimized(hwnd, allocator);

    return .{
        .process_running = true,
        .main_window_found = true,
        .window_visible = visible,
        .window_not_minimized = not_minimized,
    };
}

fn checkOutlookProcessRunning(allocator: std.mem.Allocator) bool {
    return runPowerShellBoolean(allocator, "$ErrorActionPreference='SilentlyContinue'; if (Get-Process -Name OUTLOOK) { '1' } else { '0' }");
}

fn getOutlookMainWindowHandle(allocator: std.mem.Allocator) ?usize {
    const result = std.process.Child.run(.{
        .allocator = allocator,
        .argv = &.{
            "powershell",
            "-NoProfile",
            "-Command",
            "$ErrorActionPreference='SilentlyContinue'; $p = Get-Process -Name OUTLOOK | Where-Object { $_.MainWindowHandle -ne 0 } | Select-Object -First 1; if ($p) { Write-Output $p.MainWindowHandle }",
        },
    }) catch return null;
    defer allocator.free(result.stdout);
    defer allocator.free(result.stderr);

    switch (result.term) {
        .Exited => |code| if (code != 0) return null,
        else => return null,
    }

    const trimmed = std.mem.trim(u8, result.stdout, " \r\n\t");
    if (trimmed.len == 0) return null;

    return std.fmt.parseInt(usize, trimmed, 10) catch null;
}

fn isWindowVisible(hwnd: usize, allocator: std.mem.Allocator) bool {
    var script_buf: [256]u8 = undefined;
    const script = std.fmt.bufPrint(
        &script_buf,
        "$ErrorActionPreference='SilentlyContinue'; Add-Type -Name W -Namespace N -MemberDefinition '[DllImport(\"user32.dll\")] public static extern bool IsWindowVisible(System.IntPtr hWnd);'; if ([N.W]::IsWindowVisible([IntPtr]{d})) {{ '1' }} else {{ '0' }}",
        .{hwnd},
    ) catch return false;

    return runPowerShellBoolean(allocator, script);
}

fn isWindowNotMinimized(hwnd: usize, allocator: std.mem.Allocator) bool {
    var script_buf: [256]u8 = undefined;
    const script = std.fmt.bufPrint(
        &script_buf,
        "$ErrorActionPreference='SilentlyContinue'; Add-Type -Name W2 -Namespace N2 -MemberDefinition '[DllImport(\"user32.dll\")] public static extern bool IsIconic(System.IntPtr hWnd);'; if ([N2.W2]::IsIconic([IntPtr]{d})) {{ '0' }} else {{ '1' }}",
        .{hwnd},
    ) catch return false;

    return runPowerShellBoolean(allocator, script);
}

fn runPowerShellBoolean(allocator: std.mem.Allocator, script: []const u8) bool {
    const result = std.process.Child.run(.{
        .allocator = allocator,
        .argv = &.{
            "powershell",
            "-NoProfile",
            "-Command",
            script,
        },
    }) catch return false;
    defer allocator.free(result.stdout);
    defer allocator.free(result.stderr);

    switch (result.term) {
        .Exited => |code| {
            if (code != 0) return false;
        },
        else => return false,
    }

    const trimmed = std.mem.trim(u8, result.stdout, " \r\n\t");
    return std.mem.eql(u8, trimmed, "1");
}
