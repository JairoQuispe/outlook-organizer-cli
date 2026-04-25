const std = @import("std");

pub const Spinner = struct {
    thread: ?std.Thread = null,
    running: std.atomic.Value(bool) = .{ .raw = false },
    mutex: std.Thread.Mutex = .{},
    status_buf: [256]u8 = undefined,
    status_len: usize = 0,
    enabled: bool = true,

    // Frames ASCII universales (renderizan en cualquier terminal, incluido cmd.exe)
    const frames = [_][]const u8{ "|", "/", "-", "\\" };
    const frame_interval_ns: u64 = 90 * std.time.ns_per_ms;

    pub fn init(enabled: bool) Spinner {
        return .{ .enabled = enabled };
    }

    pub fn start(self: *Spinner, status: []const u8) !void {
        if (!self.enabled) return;
        if (self.thread != null) return;

        self.setStatus(status);
        // Renderiza el primer frame YA, antes de spawnar el thread, para que
        // el usuario vea la animacion incluso si la tarea termina rapido.
        std.debug.print("\r\x1b[2K\x1b[36m{s}\x1b[0m {s}", .{ frames[0], status });

        self.running.store(true, .release);
        self.thread = try std.Thread.spawn(.{}, runLoop, .{self});
    }

    pub fn setStatus(self: *Spinner, status: []const u8) void {
        if (!self.enabled) return;
        self.mutex.lock();
        defer self.mutex.unlock();

        const n = @min(status.len, self.status_buf.len);
        @memcpy(self.status_buf[0..n], status[0..n]);
        self.status_len = n;
    }

    pub fn stop(self: *Spinner) void {
        if (!self.enabled) return;
        if (self.thread == null) return;

        self.running.store(false, .release);
        self.thread.?.join();
        self.thread = null;

        // Limpia la linea actual
        std.debug.print("\r\x1b[2K", .{});
    }

    /// Detiene el spinner y deja un mensaje final con marcador ASCII de exito
    pub fn stopWithSuccess(self: *Spinner, message: []const u8) void {
        self.stop();
        if (self.enabled) {
            std.debug.print("\x1b[32m[OK]\x1b[0m {s}\n", .{message});
        } else {
            std.debug.print("[OK] {s}\n", .{message});
        }
    }

    /// Detiene el spinner y deja un mensaje final con marcador ASCII de fallo
    pub fn stopWithFailure(self: *Spinner, message: []const u8) void {
        self.stop();
        if (self.enabled) {
            std.debug.print("\x1b[31m[X]\x1b[0m  {s}\n", .{message});
        } else {
            std.debug.print("[X] {s}\n", .{message});
        }
    }

    fn runLoop(self: *Spinner) void {
        var frame_idx: usize = 1;
        while (self.running.load(.acquire)) {
            self.mutex.lock();
            const status = self.status_buf[0..self.status_len];
            // Cyan para el frame del spinner
            std.debug.print("\r\x1b[2K\x1b[36m{s}\x1b[0m {s}", .{ frames[frame_idx], status });
            self.mutex.unlock();

            frame_idx = (frame_idx + 1) % frames.len;
            std.Thread.sleep(frame_interval_ns);
        }
    }
};
