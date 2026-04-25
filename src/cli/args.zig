pub const Command = enum {
    preflight,
    help,
};

pub const ParsedArgs = struct {
    command: Command,
    json: bool,
};

pub fn parseArgs(args: []const []const u8) ParsedArgs {
    var parsed = ParsedArgs{ .command = .preflight, .json = false };

    if (args.len <= 1) return parsed;

    var idx: usize = 1;
    while (idx < args.len) : (idx += 1) {
        const arg = args[idx];

        if (std.mem.eql(u8, arg, "--json")) {
            parsed.json = true;
            continue;
        }

        if (std.mem.eql(u8, arg, "preflight")) {
            parsed.command = .preflight;
            continue;
        }
        if (std.mem.eql(u8, arg, "help") or std.mem.eql(u8, arg, "--help") or std.mem.eql(u8, arg, "-h")) {
            parsed.command = .help;
            continue;
        }

        parsed.command = .help;
    }

    return parsed;
}

const std = @import("std");
