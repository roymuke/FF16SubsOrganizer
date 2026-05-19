#pragma once

#include <string>
#include <string_view>
#include <vector>
#include <filesystem>

namespace ff16 {

namespace ansi {
    inline constexpr const char* RESET   = "\033[00m";
    inline constexpr const char* ERROR   = "\033[91m";
    inline constexpr const char* WARN    = "\033[38;5;214m";
    inline constexpr const char* INFO    = "\033[38;5;75m";
    inline constexpr const char* DONE    = "\033[38;5;76m";
    inline constexpr const char* SKIP    = "\033[90m";
    inline constexpr const char* HINT    = "\033[38;5;81m";
    inline constexpr const char* OLD     = "\033[38;5;210m";
    inline constexpr const char* PATH    = "\033[48;5;235m";
    inline constexpr const char* CMD     = "\033[38;5;149m";
    inline constexpr const char* ARG     = "\033[38;5;222m";
}

// Enables ANSI/VT100 escape sequences in Windows consoles.
void enable_vt_mode();

// Trims leading/trailing ASCII whitespace.
std::string trim(std::string_view s);

// HTML entity unescape (subset: &amp; &lt; &gt; &quot; &apos; and numeric).
std::string html_unescape(std::string_view s);

// Sanitizes a sheet name for XLSX (replaces forbidden chars, max 31 chars).
std::string sanitize_sheet_name(std::string_view name);

// Splits an Excel-style column reference (e.g. "I2") into ("I", 2).
// Returns 0 column index on error.
int column_index_from_string(std::string_view col);

// Returns the parent directory's basename for a relative path.
std::string parent_basename(const std::filesystem::path& rel_path);

// Replaces backslashes with forward slashes (or vice versa) — neutral helper.
std::string to_native_string(const std::filesystem::path& p);

// Quotes a string for command-line use (Windows-style).
std::string cmd_quote(const std::string& s);

// Runs a command silently, returns exit code.
int run_silent(const std::vector<std::string>& argv);

} // namespace ff16
