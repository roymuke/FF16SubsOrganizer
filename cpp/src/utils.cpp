#include "utils.hpp"

#include <algorithm>
#include <cctype>
#include <cstdlib>
#include <sstream>

#ifdef _WIN32
#  include <windows.h>
#endif

namespace ff16 {

void enable_vt_mode() {
#ifdef _WIN32
    HANDLE hOut = GetStdHandle(STD_OUTPUT_HANDLE);
    if (hOut == INVALID_HANDLE_VALUE) return;
    DWORD mode = 0;
    if (!GetConsoleMode(hOut, &mode)) return;
    mode |= ENABLE_VIRTUAL_TERMINAL_PROCESSING;
    SetConsoleMode(hOut, mode);
    SetConsoleOutputCP(CP_UTF8);
#endif
}

std::string trim(std::string_view s) {
    size_t b = 0, e = s.size();
    while (b < e && std::isspace(static_cast<unsigned char>(s[b]))) ++b;
    while (e > b && std::isspace(static_cast<unsigned char>(s[e - 1]))) --e;
    return std::string(s.substr(b, e - b));
}

std::string html_unescape(std::string_view s) {
    std::string out;
    out.reserve(s.size());
    for (size_t i = 0; i < s.size(); ) {
        if (s[i] == '&') {
            size_t sc = s.find(';', i + 1);
            if (sc != std::string_view::npos && sc - i <= 10) {
                std::string_view ent = s.substr(i + 1, sc - i - 1);
                if (ent == "amp")       { out += '&';  i = sc + 1; continue; }
                if (ent == "lt")        { out += '<';  i = sc + 1; continue; }
                if (ent == "gt")        { out += '>';  i = sc + 1; continue; }
                if (ent == "quot")      { out += '"';  i = sc + 1; continue; }
                if (ent == "apos")      { out += '\''; i = sc + 1; continue; }
                if (!ent.empty() && ent[0] == '#') {
                    int base = 10; size_t off = 1;
                    if (ent.size() > 1 && (ent[1] == 'x' || ent[1] == 'X')) { base = 16; off = 2; }
                    try {
                        unsigned long cp = std::stoul(std::string(ent.substr(off)), nullptr, base);
                        if (cp < 0x80) {
                            out += static_cast<char>(cp);
                        } else if (cp < 0x800) {
                            out += static_cast<char>(0xC0 | (cp >> 6));
                            out += static_cast<char>(0x80 | (cp & 0x3F));
                        } else if (cp < 0x10000) {
                            out += static_cast<char>(0xE0 | (cp >> 12));
                            out += static_cast<char>(0x80 | ((cp >> 6) & 0x3F));
                            out += static_cast<char>(0x80 | (cp & 0x3F));
                        } else {
                            out += static_cast<char>(0xF0 | (cp >> 18));
                            out += static_cast<char>(0x80 | ((cp >> 12) & 0x3F));
                            out += static_cast<char>(0x80 | ((cp >> 6) & 0x3F));
                            out += static_cast<char>(0x80 | (cp & 0x3F));
                        }
                        i = sc + 1;
                        continue;
                    } catch (...) {}
                }
            }
        }
        out += s[i++];
    }
    return out;
}

std::string sanitize_sheet_name(std::string_view name) {
    std::string s;
    if (name.size() <= 31) {
        s = std::string(name);
    } else {
        s = std::string(name.substr(0, 28)) + "...";
    }
    for (auto& c : s) {
        switch (c) {
            case '/': case '\\': case '[': case ']':
            case '*': case '?': case ':': c = '_'; break;
            default: break;
        }
    }
    return s;
}

int column_index_from_string(std::string_view col) {
    std::string letters;
    for (char c : col) {
        if (std::isalpha(static_cast<unsigned char>(c)))
            letters += static_cast<char>(std::toupper(static_cast<unsigned char>(c)));
    }
    if (letters.empty()) return 0;
    int idx = 0;
    for (char c : letters) {
        idx = idx * 26 + (c - 'A' + 1);
    }
    return idx;
}

std::string parent_basename(const std::filesystem::path& rel_path) {
    auto parent = rel_path.parent_path();
    if (parent.empty()) return {};
    return parent.filename().string();
}

std::string to_native_string(const std::filesystem::path& p) {
    return p.string();
}

std::string cmd_quote(const std::string& s) {
    if (s.find_first_of(" \t\"") == std::string::npos) return s;
    std::string out = "\"";
    int slashes = 0;
    for (char c : s) {
        if (c == '\\') { ++slashes; out += c; continue; }
        if (c == '"') {
            out.append(slashes + 1, '\\');
            slashes = 0;
            out += '"';
            continue;
        }
        slashes = 0;
        out += c;
    }
    out.append(slashes, '\\');
    out += '"';
    return out;
}

int run_silent(const std::vector<std::string>& argv) {
    if (argv.empty()) return -1;
#ifdef _WIN32
    // Build a command line that suppresses output. CreateProcessW would be more
    // robust, but std::system with redirection is sufficient and avoids the
    // 32 KB CreateProcess limit we'd otherwise need to manage manually
    // (note: cmd.exe still imposes ~8 KB; callers chunk for large folders).
    std::string cmd;
    for (size_t i = 0; i < argv.size(); ++i) {
        if (i) cmd += ' ';
        cmd += cmd_quote(argv[i]);
    }
    cmd += " >NUL 2>&1";
    return std::system(cmd.c_str());
#else
    std::string cmd;
    for (size_t i = 0; i < argv.size(); ++i) {
        if (i) cmd += ' ';
        cmd += '\'' + argv[i] + '\'';
    }
    cmd += " >/dev/null 2>&1";
    return std::system(cmd.c_str());
#endif
}

} // namespace ff16
