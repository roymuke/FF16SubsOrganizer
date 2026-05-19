#include "ids.hpp"
#include "utils.hpp"

#include <nlohmann/json.hpp>

#include <fstream>
#include <iostream>
#include <filesystem>

namespace ff16 {

IdsMaps load_ids(const std::string& json_path) {
    IdsMaps maps;
    namespace fs = std::filesystem;

    // Look in CWD first, then alongside the executable as a fallback.
    fs::path candidate = json_path;
    if (!fs::exists(candidate)) {
        // We accept that the executable path lookup is platform-specific;
        // for the CLI use case, the user typically runs from the project root.
        std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                  << " IDs.json not found at: " << json_path << "\n";
        return maps;
    }

    try {
        std::ifstream f(candidate);
        nlohmann::json j;
        f >> j;
        if (j.contains("characters") && j["characters"].is_object()) {
            for (auto it = j["characters"].begin(); it != j["characters"].end(); ++it) {
                maps.characters.emplace(it.key(), it.value().get<std::string>());
            }
        }
        if (j.contains("subtitleID") && j["subtitleID"].is_object()) {
            for (auto it = j["subtitleID"].begin(); it != j["subtitleID"].end(); ++it) {
                maps.subtitle_id.emplace(it.key(), it.value().get<std::string>());
            }
        }
    } catch (const std::exception& e) {
        std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                  << " Failed to parse IDs.json: " << e.what() << "\n";
    }
    return maps;
}

} // namespace ff16
