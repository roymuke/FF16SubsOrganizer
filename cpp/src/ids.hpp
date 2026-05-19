#pragma once

#include <string>
#include <unordered_map>

namespace ff16 {

struct IdsMaps {
    std::unordered_map<std::string, std::string> characters;
    std::unordered_map<std::string, std::string> subtitle_id;
};

// Loads "IDs.json" from a candidate path; returns empty maps on failure.
IdsMaps load_ids(const std::string& json_path = "IDs.json");

} // namespace ff16
