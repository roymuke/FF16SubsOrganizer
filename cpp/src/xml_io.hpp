#pragma once

#include <pugixml.hpp>

#include <string>
#include <tuple>
#include <vector>

namespace ff16 {

// (ID, chara_id, subtype, message)
using TextEntry = std::tuple<std::string, std::string, std::string, std::string>;

// Reads <TextContent> entries from a PZD XML file. Returns empty on error.
std::vector<TextEntry> read_texts(const std::string& xml_path);

// Ensures Message/Voice/String children exist with at least empty text.
void fix_xml_fields(pugi::xml_node root);

// Writes the document with an explicit utf-16 declaration line (mirrors the
// Python script behavior), followed by the XML body serialized as UTF-8.
bool write_xml(const pugi::xml_document& doc, const std::string& path);

} // namespace ff16
