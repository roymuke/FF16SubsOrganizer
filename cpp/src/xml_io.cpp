#include "xml_io.hpp"
#include "utils.hpp"

#include <fstream>
#include <iostream>
#include <sstream>

namespace ff16 {

std::vector<TextEntry> read_texts(const std::string& xml_path) {
    std::vector<TextEntry> result;
    pugi::xml_document doc;
    pugi::xml_parse_result pr = doc.load_file(xml_path.c_str());
    if (!pr) {
        std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                  << " Error reading " << xml_path << ": " << pr.description() << "\n";
        return result;
    }

    pugi::xml_node root = doc.document_element();
    pugi::xml_node text_contents = root.child("TextContents");
    if (!text_contents) return result;

    for (pugi::xml_node tc : text_contents.children("TextContent")) {
        std::string id      = tc.attribute("ID").as_string("");
        std::string chara   = tc.attribute("Unknown2").as_string("");
        std::string subtype = tc.attribute("Unknown3").as_string("");
        std::string message = trim(tc.child("Message").text().as_string(""));
        result.emplace_back(std::move(id), std::move(chara), std::move(subtype), std::move(message));
    }
    return result;
}

void fix_xml_fields(pugi::xml_node root) {
    pugi::xml_node text_contents = root.child("TextContents");
    if (!text_contents) return;

    auto ensure_child = [](pugi::xml_node parent, const char* name) {
        pugi::xml_node n = parent.child(name);
        if (!n) n = parent.append_child(name);
        if (!n.first_child()) n.append_child(pugi::node_pcdata).set_value("");
    };

    for (pugi::xml_node tc : text_contents.children("TextContent")) {
        ensure_child(tc, "Message");
        ensure_child(tc, "Voice");
        ensure_child(tc, "String");
    }
}

bool write_xml(const pugi::xml_document& doc, const std::string& path) {
    // pugixml prints with its own declaration handling; we emit the body
    // without a leading declaration, then prepend the utf-16 declaration line
    // the original Python script writes — this matches FF16Converter's input
    // expectation while we keep the byte stream as UTF-8 on disk.
    std::ostringstream oss;
    doc.save(oss, "", pugi::format_raw | pugi::format_no_declaration, pugi::encoding_utf8);

    std::ofstream f(path, std::ios::binary);
    if (!f) {
        std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                  << " Cannot open for writing: " << path << "\n";
        return false;
    }
    f << "<?xml version=\"1.0\" encoding=\"utf-16\"?>\r\n";
    f << oss.str();
    return static_cast<bool>(f);
}

} // namespace ff16
