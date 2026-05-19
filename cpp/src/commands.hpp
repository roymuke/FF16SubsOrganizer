#pragma once

#include <string>

namespace ff16 {

struct ToXlsxArgs {
    std::string language;
    std::string japanese;
    std::string output;
    bool verbose = false;
};

struct EditXmlArgs {
    std::string file;
    std::string col;
    std::string language;
    bool verbose = false;
};

struct ConvertBatchArgs {
    std::string converter;
    std::string folder;
    std::string extension; // ".pzd" or ".xml"
    std::string moveto;    // optional
    bool verbose = false;
};

struct MoveBatchArgs {
    std::string folder;
    std::string moveto;
    std::string extension; // ".pzd" or ".xml"
    bool verbose = false;
};

int cmd_to_xlsx(const ToXlsxArgs& a);
int cmd_edit_xml(const EditXmlArgs& a);
int cmd_convert_batch(const ConvertBatchArgs& a);
int cmd_move_batch(const MoveBatchArgs& a);

} // namespace ff16
