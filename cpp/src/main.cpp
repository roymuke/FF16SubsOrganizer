#include "commands.hpp"
#include "utils.hpp"

#include <CLI/CLI.hpp>

#include <iostream>
#include <string>

namespace {

const char* kBanner = "\033[38;5;81m\n"
" +----------------------------------------------+\n"
" | FFXVI Subtitle Organizer v1.4 (C++ port)     |\n"
" | by Roysu                                     |\n"
" +----------------------------------------------+\n"
" | https://github.com/roymuke/FF16SubsOrganizer |\n"
" +----------------------------------------------+\033[00m";

const char* kExamples =
"\nexamples:\n"
"  # Export subtitles to XLSX for editing\n"
"  > FF16SubsOrganizer to-xlsx -l \"C:/path/to/0007.en.XML\" -j \"C:/path/to/0007.ja/nxd/txt\" -o \"file.xlsx\"\n\n"
"  # Apply translations from XLSX back to XML\n"
"  > FF16SubsOrganizer edit-xml -f \"file.xlsx\" -col I2 -l \"C:/path/to/0007.en.XML\"\n\n"
"  # Convert in batch PZD->XML or XML->PZD\n"
"  > FF16SubsOrganizer convert-batch -c \"FF16Converter.exe\" -f \"C:/path/to/folder\" --pzd -m \"C:/dest\"\n\n"
"  # Move files to another destination by extension\n"
"  > FF16SubsOrganizer move-batch -f \"C:/src\" --pzd -m \"C:/dest\"";

} // namespace

int main(int argc, char** argv) {
    ff16::enable_vt_mode();

    CLI::App app{std::string(kBanner) + "\n"};
    app.footer(kExamples);
    app.require_subcommand(1);

    // ---------- to-xlsx ----------
    auto* xlsx = app.add_subcommand("to-xlsx", "Export subtitles to XLSX file.");
    ff16::ToXlsxArgs to_xlsx{};
    to_xlsx.output = "ff16_subtitles.xlsx";
    xlsx->add_option("-l,--language", to_xlsx.language, "Path to language subs folder to translate")
        ->required();
    xlsx->add_option("-j,--japanese", to_xlsx.japanese, "Path to Japanese subs folder")
        ->required();
    xlsx->add_option("-o,--output", to_xlsx.output, "Output xlsx file");
    xlsx->add_flag("-v,--verbose", to_xlsx.verbose, "Show detailed output messages");

    // ---------- edit-xml ----------
    auto* edit = app.add_subcommand("edit-xml", "Gets translations from XLSX back to XML files.");
    ff16::EditXmlArgs edit_args{};
    edit->add_option("-f,--file", edit_args.file, "XLSX file path")->required();
    edit->add_option("-col,--col", edit_args.col, "Column with new translations (e.g. I2)")
        ->required();
    edit->add_option("-l,--language", edit_args.language,
                     "Path to language to translate folder (e.g. C:\\...\\0007.en\\nxd\\text)")
        ->required();
    edit->add_flag("-v,--verbose", edit_args.verbose, "Show detailed output messages");

    // ---------- convert-batch ----------
    auto* conv = app.add_subcommand("convert-batch",
                                    "Convert files to another format, pzd->xml OR xml->pzd.");
    ff16::ConvertBatchArgs conv_args{};
    conv->add_option("-c,--converter", conv_args.converter, "Path to FF16Converter.exe")->required();
    conv->add_option("-f,--folder", conv_args.folder, "Path to language folder")->required();
    auto* conv_pzd = conv->add_flag_callback("--pzd",
        [&]() { conv_args.extension = ".pzd"; },
        "Extension to convert (pzd -> xml).");
    auto* conv_xml = conv->add_flag_callback("--xml",
        [&]() { conv_args.extension = ".xml"; },
        "Extension to convert (xml -> pzd).");
    conv_pzd->excludes(conv_xml);
    conv_xml->excludes(conv_pzd);
    conv->add_option("-m,--moveto", conv_args.moveto, "Path to converted files folder destination.");
    conv->add_flag("-v,--verbose", conv_args.verbose, "Show detailed output messages");

    // ---------- move-batch ----------
    auto* move = app.add_subcommand("move-batch", "Move files to another destination.");
    ff16::MoveBatchArgs move_args{};
    move->add_option("-f,--folder", move_args.folder, "Path to parent folder.")->required();
    auto* mv_pzd = move->add_flag_callback("--pzd",
        [&]() { move_args.extension = ".pzd"; },
        "Move PZD files.");
    auto* mv_xml = move->add_flag_callback("--xml",
        [&]() { move_args.extension = ".xml"; },
        "Move XML files.");
    mv_pzd->excludes(mv_xml);
    mv_xml->excludes(mv_pzd);
    move->add_option("-m,--moveto", move_args.moveto, "Path to folder destination.")->required();
    move->add_flag("-v,--verbose", move_args.verbose, "Show detailed output messages");

    CLI11_PARSE(app, argc, argv);

    try {
        if (xlsx->parsed())      return ff16::cmd_to_xlsx(to_xlsx);
        if (edit->parsed())      return ff16::cmd_edit_xml(edit_args);
        if (conv->parsed()) {
            int rc = ff16::cmd_convert_batch(conv_args);
            if (rc == 0 && !conv_args.moveto.empty() && !conv_args.extension.empty()) {
                ff16::MoveBatchArgs m;
                m.folder = conv_args.folder;
                m.moveto = conv_args.moveto;
                m.extension = conv_args.extension;
                m.verbose = conv_args.verbose;
                return ff16::cmd_move_batch(m);
            }
            return rc;
        }
        if (move->parsed())      return ff16::cmd_move_batch(move_args);
    } catch (const std::exception& e) {
        std::cerr << " \033[91m[ERROR]\033[00m " << e.what() << "\n";
        return 1;
    }

    return 0;
}
