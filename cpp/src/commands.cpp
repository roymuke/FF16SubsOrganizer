#include "commands.hpp"
#include "ids.hpp"
#include "utils.hpp"
#include "xml_io.hpp"

#include <OpenXLSX.hpp>
#include <pugixml.hpp>

#include <algorithm>
#include <chrono>
#include <cstdio>
#include <filesystem>
#include <iostream>
#include <map>
#include <set>
#include <sstream>
#include <string>
#include <system_error>
#include <unordered_map>
#include <vector>

namespace fs = std::filesystem;

namespace ff16 {

// (subdir, filename, msg_id, sub_type, chara_name, chara_id, en_text, jp_text)
struct TableRow {
    std::string subdir, filename, msg_id, sub_type, chara_name, chara_id, en_text, jp_text;
};

static std::vector<TableRow> collect_table(const std::string& lang_root,
                                           const std::string& jap_root) {
    std::vector<TableRow> rows;
    IdsMaps ids = load_ids();

    std::error_code ec;
    for (auto it = fs::recursive_directory_iterator(lang_root, ec);
         it != fs::recursive_directory_iterator(); it.increment(ec)) {
        if (ec) break;
        if (!it->is_regular_file()) continue;
        const fs::path& lang_path = it->path();
        if (lang_path.extension() != ".xml") continue;

        fs::path rel = fs::relative(lang_path, lang_root, ec);
        if (ec) { ec.clear(); continue; }
        fs::path jap_path = fs::path(jap_root) / rel;
        if (!fs::exists(jap_path)) {
            std::cout << " " << ansi::WARN << "[WARNING]" << ansi::RESET
                      << " Japanese path not found: " << jap_path.string() << "\n";
            continue;
        }

        auto lang_data = read_texts(lang_path.string());
        auto jap_data  = read_texts(jap_path.string());

        std::string subdir = parent_basename(rel);
        std::string filename = lang_path.stem().string();
        if (filename.size() >= 4 && filename.compare(filename.size() - 4, 4, ".pzd") == 0)
            filename.erase(filename.size() - 4);

        for (size_t i = 0; i < lang_data.size(); ++i) {
            const auto& [id_msg, chara_id, subtype, en_msg] = lang_data[i];
            std::string jp_msg = i < jap_data.size() ? std::get<3>(jap_data[i]) : std::string{};
            auto cit = ids.characters.find(chara_id);
            auto sit = ids.subtitle_id.find(subtype);
            rows.push_back(TableRow{
                subdir, filename, id_msg,
                sit != ids.subtitle_id.end() ? sit->second : std::string{},
                cit != ids.characters.end() ? cit->second : std::string{},
                chara_id, en_msg, jp_msg
            });
        }
    }
    return rows;
}

// ---------------------------------------------------------------------------
// to-xlsx
// ---------------------------------------------------------------------------

int cmd_to_xlsx(const ToXlsxArgs& a) {
    std::cout << "> Exporting to XLSX: " << ansi::PATH << a.output << ansi::RESET << "\n";
    auto table_rows = collect_table(a.language, a.japanese);

    // Preserve insertion order of subdirs, like Python's dict iteration.
    std::vector<std::string> subdir_order;
    std::unordered_map<std::string, std::vector<const TableRow*>> rows_by_subdir;
    for (const auto& r : table_rows) {
        auto [it, inserted] = rows_by_subdir.try_emplace(r.subdir);
        if (inserted) subdir_order.push_back(r.subdir);
        it->second.push_back(&r);
    }

    std::cout << "> Generating file...\n";

    try {
        using namespace OpenXLSX;
        // Remove any existing file so we start clean.
        std::error_code ec;
        fs::remove(a.output, ec);

        XLDocument doc;
        doc.create(a.output);
        XLWorkbook wb = doc.workbook();

        // Rename the default sheet to STATS.
        // OpenXLSX creates a "Sheet1" by default.
        std::string default_name = wb.sheetNames().front();
        if (default_name != "STATS") wb.sheet(default_name).setName("STATS");

        // STATS metadata for later (sheet_name, lines_count)
        std::vector<std::pair<std::string, size_t>> stats;

        for (const auto& subdir : subdir_order) {
            const auto& rows = rows_by_subdir[subdir];
            if (a.verbose)
                std::cout << " " << ansi::INFO << "[INFO]" << ansi::RESET
                          << " Processing: " << subdir << "\n";

            std::string sheet_name = sanitize_sheet_name(subdir);
            // Avoid name collisions with STATS or duplicates.
            std::string base = sheet_name;
            int suffix = 1;
            const auto names = wb.sheetNames();
            while (std::find(names.begin(), names.end(), sheet_name) != names.end()) {
                sheet_name = base + "_" + std::to_string(suffix++);
                if (sheet_name.size() > 31) sheet_name = sheet_name.substr(0, 31);
            }

            wb.addWorksheet(sheet_name);
            XLWorksheet ws = wb.worksheet(sheet_name);

            // Header.
            ws.cell("A1").value() = "Folder";
            ws.cell("B1").value() = "Filename";
            ws.cell("C1").value() = "ID";
            ws.cell("D1").value() = "Sub Type";
            ws.cell("E1").value() = "Character";
            ws.cell("F1").value() = "Character ID";
            ws.cell("G1").value() = "Original Text";
            ws.cell("H1").value() = "Japanese";
            ws.cell("I1").value() = "Retranslation";

            // Column widths (mirrors openpyxl widths in the Python).
            ws.column("G").setWidth(45);
            ws.column("H").setWidth(60);
            ws.column("I").setWidth(59);
            ws.column("J").setWidth(30);

            // Data rows.
            uint32_t row = 2;
            for (const TableRow* r : rows) {
                ws.cell(XLCellReference(row, 1)).value() = r->subdir;
                ws.cell(XLCellReference(row, 2)).value() = r->filename;
                ws.cell(XLCellReference(row, 3)).value() = r->msg_id;
                ws.cell(XLCellReference(row, 4)).value() = r->sub_type;
                ws.cell(XLCellReference(row, 5)).value() = r->chara_name;
                ws.cell(XLCellReference(row, 6)).value() = r->chara_id;
                ws.cell(XLCellReference(row, 7)).value() = r->en_text;
                ws.cell(XLCellReference(row, 8)).value() = r->jp_text;
                ws.cell(XLCellReference(row, 9)).value() = "";
                ++row;
            }

            // Approximate the openpyxl "B" width heuristic from the first data row.
            if (!rows.empty()) {
                double w = (rows.front()->filename.size() + 0.5) * 1.1207692307692307;
                ws.column("B").setWidth(w);
            }

            stats.emplace_back(sheet_name, rows.size());
        }

        // ----- STATS sheet -----
        XLWorksheet ss = wb.worksheet("STATS");
        ss.cell("A1").value() = "FFXVI Subtitle Translation Progress Sheet";
        ss.cell("A2").value() = "Made with FF16SubsOrganizer";

        ss.column("A").setWidth(12);
        ss.column("C").setWidth(10);
        ss.column("D").setWidth(50);

        // Number of populated sheet rows in the STATS table (11..len_sheets).
        size_t len_sheets = 9 + wb.sheetNames().size();
        auto fmt = [&](const std::string& key, const std::string& kind) {
            std::ostringstream o;
            o << "SUMPRODUCT(COUNTIF(INDIRECT(A11:A" << len_sheets
              << "&B11:B" << len_sheets << ");\"" << key << "\"))";
            (void)kind;
            return o.str();
        };
        auto pending = [&](const std::string& key, int total_row) {
            std::ostringstream o;
            o << "B" << total_row << "-SUMPRODUCT(COUNTIFS(INDIRECT(A11:A" << len_sheets
              << "&B11:B" << len_sheets << ");\"" << key
              << "\";INDIRECT(A11:A" << len_sheets << "&C11:C" << len_sheets << ");\"\"))";
            return o.str();
        };

        // Progress table header (row 4).
        ss.cell("A4").value() = "SubtitleType";
        ss.cell("B4").value() = "Lines";
        ss.cell("C4").value() = "Translated";
        ss.cell("D4").value() = "Progress";

        // Rows 5..7: Normal / SFX / Hidden.
        const char* labels[3]      = { "Normal", "SFX", "Hidden" };
        const int   progress_rows[3] = { 5, 6, 7 };
        for (int i = 0; i < 3; ++i) {
            int r = progress_rows[i];
            ss.cell(XLCellReference(r, 1)).value() = labels[i];
            ss.cell(XLCellReference(r, 2)).formula() = fmt(labels[i], "lines");
            ss.cell(XLCellReference(r, 3)).formula() = pending(labels[i], r);
            std::ostringstream prog; prog << "C" << r << "/B" << r;
            ss.cell(XLCellReference(r, 4)).formula() = prog.str();
        }
        // TOTAL row 8.
        ss.cell("A8").value() = "TOTAL";
        ss.cell("B8").formula() = "SUM(B5:B7)";
        ss.cell("C8").formula() = "SUM(C5:C7)";
        ss.cell("D8").formula() = "C8/B8";

        // Per-sheet references starting at row 11.
        uint32_t r = 11;
        for (const auto& [name, count] : stats) {
            // A: "<sheet_name>"!  B: "$D$2:$D$<n+1>"  C: "$I$2:$I$<n+1>"
            std::string sheet_ref = "'" + name + "'!";
            std::ostringstream b, c;
            b << "$D$2:$D$" << (count + 1);
            c << "$I$2:$I$" << (count + 1);
            ss.cell(XLCellReference(r, 1)).value() = sheet_ref;
            ss.cell(XLCellReference(r, 2)).value() = b.str();
            ss.cell(XLCellReference(r, 3)).value() = c.str();
            ss.row(r).setHidden(true);
            ++r;
        }

        doc.save();
        doc.close();

        std::cout << " " << ansi::DONE << "[DONE]" << ansi::RESET
                  << " XLSX file generated in: " << ansi::PATH << a.output << ansi::RESET << "\n";
        std::cout << " " << ansi::HINT
                  << "[INSTRUCTION] Edit the 'Retranslation' column (I) on each sheet. "
                     "Once done, use 'edit-xml' to apply changes." << ansi::RESET << "\n";
        return 0;
    } catch (const std::exception& e) {
        std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                  << " Couldn't generate file: " << e.what() << "\n";
        return 1;
    }
}

// ---------------------------------------------------------------------------
// edit-xml
// ---------------------------------------------------------------------------

int cmd_edit_xml(const EditXmlArgs& a) {
    std::cout << "> Applying translations from: " << ansi::PATH << a.file << ansi::RESET << "\n";

    int col_idx = column_index_from_string(a.col);
    if (col_idx <= 0) {
        std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                  << " Invalid column reference: " << a.col << "\n";
        return 1;
    }

    try {
        using namespace OpenXLSX;
        XLDocument doc;
        doc.open(a.file);
        XLWorkbook wb = doc.workbook();

        auto sheet_names = wb.sheetNames();
        if (sheet_names.size() < 2) {
            std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                      << " XLSX has no data sheets.\n";
            doc.close();
            return 1;
        }

        int changes_made = 0;
        std::set<std::string> files_processed;
        std::cout << "> Processing translations...\n";

        // Skip the first sheet (STATS), like the Python version.
        for (size_t s = 1; s < sheet_names.size(); ++s) {
            XLWorksheet ws = wb.worksheet(sheet_names[s]);
            uint32_t last_row = ws.rowCount();
            for (uint32_t r = 2; r <= last_row; ++r) {
                XLCell trans_cell = ws.cell(XLCellReference(r, col_idx));
                if (trans_cell.value().type() == XLValueType::Empty) continue;

                std::string subdir   = ws.cell(XLCellReference(r, 1)).value().get<std::string>();
                std::string filename = ws.cell(XLCellReference(r, 2)).value().get<std::string>();

                // Cell C may have been stored as integer or string.
                std::string msg_id;
                XLCell id_cell = ws.cell(XLCellReference(r, 3));
                switch (id_cell.value().type()) {
                    case XLValueType::String:
                        msg_id = id_cell.value().get<std::string>(); break;
                    case XLValueType::Integer:
                        msg_id = std::to_string(id_cell.value().get<int64_t>()); break;
                    case XLValueType::Float: {
                        std::ostringstream o; o << id_cell.value().get<double>();
                        msg_id = o.str(); break;
                    }
                    default: break;
                }
                if (filename.empty() || msg_id.empty()) continue;

                std::string new_translation;
                switch (trans_cell.value().type()) {
                    case XLValueType::String:
                        new_translation = trans_cell.value().get<std::string>(); break;
                    case XLValueType::Integer:
                        new_translation = std::to_string(trans_cell.value().get<int64_t>()); break;
                    case XLValueType::Float: {
                        std::ostringstream o; o << trans_cell.value().get<double>();
                        new_translation = o.str(); break;
                    }
                    default: continue;
                }
                std::string new_trim = trim(new_translation);

                fs::path xml_path = subdir.empty()
                    ? fs::path(a.language) / (filename + ".pzd.xml")
                    : fs::path(a.language) / subdir / (filename + ".pzd.xml");

                if (!fs::exists(xml_path)) {
                    std::cout << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                              << " File not found: " << xml_path.string() << "\n";
                    continue;
                }

                pugi::xml_document xdoc;
                pugi::xml_parse_result pr = xdoc.load_file(xml_path.string().c_str());
                if (!pr) {
                    std::cout << " " << ansi::WARN << "[WARNING]" << ansi::RESET
                              << " Could not process " << xml_path.string()
                              << ": " << pr.description() << "\n";
                    continue;
                }

                pugi::xml_node root = xdoc.document_element();
                if (std::string(root.name()) == "PzdFile") {
                    if (!root.attribute("xmlns:xsi"))
                        root.append_attribute("xmlns:xsi") = "http://www.w3.org/2001/XMLSchema-instance";
                    if (!root.attribute("xmlns:xsd"))
                        root.append_attribute("xmlns:xsd") = "http://www.w3.org/2001/XMLSchema";
                }

                pugi::xml_node tcs = root.child("TextContents");
                if (tcs) {
                    for (pugi::xml_node tc : tcs.children("TextContent")) {
                        if (std::string(tc.attribute("ID").as_string("")) != msg_id) continue;
                        pugi::xml_node msg = tc.child("Message");
                        if (!msg) break;
                        std::string old_text = msg.text().as_string("");
                        std::string applied = html_unescape(new_trim);

                        // Ensure the node has a text child.
                        if (!msg.first_child())
                            msg.append_child(pugi::node_pcdata).set_value(applied.c_str());
                        else
                            msg.text().set(applied.c_str());

                        if (old_text == new_translation) {
                            if (a.verbose)
                                std::cout << " " << ansi::SKIP
                                          << "[SKIP] Message " << msg_id
                                          << " already translated, skipping." << ansi::RESET << "\n";
                        } else {
                            if (a.verbose)
                                std::cout << " " << ansi::INFO << "[INFO]" << ansi::RESET
                                          << " " << filename << " (ID: " << msg_id << "): "
                                          << ansi::OLD << "\"" << old_text << "\"" << ansi::RESET
                                          << " -> " << ansi::HINT << "\"" << new_trim << "\""
                                          << ansi::RESET << "\n";
                            ++changes_made;
                            files_processed.insert(xml_path.string());
                        }
                        break;
                    }
                }
                write_xml(xdoc, xml_path.string());
            }
        }
        doc.close();

        std::cout << "\n " << ansi::DONE << "[DONE]" << ansi::RESET << " Summary:\n";
        std::cout << "   - " << changes_made << " translations applied.\n";
        std::cout << "   - " << files_processed.size() << " files modified.\n";
        return 0;
    } catch (const std::exception& e) {
        std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                  << " Error reading XLSX file: " << e.what() << "\n";
        return 1;
    }
}

// ---------------------------------------------------------------------------
// convert-batch
// ---------------------------------------------------------------------------

static std::string format_elapsed(double seconds) {
    int total = static_cast<int>(seconds);
    int h = total / 3600;
    int m = (total % 3600) / 60;
    int s = total % 60;
    char buf[16];
    std::snprintf(buf, sizeof(buf), "%02d:%02d:%02d", h, m, s);
    return buf;
}

int cmd_convert_batch(const ConvertBatchArgs& a) {
    fs::path lang_path = a.folder;
    if (!fs::exists(lang_path)) {
        std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                  << " Folder " << lang_path.string() << " does not exist\n";
        return 1;
    }
    if (!fs::exists(a.converter)) {
        std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                  << " Converter " << a.converter << " does not exist\n";
        return 1;
    }
    if (a.extension.empty()) {
        std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                  << " Extension not set.\n";
        return 1;
    }

    std::cout << "> Converting files in: " << ansi::PATH << lang_path.string() << ansi::RESET
              << "\n> Processing. This may take a while...\n";

    // Lowercase extension comparison helper.
    auto ext_eq = [&](const fs::path& p) {
        std::string e = p.extension().string();
        std::transform(e.begin(), e.end(), e.begin(),
                       [](unsigned char c){ return std::tolower(c); });
        return e == a.extension;
    };

    std::vector<fs::path> files_to_convert;
    std::error_code ec;
    for (auto it = fs::recursive_directory_iterator(lang_path, ec);
         it != fs::recursive_directory_iterator(); it.increment(ec)) {
        if (ec) break;
        if (it->is_regular_file() && ext_eq(it->path()))
            files_to_convert.push_back(it->path());
    }

    if (a.verbose)
        std::cout << " " << ansi::INFO << "[INFO]" << ansi::RESET
                  << " " << files_to_convert.size() << " files to convert\n";

    std::map<std::string, std::vector<fs::path>> folder_group;
    for (const auto& file : files_to_convert) {
        if (a.extension == ".pzd") {
            fs::path has_xml = file.string() + ".xml";
            if (fs::exists(has_xml)) {
                if (a.verbose)
                    std::cout << " " << ansi::SKIP << "[SKIP] "
                              << has_xml.filename().string()
                              << " already exists, skipping." << ansi::RESET << "\n";
                continue;
            }
        } else {
            fs::path has_pzd = file.string() + "RB.pzd";
            if (fs::exists(has_pzd)) {
                if (a.verbose)
                    std::cout << " " << ansi::SKIP << "[SKIP] "
                              << has_pzd.filename().string()
                              << " already exists, skipping." << ansi::RESET << "\n";
                continue;
            }
        }
        folder_group[file.parent_path().filename().string()].push_back(file);
    }

    auto start = std::chrono::steady_clock::now();

    for (auto& [folder, files] : folder_group) {
        if (a.verbose)
            std::cout << " " << ansi::INFO << "[INFO]" << ansi::RESET
                      << " Converting files on: " << ansi::HINT << folder << ansi::RESET << "\n";

        auto run_chunk = [&](size_t off, size_t len) {
            std::vector<std::string> argv;
            argv.reserve(len + 1);
            argv.push_back(a.converter);
            for (size_t i = 0; i < len; ++i) argv.push_back(files[off + i].string());
            int rc = run_silent(argv);
            if (rc != 0)
                std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                          << " Converter returned " << rc << " for folder " << folder << "\n";
        };

        if (folder == "defaultq" || folder == "simpleq") {
            for (size_t i = 0; i < files.size(); i += 400)
                run_chunk(i, std::min<size_t>(400, files.size() - i));
        } else {
            // Chunk to stay well below Windows command-line limits.
            constexpr size_t kChunk = 400;
            for (size_t i = 0; i < files.size(); i += kChunk)
                run_chunk(i, std::min(kChunk, files.size() - i));
        }
    }

    auto end = std::chrono::steady_clock::now();
    double elapsed = std::chrono::duration<double>(end - start).count();
    std::cout << " " << ansi::DONE << "[DONE]" << ansi::RESET
              << " Files converted in " << format_elapsed(elapsed) << "\n";
    return 0;
}

// ---------------------------------------------------------------------------
// move-batch
// ---------------------------------------------------------------------------

int cmd_move_batch(const MoveBatchArgs& a) {
    fs::path src = a.folder;
    fs::path dst = a.moveto;
    std::error_code ec;
    fs::create_directories(dst, ec);

    std::cout << "> Moving files to: " << ansi::PATH << dst.string() << ansi::RESET
              << "\n> Processing. This may take a while...\n";

    auto move_file = [](const fs::path& from, const fs::path& to) {
        std::error_code mec;
        fs::create_directories(to.parent_path(), mec);
        fs::rename(from, to, mec);
        if (mec) {
            // Fall back to copy + remove (cross-device).
            mec.clear();
            fs::copy_file(from, to, fs::copy_options::overwrite_existing, mec);
            if (!mec) fs::remove(from, mec);
        }
        return !mec;
    };

    if (a.extension == ".xml") {
        std::vector<fs::path> files;
        for (auto it = fs::recursive_directory_iterator(src, ec);
             it != fs::recursive_directory_iterator(); it.increment(ec)) {
            if (ec) break;
            if (!it->is_regular_file()) continue;
            std::string name = it->path().filename().string();
            if (name.size() >= 6 && name.compare(name.size() - 6, 6, "RB.pzd") == 0)
                files.push_back(it->path());
        }
        if (files.empty()) {
            std::cout << " " << ansi::DONE << "[DONE]" << ansi::RESET
                      << " Move operation completed.\n";
            return 0;
        }
        for (const auto& file : files) {
            try {
                fs::path rel = fs::relative(file, src, ec);
                std::string original_name = file.filename().string();
                // Replace ".pzd.xmlRB.pzd" -> ".pzd"
                const std::string from = ".pzd.xmlRB.pzd";
                auto pos = original_name.find(from);
                if (pos != std::string::npos)
                    original_name.replace(pos, from.size(), ".pzd");
                fs::path destination_file = dst / rel.parent_path() / original_name;
                if (fs::exists(destination_file)) {
                    if (a.verbose)
                        std::cout << " " << ansi::SKIP << "[SKIP] " << original_name
                                  << " already exists, skipping." << ansi::RESET << "\n";
                    continue;
                }
                if (!move_file(file, destination_file)) {
                    std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                              << " Error moving " << file.filename().string() << "\n";
                    continue;
                }
                if (a.verbose)
                    std::cout << " " << ansi::INFO << "[INFO]" << ansi::RESET
                              << " Moved: " << ansi::HINT << file.filename().string() << ansi::RESET
                              << " to " << ansi::PATH << destination_file.parent_path().string()
                              << ansi::RESET << "\n";
            } catch (const std::exception& e) {
                std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                          << " Error moving " << file.filename().string() << ": " << e.what() << "\n";
            }
        }
    } else if (a.extension == ".pzd") {
        std::vector<fs::path> files;
        for (auto it = fs::recursive_directory_iterator(src, ec);
             it != fs::recursive_directory_iterator(); it.increment(ec)) {
            if (ec) break;
            if (!it->is_regular_file()) continue;
            std::string name = it->path().filename().string();
            if (name.size() >= 8 && name.compare(name.size() - 8, 8, ".pzd.xml") == 0)
                files.push_back(it->path());
        }
        if (files.empty()) {
            std::cout << " " << ansi::DONE << "[DONE]" << ansi::RESET
                      << " Move operation completed.\n";
            return 0;
        }
        for (const auto& file : files) {
            try {
                fs::path rel = fs::relative(file, src, ec);
                // Mirror the Python: destination_file = (dst / rel.parent_path()).parent_path() / rel
                fs::path destination_folder = dst / rel.parent_path();
                fs::path destination_file = destination_folder.parent_path() / rel;
                if (fs::exists(destination_file)) {
                    if (a.verbose)
                        std::cout << " " << ansi::SKIP << "[SKIP] " << rel.string()
                                  << " already exists, skipping." << ansi::RESET << "\n";
                    continue;
                }
                fs::path target = destination_folder / file.filename();
                if (!move_file(file, target)) {
                    std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                              << " Error moving " << file.filename().string() << "\n";
                    continue;
                }
                if (a.verbose)
                    std::cout << " " << ansi::INFO << "[INFO]" << ansi::RESET
                              << " Moved: " << ansi::HINT << file.filename().string() << ansi::RESET
                              << " to " << ansi::PATH << destination_folder.string()
                              << ansi::RESET << "\n";
            } catch (const std::exception& e) {
                std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                          << " Error moving " << file.filename().string() << ": " << e.what() << "\n";
            }
        }
    } else {
        std::cerr << " " << ansi::ERROR << "[ERROR]" << ansi::RESET
                  << " Unknown or missing extension. Use --xml or --pzd.\n";
        return 1;
    }

    std::cout << " " << ansi::DONE << "[DONE]" << ansi::RESET << " Move operation completed.\n";
    return 0;
}

} // namespace ff16
