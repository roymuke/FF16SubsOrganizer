import os, argparse, openpyxl.utils, json, subprocess
import xml.etree.ElementTree as ET
from html import escape, unescape
from collections import defaultdict
from openpyxl import Workbook, load_workbook

def get_ids():
    with open("IDs.json","r") as f:
        jsonIds = json.load(f)
    chara = jsonIds["characters"]
    subtp = jsonIds["subtitleID"]
    return chara, subtp

def read_texts(xml_path):
    try:
        with open(xml_path, "r", encoding="utf-8") as f:
            content = f.read()
        root = ET.fromstring(content)
        text_contents = root.find("TextContents")
        if text_contents is None:
            return []
        result = []
        for text_content in text_contents.findall("TextContent"):
            content_id = text_content.get("ID", "")
            message = text_content.findtext("Message", default="").strip()
            chara_id = text_content.get("Unknown2", "")
            subtype = text_content.get("Unknown3", "")
            result.append((content_id, chara_id, subtype, message))
        return result
    except Exception as e:
        print(f" \033[91m[ERROR]\033[00m Error reading {xml_path}: {e}")
        return []

def fix_xml_fields(root):
    text_contents = root.find("TextContents")
    if text_contents is None:
        return
    for text_content in text_contents.findall("TextContent"):
        message_elem = text_content.find("Message")
        if message_elem is None:
            message_elem = ET.SubElement(text_content, "Message")
        if message_elem.text is None:
            message_elem.text = ""
        voice_elem = text_content.find("Voice")
        if voice_elem is None:
            voice_elem = ET.SubElement(text_content, "Voice")
        if voice_elem.text is None:
            voice_elem.text = ""
        string_elem = text_content.find("String")
        if string_elem is None:
            string_elem = ET.SubElement(text_content, "String")
        if string_elem.text is None:
            string_elem.text = ""

def write_xml(tree, path):
    root = tree.getroot()
    fix_xml_fields(root)
    xml_body = ET.tostring(root, encoding="unicode", method="xml")
    xml_content = '<?xml version="1.0" encoding="utf-16"?>\r\n' + xml_body
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(xml_content)

def collect_table(lang_root, jap_root):
    table_rows = []
    characters, subtitleID = get_ids()
    for root_dir, _, files in os.walk(lang_root):
        for file in files:
            if not file.endswith(".xml"):
                continue
            lang_path = os.path.join(root_dir, file)
            rel_path = os.path.relpath(lang_path, lang_root)
            subdir = os.path.basename(os.path.dirname(rel_path))
            jap_path = os.path.join(jap_root, rel_path)
            if not os.path.exists(jap_path):
                print(f" \033[38;5;214m[WARNING]\033[00m Japanese path not found: {jap_path}")
                continue
            lang_data = read_texts(lang_path)
            jap_data = read_texts(jap_path)
            filename = os.path.splitext(file)[0]
            if filename.endswith(".pzd"):
                filename = filename[:-4]
            for idx, (id_msg, chara_id, subtype, en_msg) in enumerate(lang_data):
                jp_msg = jap_data[idx][3] if idx < len(jap_data) else ""
                chara_name = characters.get(chara_id, "")
                sub_type = subtitleID.get(subtype, "")
                table_rows.append((subdir, filename, id_msg, sub_type, chara_name, chara_id, en_msg, jp_msg))
    return table_rows

def safe_html(text):
    return escape(text, quote=False).replace("&", "&amp;")

# Command: to-html
def export_html(table_rows, output):
    html_content = [
        "<html><head><meta charset='UTF-8'><style>",
        "table {border-collapse: collapse; width:100%;}",
        "th, td {border: 1px solid black; padding: 4px; vertical-align: top;}",
        "th {background: #ddd;}",
        "</style></head><body>",
        "<table>",
        "<tr><th>Filename</th><th>ID</th><th>Sub type</th><th>Chara</th><th>Chara ID</th><th>Original translation</th><th>Japanese</th></tr>"
    ]
    grouped = defaultdict(list)
    for subdir, filename, msg_id, subtype, chara_name, chara_id, en, jp in table_rows:
        grouped[filename].append((msg_id, subtype, chara_name, chara_id, en, jp))
    for filename, rows in grouped.items():
        first_row = True
        for msg_id, subtype, chara_name, chara_id, en_text, jp_text in rows:
            html_content.append("<tr>")
            if first_row:
                html_content.append(f"<td class='filename'><code>{filename}</code></td>")
                first_row = False
            else:
                html_content.append("<td></td>")
            html_content.append(f"<td>{msg_id}</td>")
            html_content.append(f"<td>{subtype}</td>")
            html_content.append(f"<td>{chara_name}</td>")
            html_content.append(f"<td>{chara_id}</td>")
            html_content.append(f"<td>{safe_html(en_text)}</td>")
            html_content.append(f"<td>{safe_html(jp_text)}</td>")
            html_content.append("</tr>")
    html_content.extend(["</table>", "</body></html>"])
    with open(output, "w", encoding="utf-8") as f:
        f.write("\n".join(html_content))
    print(f" \033[38;5;76m[DONE]\033[00m HTML generated in: \033[48;5;235m{output}\033[00m")

# Command: to-xlsx
def export_xlsx(table_rows, output):
    wb = Workbook()
    ws = wb.active
    ws.title = "Retranslation"
    ws.append(["Folder", "Filename", "ID", "Sub Type", "Character", "Character ID", "Original Text", "Japanese", "Retranslation"])
    for subdir, filename, msg_id, subtype, chara_name, chara_id, en_text, jp_text in table_rows:
        ws.append([subdir, filename, msg_id, subtype, chara_name, chara_id,en_text, jp_text, ""])
    wb.save(output)
    print(f" \033[38;5;76m[DONE]\033[00m XLSX file generated in: \033[48;5;235m{output}\033[00m")
    print(f" \033[38;5;81m[INSTRUCTION] Edit the 'Retranslation' column (I) and then use 'edit-xml' to apply changes.\033[00m")

# Command: edit-xml
def edit_xml(xlsx_path, col_reference, lang_root):
    try:
        wb = load_workbook(xlsx_path)
        ws = wb.active
        col_letter = ''.join(filter(str.isalpha, col_reference.upper()))
        col_idx = openpyxl.utils.column_index_from_string(col_letter)
        changes_made = 0
        files_processed = set()
        print("> Processing translations...")
        for row in ws.iter_rows(min_row=2, values_only=False):
            if not row or len(row) < col_idx:
                continue
            subdir = row[0].value if row[0].value else ""
            filename = row[1].value if row[1].value else ""
            msg_id = str(row[2].value) if row[2].value else ""
            new_translation = row[col_idx - 1].value if row[col_idx - 1].value else ""
            if not new_translation or not new_translation.strip(): continue
            xml_filename = f"{filename}.pzd.xml"
            xml_path = os.path.join(lang_root, subdir, xml_filename) if subdir else os.path.join(lang_root, xml_filename)
            if not os.path.exists(xml_path):
                print(f" \033[91m[ERROR]\033[00m File not found: {xml_path}")
                continue
            try:
                with open(xml_path, "r", encoding="utf-8") as f:
                    content = f.read()
                root = ET.fromstring(content)
                if root.tag == "PzdFile":
                    root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
                    root.set("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
                text_contents = root.find("TextContents")
                if text_contents is not None:
                    for text_content in text_contents.findall("TextContent"):
                        if text_content.get("ID") == msg_id:
                            message_elem = text_content.find("Message")
                            if message_elem is not None:
                                old_text = message_elem.text or ""
                                message_elem.text = unescape(new_translation.strip())
                                if old_text == new_translation:
                                    print(f" \033[90m[SKIP] Message {msg_id} already translated, skipping.\033[00m") 
                                    continue
                                else:
                                    print(f" \033[38;5;75m[INFO]\033[00m {filename} (ID: {msg_id}): \033[38;5;210m\"{old_text}\"\033[00m -> \033[38;5;81m\"{new_translation.strip()}\"\033[00m")
                                    changes_made += 1
                                    files_processed.add(xml_path)
                                break
                tree = ET.ElementTree(root)
                write_xml(tree, xml_path)
            except Exception as e:
                print(f" \033[91m[ERROR]\033[00m Error processing {xml_path}: {e}")
                continue
        print(f"\n \033[38;5;76m[DONE]\033[00m Summary:")
        print(f"   • {changes_made} translations applied.")
        print(f"   • {len(files_processed)} files modified.")
    except Exception as e:
        print(f" \033[91m[ERROR]\033[00m Error reading XLSX file: {e}")

# Command: convert-batch
def convert_batch(ff16converter, lang_path, valid_ext):
    from pathlib import Path
    from time import perf_counter
    lang_path = Path(lang_path)
    if not lang_path.exists():
        print(f" \033[91m[ERROR]\033[00m Folder {lang_path} does not exist")
        return
    if not Path(ff16converter).exists():
        print(f" \033[91m[ERROR]\033[00m Converter {ff16converter} does not exist")
        return
    if not valid_ext:
        print(f" \033[91m[ERROR]\033[00m Extension not set.")
        return
    if valid_ext == ".pzd": ext_convert = ".xml"
    else: ext_convert = ".pzd"
    valid_ext = [valid_ext]
    print(f"> Converting files in: \033[48;5;235m{lang_path}\033[00m")
    files_to_convert = [
        file for file in lang_path.rglob("*")
        if file.suffix.lower() in valid_ext and file.is_file()
    ]
    print(f" \033[38;5;75m[INFO]\033[00m {len(files_to_convert)} files to convert")
    start_time = perf_counter()
    for file in files_to_convert:
        try:
            if valid_ext == [".pzd"]:
                has_xml = Path(str(file) + ".xml")
                if has_xml.exists():
                    print(f" \033[90m[SKIP] {has_xml.name} already exists, skipping.\033[00m")
                    continue
            else:
                has_pzd = Path(str(file) + "RB.pzd")
                if has_pzd.exists():
                    print(f" \033[90m[SKIP] {has_pzd.name} already exists, skipping.\033[00m")
                    continue
            print(f" \033[38;5;75m[INFO]\033[00m Converting: \033[38;5;81m{file.name}\033[00m to \033[38;5;211m{ext_convert}\033[00m")
            result = subprocess.run([ff16converter, str(file)], capture_output=True, text=True)
            if result.returncode != 0:
                print(f" \033[38;5;214m[WARNING]\033[00m Conversion failed for {file.name}")
                print(f" \033[91m[ERROR]\033[00m {result.stderr}")
        except FileNotFoundError:
            print(f" \033[91m[ERROR]\033[00m Converter not found at {ff16converter}")
            break
        except Exception as e:
            print(f" \033[91m[ERROR]\033[00m Error converting {file}: {e}")
    time_lapsed = perf_counter() - start_time
    print(f" \033[38;5;76m[DONE]\033[00m Files converted in {int(time_lapsed // 3600):02d}:{int((time_lapsed % 3600) // 60):02d}:{int(time_lapsed % 60):02d}")

def move_converted(this_directory, to_directory, extension):
    from pathlib import Path
    import shutil
    Path.mkdir(Path(to_directory), parents=True, exist_ok=True)
    match extension:
        case ".xml":
            converted_files = list(Path(this_directory).rglob("*RB.pzd"))
            if not converted_files:
                return
            for file in converted_files:
                try:
                    relative_path = file.relative_to(this_directory)
                    original_name = file.name.replace(".pzd.xmlRB.pzd",".pzd")
                    destination_file = to_directory / relative_path.parent / original_name
                    Path.mkdir(destination_file.parent, parents=True, exist_ok=True)
                    if destination_file.exists():
                        print(f" \033[90m[SKIP] {original_name} already exists, skipping.\033[00m")
                        continue
                    shutil.move(str(file), str(destination_file))
                    print(f" \033[38;5;75m[INFO]\033[00m Moved: \033[38;5;81m{file.name}\033[00m to \033[48;5;235m{destination_file.parent}\033[00m")
                except Exception as e:
                    print(f" \033[91m[ERROR]\033[00m Error moving {file.name}: {e}")
        case ".pzd":
            converted_files = list(Path(this_directory).rglob("*.pzd.xml"))
            if not converted_files:
                return
            for file in converted_files:
                try:
                    relative_path = file.relative_to(this_directory)
                    destination_folder = to_directory / relative_path.parent
                    destination_file = destination_folder.parent / relative_path
                    Path.mkdir(destination_folder, parents=True, exist_ok=True)
                    if destination_file.exists():
                        print(f" \033[90m[SKIP] {relative_path} already exists, skipping.\033[00m")
                        continue
                    shutil.move(str(file), str(destination_folder))
                    print(f" \033[38;5;75m[INFO]\033[00m Moved: \033[38;5;81m{file.name}\033[00m to \033[48;5;235m{destination_folder}\033[00m")
                except Exception as e:
                    print(f" \033[91m[ERROR]\033[00m Error moving {file.name}: {e}")
    print(f" \033[38;5;76m[DONE]\033[00m Move operation completed.")

def main():
    os.system("color")
    parser = argparse.ArgumentParser(
        description="""\033[38;5;81m
 +----------------------------------------------+
 | FFXVI Subtitle Organizer v1.2                |
 | by Roysu                                     |
 +----------------------------------------------+
 | https://github.com/roymuke/FF16SubsOrganizer |
 +----------------------------------------------+\033[00m""",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  \033[90m# Export to HTML for preview\033[00m
  > python \033[38;5;149mFF16SubsOrganizer.py\033[00m to-html \033[38;5;149m-l\033[00m \033[38;5;222m"C:\path\\to\\folder\\0001.en"\033[00m \033[38;5;149m-j\033[00m \033[38;5;222m"C:\path\\to\\folder\\0001.ja"\033[00m [\033[38;5;149m-o\033[00m \033[38;5;222m"C:\custom\path\\to\\file.html"\033[00m]

  \033[90m# Export to XLSX for editing\033[00m
  > python \033[38;5;149mFF16SubsOrganizer.py\033[00m to-xlsx \033[38;5;149m-l\033[00m \033[38;5;222m"C:\path\\to\\folder\\0001.en.XML"\033[00m \033[38;5;149m-j\033[00m \033[38;5;222m"C:\path\\to\\folder\\0001.ja\\nxd\\txt"\033[00m [\033[38;5;149m-o\033[00m \033[38;5;222m"C:\custom\path\\to\\file.xlsx"\033[00m]

  \033[90m# Apply translations from XLSX back to XML\033[00m
  > python \033[38;5;149mFF16SubsOrganizer.py\033[00m edit-xml \033[38;5;149m-f\033[00m \033[38;5;222m"file.xlsx"\033[00m \033[38;5;149m-col\033[00m I2 \033[38;5;149m-l\033[00m \033[38;5;222m"C:\path\\to\\folder\\0001.en"\033[00m

  \033[90m# Convert in batch XML back to PZD\033[00m
  > python \033[38;5;149mFF16SubsOrganizer.py\033[00m convert-batch \033[38;5;149m-c\033[00m \033[38;5;222m"C:\path\\to\\FF16Converter.exe"\033[00m \033[38;5;149m-f\033[00m \033[38;5;222m"C:\path\\to\\folder\\0001.en\\nxd\\text"\033[00m \033[38;5;149m--xml\033[00m [\033[38;5;149m-m\033[00m \033[38;5;222m"C:\path\\to\moving\\folder"\033[00m]""")
    subparsers = parser.add_subparsers(dest="command", required=True, help="Available commands")
    # to-html command
    html_parser = subparsers.add_parser("to-html", help="Export subtitles to HTML table.")
    html_parser.add_argument("-l", "--language", required=True, help="Path to language subs folder to translate")
    html_parser.add_argument("-j", "--japanese", required=True, help="Path to Japanese subs folder")
    html_parser.add_argument("-o", "--output", default="ff16_subtitles.html", help="Output HTML file")
    # to-xlsx command
    xlsx_parser = subparsers.add_parser("to-xlsx", help="Export subtitles to XLSX file.")
    xlsx_parser.add_argument("-l", "--language", required=True, help="Path to language subs folder to translate")
    xlsx_parser.add_argument("-j", "--japanese", required=True, help="Path to Japanese  subsfolder")
    xlsx_parser.add_argument("-o", "--output", default="ff16_subtitles.xlsx", help="Output xlsx file")
    # edit-xml command
    edit_parser = subparsers.add_parser("edit-xml", help="Gets translations from XLSX back to XML files.")
    edit_parser.add_argument("-f", "--file", required=True, help="XLSX file path")
    edit_parser.add_argument("-col", required=True, help="Column with new translations (e.g., I2)")
    edit_parser.add_argument("-l", "--language", required=True, help="Path to language to translate folder (e.g. C:\...\0001.en\nxd\text)")
    # convert-batch command
    move_parser = subparsers.add_parser("convert-batch", help="Convert entire XML files back to PZD.")
    move_parser.add_argument("-c", "--converter", required=True, help="Path to FF16Converter.exe")
    move_parser.add_argument("-f", "--folder", required=True, help="Path to language folder")
    move_parser.add_argument("--xml", action="store_const", const=".xml", dest="extension", help="Extension to convert, i.e, XML to PZD.")
    move_parser.add_argument("--pzd", action="store_const", const=".pzd", dest="extension", help="Extension to convert, i.e, PZD to XML.")
    move_parser.add_argument("-m", "--moveto", help="Path to converted pzd files folder destination.")

    args = parser.parse_args()

    if args.command == "to-html":
        print(f"> Exporting to HTML: \033[48;5;235m{args.output}\033[00m")
        table_rows = collect_table(args.language, args.japanese)
        export_html(table_rows, args.output)
    elif args.command == "to-xlsx":
        print(f"> Exporting to XLSX: \033[48;5;235m{args.output}\033[00m")
        table_rows = collect_table(args.language, args.japanese)
        export_xlsx(table_rows, args.output)
    elif args.command == "edit-xml":
        print(f"> Applying translations from: \033[48;5;235m{args.file}\033[00m")
        edit_xml(args.file, args.col, args.language)
    elif args.command == "convert-batch":
        convert_batch(args.converter,args.folder,args.extension)
        if args.moveto and args.extension:
            print(f"> Moving files to: \033[48;5;235m{args.moveto}\033[00m")
            move_converted(args.folder, args.moveto, args.extension)

if __name__ == "__main__":
    main()