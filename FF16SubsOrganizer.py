import os
import argparse
import xml.etree.ElementTree as ET
from html import escape, unescape
from collections import defaultdict
from openpyxl import Workbook, load_workbook
import openpyxl.utils

# ============================================================
# Character and Subtitle Type Mappings
# ============================================================

characters = {"90010001": "Clive", "90020001": "Clive", "100100": "Clive", "100101": "Clive", "302200": "Aevis", "302300": "Tiamat", "302400": "Biast", "90040001": "Eugen Havel", "90040002": "Barnabas", "90040003": "Sleipnir", "90040004": "Benedikta", "200101": "Benedikta", "200102": "Benedikta", "90040005": "Kupka", "200200": "Kupka", "90080001": "Joshua", "90080002": "Rodney", "301400": "Rodney", "100200": "Joshua", "100206": "Joshua", "100300": "Jill", "300800": "Anabella", "300900": "Elwin", "301200": "Tyler", "301201": "Tyler", "301300": "Wade", "301303": "Wade", "100400": "Cid"}
subtitleID = {"0": "Normal", "1": "Hidden", "2": "SFX"}

# ============================================================
# Core XML Functions
# ============================================================

# Read XML texts with proper encoding handling
def read_texts_fixed(xml_path):
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
        print(f"Error reading {xml_path}: {e}")
        return []

# Fix null/empty fields that cause C# converter errors
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

# Write XML with proper encoding and field fixes
def write_xml_fixed(tree, path):
    root = tree.getroot()
    fix_xml_fields(root)
    xml_body = ET.tostring(root, encoding="unicode", method="xml")
    xml_content = '<?xml version="1.0" encoding="utf-16"?>\r\n' + xml_body
    with open(path, "w", encoding="utf-8", newline="") as f:
        f.write(xml_content)

# ============================================================
# Data Collection and Processing
# ============================================================

# Collect subtitle data from selected language and japanese folders
def collect_table(lang_root, jap_root):
    table_rows = []
    for root_dir, _, files in os.walk(lang_root):
        for file in files:
            if not file.endswith(".xml"):
                continue
            lang_path = os.path.join(root_dir, file)
            rel_path = os.path.relpath(lang_path, lang_root)
            subdir = os.path.dirname(rel_path)
            jap_path = os.path.join(jap_root, rel_path)
            if not os.path.exists(jap_path):
                print(f"[WARNING] Japanese path not found: {jap_path}")
                continue
            lang_data = read_texts_fixed(lang_path)
            jap_data = read_texts_fixed(jap_path)
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

# ============================================================
# Command: to-html
# ============================================================

# Export subtitle data to HTML table
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
    print(f"> HTML exported to: {output}")

# ============================================================
# Command: to-xlsx
# ============================================================

# Export subtitle data to XLSX file
def export_xlsx(table_rows, output):
    wb = Workbook()
    ws = wb.active
    ws.title = "Retranslation"
    ws.append(["Folder", "Filename", "ID", "Sub Type", "Character", "Character ID", "Original Text", "Japanese", "Retranslation"])
    for subdir, filename, msg_id, subtype, chara_name, chara_id, en_text, jp_text in table_rows:
        ws.append([subdir, filename, msg_id, subtype, chara_name, chara_id,en_text, jp_text, ""])
    wb.save(output)
    print(f"> XLSX file exported to: {output}")
    print(f" (INSTRUCTION) Edit the 'Retranslation' column (I) and then use 'edit-xml' to apply changes")

# ============================================================
# Command: edit-xml
# ============================================================

# Apply translations from XLSX back to XML files with proper encoding
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
                print(f"[ERROR] File not found: {xml_path}")
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
                                print(f" > {filename} (ID: {msg_id}): '{old_text}' -> '{new_translation.strip()}'")
                                changes_made += 1
                                files_processed.add(xml_path)
                                break
                tree = ET.ElementTree(root)
                write_xml_fixed(tree, xml_path)
            except Exception as e:
                print(f"[ERROR] Error processing {xml_path}: {e}")
                continue
        print(f"\n Summary:")
        print(f"   • {changes_made} translations applied")
        print(f"   • {len(files_processed)} files modified")
    except Exception as e:
        print(f"[ERROR] Error reading XLSX file: {e}")

# ============================================================
# Main Function
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="FFXVI Subtitle Organizer v1.0",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Export to HTML for preview
  > python FF16SubsOrganizer.py to-html -l "C:\path\\to\\folder\\0001.en" -j "C:\path\\to\\folder\\0001.ja" [-o "C:\custom\path\\to\\file.html"]
  
  # Export to XLSX for editing
  > python FF16SubsOrganizer.py to-xlsx -l "C:\path\\to\\folder\\0001.en.XML" -j "C:\path\\to\\folder\\0001.ja\\nxd\\txt" [-o "C:\custom\path\\to\\file.xlsx"]
  
  # Apply translations from XLSX back to XML
  > python FF16SubsOrganizer.py edit-xml -f "file.xlsx" -col I2 -l "C:\path\\to\\folder\\0001.en"
        """)
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

    args = parser.parse_args()

    if args.command == "to-html":
        print(f"> Exporting to HTML: {args.output}")
        table_rows = collect_table(args.language, args.japanese)
        export_html(table_rows, args.output)
    elif args.command == "to-xlsx":
        print(f"> Exporting to XLSX: {args.output}")
        table_rows = collect_table(args.language, args.japanese)
        export_xlsx(table_rows, args.output)
    elif args.command == "edit-xml":
        print(f"> Applying translations from: {args.file}")
        edit_xml(args.file, args.col, args.language)

if __name__ == "__main__":
    main()