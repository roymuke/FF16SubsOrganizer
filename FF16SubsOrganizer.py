import os, sys, argparse
import xml.etree.ElementTree as ET

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE_DIR)

def get_ids():
    import json
    try:
        jfile = os.path.join(BASE_DIR, "IDs.json")
        with open(jfile,"r",encoding="utf-8") as f:
            jsonIds = json.load(f)
    except Exception as e:
        print(f" \033[91m[ERROR]\033[00m {e}")
    return jsonIds["characters"], jsonIds["subtitleID"]

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
            chara_id = text_content.get("CharacterID", text_content.get("Unknown2", ""))
            subtype = text_content.get("SubtitleType", text_content.get("Unknown3", ""))
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

def write_xml(tree, path, builtin):
    try:
        root = tree.getroot()
        fix_xml_fields(root)
        xml_body = ET.tostring(root, encoding="unicode", method="xml")
        if not builtin:
            xml_content = '<?xml version="1.0" encoding="utf-16"?>\r\n' + xml_body
        else:
            xml_content = '<?xml version="1.0" encoding="utf-8"?>\n' + xml_body
        with open(path, "w", encoding="utf-8", newline="") as f:
            f.write(xml_content)
    except Exception as e:
        print(f" \033[91m[ERROR]\033[00m {e}")

def collect_table(lang_root, jap_root):
    table_rows = []
    characters, subtitleID = get_ids()
    for root_dir, _, files in os.walk(lang_root):
        subdir = os.path.basename(root_dir)
        for file in files:
            if not file.endswith(".xml"):
                continue
            lang_path = os.path.join(root_dir, file)
            rel_path = os.path.relpath(lang_path, lang_root)
            jap_path = os.path.join(jap_root, rel_path[:-10] + "ja" + rel_path[-8:])
            if not os.path.exists(jap_path):
                print(f" \033[38;5;214m[WARNING]\033[00m Japanese path not found: {jap_path}")
                continue
            lang_data = read_texts(lang_path)
            jap_data = read_texts(jap_path)
            filename = os.path.splitext(os.path.splitext(file)[0])[0]
            for idx, (id_msg, chara_id, subtype, en_msg) in enumerate(lang_data):
                jp_msg = jap_data[idx][3] if idx < len(jap_data) else ""
                chara_name = characters.get(chara_id, "")
                sub_type = subtitleID.get(subtype, "")
                table_rows.append((subdir, filename, id_msg, sub_type, chara_name, chara_id, en_msg, jp_msg))
    return table_rows

def file_counter(lang_path, ext):
    cFiles = []
    for dir, _, files in os.walk(lang_path):
        for file in files:
            if file.endswith(ext):
                cFiles.append(os.path.join(dir, file))
    return cFiles

# Command: to-xlsx
def export_xlsx(table_rows, output, verbose):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.formatting.rule import FormulaRule, DataBarRule
    from collections import Counter, defaultdict
    from openpyxl.workbook.defined_name import DefinedName
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.comments import Comment
    rows_by_subdir = defaultdict(list); stats = list(); batch_filename = list()
    print("> Generating file...")
    try:
        for row in table_rows:
            subdir = row[0]; rows_by_subdir[subdir].append(row)
        wb = Workbook()
        stats_ws = wb.active; stats_ws.title = "STATS"
        note = Comment(text="""The 'Status' column (K) has the current tags available:
  · Accurate: when a line doesn't need translation.
  · Format: when a line is accurate but differs slightly from how it should be (generally punctuations).
  · Check: when a line needs revisions.
  · Improve: when a line has alredy been revised and seems fine, but the translation can be improved.
  · Revised: when a line has been revised (usually in lines that were marked as 'Check').
  · Omit: when a line won't be translated (mostly for SFXs since they aren't displayed, and used for translation progress stats).
  · Sin: when a localization line was too badly translated (to remember which lines were butchered beyond reason).""",
            author="Roysu")
        note.width = 410; note.height = 270
        stats_ws["L3"] = "About 'statuses'"; stats_ws["L3"].comment = note
        stList = ("Accurate","Format","Check","Improve","Revised","Omit","Sin")
        stats_ws["O2"] = "Status list (do not remove!)"
        for row, status in enumerate(stList):
            stats_ws[f"O{row+3}"] = status
        stats_ws.column_dimensions["O"].hidden = True
        dn = DefinedName("Statuses",attr_text="STATS!$O$3:$O$9")
        wb.defined_names["Statuses"] = dn
        stats_ws["P2"] = "↓ Do not remove this cell!"
        stats_ws["P3"] = os.path.splitext(next(iter(rows_by_subdir.values()))[0][1])[1]
        stats_ws.column_dimensions["P"].hidden = True
        UNKfill = PatternFill(bgColor="FFE6B8B7"); SFXfill = PatternFill(patternType="gray0625", fgColor="FFB8CCE4")
        SFXtext = Font(italic=True, color="FF1F497D"); ACCfill = PatternFill(bgColor="FFEBF1DE")
        CHKfill = PatternFill(bgColor="FFFFEB9C"); CHKtext = Font(color="FF9C6500")
        IMPfill = PatternFill(bgColor="FFCCC0DA"); IMPtext = Font(color="FF403151")
        FRMfill = PatternFill(bgColor="FFF4F2CC"); OMIfill = PatternFill(patternType="lightUp", fgColor="FF0070C0")
        EQUfill = PatternFill(bgColor="FFFABF8F"); SINfill = Font(color="FF963634")
        sfxRule = FormulaRule(formula=['$D2="SFX"'], fill=SFXfill, font=SFXtext); accRule = FormulaRule(formula=['$K2="Accurate"'], fill=ACCfill)
        chkRule = FormulaRule(formula=['$K2="Check"'], fill=CHKfill, font=CHKtext); impRule = FormulaRule(formula=['$K2="Improve"'], fill=IMPfill, font=IMPtext)
        frmRule = FormulaRule(formula=['$K2="Format"'], fill=FRMfill); omiRule = FormulaRule(formula=['$K2="Omit"'], fill=OMIfill)
        unkRule = FormulaRule(formula=['$E2=""'], fill=UNKfill); equRule = FormulaRule(formula=['$I2=$G2'], fill=EQUfill)
        sinRule = FormulaRule(formula=['$K2="Sin"'], font=SINfill)
        for subdir, rows in rows_by_subdir.items():
            if verbose:
                print(f" \033[38;5;75m[INFO]\033[00m Processing: {subdir}")
            stats.append([subdir,"!D2:D"+ str(len(rows)+1),"!I2:I"+ str(len(rows)+1),"!K2:K"+ str(len(rows)+1),"!M2:M"+ str(len(rows)+1)])
            sheet_name = subdir[:31] if len(subdir) <= 31 else subdir[:28] + "..."
            sheet_name = sheet_name.replace("/", "_").replace("\\", "_").replace("[", "_").replace("]", "_").replace("*", "_").replace("?", "_").replace(":", "_")
            ws = wb.create_sheet(title=sheet_name)
            ws.append(["Folder", "Filename", "ID", "Sub Type", "Character", "Character ID", "Original Text", "Japanese", "Retranslation", "Dialogue format", "Status", "Comment", "Quest"])
            ws.freeze_panes = "A2"
            ws.column_dimensions["G"].width = 45; ws.column_dimensions["H"].width = 60; ws.column_dimensions["I"].width = 59; ws.column_dimensions["J"].width = 30
            ws.column_dimensions["K"].width = 11.42; ws.column_dimensions["L"].width = 11.42; ws.column_dimensions["M"].width = 11.42
            ws.sheet_view.zoomScale = 80
            for subdir_name, filename, msg_id, subtype, chara_name, chara_id, en_text, jp_text in rows:
                ws.append([subdir_name, os.path.splitext(filename)[0], msg_id, subtype, chara_name, chara_id, en_text, jp_text, ""])
                batch_filename.append(os.path.splitext(filename)[0])
            ws.column_dimensions["B"].width = (len(str(ws["B2"].value)) + 0.5) * 1.1207692307692307
            row_filename_len = Counter(batch_filename); start_row = 2
            for i, count in enumerate(row_filename_len.items()):
                if i % 2 == 0:
                    for j in range(start_row, start_row + count[1]):
                        for col_num in range(1,12):
                            cell = ws.cell(row=j, column=col_num)
                            cell.fill = PatternFill(fill_type="solid", start_color="FFF2F2F2")
                else:
                    start_row = start_row + count[1]
                    continue
                start_row = start_row + count[1]
            for cRow in range(2,len(batch_filename) + 2):
                ws[f"J{cRow}"] = f'=CONCATENATE(IF(E{cRow}<>"",E{cRow},"Unknown"),": ",H{cRow})'
            lastRow = len(batch_filename)+1
            dv = DataValidation(type="list",formula1="Statuses",allow_blank=True)
            ws.add_data_validation(dv); dv.add(f"K2:K{lastRow}")
            ws.conditional_formatting.add(f"A2:K{lastRow}",sinRule); ws.conditional_formatting.add(f"I2:I{lastRow}",equRule)
            ws.conditional_formatting.add(f"E2:F{lastRow}",unkRule); ws.conditional_formatting.add(f"I2:I{lastRow}",omiRule)
            ws.conditional_formatting.add(f"A2:K{lastRow}",frmRule); ws.conditional_formatting.add(f"A2:K{lastRow}",impRule)
            ws.conditional_formatting.add(f"A2:K{lastRow}",chkRule); ws.conditional_formatting.add(f"A2:K{lastRow}",accRule)
            ws.conditional_formatting.add(f"A2:K{lastRow}",sfxRule)
            batch_filename = list()
        stats_ws["A1"] = "FFXVI Subtitle Translation Progress Sheet"; stats_ws["A1"].font = Font(size="22"); stats_ws.merge_cells("A1:D1")
        stats_ws["A2"] = "Made with FF16SubsOrganizer"; stats_ws["A2"].font = Font(size="10"); stats_ws["A2"].alignment = Alignment(horizontal="right")
        stats_ws.merge_cells("A2:D2"); stats_ws.column_dimensions["A"].width = 12; stats_ws.column_dimensions["B"].width = 11.85
        stats_ws.column_dimensions["C"].width = 10; stats_ws.column_dimensions["D"].width = 50; len_sheets = 9 + len(wb.sheetnames)
        sidequests = "="; ifSidequests = ["=","=","="]
        i, j, s = 4, 11, 6
        for _, row_data in enumerate(stats):
            for columna, valor in enumerate(row_data):
                stats_ws.cell(row=j, column=columna + 1, value=valor)
            j += 1
        for row in range(11,len_sheets+1):
            if stats_ws[f"A{row}"].value == "cutq":
                sidequests+=f'SUMPRODUCT(COUNTIF(INDIRECT(A{row}&E{row}),"Subquest*"))+\n'
                ifSidequests[0] += f'SUMPRODUCT(COUNTIFS(INDIRECT(A{row}&B{row}),"Normal",INDIRECT(A{row}&E{row}),"Subquest*"))+\n'
                ifSidequests[1] += f'SUMPRODUCT(COUNTIFS(INDIRECT(A{row}&B{row}),"SFX",INDIRECT(A{row}&E{row}),"Subquest*"))+\n'
                ifSidequests[2] += f'SUMPRODUCT(COUNTIFS(INDIRECT(A{row}&B{row}),"Hidden",INDIRECT(A{row}&E{row}),"Subquest*"))+\n'
            if stats_ws[f"A{row}"].value == "partyq":
                sidequests+=f'SUMPRODUCT(COUNTIF(INDIRECT(A{row}&E{row}),"Subquest*"))+\n'
                ifSidequests[0] += f'SUMPRODUCT(COUNTIFS(INDIRECT(A{row}&B{row}),"Normal",INDIRECT(A{row}&E{row}),"Subquest*"))+\n'
                ifSidequests[1] += f'SUMPRODUCT(COUNTIFS(INDIRECT(A{row}&B{row}),"SFX",INDIRECT(A{row}&E{row}),"Subquest*"))+\n'
                ifSidequests[2] += f'SUMPRODUCT(COUNTIFS(INDIRECT(A{row}&B{row}),"Hidden",INDIRECT(A{row}&E{row}),"Subquest*"))+\n'
            if stats_ws[f"A{row}"].value == "simpleq":
                sidequests+=f'SUMPRODUCT(COUNTIF(INDIRECT(A{row}&E{row}),"Subquest*"))'
                ifSidequests[0] += f'SUMPRODUCT(COUNTIFS(INDIRECT(A{row}&B{row}),"Normal",INDIRECT(A{row}&E{row}),"Subquest*"))'
                ifSidequests[1] += f'SUMPRODUCT(COUNTIFS(INDIRECT(A{row}&B{row}),"SFX",INDIRECT(A{row}&E{row}),"Subquest*"))'
                ifSidequests[2] += f'SUMPRODUCT(COUNTIFS(INDIRECT(A{row}&B{row}),"Hidden",INDIRECT(A{row}&E{row}),"Subquest*"))'
        if sidequests.endswith("+\n"): sidequests = sidequests[:-2]
        if ifSidequests[0].endswith("+\n"): ifSidequests[0] = ifSidequests[0][:-2]
        if ifSidequests[1].endswith("+\n"): ifSidequests[1] = ifSidequests[1][:-2]
        if ifSidequests[2].endswith("+\n"): ifSidequests[2] = ifSidequests[2][:-2]
        if sidequests == "=": sidequests = ""; ifSidequests = ["","",""]
        progress = [['Subtitle' 'Type','Lines','Translated','Progress','','Accurate','To Check','Mainquest','Subquest','To Improve','Omitted'],
            ['Normal',f'=SUMPRODUCT(COUNTIF(INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Normal"))',
             f'=B5-SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Normal",INDIRECT(A11:A{len_sheets}&C11:C{len_sheets}),""))','=C5/B5','',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Accurate",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Normal"))',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Check",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Normal"))',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Normal",INDIRECT(A11:A{len_sheets}&E11:E{len_sheets}),"Mainquest:*"))',
             ifSidequests[0],
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Improve",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Normal"))',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Omit",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Normal"))'],
            ['SFX',f'=SUMPRODUCT(COUNTIF(INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"SFX"))',
             f'=B6-SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"SFX",INDIRECT(A11:A{len_sheets}&C11:C{len_sheets}),""))','=C6/B6','',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Accurate",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"SFX"))',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Check",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"SFX"))',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"SFX",INDIRECT(A11:A{len_sheets}&E11:E{len_sheets}),"Mainquest:*"))',
             ifSidequests[1],
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Improve",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"SFX"))',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Omit",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"SFX"))'],
            ['Hidden',f'=SUMPRODUCT(COUNTIF(INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Hidden"))',
             f'=B7-SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Hidden",INDIRECT(A11:A{len_sheets}&C11:C{len_sheets}),""))','=C7/B7','',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Accurate",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Hidden"))',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Check",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Hidden"))',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Hidden",INDIRECT(A11:A{len_sheets}&E11:E{len_sheets}),"Mainquest:*"))',
             ifSidequests[2],
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Improve",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Hidden"))',
             f'=SUMPRODUCT(COUNTIFS(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Omit",INDIRECT(A11:A{len_sheets}&B11:B{len_sheets}),"Hidden"))'],
            ['TOTAL','=SUM(B5:B7)','=SUM(C5:C7)','=(C8+F8+K8)/B8','','=SUM(F5:F7)','=SUM(G5:G7)',f'=SUMPRODUCT(COUNTIF(INDIRECT(A11:A{len_sheets}&E11:E{len_sheets}),"Mainquest:*"))',
             f"{sidequests}",'=SUM(J5:J7)','=SUM(K5:K7)']]
        barRule = DataBarRule(start_type="num",start_value=0,end_type="num",end_value=1,color="FF63C384",minLength=0,maxLength=100)
        stats_ws.conditional_formatting.add("D5:D9", barRule)
        stats_ws.conditional_formatting.add("H9:I9", barRule)
        stats_ws["C9"] = "Main"; stats_ws["C9"].alignment = Alignment(horizontal='right'); stats_ws["D9"] = "=1-(G8/C8)"
        for row in range(5,10):
            stats_ws[f"D{row}"].number_format = "0.00%"
        for _, row_data in enumerate(progress):
            for columna, valor in enumerate(row_data):
                stats_ws.cell(row=i, column=columna + 1, value=valor)
            i += 1
        styles = [PatternFill(fill_type="solid", start_color="FFEBF1DE"),PatternFill(fill_type="solid", start_color="FFFFEB9C"),
                  PatternFill(fill_type="solid", start_color="FFC0504D"),PatternFill(fill_type="solid", start_color="FF9BBB59"),
                  PatternFill(fill_type="solid", start_color="FFCCC0DA"),PatternFill(fill_type="solid", start_color="FFF2F2F2")]
        for style in styles:
            stats_ws.cell(row=4,column=s).fill = style
            s += 1
        for col in range(1,12):
            if col == 5: continue
            if col < 6:
                stats_ws.cell(row=4,column=col).font = Font(bold=True)
            stats_ws.cell(row=8,column=col).font = Font(bold=True)
        stats_ws["G4"].font = Font(color="FF9C6500"); stats_ws["H4"].font = stats_ws["I4"].font = Font(color="FFFFFFFF")
        stats_ws["J4"].font = Font(color="FF403151"); stats_ws["K4"].font = Font(color="FF3F3F3F",bold=True)
        stats_ws["F10"]= "Revised:"; stats_ws["G10"] = f'=SUMPRODUCT(COUNTIF(INDIRECT(A11:A{len_sheets}&D11:D{len_sheets}),"Revised"))+J8'
        for row_num in range(11, len_sheets + 1):
            stats_ws.row_dimensions[row_num].hidden = True
        sheet_progress = len_sheets + 3; diffe = len(stats)+2; name = 0
        for row in range(sheet_progress,sheet_progress+len(stats)):
            stats_ws[f"A{row}"] = stats[name][0]
            name += 1
        for row in range(sheet_progress,sheet_progress+len(stats)):
            stats_ws[f"B{row}"] = f'=SUMPRODUCT(COUNTIF(INDIRECT(A{row-diffe}&C{row-diffe}),""))+C{row}'
            stats_ws[f"C{row}"] = f'=SUMPRODUCT(COUNTIF(INDIRECT(A{row-diffe}&C{row-diffe}),"*"))'
            stats_ws[f"D{row}"] = f'=(C{row}+F{row}+SUMPRODUCT(COUNTIF(INDIRECT(A{row-diffe}&D{row-diffe}),"Omit"))-G{row})/B{row}'
            stats_ws[f"D{row}"].number_format = "0.00%"
            stats_ws[f"F{row}"] = f'=SUMPRODUCT(COUNTIF(INDIRECT(A{row-diffe}&D{row-diffe}),"Accurate"))'
            stats_ws[f"G{row}"] = f'=SUMPRODUCT(COUNTIF(INDIRECT(A{row-diffe}&D{row-diffe}),"Check"))'
            stats_ws[f"H{row}"] = f'=SUMPRODUCT(COUNTIFS(INDIRECT(A{row-diffe}&C{row-diffe}),"*",INDIRECT(A{row-diffe}&E{row-diffe}),"Mainquest:*"))'
            stats_ws[f"I{row}"] = f'=SUMPRODUCT(COUNTIFS(INDIRECT(A{row-diffe}&C{row-diffe}),"*",INDIRECT(A{row-diffe}&E{row-diffe}),"Subquest:*"))'
        bbarRule = DataBarRule(start_type="num",start_value=0,end_type="num",end_value=1,color="FF008AEF",minLength=0,maxLength=100)
        stats_ws.conditional_formatting.add(F"D{sheet_progress}:D{sheet_progress+len(stats)-1}", bbarRule)
        for row in range(sheet_progress+len(stats)+2,sheet_progress+(len(stats)*2)+2):
            stats_ws[f"H{row}"] = f'=SUMPRODUCT(COUNTIFS(INDIRECT(A{row-(diffe*2)}&D{row-(diffe*2)}),"Check",INDIRECT(A{row-(diffe*2)}&E{row-(diffe*2)}),"Mainquest:*"))'
            stats_ws[f"I{row}"] = f'=SUMPRODUCT(COUNTIFS(INDIRECT(A{row-(diffe*2)}&D{row-(diffe*2)}),"Check",INDIRECT(A{row-(diffe*2)}&E{row-(diffe*2)}),"Subquest:*"))'
            stats_ws.row_dimensions[row].hidden = True
        stats_ws["H9"] = f'=SUMPRODUCT(H{sheet_progress}:H{sheet_progress+len(stats)-1}-H{sheet_progress+len(stats)+2}:H{sheet_progress+(len(stats)*2)+1})/SUM(H{sheet_progress}:H{sheet_progress+len(stats)-1})'
        stats_ws["I9"] = f'=SUMPRODUCT(I{sheet_progress}:I{sheet_progress+len(stats)-1}-I{sheet_progress+len(stats)+2}:I{sheet_progress+(len(stats)*2)+1})/SUM(I{sheet_progress}:I{sheet_progress+len(stats)-1})'
        stats_ws["H9"].number_format = stats_ws["I9"].number_format = "0.00%"
        wb.save(output)
        print(f" \033[38;5;76m[DONE]\033[00m XLSX file generated in: \033[48;5;235m{output}\033[00m")
        print(f" \033[38;5;81m[INSTRUCTION] Edit the 'Retranslation' column (I) on each sheet. Once done, use 'edit-xml' to apply changes.\033[00m")
    except Exception as e:
        print(f" \033[91m[ERROR]\033[00m Couldn't generate file: {e}")

# Command: edit-xml
def edit_xml(xlsx_path, col_reference, lang_root, verbose):
    from openpyxl import load_workbook
    from openpyxl.utils import column_index_from_string
    from html import unescape
    try:
        wb = load_workbook(xlsx_path)
        all_sheets = {}
        lan = wb["STATS"]["P3"].value
        for sheet_name in wb.sheetnames[1:]:
            sheet = wb[sheet_name]
            sheet_data = []
            for row in sheet.iter_rows(min_row=2,values_only=True):
                sheet_data.append(list(row))
            all_sheets[sheet_name] = sheet_data
        col_letter = "".join(filter(str.isalpha, col_reference.upper()))
        col_idx = column_index_from_string(col_letter)
        changes_made = 0
        files_processed = set()
        print("> Processing translations...")
        for row, data in all_sheets.items():
            for item in data:
                if item[col_idx - 1] is not None:
                    subdir = item[0] if item[0] else ""
                    filename = item[1] if item[1] else ""
                    msg_id = str(item[2]) if item[2] else ""
                    new_translation = item[col_idx - 1] if item[col_idx - 1] else ""
                    xml_filename = f"{filename}{lan}.pzd.xml"
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
                                            if verbose: print(f" \033[90m[SKIP] Message {msg_id} already translated, skipping.\033[00m")
                                            continue
                                        else:
                                            if verbose: print(f" \033[38;5;75m[INFO]\033[00m {filename} (ID: {msg_id}): \033[38;5;210m\"{old_text}\"\033[00m -> \033[38;5;81m\"{new_translation.strip()}\"\033[00m")
                                            changes_made += 1
                                            files_processed.add(xml_path)
                                        break
                        conv_by = False
                        if root.find("PzdHeader").find("FF16SubsOrganizer") is not None:
                            conv_by = True
                        tree = ET.ElementTree(root)
                        write_xml(tree, xml_path, conv_by)
                    except Exception as e:
                        print(f" \033[38;5;214m[WARNING]\033[00m Could not process {xml_path}: {e}")
                        continue
        print(f"\n \033[38;5;76m[DONE]\033[00m Summary:")
        print(f"   · {changes_made} translations applied.")
        print(f"   · {len(files_processed)} files modified.")
    except Exception as e:
        print(f" \033[91m[ERROR]\033[00m Error reading XLSX file: {e}")

# Command: convert-batch
def convert_batch(ff16converter, builtin, lang_path, valid_ext, verbose):
    from time import perf_counter
    from collections import defaultdict
    if builtin: from ff16pzd import convert_file
    else: import subprocess
    if not os.path.exists(lang_path):
        print(f" \033[91m[ERROR]\033[00m Folder {lang_path} does not exist")
        return False
    if not builtin:
        if not os.path.exists(ff16converter):
            print(f" \033[91m[ERROR]\033[00m Converter {ff16converter} does not exist")
            return False
    print(f"> Converting files in: \033[48;5;235m{lang_path}\033[00m\n> Processing. This may take a while...")
    files_to_convert = file_counter(lang_path, valid_ext)
    if verbose: print(f" \033[38;5;75m[INFO]\033[00m {len(files_to_convert)} files to convert")
    folder_group = defaultdict(list); start_time = perf_counter()
    if valid_ext == "pzd": ext = ".xml"
    else: ext = "RB.pzd"
    for file in files_to_convert:
        if os.path.exists(file + ext):
            print(f" \033[90m[SKIP] {os.path.basename(file)} already exists, skipping.\033[00m")
            continue
        folder_group[os.path.basename(os.path.dirname(file))].append(file)
    for folder, files in folder_group.items():
        try:
            if verbose: print(f" \033[38;5;75m[INFO]\033[00m Converting files on: \033[38;5;81m{folder}\033[00m")
            if folder == "defaultq" or folder == "simpleq":
                helper = []
                for i in range(0, len(files), 400):
                    helper.append(files[i:i + 400])
                for chunk in helper:
                    if builtin:
                        for file in chunk:
                            convert_file(file, verbose)
                    else: subprocess.run([ff16converter] + [str(file) for file in chunk], capture_output=True, text=True)
            else:
                if builtin:
                    for file in files:
                        convert_file(file, verbose)
                else: subprocess.run([ff16converter] + [str(file) for file in files], capture_output=True, text=True)
        except Exception as e:
            print(f" \033[91m[ERROR]\033[00m Error converting: {e}")
    time_lapsed = perf_counter() - start_time
    print(f" \033[38;5;76m[DONE]\033[00m Files converted in {int(time_lapsed // 3600):02d}:{int((time_lapsed % 3600) // 60):02d}:{int(time_lapsed % 60):02d}.{int((time_lapsed % 1) * 1000):03d}")
    return True

# Command: move-batch
def move_converted(this_directory, to_directory, extension, verbose):
    from pathlib import Path
    if not os.path.exists(to_directory):
        print(f" \033[38;5;75m[INFO]\033[00m Directory \033[48;5;235m{to_directory}\033[00m not found\n> Creating directory {to_directory}...")
    Path.mkdir(Path(to_directory), parents=True, exist_ok=True)
    print(f"> Moving files to: \033[48;5;235m{to_directory}\033[00m\n> Processing. This may take a while...")
    match extension:
        case ".xml":
            converted_files = file_counter(this_directory, "RB.pzd")
            if not converted_files:
                print(f" \033[90m[SKIP] Converted files not found\033[00m")
                return
            for file in converted_files:
                try:
                    folder = os.path.basename(os.path.dirname(file))
                    original_name = os.path.basename(file).replace(".pzd.xmlRB.pzd",".pzd")
                    destination_file = os.path.join(to_directory, folder, original_name)
                    Path.mkdir(Path(os.path.join(to_directory, folder)), parents=True, exist_ok=True)
                    os.replace(file, destination_file)
                    if verbose: print(f" \033[38;5;75m[INFO]\033[00m Moved: \033[38;5;81m{os.path.basename(file)}\033[00m to \033[48;5;235m{os.path.dirname(destination_file)}\033[00m")
                except Exception as e:
                    print(f" \033[91m[ERROR]\033[00m Error moving {os.path.basename(file)}: {e}")
        case ".pzd":
            converted_files = file_counter(this_directory, ".pzd.xml")
            if not converted_files:
                print(f" \033[90m[SKIP] Converted files not found\033[00m")
                return
            for file in converted_files:
                try:
                    destination_folder = os.path.join(to_directory, os.path.basename(os.path.dirname(file)))
                    destination_file = os.path.join(destination_folder, os.path.basename(file))
                    Path.mkdir(Path(destination_folder), parents=True, exist_ok=True)
                    os.replace(str(file), str(destination_file))
                    if verbose: print(f" \033[38;5;75m[INFO]\033[00m Moved: \033[38;5;81m{os.path.basename(file)}\033[00m to \033[48;5;235m{destination_folder}\033[00m")
                except Exception as e:
                    print(f" \033[91m[ERROR]\033[00m Error moving {os.path.basename(file)}: {e}")
    print(f" \033[38;5;76m[DONE]\033[00m Move operation completed.")
    return True

def main():
    os.system("color")
    creditos = """\033[38;5;81m
 +----------------------------------------------+
 | FFXVI Subtitle Organizer v1.5                |
 | by Roysu                                     |
 +----------------------------------------------+
 | https://github.com/roymuke/FF16SubsOrganizer |
 +----------------------------------------------+\033[00m"""
    creditos = creditos.replace("\n","",1)
    parser = argparse.ArgumentParser(
        description=creditos,
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
examples:
  \033[90m# Export subtitles to XLSX for editing\033[00m
  > python \033[38;5;149mFF16SubsOrganizer.py\033[00m to-xlsx \033[38;5;149m-l\033[00m "C:\\modding\\0007.en.XML" \033[38;5;149m-j\033[00m "C:\\modding\\0007.jaXML\\nxd\\txt" \033[38;5;149m-o\033[00m "C:\\modding\\ff16_subs.xlsx"

  \033[90m# Apply translations from XLSX back to XML\033[00m
  > python \033[38;5;149mFF16SubsOrganizer.py\033[00m edit-xml \033[38;5;149m-f\033[00m "C:\\modding\\ff16_subs.xlsx" \033[38;5;149m-col\033[00m I2 \033[38;5;149m-l\033[00m "C:\\modding\\0007.en.XML"

  \033[90m# Convert in batch PZD->XML or XML->PZD\033[00m
  > python \033[38;5;149mFF16SubsOrganizer.py\033[00m convert-batch \033[38;5;149m-c\033[00m "C:\\FF16Converter\\FF16Converter.exe" \033[38;5;149m-f\033[00m "C:\\modding\\0007.en\\nxd\\text" \033[38;5;149m--pzd\033[00m \033[38;5;149m-m\033[00m "C:\\modding\\0007.en.XML"

  \033[90m# Move files to another destination by extension\033[00m
  > python \033[38;5;149mFF16SubsOrganizer.py\033[00m move-batch \033[38;5;149m-f\033[00m "C:\\modding\\0007.en.XML" \033[38;5;149m--xml\033[00m \033[38;5;149m-m\033[00m "C:\\modding\\0007.en.PZD\"""")
    subparsers = parser.add_subparsers(dest="command", required=True, help="AVAILABLE COMMANDS")
    # to-xlsx command
    xlsx_parser = subparsers.add_parser("to-xlsx", help="export subtitles to XLSX file")
    xlsx_parser.add_argument("-l", "--language", required=True, help="path to selected language subtitles folder")
    xlsx_parser.add_argument("-j", "--japanese", required=True, help="path to japanese subtitles folder")
    xlsx_parser.add_argument("-o", "--output", default="ff16_subtitles.xlsx", help="path to output xlsx file")
    xlsx_parser.add_argument("-v", "--verbose", action="store_true", help="show detailed output messages")
    # edit-xml command
    edit_parser = subparsers.add_parser("edit-xml", help="gets translations from XLSX back to XML files")
    edit_parser.add_argument("-f", "--file", required=True, help="path to xlsx file")
    edit_parser.add_argument("-col", required=True, help="column with new translation (e.g. I2)")
    edit_parser.add_argument("-l", "--language", required=True, help="path to selected language to translate")
    edit_parser.add_argument("-v", "--verbose", action="store_true", help="show detailed output messages")
    # convert-batch command
    batch_parser = subparsers.add_parser("convert-batch", help="convert files to another format, pzd->xml OR xml->pzd")
    converter_group = batch_parser.add_mutually_exclusive_group(required=True)
    converter_group.add_argument("-c", "--converter", help="path to FF16Converter.exe")
    converter_group.add_argument("-b", "--builtin", action="store_true", help="use FF16SubsOrganizer built-in converter")
    batch_parser.add_argument("-f", "--folder", required=True, help="path to language folder")
    extension_group = batch_parser.add_mutually_exclusive_group(required=True)
    extension_group.add_argument("--pzd", action="store_const", const=".pzd", dest="extension", help="extension to convert (pzd -> xml)")
    extension_group.add_argument("--xml", action="store_const", const=".xml", dest="extension", help="extension to convert (xml -> pzd)")
    batch_parser.add_argument("-m", "--moveto", help="path to converted files folder destination")
    batch_parser.add_argument("-v", "--verbose", action="store_true", help="show detailed output messages")
    # move-batch command
    move_parser = subparsers.add_parser("move-batch", help="move files to another destination")
    move_parser.add_argument("-f", "--folder", required=True, help="path to folder of misplaced files")
    extension_grp = move_parser.add_mutually_exclusive_group(required=True)
    extension_grp.add_argument("--pzd", action="store_const", const=".pzd", dest="extension", help="move XML files (converted pzds)")
    extension_grp.add_argument("--xml", action="store_const", const=".xml", dest="extension", help="move PZD files (converted xmls)")
    move_parser.add_argument("-m", "--moveto", required=True, help="path to folder destination")
    move_parser.add_argument("-v", "--verbose", action="store_true", help="show detailed output messages")

    args = parser.parse_args()

    if args.command == "to-xlsx":
        print(f"> Exporting to XLSX: \033[48;5;235m{args.output}\033[00m")
        table_rows = collect_table(args.language, args.japanese)
        export_xlsx(table_rows, args.output, args.verbose)
    elif args.command == "edit-xml":
        print(f"> Applying translations from: \033[48;5;235m{args.file}\033[00m")
        edit_xml(args.file, args.col, args.language, args.verbose)
    elif args.command == "convert-batch":
        cnvtd = convert_batch(args.converter,args.builtin,args.folder,args.extension, args.verbose)
        if args.moveto and args.extension and cnvtd:
            move_converted(args.folder, args.moveto, args.extension, args.verbose)
    elif args.command == "move-batch":
        move_converted(args.folder, args.moveto, args.extension, args.verbose)

if __name__ == "__main__":
    main()