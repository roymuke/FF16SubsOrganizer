# Python port of 'PzdFile.cs' from FF16Converter by KillzXGaming.
# https://github.com/KillzXGaming/FF16Converter/blob/main/PZDF/PzdFile.cs
from __future__ import annotations
from dataclasses import dataclass
from typing import BinaryIO, Dict, List, Tuple
import struct, xml.etree.ElementTree as ET

_HEADER_FORMAT = "<4s11I"
HEADER_SIZE = struct.calcsize(_HEADER_FORMAT)# 48

@dataclass
class Header:
    magic: bytes = b"PZDF"
    version: int = 2
    padding1: int = 0
    padding2: int = 0
    padding3: int = 0
    padding4: int = 0
    padding5: int = 0
    padding6: int = 0
    text_content_offset: int = 0
    text_content_count: int = 0
    padding7: int = 0
    padding8: int = 0

    def pack(self) -> bytes:
        return struct.pack(
            _HEADER_FORMAT,
            self.magic.ljust(4, b"\x00")[:4],
            self.version,
            self.padding1,
            self.padding2,
            self.padding3,
            self.padding4,
            self.padding5,
            self.padding6,
            self.text_content_offset,
            self.text_content_count,
            self.padding7,
            self.padding8,
        )

    @classmethod
    def unpack(cls, data: bytes) -> "Header":
        vals = struct.unpack(_HEADER_FORMAT, data)
        return cls(magic=vals[0], version=vals[1], padding1=vals[2], padding2=vals[3],
                    padding3=vals[4], padding4=vals[5], padding5=vals[6], padding6=vals[7],
                    text_content_offset=vals[8], text_content_count=vals[9],
                    padding7=vals[10], padding8=vals[11])

@dataclass
class TextContent:
    id: int = 0
    unknown1: int = 0
    character_id: int = 0
    subtitle_type: int = 0
    unknown4: int = 0
    message: str = ""
    voice: str = ""
    string: str = ""

@dataclass
class NexSerialization:
    struct_name: str = ""
    size: int = 0

# Binary helpers
def _read_u32(f: BinaryIO) -> int:
    return struct.unpack("<I", f.read(4))[0]

def _write_u32(f: BinaryIO, value: int) -> None:
    f.write(struct.pack("<I", value & 0xFFFFFFFF))

def _write_i32(f: BinaryIO, value: int) -> None:
    f.write(struct.pack("<i", value))

def _read_cstring(f: BinaryIO) -> str:
    chunk = bytearray()
    while True:
        b = f.read(1)
        if not b or b == b"\x00":
            break
        chunk += b
    return chunk.decode("utf-8")

def _get_string(f: BinaryIO, pos_start: int) -> str:
    offset = _read_u32(f) + pos_start
    cur = f.tell()
    f.seek(offset)
    s = _read_cstring(f)
    f.seek(cur)
    return s

def _align(f: BinaryIO, alignment: int) -> None:
    pad = (-f.tell()) % alignment
    if pad:
        f.write(b"\x00" * pad)

# PzdFile
class PzdFile:
    def __init__(self, path_or_stream=None):
        self.text_contents: List[TextContent] = []
        self.serialization: List[NexSerialization] = []
        self.header: Header = Header()
        if path_or_stream is None: return
        if isinstance(path_or_stream, (str,)):
            with open(path_or_stream, "rb") as f:
                self._read(f)
        else:
            self._read(path_or_stream)

    def _read(self, f: BinaryIO) -> None:
        f.seek(0)
        self.header = Header.unpack(f.read(HEADER_SIZE))
        f.seek(self.header.text_content_offset)
        for _ in range(self.header.text_content_count):
            start_pos = f.tell()
            tc = TextContent()
            tc.id = _read_u32(f)
            tc.message = _get_string(f, start_pos)
            tc.unknown1 = _read_u32(f)
            tc.character_id = _read_u32(f)
            tc.voice = _get_string(f, start_pos)
            tc.subtitle_type = _read_u32(f)
            tc.string = _get_string(f, start_pos)
            tc.unknown4 = _read_u32(f)
            self.text_contents.append(tc)
        pos = f.tell()
        serialize_offset = _read_u32(f)
        serialize_count = _read_u32(f)
        for _ in range(6): 
            _read_u32(f)
        f.seek(pos + serialize_offset)
        for _ in range(serialize_count):
            start_pos = f.tell()
            ns = NexSerialization()
            ns.struct_name = _get_string(f, start_pos)
            ns.size = _read_u32(f)
            self.serialization.append(ns)

    def save(self, path_or_stream) -> None:
        if isinstance(path_or_stream, str):
            with open(path_or_stream, "wb") as f:
                self._write(f)
        else:
            self._write(path_or_stream)

    def _write(self, f: BinaryIO) -> None:
        self.header.magic = b"PZDF"
        self.header.text_content_offset = HEADER_SIZE
        self.header.text_content_count = len(self.text_contents)
        saved_strings: Dict[str, List[Tuple[int, int]]] = {}

        def save_string(s: str, relative_pos: int) -> None:
            saved_strings.setdefault(s, []).append((f.tell(), relative_pos))
            _write_u32(f, 0)

        f.write(self.header.pack())
        f.seek(self.header.text_content_offset)
        for tc in self.text_contents:
            relative_pos = f.tell()
            _write_u32(f, tc.id)
            save_string(tc.message, relative_pos)
            _write_u32(f, tc.unknown1)
            _write_u32(f, tc.character_id)
            save_string(tc.voice, relative_pos)
            _write_u32(f, tc.subtitle_type)
            save_string(tc.string, relative_pos)
            _write_u32(f, tc.unknown4)
        serialize_pos = f.tell()
        _write_u32(f, 32)
        _write_i32(f, len(self.serialization))
        for _ in range(6):
            _write_u32(f, 0)
        for ns in self.serialization:
            relative_pos = f.tell()
            save_string(ns.struct_name, relative_pos)
            _write_u32(f, ns.size)
        _align(f, 16)
        for text, locations in saved_strings.items():
            target = f.tell()
            for abs_ofs, relative_pos in locations:
                cur = f.tell()
                f.seek(abs_ofs)
                _write_u32(f, target - relative_pos)
                f.seek(cur)
            f.write(text.encode("utf-8"))
            f.write(b"\x00")
        _align(f, 4)
        sect_ofs = serialize_pos - f.tell()
        f.write(b"BVLD")
        _write_i32(f, 1)
        _write_i32(f, sect_ofs)
        _write_i32(f, 0)

    # PZD->XML
    def to_xml(self) -> str:
        root = ET.Element("PzdFile")
        root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
        root.set("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
        contents_el = ET.SubElement(root, "TextContents")
        for tc in self.text_contents:
            tc_el = ET.SubElement(contents_el, "TextContent")
            tc_el.set("ID", str(tc.id))
            tc_el.set("Unknown1", str(tc.unknown1))
            tc_el.set("CharacterID", str(tc.character_id))
            tc_el.set("SubtitleType", str(tc.subtitle_type))
            tc_el.set("Unknown4", str(tc.unknown4))
            ET.SubElement(tc_el, "Message").text = tc.message
            ET.SubElement(tc_el, "Voice").text = tc.voice
            ET.SubElement(tc_el, "String").text = tc.string
        ser_el = ET.SubElement(root, "Serialization")
        for ns in self.serialization:
            ns_el = ET.SubElement(ser_el, "NexSerialization")
            ET.SubElement(ns_el, "Struct").text = ns.struct_name
            ET.SubElement(ns_el, "Size").text = str(ns.size)
        hdr = ET.SubElement(root, "PzdHeader")
        ET.SubElement(hdr, "Version").text = str(self.header.version)
        ET.SubElement(hdr, "TextContentCount").text = str(self.header.text_content_count)
        ET.SubElement(hdr, "FF16SubsOrganizer").text = "1"
        ET.indent(root, space="  ", level=0)
        body = ET.tostring(root, encoding="unicode", xml_declaration=True).splitlines()
        body[0] = '<?xml version="1.0" encoding="utf-8"?>'
        return "\n".join(body)

    # XML->PZD
    def from_xml(self, xml_text: str) -> None:
        self.text_contents = []; self.serialization = []
        root = ET.fromstring(xml_text)
        hdr_el = root.find("PzdHeader")
        if hdr_el is not None:
            self.header.version = int(hdr_el.find("Version").text)
            self.header.text_content_count = int(hdr_el.find("TextContentCount").text)
        contents_el = root.find("TextContents")
        for tc_el in (contents_el if contents_el is not None else []):
            character_id = tc_el.get("CharacterID", tc_el.get("Unknown2", "0"))
            subtitle_type = tc_el.get("SubtitleType", tc_el.get("Unknown3", "0"))
            self.text_contents.append(TextContent(
                id=int(tc_el.get("ID", "0")),
                unknown1=int(tc_el.get("Unknown1", "0")),
                character_id=int(character_id),
                subtitle_type=int(subtitle_type),
                unknown4=int(tc_el.get("Unknown4", "0")),
                message=(tc_el.findtext("Message") or ""),
                voice=(tc_el.findtext("Voice") or ""),
                string=(tc_el.findtext("String") or "")))
        ser_el = root.find("Serialization")
        for ns_el in (ser_el if ser_el is not None else []):
            self.serialization.append(NexSerialization(
                struct_name=(ns_el.findtext("Struct") or ""),
                size=int(ns_el.findtext("Size") or "0")))

def convert_file(path, verbose):
    import os
    if path.endswith(".pzd.xml"):
        directory = os.path.dirname(path) or "."
        filename = os.path.basename(path)
        out_path = os.path.join(directory, f"{filename}RB.pzd")
        pzd = PzdFile()
        pzd.from_xml(open(path, encoding="utf-8").read())
        pzd.save(out_path)
        if verbose: print(f" \033[38;5;75m[INFO]\033[00m {os.path.basename(path)}  ->  {os.path.basename(out_path)}")
    elif path.endswith(".pzd"):
        out_path = path + ".xml"
        pzd = PzdFile(path)
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(pzd.to_xml())
        if verbose: print(f" \033[38;5;75m[INFO]\033[00m {os.path.basename(path)}  ->  {os.path.basename(out_path)} ({len(pzd.text_contents)} entries)")
    else:
        print(f"  \033[90m[SKIP] Unsupported extension: {path}\033[00m")

if __name__ == "__main__":
    import sys, os
    if len(sys.argv) < 2:
        print("Usage: python ff16pzd.py <file1.pzd> [file2.pzd] ...")
        print("       python ff16pzd.py <file1.pzd.xml> [file2.pzd.xml] ...")
        print("Example:\n> python ff16pzd.py \"C:\\subs\\cutm\\cutm1000200.en.pzd\" \"C:\\subs\\cutm\\cutm1000800.en.pzd\"")
        sys.exit(1)
    files = sys.argv[1:]
    print(f"> Processing {len(files)} file(s)...")
    for path in files:
        try:
            convert_file(path, True)
        except Exception as e:
            print(f"  \033[91m[ERROR]\033[00m {os.path.basename(path)}: {e}")
    print(" [DONE]")