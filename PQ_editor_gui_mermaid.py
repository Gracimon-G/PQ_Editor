"""
PQ_editor_gui_mermaid.py — Read & edit Power Query (M) code without Excel open or freezing.

Slogan: Unchained M Engine · Zero Locks, Zero Freezes
"""

from __future__ import annotations

import base64
import html
import io
import os
import re
import shutil
import struct
import tempfile
import urllib.parse
import uuid
import webbrowser
import xml.etree.ElementTree as ET
import zipfile
from io import BytesIO
from pathlib import Path

_SECTION_PART = "Formulas/Section1.m"


# ============================================================ THEMES ============================================================
THEMES: dict[str, dict] = {
    "light": {
        "name": "light",
        "bg": "#F3F2F1",
        "card_bg": "#EAE8E6",
        "fg": "#201F1E",
        "header_bg": "#0E703C",
        "header_fg": "#FFFFFF",
        "header_sub": "#C8E6C9",
        "badge_bg": "#0B5630",
        "badge_fg": "#FFFFFF",
        "toolbar_bg": "#F3F2F1",
        "sidebar_bg": "#FFFFFF",
        "sidebar_fg": "#201F1E",
        "gutter_bg": "#F3F2F1",
        "gutter_fg": "#707070",
        "editor_bg": "#FFFFFF",
        "editor_fg": "#000000",
        "editor_insert": "#000000",
        "editor_select": "#ADD6FF",
        "input_bg": "#FFFFFF",
        "input_fg": "#000000",
        "border_color": "#D2D0CE",
        "btn_bg": "#E1DFDD",
        "btn_active": "#EDEBE9",
        "menu_active_bg": "#0E703C",
        "menu_active_fg": "#FFFFFF",
        "status_bg": "#F3F2F1",
        "status_fg": "#201F1E",
        "clean_fg": "#0E703C",
        "dirty_fg": "#D83B01",
        "canvas_bg": "#FFFFFF",
        "line_color": "#707070",
        "node_src_fill": "#E2EFDA",
        "node_src_outline": "#375623",
        "node_dest_fill": "#DDEBF7",
        "node_dest_outline": "#41719C",
        "node_text": "#201F1E",
        "colors": {
            "keyword": "#0000CD",
            "string": "#A31515",
            "comment": "#008000",
            "number": "#098658",
            "atname": "#795E26",
            "find_match": "#FFF2A8",
            "find_active": "#FF8C00",
        },
    },
    "dark": {
        "name": "dark",
        "bg": "#1E1E1E",
        "card_bg": "#2D2D30",
        "fg": "#D4D4D4",
        "header_bg": "#104E2D",
        "header_fg": "#FFFFFF",
        "header_sub": "#A3D9B5",
        "badge_bg": "#0B361F",
        "badge_fg": "#E0E0E0",
        "toolbar_bg": "#252526",
        "sidebar_bg": "#252526",
        "sidebar_fg": "#D4D4D4",
        "gutter_bg": "#1E1E1E",
        "gutter_fg": "#858585",
        "editor_bg": "#1E1E1E",
        "editor_fg": "#D4D4D4",
        "editor_insert": "#FFFFFF",
        "editor_select": "#264F78",
        "input_bg": "#2D2D2D",
        "input_fg": "#E0E0E0",
        "border_color": "#4A4A4C",
        "btn_bg": "#333333",
        "btn_active": "#3E3E42",
        "menu_active_bg": "#0E703C",
        "menu_active_fg": "#FFFFFF",
        "status_bg": "#252526",
        "status_fg": "#D4D4D4",
        "clean_fg": "#6A9955",
        "dirty_fg": "#E5A93C",
        "canvas_bg": "#1E1E1E",
        "line_color": "#888888",
        "node_src_fill": "#203A20",
        "node_src_outline": "#3F753F",
        "node_dest_fill": "#1B3047",
        "node_dest_outline": "#36618F",
        "node_text": "#E0E0E0",
        "colors": {
            "keyword": "#569CD6",
            "string": "#CE9178",
            "comment": "#6A9955",
            "number": "#B5CEA8",
            "atname": "#DCDCAA",
            "find_match": "#495057",
            "find_active": "#E5A93C",
        },
    }
}


# ============================================================ ENGINE ============================================================
class PowerQueryError(Exception):
    """Raised for any DataMashup read/parse/write failure."""


class ExcelPowerQuery:
    """Read and edit the Power Query M code inside an Excel workbook."""

    def __init__(self, xlsx_path: str | Path):
        self.path = Path(xlsx_path).resolve()
        if not self.path.exists():
            raise PowerQueryError(f"File not found: {self.path}")
        if self.path.suffix.lower() not in {".xlsx", ".xlsm", ".xlsb"}:
            raise PowerQueryError(f"Unsupported extension: {self.path.suffix}")
        self._item_name: str | None = None
        self._item_xml: bytes | None = None
        self._item_encoding: str = "utf-8"
        self._envelope: bytes | None = None
        self.query_groups: dict[str, str] = {}        # query_name -> group_name
        self.query_destinations: dict[str, str] = {}  # query_name -> destination description
        self._load()
        self._inspect_excel_metadata()

    # ---------- public API ----------
    def read(self) -> str:
        with zipfile.ZipFile(BytesIO(self._inner_zip_bytes())) as z:
            if _SECTION_PART not in z.namelist():
                raise PowerQueryError(f"{_SECTION_PART} not found in DataMashup package.")
            return z.read(_SECTION_PART).decode("utf-8")

    def write(self, new_m: str, out_path: str | Path | None = None) -> Path:
        out = Path(out_path).resolve() if out_path else self.path
        # 1. Automatic backup creation before modifying existing file
        if out.exists():
            bak = out.with_suffix(out.suffix + ".bak")
            try:
                shutil.copy2(out, bak)
            except PermissionError:
                raise PermissionError(f"File locked by Excel: {out.name}")
            except Exception:
                pass

        # 2. Rebuild mashup binary
        new_item_xml = self._reencode_item_xml(self._rebuild_envelope(new_m.encode("utf-8")))

        # 3. Atomic write via temporary file
        self._rewrite_workbook_atomic(out, new_item_xml)
        return out

    def backup_exists(self) -> bool:
        bak = self.path.with_suffix(self.path.suffix + ".bak")
        return bak.exists()

    def restore_backup(self) -> Path:
        bak = self.path.with_suffix(self.path.suffix + ".bak")
        if not bak.exists():
            raise PowerQueryError(f"No backup file found at: {bak}")
        try:
            shutil.copy2(bak, self.path)
        except PermissionError:
            raise PermissionError(f"File locked by Excel: {self.path.name}")
        self._load()
        self._inspect_excel_metadata()
        return self.path

    # ---------- internals ----------
    def _load(self) -> None:
        with zipfile.ZipFile(self.path) as z:
            for name in z.namelist():
                if not (name.startswith("customXml/") and name.endswith(".xml")):
                    continue
                raw = z.read(name)
                m = re.search(r"<DataMashup[^>]*>(.*?)</DataMashup>", self._decode_xml(raw), re.DOTALL)
                if not m:
                    continue
                self._item_name, self._item_xml = name, raw
                self._item_encoding = self._xml_encoding(raw)
                self._envelope = base64.b64decode(re.sub(r"\s+", "", m.group(1)))
                self._inspect_mashup_groups()
                return
        raise PowerQueryError("No Power Query (DataMashup) part found in this workbook.")

    def _sections(self) -> tuple[int, bytes, int, int]:
        buf = self._envelope
        if not buf or len(buf) < 8:
            raise PowerQueryError("Corrupt DataMashup: buffer too short.")
        version, pkg_len = struct.unpack_from("<ii", buf, 0)
        start, end = 8, 8 + pkg_len
        if pkg_len <= 0 or end > len(buf):
            raise PowerQueryError("Corrupt DataMashup: invalid package length.")
        return version, buf[start:end], start, end

    def _inner_zip_bytes(self) -> bytes:
        return self._sections()[1]

    def _inspect_mashup_groups(self) -> None:
        """Extract Query Groups (folders) from LocalPackageMetadataFile or Metadata/Section1.xml."""
        try:
            # 1. Primary: LocalPackageMetadataFile in DataMashup envelope (covers .xlsx, .xlsm, .xlsb)
            if self._envelope:
                idx = self._envelope.find(b"<LocalPackageMetadataFile")
                if idx != -1:
                    end = self._envelope.find(b"</LocalPackageMetadataFile>", idx)
                    if end != -1:
                        raw_xml = self._envelope[idx : end + len("</LocalPackageMetadataFile>")].decode("utf-8", errors="replace")
                        root = ET.fromstring(raw_xml)
                        groups_raw = None
                        for entry in root.findall(".//Entry"):
                            if entry.attrib.get("Type") == "QueryGroups":
                                val = entry.attrib.get("Value", "")
                                if val.startswith("s"):
                                    try:
                                        groups_raw = base64.b64decode(val[1:])
                                    except Exception:
                                        pass

                        for item in root.findall(".//Item"):
                            loc = item.find("ItemLocation")
                            if loc is not None:
                                itype = loc.find("ItemType")
                                ipath = loc.find("ItemPath")
                                if itype is not None and itype.text == "Formula" and ipath is not None and ipath.text:
                                    parts = ipath.text.split("/")
                                    if len(parts) == 2 and parts[0] == "Section1":
                                        qname = urllib.parse.unquote(parts[1])
                                        for entry in item.iter("Entry"):
                                            if entry.attrib.get("Type") == "QueryGroupID" and groups_raw:
                                                gid_val = entry.attrib.get("Value", "")
                                                if gid_val.startswith("s"):
                                                    try:
                                                        g_bytes = uuid.UUID(gid_val[1:]).bytes_le
                                                        g_pos = groups_raw.find(g_bytes)
                                                        if g_pos != -1:
                                                            nlen = groups_raw[g_pos + 16]
                                                            gname = groups_raw[g_pos + 17 : g_pos + 17 + nlen].decode("utf-8", errors="replace")
                                                            if gname:
                                                                self.query_groups[qname] = gname
                                                    except Exception:
                                                        pass
                        if self.query_groups:
                            return
        except Exception:
            pass

        # 2. Fallback: Metadata/Section1.xml or Config/Package.xml in inner zip
        try:
            with zipfile.ZipFile(BytesIO(self._inner_zip_bytes())) as z:
                namelist = z.namelist()
                meta_file = next((f for f in namelist if f.endswith("Metadata/Section1.xml")), None)
                if meta_file:
                    raw_xml = z.read(meta_file)
                    root = ET.fromstring(raw_xml)
                    groups = {}
                    for g in root.iter():
                        if g.tag.endswith("QueryGroup"):
                            gid = g.attrib.get("GroupId") or g.attrib.get("id")
                            gname = g.attrib.get("Name") or g.attrib.get("name")
                            if gid and gname:
                                groups[gid] = gname
                    for item in root.iter():
                        if item.tag.endswith("Item"):
                            qname = None
                            gid = None
                            for child in item:
                                if child.tag.endswith("Name"):
                                    qname = child.text
                                elif child.tag.endswith("QueryGroupID"):
                                    gid = child.text
                            if qname and gid and gid in groups:
                                self.query_groups[qname] = groups[gid]
        except Exception:
            pass

    def _inspect_excel_metadata(self) -> None:
        """Inspect LocalPackageMetadataFile, tables, sheets, and connections to determine load destinations."""
        table_to_sheet: dict[str, str] = {}
        try:
            with zipfile.ZipFile(self.path) as z:
                namelist = set(z.namelist())
                is_bin = any(f.endswith(".bin") for f in namelist if f.startswith("xl/"))

                if is_bin:
                    # --- XLSB BINARY WORKBOOK ---
                    # 1. Parse sheet rId -> sheet name from xl/workbook.bin
                    rid_to_name: dict[str, str] = {}
                    if "xl/workbook.bin" in namelist:
                        raw_wb = z.read("xl/workbook.bin")
                        pattern = rb"r\x00I\x00d\x00(?:[0-9]\x00)+"
                        for m in re.finditer(pattern, raw_wb):
                            rid = m.group(0).decode("utf-16le")
                            offset = m.end()
                            if offset + 4 <= len(raw_wb):
                                str_len, = struct.unpack_from("<I", raw_wb, offset)
                                if 0 < str_len < 200 and offset + 4 + str_len * 2 <= len(raw_wb):
                                    sheet_name = raw_wb[offset + 4 : offset + 4 + str_len * 2].decode("utf-16le", errors="replace")
                                    rid_to_name[rid] = sheet_name

                    # 2. Parse sheet rId -> sheet file from xl/_rels/workbook.bin.rels
                    sheet_target_to_name: dict[str, str] = {}
                    if "xl/_rels/workbook.bin.rels" in namelist:
                        root = ET.fromstring(z.read("xl/_rels/workbook.bin.rels"))
                        for rel in root.iter():
                            if rel.tag.endswith("Relationship"):
                                rid = rel.attrib.get("Id")
                                tgt = rel.attrib.get("Target", "")
                                if "worksheets/" in tgt and rid in rid_to_name:
                                    sheet_file = "xl/" + tgt.replace("../", "").lstrip("/")
                                    sheet_target_to_name[sheet_file] = rid_to_name[rid]

                    # 3. Map table bin to sheet name via xl/worksheets/_rels/*.bin.rels
                    table_file_to_sheet: dict[str, str] = {}
                    for fname in namelist:
                        if fname.startswith("xl/worksheets/_rels/") and fname.endswith(".bin.rels"):
                            sheet_file = "xl/worksheets/" + fname.replace("xl/worksheets/_rels/", "").replace(".bin.rels", ".bin")
                            sname = sheet_target_to_name.get(sheet_file, "Sheet")
                            try:
                                rels_root = ET.fromstring(z.read(fname))
                                for rel in rels_root.iter():
                                    if rel.tag.endswith("Relationship") and "table" in rel.attrib.get("Type", ""):
                                        tgt = rel.attrib.get("Target", "")
                                        tfile = "xl/" + tgt.replace("../", "").lstrip("/")
                                        table_file_to_sheet[tfile] = sname
                            except Exception:
                                pass

                    # 4. Extract table name from table*.bin
                    for tfile, sname in table_file_to_sheet.items():
                        if tfile in namelist:
                            raw_tbl = z.read(tfile)
                            matches = re.findall(rb"(?:[\x20-\x7e]\x00){2,}", raw_tbl)
                            if matches:
                                tbl_name = matches[0].decode("utf-16le")
                                table_to_sheet[tbl_name] = sname

                else:
                    # --- XLSX / XLSM OPENXML WORKBOOK ---
                    sheet_target_to_name: dict[str, str] = {}
                    if "xl/_rels/workbook.xml.rels" in namelist and "xl/workbook.xml" in namelist:
                        rid_to_target: dict[str, str] = {}
                        wb_rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
                        for rel in wb_rels.iter():
                            if rel.tag.endswith("Relationship"):
                                rid = rel.attrib.get("Id")
                                tgt = rel.attrib.get("Target", "")
                                if rid and tgt:
                                    if not tgt.startswith("xl/"):
                                        tgt = "xl/" + tgt.lstrip("/")
                                    rid_to_target[rid] = tgt
                        wb = ET.fromstring(z.read("xl/workbook.xml"))
                        for sheet in wb.iter():
                            if sheet.tag.endswith("sheet"):
                                sname = sheet.attrib.get("name")
                                rid = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") or sheet.attrib.get("r:id")
                                if sname and rid in rid_to_target:
                                    sheet_target_to_name[rid_to_target[rid]] = sname

                    for fname in namelist:
                        if fname.startswith("xl/worksheets/_rels/") and fname.endswith(".xml.rels"):
                            sheet_file = "xl/worksheets/" + fname.replace("xl/worksheets/_rels/", "").replace(".xml.rels", ".xml")
                            sname = sheet_target_to_name.get(sheet_file, "Sheet")
                            try:
                                rels_root = ET.fromstring(z.read(fname))
                                for rel in rels_root.iter():
                                    if rel.tag.endswith("Relationship") and "table" in rel.attrib.get("Type", ""):
                                        tgt = rel.attrib.get("Target", "")
                                        tfile = "xl/" + tgt.replace("../", "").lstrip("/")
                                        if tfile in namelist:
                                            t_root = ET.fromstring(z.read(tfile))
                                            tbl_name = t_root.attrib.get("displayName") or t_root.attrib.get("name", "Table")
                                            table_to_sheet[tbl_name] = sname
                            except Exception:
                                pass
        except Exception:
            pass

        # --- DataMashup Envelope Metadata (Primary Source of Truth) ---
        try:
            if self._envelope:
                idx = self._envelope.find(b"<LocalPackageMetadataFile")
                if idx != -1:
                    end = self._envelope.find(b"</LocalPackageMetadataFile>", idx)
                    if end != -1:
                        raw_xml = self._envelope[idx : end + len("</LocalPackageMetadataFile>")].decode("utf-8", errors="replace")
                        root = ET.fromstring(raw_xml)
                        for item in root.findall(".//Item"):
                            loc = item.find("ItemLocation")
                            if loc is not None:
                                itype = loc.find("ItemType")
                                ipath = loc.find("ItemPath")
                                if itype is not None and itype.text == "Formula" and ipath is not None and ipath.text:
                                    parts = ipath.text.split("/")
                                    if len(parts) == 2 and parts[0] == "Section1":
                                        qname = urllib.parse.unquote(parts[1])
                                        entries = {e.attrib.get("Type"): e.attrib.get("Value") for e in item.iter("Entry")}

                                        fill_en = entries.get("FillEnabled") == "l1"
                                        fill_obj = entries.get("FillObjectType", "")
                                        fill_target = entries.get("FillTarget")
                                        target_name = fill_target[1:] if fill_target and fill_target.startswith("s") else None
                                        dm_en = entries.get("FillToDataModelEnabled") == "l1" or entries.get("AddedToDataModel") == "l1"

                                        is_table = (fill_obj == "sTable") and (target_name or (qname in table_to_sheet))
                                        if is_table:
                                            tbl_lookup = target_name or qname
                                            if tbl_lookup in table_to_sheet:
                                                sheet = table_to_sheet[tbl_lookup]
                                                dest = f"Table: {sheet}!{tbl_lookup}"
                                            elif target_name:
                                                dest = f"Table: {target_name}"
                                            else:
                                                dest = "Table"
                                            if dm_en:
                                                dest += " + Data Model"
                                        elif fill_obj == "sPivotTable":
                                            dest = "PivotTable + Data Model" if dm_en else "PivotTable"
                                        elif dm_en:
                                            dest = "Data Model"
                                        else:
                                            dest = "Connection Only"

                                        self.query_destinations[qname] = dest
        except Exception:
            pass

        # --- Fallback: Any queries in table_to_sheet not yet marked as table ---
        for tbl_name, sname in table_to_sheet.items():
            if tbl_name in self.query_destinations and self.query_destinations[tbl_name] == "Connection Only":
                self.query_destinations[tbl_name] = f"Table: {sname}!{tbl_name}"

    def _rebuild_envelope(self, new_section: bytes) -> bytes:
        version, pkg, _, end = self._sections()
        tail = self._envelope[end:]
        new_pkg = BytesIO()
        with zipfile.ZipFile(BytesIO(pkg)) as zin, zipfile.ZipFile(new_pkg, "w", zipfile.ZIP_DEFLATED) as zout:
            replaced = False
            for info in zin.infolist():
                data = new_section if info.filename == _SECTION_PART else zin.read(info.filename)
                replaced = replaced or info.filename == _SECTION_PART
                zout.writestr(info, data)
            if not replaced:
                raise PowerQueryError(f"{_SECTION_PART} not present; cannot edit.")
        new_pkg_bytes = new_pkg.getvalue()
        return struct.pack("<ii", version, len(new_pkg_bytes)) + new_pkg_bytes + tail

    def _reencode_item_xml(self, envelope: bytes) -> bytes:
        b64 = base64.b64encode(envelope).decode("ascii")
        text = re.sub(r"(<DataMashup[^>]*>).*?(</DataMashup>)",
                      lambda m: m.group(1) + b64 + m.group(2),
                      self._decode_xml(self._item_xml), flags=re.DOTALL)
        return text.encode(self._item_encoding)

    def _rewrite_workbook_atomic(self, out: Path, new_item_xml: bytes) -> None:
        """Safely write new workbook via a tempfile then atomic replace."""
        temp_dir = out.parent
        temp_dir.mkdir(parents=True, exist_ok=True)
        fd, temp_file_str = tempfile.mkstemp(dir=temp_dir, prefix="pq_", suffix=".tmp")
        os.close(fd)
        temp_path = Path(temp_file_str)

        try:
            with zipfile.ZipFile(self.path, "r") as zin, zipfile.ZipFile(temp_path, "w", zipfile.ZIP_DEFLATED) as zout:
                for info in zin.infolist():
                    if info.filename == self._item_name:
                        zout.writestr(info, new_item_xml)
                    else:
                        zout.writestr(info, zin.read(info.filename))
            # Test that temp_path is valid zip before replacing
            with zipfile.ZipFile(temp_path, "r") as test_z:
                _ = test_z.namelist()
            os.replace(temp_path, out)
        except Exception:
            if temp_path.exists():
                try:
                    temp_path.unlink()
                except Exception:
                    pass
            raise

    @staticmethod
    def _xml_encoding(raw: bytes) -> str:
        if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
            return "utf-16"
        m = re.match(rb"<\?xml[^>]*encoding=['\"]([\w-]+)['\"]", raw)
        return m.group(1).decode("ascii").lower() if m else "utf-8"

    @classmethod
    def _decode_xml(cls, raw: bytes) -> str:
        return raw.decode(cls._xml_encoding(raw), errors="replace")


# ============================================================ M PARSING & VALIDATION ============================================================
_DECL_RE = re.compile(r'(?:shared\s+)?(#"(?P<q>[^"]+)"|(?P<n>[A-Za-z_][\w.]*))\s*=(?!=)')
_SKIP_NAMES = {"section"}

_JOIN_PATTERN = re.compile(
    r'Table\.(NestedJoin|Join)\s*\(\s*'
    r'(?P<t1>#"[^"]+"|[A-Za-z_][\w.]*)\s*,\s*'
    r'(?P<k1>\{[^}]*\}|"[^"]+")\s*,\s*'
    r'(?P<t2>#"[^"]+"|[A-Za-z_][\w.]*)\s*,\s*'
    r'(?P<k2>\{[^}]*\}|"[^"]+")'
    r'(?:[^;)]*?(?:JoinKind\.(?P<kind>[A-Za-z]+)))?',
    re.DOTALL
)

_FIELD_ACCESS_PATTERN = re.compile(
    r'(?P<tbl>#"[^"]+"|[A-Za-z_][\w.]*)'
    r'\['
    r'(?P<col>#"[^"]+"|[A-Za-z_][\w.]*)'
    r'\]'
)


def validate_m_syntax(text: str) -> list[str]:
    """
    Check balanced quotes, comments, brackets, and let...in statements.
    Accurately accounts for bracket depth so that column/field references like
    [Value in Obj. Crcy] are not falsely flagged as mismatched 'in' keywords.
    """
    errors = []
    stack = []      # (char, line_no)
    let_stack = []  # (line_no, bracket_depth)
    i = 0
    n = len(text)
    line_no = 1

    while i < n:
        c = text[i]
        if c == '\n':
            line_no += 1
            i += 1
            continue

        # Quoted identifier #"..."
        if c == '#' and i + 1 < n and text[i + 1] == '"':
            start_line = line_no
            j = i + 2
            closed = False
            while j < n:
                if text[j] == '\n':
                    line_no += 1
                elif text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        j += 2
                        continue
                    closed = True
                    j += 1
                    break
                j += 1
            if not closed:
                errors.append(f"Line {start_line}: Unclosed quoted identifier #\"...")
            i = j
            continue

        # String literal "..."
        if c == '"':
            start_line = line_no
            j = i + 1
            closed = False
            while j < n:
                if text[j] == '\n':
                    line_no += 1
                elif text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"':
                        j += 2
                        continue
                    closed = True
                    j += 1
                    break
                j += 1
            if not closed:
                errors.append(f"Line {start_line}: Unclosed string literal '\"...'")
            i = j
            continue

        # Single-line comment //
        if c == '/' and i + 1 < n and text[i + 1] == '/':
            j = text.find('\n', i)
            if j == -1:
                break
            line_no += 1
            i = j + 1
            continue

        # Block comment /* ... */
        if c == '/' and i + 1 < n and text[i + 1] == '*':
            start_line = line_no
            j = text.find('*/', i + 2)
            if j == -1:
                errors.append(f"Line {start_line}: Unclosed block comment '/*...'")
                break
            line_no += text.count('\n', i, j + 2)
            i = j + 2
            continue

        # Brackets
        if c in '([{':
            stack.append((c, line_no))
        elif c in ')]}':
            matching = {')': '(', ']': '[', '}': '{'}[c]
            if not stack:
                errors.append(f"Line {line_no}: Unexpected closing bracket '{c}'")
            else:
                top, top_line = stack.pop()
                if top != matching:
                    errors.append(f"Line {line_no}: Mismatched bracket '{c}', expected closing for '{top}' from line {top_line}")

        # Identifiers let / in
        if c.isalpha() or c == '_':
            j = i
            while j < n and (text[j].isalnum() or text[j] in '._'):
                j += 1
            word = text[i:j]
            prev_ok = (i == 0 or not (text[i - 1].isalnum() or text[i - 1] in '._'))
            next_ok = (j == n or not (text[j].isalnum() or text[j] in '._'))
            if prev_ok and next_ok:
                current_depth = len(stack)
                inside_bracket = stack and stack[-1][0] == '['

                if word == 'let':
                    let_stack.append((line_no, current_depth))
                elif word == 'in':
                    # 'in' can only close a 'let' if it occurs at the exact same bracket depth.
                    if let_stack and let_stack[-1][1] == current_depth:
                        let_stack.pop()
                    elif not inside_bracket:
                        errors.append(f"Line {line_no}: Unexpected 'in' without preceding 'let'")
            i = j
            continue

        i += 1

    for char, l_no in stack:
        errors.append(f"Line {l_no}: Unclosed bracket '{char}'")
    for l_no, _depth in let_stack:
        errors.append(f"Line {l_no}: Unclosed 'let' missing matching 'in'")

    return errors


def _skip_trivia(s: str, i: int = 0) -> int:
    n = len(s)
    while i < n:
        if s[i] in ' \t\r\n': i += 1; continue
        if s.startswith('//', i): j = s.find('\n', i); i = n if j == -1 else j + 1; continue
        if s.startswith('/*', i): j = s.find('*/', i + 2); i = n if j == -1 else j + 2; continue
        break
    return i


def _top_level_statements(text: str):
    depth = 0; i = 0; n = len(text); start = 0; in_str = False
    while i < n:
        c = text[i]
        if in_str:
            if c == '"':
                if i + 1 < n and text[i + 1] == '"': i += 2; continue
                in_str = False
            i += 1; continue
        if c == '/' and i + 1 < n and text[i + 1] == '/':
            j = text.find('\n', i); i = n if j == -1 else j; continue
        if c == '/' and i + 1 < n and text[i + 1] == '*':
            j = text.find('*/', i + 2); i = n if j == -1 else j + 2; continue
        if c == '#' and i + 1 < n and text[i + 1] == '"':
            j = i + 2
            while j < n:
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"': j += 2; continue
                    break
                j += 1
            i = j + 1; continue
        if c == '"': in_str = True; i += 1; continue
        if c in '([{': depth += 1
        elif c in ')]}': depth -= 1
        elif c == ';' and depth == 0:
            yield start, i; start = i + 1
        i += 1
    if text[start:n].strip(): yield start, n


def parse_query_spans(m_text: str):
    out = []
    for s, e in _top_level_statements(m_text):
        stmt = m_text[s:e]
        k = _skip_trivia(stmt)
        mt = _DECL_RE.match(stmt, k)
        if not mt or mt.start() != k:
            continue
        name = mt.group("q") or mt.group("n")
        if name.lower() in _SKIP_NAMES:
            continue
        out.append((name, s + mt.start(1), s, e))
    return out


def parse_queries(m_text: str) -> list[tuple[str, int]]:
    return [(n, off) for (n, off, _s, _e) in parse_query_spans(m_text)]


def _strip_noncode(text: str) -> str:
    out = list(text); i = 0; n = len(text)
    while i < n:
        c = text[i]
        if c == '#' and i + 1 < n and text[i + 1] == '"':
            j = i + 2
            while j < n:
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"': j += 2; continue
                    break
                j += 1
            i = j + 1; continue
        if c == '"':
            out[i] = ' '; j = i + 1
            while j < n:
                if text[j] == '"':
                    if j + 1 < n and text[j + 1] == '"': out[j] = out[j + 1] = ' '; j += 2; continue
                    out[j] = ' '; j += 1; break
                out[j] = ' '; j += 1
            i = j; continue
        if c == '/' and i + 1 < n and text[i + 1] == '/':
            j = text.find('\n', i); j = n if j == -1 else j
            for k in range(i, j): out[k] = ' '
            i = j; continue
        if c == '/' and i + 1 < n and text[i + 1] == '*':
            j = text.find('*/', i + 2); j = n if j == -1 else j + 2
            for k in range(i, min(j, n)): out[k] = ' '
            i = j; continue
        i += 1
    return ''.join(out)


def _references(area: str, other: str) -> bool:
    if re.fullmatch(r'[A-Za-z_][\w.]*', other):
        return re.search(r'(?<![\w.#"])' + re.escape(other) + r'(?![\w.])', area) is not None
    return ('#"' + other + '"') in area


def build_dependencies(m_text: str):
    spans = parse_query_spans(m_text)
    names = [n for (n, _o, _s, _e) in spans]
    names_set = set(names)
    deps = {n: set() for n in names}
    join_details: dict[tuple[str, str], list[dict]] = {}
    field_details: dict[tuple[str, str], set[str]] = {}

    for (name, _o, s, e) in spans:
        stmt = m_text[s:e]
        mt = _DECL_RE.match(stmt, _skip_trivia(stmt))
        rhs_raw = stmt[mt.end():] if mt else ""
        rhs_clean = _strip_noncode(rhs_raw)

        # 1. Standard references
        for other in names:
            if other != name and _references(rhs_clean, other):
                deps[name].add(other)

        # 2. Join detection
        for jm in _JOIN_PATTERN.finditer(rhs_raw):
            t1 = jm.group("t1").strip('#"')
            t2 = jm.group("t2").strip('#"')
            k1 = [x.strip(' "\t\r\n') for x in re.findall(r'"([^"]+)"', jm.group("k1"))]
            k2 = [x.strip(' "\t\r\n') for x in re.findall(r'"([^"]+)"', jm.group("k2"))]
            kind = jm.group("kind") or ("LeftOuter" if jm.group(1) == "NestedJoin" else "Inner")

            if t2 in names_set and t2 != name:
                deps[name].add(t2)
                join_details.setdefault((t2, name), []).append({
                    "kind": kind, "source_keys": k2, "target_keys": k1
                })
            if t1 in names_set and t1 != name:
                deps[name].add(t1)
                join_details.setdefault((t1, name), []).append({
                    "kind": kind, "source_keys": k1, "target_keys": k2
                })

        # 3. Field access detection
        for fm in _FIELD_ACCESS_PATTERN.finditer(rhs_raw):
            tbl = fm.group("tbl").strip('#"')
            col = fm.group("col").strip('#"')
            if tbl in names_set and tbl != name:
                deps[name].add(tbl)
                field_details.setdefault((tbl, name), set()).add(col)

    return names, deps, join_details, field_details


def dependencies_to_mermaid(names, deps, join_details=None, field_details=None) -> str:
    ids = {n: f"q{i}" for i, n in enumerate(names)}
    lines = [
        "graph LR",
        "    classDef default fill:#F8F9FA,stroke:#0E703C,stroke-width:1.5px,color:#201F1E;",
        "    classDef source fill:#E2EFDA,stroke:#375623,stroke-width:2px,color:#201F1E;"
    ]
    for n in names:
        safe_name = n.replace('"', '&quot;')
        lines.append(f'    {ids[n]}["{safe_name}"]')
        if not deps.get(n):
            lines.append(f'    class {ids[n]} source;')

    for n in names:
        for d in sorted(deps.get(n, ())):
            label_parts = []
            if join_details and (d, n) in join_details:
                for j in join_details[(d, n)]:
                    k_str = f"[{', '.join(j['source_keys'])}] = [{', '.join(j['target_keys'])}]" if j['source_keys'] else ""
                    label_parts.append(f"Join: {j['kind']}<br>{k_str}".strip())
            elif field_details and (d, n) in field_details:
                cols = field_details[(d, n)]
                label_parts.append(f"Fields: [{', '.join(sorted(cols))}]")

            edge_label = f'|"{ "<br>".join(label_parts) }"|' if label_parts else ""
            lines.append(f'    {ids[d]} -->{edge_label} {ids[n]}')

    return "\n".join(lines)


def compute_layers(names, deps) -> dict[str, int]:
    memo: dict[str, int] = {}; visiting: set[str] = set(); known = set(names)
    def layer(n: str) -> int:
        if n in memo: return memo[n]
        if n in visiting: return 0
        visiting.add(n); L = 0
        for d in deps.get(n, ()):
            if d in known: L = max(L, layer(d) + 1)
        visiting.discard(n); memo[n] = L; return L
    return {n: layer(n) for n in names}


def query_bodies(m_text: str) -> list[tuple[str, str]]:
    out = []
    for (name, _o, s, e) in parse_query_spans(m_text):
        stmt = m_text[s:e]
        mt = _DECL_RE.match(stmt, _skip_trivia(stmt))
        rhs = stmt[mt.end():].strip() if mt else ""
        out.append((name, rhs))
    return out


# ============================================================ GUI ============================================================
def launch_gui() -> None:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    FILETYPES = [("Excel workbooks", "*.xlsx *.xlsm *.xlsb"), ("All files", "*.*")]

    M_KEYWORDS = {
        "and", "as", "each", "else", "error", "false", "if", "in", "is", "let",
        "meta", "not", "otherwise", "or", "section", "shared", "then", "true",
        "try", "type", "nullable", "optional",
    }
    _KW_RE = re.compile(r"\b(" + "|".join(sorted(M_KEYWORDS, key=len, reverse=True)) + r")\b")
    _STR_RE = re.compile(r'"(?:[^"]|"")*"')
    _CMT_RE = re.compile(r"//[^\n]*|/\*.*?\*/", re.DOTALL)
    _NUM_RE = re.compile(r"\b\d+(?:\.\d+)?\b")
    _ATN_RE = re.compile(r'#"[^"]+"')

    state: dict[str, object] = {
        "pq": None,
        "loaded_text": "",
        "queries": [],
        "hl_job": None,
        "find_matches": [],
        "find_idx": -1,
        "theme_name": "light",
    }

    root = tk.Tk()
    root.title("PQ Editor — Standalone M Engine & Visualizer")
    root.geometry("1120x760")
    root.minsize(860, 560)

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass

    root.columnconfigure(0, weight=1)
    root.rowconfigure(2, weight=1)

    # --- 1. Header Banner ---
    header = ttk.Frame(root, style="Header.TFrame", padding=(12, 10, 12, 10))
    header.grid(row=0, column=0, sticky="ew")
    header.columnconfigure(1, weight=1)

    ttk.Label(header, text="PQ Editor", style="HeaderTitle.TLabel").grid(row=0, column=0, sticky="w")
    ttk.Label(header, text="Unchained M Engine · Zero Locks, Zero Freezes", style="HeaderSub.TLabel").grid(row=0, column=1, padx=(14, 0), sticky="w")

    # Header Controls (Theme toggle + File badge)
    hdr_right = ttk.Frame(header, style="Header.TFrame")
    hdr_right.grid(row=0, column=2, sticky="e")

    btn_theme = ttk.Button(hdr_right, text="🌙 Dark Mode", command=lambda: _toggle_theme())
    btn_theme.pack(side="left", padx=(0, 10))

    file_badge = ttk.Label(hdr_right, text="No File Open", style="Badge.TLabel")
    file_badge.pack(side="left")

    # --- 2. Action Toolbar (Streamlined & Combined) ---
    tb = ttk.Frame(root, style="Toolbar.TFrame", padding=(8, 6, 8, 6))
    tb.grid(row=1, column=0, sticky="ew")

    btn_open = ttk.Button(tb, text="📂 Open…", command=lambda: _choose())
    btn_open.grid(row=0, column=0, padx=2)

    # Combined Save ▾
    mb_save = ttk.Menubutton(tb, text="💾 Save ▾", state="disabled")
    menu_save = tk.Menu(mb_save, tearoff=False)
    menu_save.add_command(label="Save (in place)        Ctrl+S", command=lambda: _save(False))
    menu_save.add_command(label="Save As…", command=lambda: _save(True))
    menu_save.add_separator()
    menu_save.add_command(label="🔄 Restore Backup (.bak)", command=lambda: _restore_backup())
    menu_save.add_command(label="Reload from Disk", command=lambda: _reload())
    mb_save.configure(menu=menu_save)
    mb_save.grid(row=0, column=1, padx=2)

    # Find Button
    btn_find = ttk.Button(tb, text="🔍 Find (Ctrl+F)", command=lambda: _toggle_find_bar(True), state="disabled")
    btn_find.grid(row=0, column=2, padx=2)

    # Combined Dependencies ▾
    mb_dep = ttk.Menubutton(tb, text="📊 Dependencies ▾", state="disabled")
    menu_dep = tk.Menu(mb_dep, tearoff=False)
    menu_dep.add_command(label="Interactive Canvas View…", command=lambda: _show_dependencies())
    menu_dep.add_command(label="🌐 Web Graph in Browser…", command=lambda: _open_browser_graph())
    menu_dep.add_separator()
    menu_dep.add_command(label="Copy Mermaid Code", command=lambda: _copy_mermaid_direct())
    menu_dep.add_command(label="Save Mermaid (.mmd)…", command=lambda: _save_mermaid_direct())
    mb_dep.configure(menu=menu_dep)
    mb_dep.grid(row=0, column=3, padx=2)

    # Combined Export ▾
    mb_exp = ttk.Menubutton(tb, text="📤 Export ▾", state="disabled")
    menu_exp = tk.Menu(mb_exp, tearoff=False)
    menu_exp.add_command(label="Export Single File… (.m / .pq / .txt)", command=lambda: _export_single())
    menu_exp.add_command(label="Export Split (per query)…", command=lambda: _export_split())
    mb_exp.configure(menu=menu_exp)
    mb_exp.grid(row=0, column=4, padx=2)

    # --- 3. Main Body: Split PanedWindow (Sidebar + Editor) ---
    paned = ttk.PanedWindow(root, orient="horizontal")
    paned.grid(row=2, column=0, sticky="nsew", padx=6, pady=4)

    # --- Left Pane: Queries Treeview Sidebar ---
    sidebar_frame = ttk.Frame(paned, padding=(4, 4, 4, 4))
    sidebar_frame.columnconfigure(0, weight=1)
    sidebar_frame.rowconfigure(2, weight=1)
    paned.add(sidebar_frame, weight=1)

    # Search query filter entry
    filter_box = ttk.Frame(sidebar_frame)
    filter_box.grid(row=0, column=0, sticky="ew", pady=(0, 4))
    filter_box.columnconfigure(0, weight=1)
    search_var = tk.StringVar()
    search_entry = ttk.Entry(filter_box, textvariable=search_var)
    search_entry.grid(row=0, column=0, sticky="ew")
    search_entry.insert(0, "Search queries…")

    def _on_search_focus_in(e):
        if search_var.get() == "Search queries…":
            search_entry.delete(0, "end")
    def _on_search_focus_out(e):
        if not search_var.get().strip():
            search_entry.insert(0, "Search queries…")
    search_entry.bind("<FocusIn>", _on_search_focus_in)
    search_entry.bind("<FocusOut>", _on_search_focus_out)
    search_entry.bind("<KeyRelease>", lambda e: _filter_treeview())

    btn_clear_search = ttk.Button(filter_box, text="✕", width=3, command=lambda: (search_var.set(""), _filter_treeview()))
    btn_clear_search.grid(row=0, column=1, padx=(2, 0))

    # Treeview
    tree_scroll = ttk.Scrollbar(sidebar_frame, orient="vertical")
    tree = ttk.Treeview(sidebar_frame, columns=("destination",), show="tree headings", selectmode="browse", yscrollcommand=tree_scroll.set, style="Sidebar.Treeview")
    tree.heading("#0", text="Queries / Groups")
    tree.heading("destination", text="Excel Destination")
    tree.column("#0", width=190, stretch=True)
    tree.column("destination", width=140, stretch=True)
    tree_scroll.config(command=tree.yview)

    tree.grid(row=2, column=0, sticky="nsew")
    tree_scroll.grid(row=2, column=1, sticky="ns")

    # --- Right Pane: M Code Editor ---
    editor_container = ttk.Frame(paned, padding=(4, 4, 4, 4))
    editor_container.columnconfigure(0, weight=1)
    editor_container.rowconfigure(1, weight=1)
    paned.add(editor_container, weight=4)

    # Top Location & Breadcrumb bar
    loc_bar = ttk.Frame(editor_container)
    loc_bar.grid(row=0, column=0, sticky="ew", pady=(0, 4))
    loc_bar.columnconfigure(1, weight=1)

    current_query_lbl = ttk.Label(loc_bar, text="Section1.m", font=("Segoe UI", 9, "bold"))
    current_query_lbl.grid(row=0, column=0, sticky="w")
    dest_pill = ttk.Label(loc_bar, text="", foreground="#505050")
    dest_pill.grid(row=0, column=1, padx=(10, 0), sticky="w")

    # Options: Line Numbers & Wrap toggles
    line_nums_var = tk.BooleanVar(value=True)
    wrap_var = tk.BooleanVar(value=False)

    chk_wrap = ttk.Checkbutton(loc_bar, text="Wrap Text", variable=wrap_var, command=lambda: _toggle_wrap())
    chk_lines = ttk.Checkbutton(loc_bar, text="Line Numbers", variable=line_nums_var, command=lambda: _toggle_line_numbers())
    chk_wrap.grid(row=0, column=2, padx=4)
    chk_lines.grid(row=0, column=3, padx=4)

    # Editor Wrap Frame
    ewrap = ttk.Frame(editor_container)
    ewrap.grid(row=1, column=0, sticky="nsew")
    ewrap.columnconfigure(1, weight=1)
    ewrap.rowconfigure(0, weight=1)

    # Line Numbers Canvas
    line_canvas = tk.Canvas(ewrap, width=44, bg="#F3F2F1", highlightthickness=0)
    line_canvas.grid(row=0, column=0, sticky="ns")

    # Text Widget
    editor = tk.Text(ewrap, wrap="none", undo=True, font=("Consolas", 10), state="disabled")
    editor.grid(row=0, column=1, sticky="nsew")

    ysb = ttk.Scrollbar(ewrap, orient="vertical", command=editor.yview)
    ysb.grid(row=0, column=2, sticky="ns")
    xsb = ttk.Scrollbar(ewrap, orient="horizontal", command=editor.xview)
    xsb.grid(row=1, column=1, sticky="ew")

    def _sync_scroll(*args):
        ysb.set(*args)
        _redraw_line_numbers()
    editor.configure(yscrollcommand=_sync_scroll, xscrollcommand=xsb.set)

    # --- Find & Replace Bar (Collapsible) ---
    find_frame = ttk.Frame(editor_container, style="Find.TFrame", padding=(8, 6, 8, 6))
    find_frame.columnconfigure(1, weight=1)

    ttk.Label(find_frame, text="Find:", style="Find.TLabel").grid(row=0, column=0, padx=(0, 4))
    find_var = tk.StringVar()
    find_entry = ttk.Entry(find_frame, textvariable=find_var, width=24)
    find_entry.grid(row=0, column=1, sticky="ew")

    match_lbl = ttk.Label(find_frame, text="", style="Find.TLabel", width=14)
    match_lbl.grid(row=0, column=2, padx=4)

    btn_find_prev = ttk.Button(find_frame, text="▲ Prev", width=6, command=lambda: _find_step(-1))
    btn_find_next = ttk.Button(find_frame, text="▼ Next", width=6, command=lambda: _find_step(1))
    btn_find_prev.grid(row=0, column=3, padx=2)
    btn_find_next.grid(row=0, column=4, padx=2)

    ttk.Label(find_frame, text="Replace:", style="Find.TLabel").grid(row=1, column=0, padx=(0, 4), pady=(4, 0))
    replace_var = tk.StringVar()
    replace_entry = ttk.Entry(find_frame, textvariable=replace_var, width=24)
    replace_entry.grid(row=1, column=1, sticky="ew", pady=(4, 0))

    btn_replace_one = ttk.Button(find_frame, text="Replace", width=8, command=lambda: _replace_one())
    btn_replace_all = ttk.Button(find_frame, text="Replace All", width=10, command=lambda: _replace_all())
    btn_replace_one.grid(row=1, column=3, padx=2, pady=(4, 0))
    btn_replace_all.grid(row=1, column=4, padx=2, pady=(4, 0))

    btn_close_find = ttk.Button(find_frame, text="✕", width=3, command=lambda: _toggle_find_bar(False))
    btn_close_find.grid(row=0, column=5, padx=(6, 0))

    def _on_find_key(event=None) -> None:
        if event and event.keysym in ("Return", "KP_Enter", "Shift_L", "Shift_R", "Control_L", "Control_R", "Alt_L", "Alt_R", "Escape", "Up", "Down", "Tab"):
            return
        term = find_var.get()
        if term != state.get("last_find_term"):
            state["last_find_term"] = term
            _do_find()

    find_entry.bind("<Return>", lambda e: (_find_step(1), "break")[1])
    find_entry.bind("<Shift-Return>", lambda e: (_find_step(-1), "break")[1])
    find_entry.bind("<KeyRelease>", _on_find_key)
    find_entry.bind("<Escape>", lambda e: _toggle_find_bar(False))
    replace_entry.bind("<Escape>", lambda e: _toggle_find_bar(False))

    # --- 4. Status Bar ---
    status_bar = ttk.Frame(root, style="Status.TFrame", relief="sunken", padding=(8, 3))
    status_bar.grid(row=3, column=0, sticky="ew")
    status_bar.columnconfigure(0, weight=1)

    status = tk.StringVar(value="Ready. Click 'Open…' to choose an Excel workbook.")
    status_lbl = ttk.Label(status_bar, style="Status.TLabel", textvariable=status, anchor="w")
    status_lbl.grid(row=0, column=0, sticky="w")

    cur_pos_lbl = ttk.Label(status_bar, style="Status.TLabel", text="Ln 1, Col 1", anchor="e")
    cur_pos_lbl.grid(row=0, column=1, padx=(10, 4), sticky="e")

    dirty_lbl = ttk.Label(status_bar, style="Status.TLabel", text="✓ Clean", anchor="e", font=("Segoe UI", 9, "bold"))
    dirty_lbl.grid(row=0, column=2, padx=(4, 8), sticky="e")

    # ---------- THEME APPLICATION ----------
    def _apply_theme(t_name: str) -> None:
        state["theme_name"] = t_name
        t = THEMES[t_name]

        root.configure(bg=t["bg"])

        # Base styles
        style.configure(".", font=("Segoe UI", 9), background=t["bg"], foreground=t["fg"])
        style.configure("TFrame", background=t["bg"])
        style.configure("TLabel", background=t["bg"], foreground=t["fg"])
        style.configure("TCheckbutton", background=t["bg"], foreground=t["fg"])
        style.map("TCheckbutton", background=[("active", t["bg"])], foreground=[("active", t["fg"])])

        # Header banner
        style.configure("Header.TFrame", background=t["header_bg"])
        style.configure("HeaderTitle.TLabel", background=t["header_bg"], foreground=t["header_fg"], font=("Segoe UI", 13, "bold"))
        style.configure("HeaderSub.TLabel", background=t["header_bg"], foreground=t["header_sub"], font=("Segoe UI", 9))
        style.configure("Badge.TLabel", background=t["badge_bg"], foreground=t["badge_fg"], font=("Segoe UI", 8, "bold"), padding=(6, 2))

        # Toolbar
        style.configure("Toolbar.TFrame", background=t["toolbar_bg"])

        # Treeview sidebar
        style.configure("Sidebar.Treeview", background=t["sidebar_bg"], foreground=t["sidebar_fg"], fieldbackground=t["sidebar_bg"], font=("Segoe UI", 9), rowheight=24)
        style.configure("Sidebar.Treeview.Heading", background=t["toolbar_bg"], foreground=t["fg"], font=("Segoe UI", 9, "bold"))

        # Entries (Search query bar, Find & Replace entry bars)
        style.configure("TEntry",
            fieldbackground=t["input_bg"],
            foreground=t["input_fg"],
            insertcolor=t["input_fg"],
            bordercolor=t["border_color"],
            lightcolor=t["border_color"],
            darkcolor=t["border_color"]
        )
        style.map("TEntry",
            fieldbackground=[("focus", t["input_bg"]), ("!disabled", t["input_bg"])],
            foreground=[("focus", t["input_fg"]), ("!disabled", t["input_fg"])]
        )

        # Buttons & Menubuttons
        style.configure("TButton", background=t["btn_bg"], foreground=t["fg"], bordercolor=t["border_color"])
        style.map("TButton", background=[("active", t["btn_active"])], foreground=[("active", t["fg"])])
        style.configure("TMenubutton", background=t["btn_bg"], foreground=t["fg"], bordercolor=t["border_color"])
        style.map("TMenubutton", background=[("active", t["btn_active"])], foreground=[("active", t["fg"])])

        # Find & Replace Card Bar
        style.configure("Find.TFrame", background=t["card_bg"])
        style.configure("Find.TLabel", background=t["card_bg"], foreground=t["fg"])

        # Status bar
        style.configure("Status.TFrame", background=t["status_bg"])
        style.configure("Status.TLabel", background=t["status_bg"], foreground=t["status_fg"])

        # Editor styling
        editor.configure(
            bg=t["editor_bg"],
            fg=t["editor_fg"],
            insertbackground=t["editor_insert"],
            selectbackground=t["editor_select"]
        )
        line_canvas.configure(bg=t["gutter_bg"])

        # Menus
        for m in (menu_save, menu_dep, menu_exp):
            try:
                m.configure(
                    bg=t["card_bg"],
                    fg=t["fg"],
                    activebackground=t["menu_active_bg"],
                    activeforeground=t["menu_active_fg"]
                )
            except Exception:
                pass

        # Reconfigure syntax tags
        for tag, color in t["colors"].items():
            if tag == "find_match":
                editor.tag_configure(tag, background=color)
            elif tag == "find_active":
                editor.tag_configure(tag, background=color, foreground="white")
            else:
                editor.tag_configure(tag, foreground=color)
        editor.tag_configure("qhit", background="#FFF2A8" if t_name == "light" else "#5F5010")

        # Labels & Buttons
        btn_theme.config(text="☀️ Light Mode" if t_name == "dark" else "🌙 Dark Mode")
        _redraw_line_numbers()
        _update_status()
        if editor.get("1.0", "end-1c"):
            _highlight()

    def _toggle_theme() -> None:
        new_theme = "dark" if state["theme_name"] == "light" else "light"
        _apply_theme(new_theme)

    # ---------- HELPER FUNCTIONS ----------
    def _set_enabled(on: bool) -> None:
        mb_save.configure(state="normal" if on else "disabled")
        btn_find.configure(state="normal" if on else "disabled")
        mb_dep.configure(state="normal" if on else "disabled")
        mb_exp.configure(state="normal" if on else "disabled")
        editor.configure(state="normal" if on else "disabled")

    def _update_status() -> None:
        t = THEMES[state["theme_name"]]
        try:
            pos = editor.index("insert")
            line, col = pos.split(".")
            cur_pos_lbl.config(text=f"Ln {line}, Col {int(col) + 1}", foreground=t["gutter_fg"])
        except Exception:
            pass
        if _dirty():
            dirty_lbl.config(text="● Modified", foreground=t["dirty_fg"])
        else:
            dirty_lbl.config(text="✓ Clean", foreground=t["clean_fg"])

    def _dirty() -> bool:
        return editor.get("1.0", "end-1c") != state["loaded_text"]

    def _toggle_wrap() -> None:
        if wrap_var.get():
            editor.configure(wrap="word")
            xsb.grid_remove()
        else:
            editor.configure(wrap="none")
            xsb.grid()
        _redraw_line_numbers()

    def _toggle_line_numbers() -> None:
        if line_nums_var.get():
            line_canvas.grid()
            _redraw_line_numbers()
        else:
            line_canvas.grid_remove()

    def _redraw_line_numbers() -> None:
        if not line_nums_var.get():
            return
        t = THEMES[state["theme_name"]]
        line_canvas.delete("all")
        i = editor.index("@0,0")
        w = line_canvas.winfo_width() or 44
        while True:
            dline = editor.dlineinfo(i)
            if dline is None:
                break
            y = dline[1]
            line_num = str(i).split(".")[0]
            line_canvas.create_text(
                w - 8, y + dline[3] // 2,
                anchor="e", text=line_num, fill=t["gutter_fg"], font=("Consolas", 9)
            )
            i = editor.index(f"{i}+1line")

    def _highlight() -> None:
        """Full-document high-fidelity syntax highlighting."""
        text = editor.get("1.0", "end-1c")
        for tag in ("comment", "string", "atname", "keyword", "number"):
            editor.tag_remove(tag, "1.0", "end")

        spans: list[tuple[int, int, str]] = []
        for rx, tag in ((_CMT_RE, "comment"), (_STR_RE, "string"), (_ATN_RE, "atname")):
            spans += [(m.start(), m.end(), tag) for m in rx.finditer(text)]

        spans.sort(key=lambda x: x[0])
        def _covered(a: int, b: int) -> bool:
            for s, e, _ in spans:
                if s <= a and b <= e:
                    return True
                if s > b:
                    break
            return False

        for m in _KW_RE.finditer(text):
            if not _covered(m.start(), m.end()):
                spans.append((m.start(), m.end(), "keyword"))
        for m in _NUM_RE.finditer(text):
            if not _covered(m.start(), m.end()):
                spans.append((m.start(), m.end(), "number"))

        for a, b, tag in spans:
            editor.tag_add(tag, f"1.0+{a}c", f"1.0+{b}c")

    def _schedule_highlight() -> None:
        if state["hl_job"]:
            editor.after_cancel(state["hl_job"])
        state["hl_job"] = editor.after(180, lambda: (_highlight(), _refresh_query_list(), _update_status()))

    def _refresh_query_list() -> None:
        queries = parse_queries(editor.get("1.0", "end-1c"))
        state["queries"] = queries
        _populate_treeview()

    def _populate_treeview() -> None:
        for item in tree.get_children():
            tree.delete(item)

        pq: ExcelPowerQuery = state["pq"]
        groups = pq.query_groups if pq else {}
        destinations = pq.query_destinations if pq else {}
        filter_text = search_var.get().strip().lower()
        if filter_text == "search queries…":
            filter_text = ""

        # Group queries
        by_group: dict[str, list[tuple[str, int]]] = {}
        for name, off in state["queries"]:
            if filter_text and filter_text not in name.lower():
                continue
            grp = groups.get(name, "Other Queries")
            by_group.setdefault(grp, []).append((name, off))

        for grp in sorted(by_group):
            grp_node = tree.insert("", "end", text=f"📁 {grp}", open=True)
            for name, off in sorted(by_group[grp], key=lambda x: x[0].lower()):
                dest = destinations.get(name, "Connection Only")
                icon = "📊" if "Table:" in dest else ("📈" if "Pivot" in dest else "🔗")
                tree.insert(grp_node, "end", iid=name, text=f"{icon} {name}", values=(dest,))

    def _filter_treeview() -> None:
        _populate_treeview()

    def _on_tree_select(event) -> None:
        selected = tree.selection()
        if not selected:
            return
        target_name = selected[0]
        if tree.get_children(target_name):
            return
        _jump_to_query(target_name)

    tree.bind("<<TreeviewSelect>>", _on_tree_select)

    def _jump_to_query(name: str) -> None:
        for n, off in state["queries"]:
            if n == name:
                idx = f"1.0+{off}c"
                editor.tag_remove("qhit", "1.0", "end")
                editor.tag_add("qhit", f"{idx} linestart", f"{idx} lineend")
                editor.see(idx)
                editor.mark_set("insert", idx)
                current_query_lbl.config(text=f"Query: {n}")
                pq = state["pq"]
                dest = pq.query_destinations.get(n, "Connection Only") if pq else ""
                dest_pill.config(text=f"[{dest}]")
                editor.focus_set()
                status.set(f"Jumped to query: {n}")
                _redraw_line_numbers()
                return

    # --- Find / Replace Implementation ---
    def _toggle_find_bar(show: bool) -> None:
        if show:
            find_frame.grid(row=2, column=0, sticky="ew", pady=(4, 0))
            find_entry.focus_set()
            find_entry.select_range(0, "end")
            _do_find()
        else:
            find_frame.grid_remove()
            editor.tag_remove("find_match", "1.0", "end")
            editor.tag_remove("find_active", "1.0", "end")
            editor.focus_set()

    def _do_find() -> None:
        editor.tag_remove("find_match", "1.0", "end")
        editor.tag_remove("find_active", "1.0", "end")
        term = find_var.get()
        state["find_matches"] = []
        state["find_idx"] = -1
        if not term:
            match_lbl.config(text="")
            return

        start = "1.0"
        while True:
            pos = editor.search(term, start, stopindex="end", nocase=True)
            if not pos:
                break
            end = f"{pos}+{len(term)}c"
            editor.tag_add("find_match", pos, end)
            state["find_matches"].append((pos, end))
            start = end

        matches = state["find_matches"]
        if matches:
            state["find_idx"] = 0
            _highlight_active_find()
        else:
            match_lbl.config(text="0 matches")

    def _highlight_active_find() -> None:
        matches = state["find_matches"]
        idx = state["find_idx"]
        if 0 <= idx < len(matches):
            editor.tag_remove("find_active", "1.0", "end")
            pos, end = matches[idx]
            editor.tag_add("find_active", pos, end)
            editor.see(pos)
            editor.mark_set("insert", pos)
            match_lbl.config(text=f"{idx + 1} of {len(matches)}")
            _redraw_line_numbers()
        else:
            match_lbl.config(text=f"{len(matches)} matches")

    def _find_step(direction: int) -> None:
        matches = state["find_matches"]
        if not matches:
            return
        state["find_idx"] = (state["find_idx"] + direction) % len(matches)
        _highlight_active_find()

    def _replace_one() -> None:
        matches = state["find_matches"]
        idx = state["find_idx"]
        rep = replace_var.get()
        if matches and 0 <= idx < len(matches):
            pos, end = matches[idx]
            editor.delete(pos, end)
            editor.insert(pos, rep)
            _do_find()

    def _replace_all() -> None:
        term = find_var.get()
        rep = replace_var.get()
        if not term:
            return
        content = editor.get("1.0", "end-1c")
        new_content, count = re.subn(re.escape(term), rep, content, flags=re.IGNORECASE)
        if count:
            editor.delete("1.0", "end")
            editor.insert("1.0", new_content)
            _do_find()
            messagebox.showinfo("Replace All", f"Replaced {count} occurrence(s).")

    # Unmap Ctrl+/ from SelectAll so it doesn't select the whole document
    try:
        root.event_delete("<<SelectAll>>", "<Control-Key-/>")
    except Exception:
        pass

    # --- Comment / Uncomment (Ctrl+/) ---
    def _toggle_comment(event=None):
        had_selection = False
        try:
            sel_first = editor.index("sel.first")
            sel_last = editor.index("sel.last")
            start_line = int(sel_first.split(".")[0])
            end_line = int(sel_last.split(".")[0])
            if int(sel_last.split(".")[1]) == 0 and end_line > start_line:
                end_line -= 1
            had_selection = True
        except tk.TclError:
            cur_line = int(editor.index("insert").split(".")[0])
            start_line = end_line = cur_line

        all_commented = True
        for l in range(start_line, end_line + 1):
            line_str = editor.get(f"{l}.0", f"{l}.end")
            if line_str.strip() and not line_str.lstrip().startswith("//"):
                all_commented = False
                break

        for l in range(start_line, end_line + 1):
            line_str = editor.get(f"{l}.0", f"{l}.end")
            if all_commented:
                new_str = re.sub(r"^(\s*)//\s?", r"\1", line_str)
            else:
                new_str = f"// {line_str}"
            editor.delete(f"{l}.0", f"{l}.end")
            editor.insert(f"{l}.0", new_str)

        if had_selection:
            editor.tag_remove("sel", "1.0", "end")
            editor.tag_add("sel", f"{start_line}.0", f"{end_line}.end")

        _schedule_highlight()
        _redraw_line_numbers()
        _update_status()
        return "break"

    # --- Uncomment Lines (Ctrl+Shift+/) ---
    def _uncomment_lines(event=None):
        had_selection = False
        try:
            sel_first = editor.index("sel.first")
            sel_last = editor.index("sel.last")
            start_line = int(sel_first.split(".")[0])
            end_line = int(sel_last.split(".")[0])
            if int(sel_last.split(".")[1]) == 0 and end_line > start_line:
                end_line -= 1
            had_selection = True
        except tk.TclError:
            cur_line = int(editor.index("insert").split(".")[0])
            start_line = end_line = cur_line

        for l in range(start_line, end_line + 1):
            line_str = editor.get(f"{l}.0", f"{l}.end")
            new_str = re.sub(r"^(\s*)//\s?", r"\1", line_str)
            if new_str != line_str:
                editor.delete(f"{l}.0", f"{l}.end")
                editor.insert(f"{l}.0", new_str)

        if had_selection:
            editor.tag_remove("sel", "1.0", "end")
            editor.tag_add("sel", f"{start_line}.0", f"{end_line}.end")

        _schedule_highlight()
        _redraw_line_numbers()
        _update_status()
        return "break"

    # Shortcuts
    root.bind("<Control-s>", lambda e: (_save(False), "break")[1])
    root.bind("<Control-f>", lambda e: (_toggle_find_bar(True), "break")[1])
    root.bind("<Control-h>", lambda e: (_toggle_find_bar(True), replace_entry.focus_set(), "break")[1])
    for seq in ("<Control-slash>", "<Control-Key-slash>", "<Control-KP_Divide>"):
        editor.bind(seq, _toggle_comment)
        root.bind(seq, _toggle_comment)
    for seq in (
        "<Control-Shift-slash>",
        "<Control-Shift-Key-slash>",
        "<Control-question>",
        "<Control-Key-question>",
        "<Control-Shift-question>",
        "<Control-Shift-KP_Divide>",
    ):
        editor.bind(seq, _uncomment_lines)
        root.bind(seq, _uncomment_lines)
    editor.bind("<KeyRelease>", lambda e: (_schedule_highlight(), _redraw_line_numbers(), _update_status()))
    editor.bind("<MouseWheel>", lambda e: root.after_idle(_redraw_line_numbers))
    editor.bind("<Configure>", lambda e: root.after_idle(_redraw_line_numbers))

    # ---------- FILE ACTIONS ----------
    def _load_text(text: str) -> None:
        editor.configure(state="normal")
        editor.delete("1.0", "end")
        editor.insert("1.0", text)
        editor.edit_reset()
        state["loaded_text"] = text
        _highlight()
        _refresh_query_list()
        _redraw_line_numbers()
        _update_status()

    def _choose() -> None:
        fp = filedialog.askopenfilename(title="Select an Excel workbook", filetypes=FILETYPES)
        if fp:
            _open(fp)

    def _open(fp: str) -> None:
        try:
            pq = ExcelPowerQuery(fp)
            text = pq.read()
        except PowerQueryError as e:
            messagebox.showerror("Cannot open", str(e))
            status.set(f"Error: {e}")
            return
        except Exception as e:
            messagebox.showerror("Unexpected error", repr(e))
            status.set(f"Error: {e}")
            return

        state["pq"] = pq
        file_badge.config(text=pq.path.name)
        _set_enabled(True)
        _load_text(text)
        n = len(state["queries"])
        status.set(f"Loaded {len(text.splitlines())} lines, {n} quer{'y' if n == 1 else 'ies'} from {pq.path.name}")

    def _reload() -> None:
        pq = state["pq"]
        if pq:
            if _dirty() and not messagebox.askyesno("Discard changes", "Discard unsaved changes and reload from disk?"):
                return
            _open(str(pq.path))

    def _restore_backup() -> None:
        pq: ExcelPowerQuery = state["pq"]
        if not pq or not pq.backup_exists():
            messagebox.showinfo("No Backup", "No .bak file exists for this workbook.")
            return
        bak_file = pq.path.with_suffix(pq.path.suffix + ".bak")
        if messagebox.askyesno("Restore Backup", f"Restore workbook from:\n{bak_file.name}?\n\nThis will replace the current file with the backup state."):
            try:
                pq.restore_backup()
                _open(str(pq.path))
                messagebox.showinfo("Restored", f"Workbook restored successfully from {bak_file.name}.")
            except PermissionError:
                messagebox.showerror(
                    "File Locked by Excel",
                    f"Cannot restore '{pq.path.name}' because it is currently open in Excel.\n\n"
                    "Please close the file in Excel first, then try restoring again."
                )
            except Exception as e:
                messagebox.showerror("Restore Failed", str(e))

    def _save(as_new: bool) -> None:
        pq: ExcelPowerQuery = state["pq"]
        if not pq:
            return
        new_m = editor.get("1.0", "end-1c")

        # 1. Pre-Save Syntax Validation Check
        syntax_errors = validate_m_syntax(new_m)
        if syntax_errors:
            msg = "Syntax analysis detected potential issues in your Power Query M code:\n\n"
            msg += "\n".join(f"• {err}" for err in syntax_errors[:6])
            if len(syntax_errors) > 6:
                msg += f"\n... and {len(syntax_errors) - 6} more errors."
            msg += "\n\nSaving invalid M code may cause Excel to fail on open or refresh.\nDo you want to save anyway?"
            if not messagebox.askyesno("Syntax Warning Before Save", msg, icon="warning"):
                m_line = re.search(r"Line (\d+)", syntax_errors[0])
                if m_line:
                    l_target = f"{m_line.group(1)}.0"
                    editor.see(l_target)
                    editor.mark_set("insert", l_target)
                    _redraw_line_numbers()
                return

        out = None
        if as_new:
            out = filedialog.asksaveasfilename(
                title="Save edited workbook as", defaultextension=pq.path.suffix,
                initialfile=pq.path.stem + "_edited" + pq.path.suffix, filetypes=FILETYPES
            )
            if not out:
                return

        try:
            dest = pq.write(new_m, out_path=out)
        except PermissionError:
            target_name = Path(out).name if out else pq.path.name
            msg = (
                f"Cannot save to '{target_name}' because the file is currently open and locked by Excel.\n\n"
                f"Please close '{target_name}' in Excel and try saving again,\n"
                f"or use 'Save As…' to save your changes to a new file."
            )
            messagebox.showerror("File Locked by Excel", msg)
            status.set(f"Save failed: '{target_name}' is locked by Excel.")
            return
        except PowerQueryError as e:
            messagebox.showerror("Save failed", str(e))
            status.set(f"Save failed: {e}")
            return
        except Exception as e:
            messagebox.showerror("Unexpected error", repr(e))
            status.set(f"Save failed: {e}")
            return

        state["loaded_text"] = new_m
        if not as_new:
            state["pq"] = ExcelPowerQuery(str(dest))
        _set_enabled(True)
        _update_status()
        messagebox.showinfo("Saved", f"Power Query updated safely:\n{dest}\n\nBackup (.bak) preserved. Open in Excel and Refresh to apply.")
        status.set(f"Saved safely → {dest.name}")

    # ---------- EXPORTS ----------
    def _export_single() -> None:
        pq = state["pq"]
        if not pq: return
        default = (pq.path.stem if pq else "queries") + ".m"
        out = filedialog.asksaveasfilename(
            title="Export all M code to one file",
            defaultextension=".m",
            initialfile=default,
            filetypes=[
                ("Power Query M (*.m)", "*.m"),
                ("Power Query (*.pq)", "*.pq"),
                ("Text File (*.txt)", "*.txt"),
                ("All files (*.*)", "*.*")
            ]
        )
        if not out: return
        try:
            Path(out).write_text(editor.get("1.0", "end-1c"), encoding="utf-8")
        except Exception as e:
            messagebox.showerror("Export failed", repr(e)); return
        status.set(f"Exported single file → {out}")
        messagebox.showinfo("Exported", f"All queries written to:\n{out}")

    def _export_split() -> None:
        pq = state["pq"]
        if not pq: return
        bodies = query_bodies(editor.get("1.0", "end-1c"))
        if not bodies:
            messagebox.showinfo("Nothing to split", "No queries were detected."); return

        dlg = tk.Toplevel(root)
        dlg.title("Export Split Queries")
        dlg.geometry("360x220")
        dlg.resizable(False, False)
        dlg.transient(root)
        dlg.grab_set()

        dlg.update_idletasks()
        x = root.winfo_x() + (root.winfo_width() - 360) // 2
        y = root.winfo_y() + (root.winfo_height() - 220) // 2
        dlg.geometry(f"+{x}+{y}")

        ttk.Label(dlg, text=f"Export {len(bodies)} queries to separate files", font=("Segoe UI", 10, "bold")).pack(pady=(12, 6))

        ext_frame = ttk.LabelFrame(dlg, text="Select File Extension", padding=10)
        ext_frame.pack(fill="x", padx=16, pady=4)

        ext_var = tk.StringVar(value=".pq")
        ttk.Radiobutton(ext_frame, text=".pq   (Power Query file)", value=".pq", variable=ext_var).pack(anchor="w", pady=2)
        ttk.Radiobutton(ext_frame, text=".m    (Power Query M script)", value=".m", variable=ext_var).pack(anchor="w", pady=2)
        ttk.Radiobutton(ext_frame, text=".txt  (Plain text file)", value=".txt", variable=ext_var).pack(anchor="w", pady=2)

        def _do_split_export():
            ext = ext_var.get()
            dlg.destroy()
            folder = filedialog.askdirectory(title=f"Choose a folder for split {ext} files")
            if not folder: return
            dest = Path(folder)
            written = 0
            try:
                for name, rhs in bodies:
                    safe = re.sub(r'[^\w.\- ]', "_", name).strip() or f"query_{written}"
                    (dest / f"{safe}{ext}").write_text(f"// Power Query: {name}\n{rhs}\n", encoding="utf-8")
                    written += 1
                status.set(f"Split export → {written} file(s) ({ext}) in {dest}")
                messagebox.showinfo("Exported", f"Wrote {written} {ext} file(s) to:\n{dest}")
            except Exception as e:
                messagebox.showerror("Split export failed", repr(e))

        btn_frame = ttk.Frame(dlg)
        btn_frame.pack(fill="x", padx=16, pady=(12, 8))
        ttk.Button(btn_frame, text="Choose Folder & Export…", command=_do_split_export).pack(side="right", padx=(6, 0))
        ttk.Button(btn_frame, text="Cancel", command=dlg.destroy).pack(side="right")

    # ---------- DEPENDENCIES & GRAPH VISUALIZATION ----------
    def _get_dep_data():
        return build_dependencies(editor.get("1.0", "end-1c"))

    def _copy_mermaid_direct() -> None:
        names, deps, join_details, field_details = _get_dep_data()
        if not names:
            messagebox.showinfo("No queries", "No queries detected to visualize.")
            return
        mmd = dependencies_to_mermaid(names, deps, join_details, field_details)
        root.clipboard_clear()
        root.clipboard_append(mmd)
        status.set("Mermaid code copied to clipboard.")
        messagebox.showinfo("Copied", "Mermaid diagram code copied to clipboard!")

    def _save_mermaid_direct() -> None:
        names, deps, join_details, field_details = _get_dep_data()
        if not names:
            messagebox.showinfo("No queries", "No queries detected to visualize.")
            return
        mmd = dependencies_to_mermaid(names, deps, join_details, field_details)
        fp = filedialog.asksaveasfilename(
            title="Save Mermaid diagram", defaultextension=".mmd",
            initialfile="query_dependencies.mmd",
            filetypes=[("Mermaid", "*.mmd *.md"), ("Text", "*.txt")]
        )
        if fp:
            Path(fp).write_text(mmd, encoding="utf-8")
            status.set(f"Mermaid saved → {fp}")

    def _open_browser_graph() -> None:
        names, deps, join_details, field_details = _get_dep_data()
        if not names:
            messagebox.showinfo("No queries", "No queries detected to visualize.")
            return

        mmd = dependencies_to_mermaid(names, deps, join_details, field_details)
        html_code = f"""<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8">
    <title>PQ Editor — Dependencies & Join Graph</title>
    <script src="https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.min.js"></script>
    <style>
        body {{
            font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
            margin: 0; padding: 24px; background: #FAF9F8; color: #201F1E;
        }}
        .header {{
            display: flex; align-items: center; justify-content: space-between;
            border-bottom: 2px solid #0E703C; padding-bottom: 12px; margin-bottom: 20px;
        }}
        h2 {{ margin: 0; color: #0E703C; font-weight: 600; }}
        .badge {{
            background: #0E703C; color: white; padding: 5px 12px; border-radius: 14px;
            font-size: 13px; font-weight: bold;
        }}
        .card {{
            background: white; padding: 20px; border-radius: 8px;
            box-shadow: 0 2px 10px rgba(0,0,0,0.08); overflow: auto; min-height: 480px;
        }}
        .legend {{
            margin-top: 15px; font-size: 12px; color: #605E5C; display: flex; gap: 20px;
        }}
        .legend span {{ display: inline-flex; align-items: center; gap: 6px; }}
        .box-src {{ width: 14px; height: 14px; background: #E2EFDA; border: 1.5px solid #375623; border-radius: 3px; }}
        .box-mid {{ width: 14px; height: 14px; background: #F8F9FA; border: 1.5px solid #0E703C; border-radius: 3px; }}
    </style>
</head>
<body>
    <div class="header">
        <h2>PQ Editor — Dataflow & Table Joins</h2>
        <span class="badge">{len(names)} Queries &bull; {sum(len(v) for v in deps.values())} Connections</span>
    </div>
    <div class="card">
        <pre class="mermaid">
{html.escape(mmd)}
        </pre>
    </div>
    <div class="legend">
        <span><div class="box-src"></div> Source Queries (No upstream)</span>
        <span><div class="box-mid"></div> Derived / Transformed Queries</span>
        <span>Edges annotated with Join Kind &amp; Key Fields</span>
    </div>
    <script>
        mermaid.initialize({{ startOnLoad: true, theme: 'default', flowchart: {{ curve: 'basis', useMaxWidth: false }} }});
    </script>
</body>
</html>"""
        try:
            tmp = tempfile.NamedTemporaryFile("w", delete=False, suffix=".html", encoding="utf-8")
            tmp.write(html_code)
            tmp.close()
            webbrowser.open(f"file://{Path(tmp.name).resolve()}")
            status.set("Opened interactive graph in web browser.")
        except Exception as e:
            messagebox.showerror("Cannot open browser", str(e))

    def _show_dependencies() -> None:
        names, deps, _join_details, _field_details = _get_dep_data()
        if not names:
            messagebox.showinfo("No queries", "No queries were detected to analyze."); return

        t = THEMES[state["theme_name"]]
        win = tk.Toplevel(root)
        win.title("PQ Editor — Query Dependencies")
        win.geometry("960x640")
        win.minsize(620, 420)
        win.columnconfigure(0, weight=1)
        win.rowconfigure(1, weight=1)

        top = ttk.Frame(win, padding=(10, 8, 10, 6))
        top.grid(row=0, column=0, sticky="ew")
        top.columnconfigure(0, weight=1)

        edges = sum(len(v) for v in deps.values())
        ttk.Label(
            top, text=f"{len(names)} queries · {edges} connections  (Pan: Left-Drag · Zoom: Mouse-Wheel · Double-click node to jump)",
            font=("Segoe UI", 9)
        ).grid(row=0, column=0, sticky="w")

        mermaid = dependencies_to_mermaid(names, deps, _join_details, _field_details)

        def _copy_mermaid():
            win.clipboard_clear(); win.clipboard_append(mermaid); status.set("Mermaid copied to clipboard.")
        def _save_mermaid():
            fp = filedialog.asksaveasfilename(
                title="Save Mermaid diagram", defaultextension=".mmd",
                initialfile="query_dependencies.mmd",
                filetypes=[("Mermaid", "*.mmd *.md"), ("Text", "*.txt")]
            )
            if fp: Path(fp).write_text(mermaid, encoding="utf-8"); status.set(f"Mermaid saved → {fp}")

        ttk.Button(top, text="🌐 Web Graph", command=_open_browser_graph).grid(row=0, column=1, padx=4)
        ttk.Button(top, text="📋 Copy Mermaid", command=_copy_mermaid).grid(row=0, column=2, padx=4)
        ttk.Button(top, text="💾 Save .mmd", command=_save_mermaid).grid(row=0, column=3)

        cwrap = ttk.Frame(win, padding=(10, 0, 10, 10))
        cwrap.grid(row=1, column=0, sticky="nsew")
        cwrap.columnconfigure(0, weight=1)
        cwrap.rowconfigure(0, weight=1)

        canvas = tk.Canvas(cwrap, bg=t["canvas_bg"], highlightthickness=0)
        canvas.grid(row=0, column=0, sticky="nsew")
        cv = ttk.Scrollbar(cwrap, orient="vertical", command=canvas.yview)
        cv.grid(row=0, column=1, sticky="ns")
        ch = ttk.Scrollbar(cwrap, orient="horizontal", command=canvas.xview)
        ch.grid(row=1, column=0, sticky="ew")
        canvas.configure(yscrollcommand=cv.set, xscrollcommand=ch.set)

        # Pan and Drag
        canvas.bind("<ButtonPress-1>", lambda e: canvas.scan_mark(e.x, e.y))
        canvas.bind("<B1-Motion>", lambda e: canvas.scan_dragto(e.x, e.y, gain=1))

        # Mouse-wheel zoom
        def _on_zoom(e):
            scale = 1.1 if (e.delta > 0 or getattr(e, "num", 0) == 4) else 0.9
            canvas.scale("all", e.x, e.y, scale, scale)
            canvas.configure(scrollregion=canvas.bbox("all"))

        canvas.bind("<MouseWheel>", _on_zoom)
        canvas.bind("<Button-4>", _on_zoom)
        canvas.bind("<Button-5>", _on_zoom)

        # Layered Layout
        layers = compute_layers(names, deps)
        by_layer: dict[int, list[str]] = {}
        for n in names: by_layer.setdefault(layers[n], []).append(n)
        COLW, ROWH, PADX, PADY, H = 260, 70, 50, 40, 36
        pos: dict[str, tuple[int, int, int, int]] = {}
        for L in sorted(by_layer):
            for row, n in enumerate(by_layer[L]):
                w = max(110, len(n) * 9 + 30)
                pos[n] = (PADX + L * COLW, PADY + row * ROWH, w, H)

        # Draw Clean Edges (No cluttering text in canvas)
        for n in names:
            for d in deps[n]:
                if d not in pos or n not in pos: continue
                x1, y1, w1, h1 = pos[d]; x2, y2, w2, h2 = pos[n]
                sx, sy = x1 + w1, y1 + h1 / 2
                ex, ey = x2, y2 + h2 / 2
                mx = (sx + ex) / 2 if ex > sx else sx + 25

                canvas.create_line(sx, sy, mx, sy, mx, ey, ex, ey,
                                   arrow="last", fill=t["line_color"], width=1.5,
                                   joinstyle="miter", capstyle="projecting")

        # Draw Nodes
        node_tag = {}
        for i, n in enumerate(names):
            x, y, w, h = pos[n]
            fill = t["node_src_fill"] if not deps[n] else t["node_dest_fill"]
            outline = t["node_src_outline"] if not deps[n] else t["node_dest_outline"]
            tag = f"node{i}"
            node_tag[tag] = n
            canvas.create_rectangle(x, y, x + w, y + h, fill=fill, outline=outline, width=1.8, tags=(tag,))
            canvas.create_text(x + w / 2, y + h / 2, text=n, font=("Segoe UI", 9, "bold"), fill=t["node_text"], tags=(tag,))
            canvas.tag_bind(tag, "<Enter>", lambda e: canvas.configure(cursor="hand2"))
            canvas.tag_bind(tag, "<Leave>", lambda e: canvas.configure(cursor=""))

        def _on_node_dblclick(event):
            item = canvas.find_withtag("current")
            if not item: return
            for tg in canvas.gettags(item[0]):
                if tg in node_tag:
                    _jump_to_query(node_tag[tg])
                    root.deiconify()
                    root.lift()
                    root.focus_force()
                    return
        canvas.tag_bind("all", "<Double-Button-1>", _on_node_dblclick)

        bbox = canvas.bbox("all")
        if bbox:
            canvas.configure(scrollregion=(bbox[0] - 40, bbox[1] - 40, bbox[2] + 40, bbox[3] + 40))

    def _on_close() -> None:
        if _dirty() and not messagebox.askyesno("Unsaved changes", "Discard unsaved M code changes?"):
            return
        root.destroy()

    # Initial theme application (Light by default)
    _apply_theme("light")

    root.protocol("WM_DELETE_WINDOW", _on_close)
    root.mainloop()


if __name__ == "__main__":
    launch_gui()
