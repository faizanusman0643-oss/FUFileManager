"""Build a genuine (minimal) vbaProject.bin for tests, following MS-OVBA / MS-CFB.

olefile can only read compound files, so this module contains a small
Compound File Binary (CFB v3, 512-byte sectors) writer plus the MS-OVBA
compression algorithm (2.4.1) and the dir / PROJECT / _VBA_PROJECT streams.
Module streams hold compressed source only (MODULEOFFSET = 0, no p-code),
and _VBA_PROJECT declares version 0xFFFF so hosts recompile from source.

Usage:
    data = build_vba_project([
        ("ThisWorkbook", "document", code),
        ("Sheet1", "document", code),
        ("Module1", "standard", code),
        ("Class1", "class", code),
    ])
"""
from __future__ import annotations

import struct
import uuid

__all__ = ["compress", "build_vba_project", "write_cfb", "module_header"]

# --------------------------------------------------------------------------
# MS-OVBA 2.4.1 compression
# --------------------------------------------------------------------------
_CHUNK = 4096


def _copy_token_help(pos_in_chunk: int):
    bit_count = 4
    while (1 << bit_count) < pos_in_chunk:
        bit_count += 1
    bit_count = max(bit_count, 4)
    length_mask = 0xFFFF >> bit_count
    offset_mask = ~length_mask & 0xFFFF
    maximum_length = length_mask + 3
    return length_mask, offset_mask, bit_count, maximum_length


def _compress_chunk(data: bytes) -> bytes:
    out = bytearray()
    pos = 0
    n = len(data)
    while pos < n:
        flag_index = len(out)
        out.append(0)
        flags = 0
        for bit in range(8):
            if pos >= n:
                break
            # find the longest match in the already-processed part of the chunk
            best_len = 0
            best_off = 0
            if pos > 0:
                _, _, bit_count, max_len = _copy_token_help(pos)
                max_len = min(max_len, n - pos)
                start = 0
                cand = pos - 1
                while cand >= start:
                    ln = 0
                    while ln < max_len and data[cand + ln] == data[pos + ln]:
                        ln += 1
                    if ln > best_len:
                        best_len, best_off = ln, pos - cand
                        if ln == max_len:
                            break
                    cand -= 1
            if best_len >= 3:
                _, _, bit_count, _ = _copy_token_help(pos)
                token = ((best_off - 1) << (16 - bit_count)) | (best_len - 3)
                out += struct.pack("<H", token)
                flags |= 1 << bit
                pos += best_len
            else:
                out.append(data[pos])
                pos += 1
        out[flag_index] = flags
    return bytes(out)


def compress(data: bytes) -> bytes:
    """Compress bytes into an MS-OVBA CompressedContainer."""
    out = bytearray(b"\x01")
    for start in range(0, len(data), _CHUNK):
        chunk = data[start:start + _CHUNK]
        comp = _compress_chunk(chunk)
        if len(comp) <= _CHUNK:
            header = 0xB000 | (len(comp) + 2 - 3)  # compressed flag + signature 0b011
            out += struct.pack("<H", header) + comp
        else:  # raw chunk: exactly 4096 bytes
            header = 0x3000 | (4098 - 3)
            out += struct.pack("<H", header) + chunk.ljust(_CHUNK, b"\x00")
    return bytes(out)


# --------------------------------------------------------------------------
# MS-CFB writer (version 3, 512-byte sectors, mini stream for < 4096 bytes)
# --------------------------------------------------------------------------
SECTOR = 512
MINI = 64
CUTOFF = 4096
FREESECT = 0xFFFFFFFF
ENDOFCHAIN = 0xFFFFFFFE
FATSECT = 0xFFFFFFFD
NOSTREAM = 0xFFFFFFFF


class _Node:
    def __init__(self, name, is_storage, data=b""):
        self.name = name
        self.is_storage = is_storage
        self.data = data
        self.children: list[_Node] = []
        self.sid = 0
        self.left = self.right = self.child = NOSTREAM
        self.start = ENDOFCHAIN
        self.size = 0


def _sort_key(name: str):
    return (len(name), name.upper())


def write_cfb(tree: dict) -> bytes:
    """tree: nested dict {name: bytes | dict}; dicts are storages."""
    root = _Node("Root Entry", True)

    def build(node, d):
        for name, val in d.items():
            if isinstance(val, dict):
                ch = _Node(name, True)
                build(ch, val)
            else:
                ch = _Node(name, False, bytes(val))
            node.children.append(ch)

    build(root, tree)

    # assign sids (pre-order)
    nodes: list[_Node] = []

    def number(node):
        node.sid = len(nodes)
        nodes.append(node)
        for ch in node.children:
            number(ch)

    number(root)

    # balanced BST of siblings
    def bst(items):
        if not items:
            return NOSTREAM
        mid = len(items) // 2
        n = items[mid]
        n.left = bst(items[:mid])
        n.right = bst(items[mid + 1:])
        return n.sid

    for n in nodes:
        if n.is_storage:
            n.child = bst(sorted(n.children, key=lambda x: _sort_key(x.name)))

    # mini stream
    mini = bytearray()
    minifat: list[int] = []
    big_streams = []
    for n in nodes:
        if n.is_storage:
            continue
        n.size = len(n.data)
        if n.size == 0:
            n.start = ENDOFCHAIN
        elif n.size < CUTOFF:
            first = len(minifat)
            cnt = (n.size + MINI - 1) // MINI
            n.start = first
            for i in range(cnt):
                minifat.append(first + i + 1 if i < cnt - 1 else ENDOFCHAIN)
            mini += n.data.ljust(cnt * MINI, b"\x00")
        else:
            big_streams.append(n)

    # sector allocation: [data sectors...] [ministream] [minifat] [dir] [fat]
    sectors: list[bytes] = []
    fat: list[int] = []

    def alloc(data: bytes) -> int:
        if not data:
            return ENDOFCHAIN
        cnt = (len(data) + SECTOR - 1) // SECTOR
        first = len(sectors)
        padded = data.ljust(cnt * SECTOR, b"\x00")
        for i in range(cnt):
            sectors.append(padded[i * SECTOR:(i + 1) * SECTOR])
            fat.append(first + i + 1 if i < cnt - 1 else ENDOFCHAIN)
        return first

    for n in big_streams:
        n.start = alloc(n.data)
    root.start = alloc(bytes(mini)) if mini else ENDOFCHAIN
    root.size = len(mini)
    minifat_bytes = b"".join(struct.pack("<I", x) for x in minifat)
    if minifat_bytes:
        minifat_bytes = minifat_bytes + b"\xff" * ((-len(minifat_bytes)) % SECTOR)
    minifat_start = alloc(minifat_bytes) if minifat_bytes else ENDOFCHAIN
    n_minifat = len(minifat_bytes) // SECTOR

    # directory
    entries = bytearray()
    for n in nodes:
        name16 = n.name.encode("utf-16-le") + b"\x00\x00"
        typ = 5 if n is root else (1 if n.is_storage else 2)
        e = name16.ljust(64, b"\x00")
        e += struct.pack("<HBB", len(name16), typ, 1)
        e += struct.pack("<III", n.left, n.right, n.child)
        e += b"\x00" * 16 + b"\x00" * 4 + b"\x00" * 16
        start = n.start if (not n.is_storage or n is root) else 0
        size = n.size if (not n.is_storage or n is root) else 0
        e += struct.pack("<IQ", start, size)
        assert len(e) == 128
        entries += e
    while len(entries) % SECTOR:
        empty = b"\x00" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM)
        entries += empty.ljust(128, b"\x00")
    dir_start = alloc(bytes(entries))

    # FAT sectors (must cover themselves)
    n_fat = 1
    while True:
        total = len(sectors) + n_fat
        if (total + 127) // 128 <= n_fat:
            break
        n_fat += 1
    if n_fat > 109:
        raise ValueError("file too large for this minimal CFB writer")
    fat_start = len(sectors)
    for i in range(n_fat):
        fat.append(FATSECT)
    fat += [FREESECT] * (n_fat * 128 - len(fat))
    fat_bytes = b"".join(struct.pack("<I", x) for x in fat)
    for i in range(n_fat):
        sectors.append(fat_bytes[i * SECTOR:(i + 1) * SECTOR])

    header = bytearray()
    header += bytes.fromhex("D0CF11E0A1B11AE1")
    header += b"\x00" * 16
    header += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6)
    header += b"\x00" * 6
    header += struct.pack("<I", 0)              # num dir sectors (0 for v3)
    header += struct.pack("<I", n_fat)
    header += struct.pack("<I", dir_start)
    header += struct.pack("<I", 0)              # transaction signature
    header += struct.pack("<I", CUTOFF)
    header += struct.pack("<I", minifat_start)
    header += struct.pack("<I", n_minifat)
    header += struct.pack("<I", ENDOFCHAIN)     # first DIFAT sector
    header += struct.pack("<I", 0)              # num DIFAT sectors
    difat = [fat_start + i for i in range(n_fat)] + [FREESECT] * (109 - n_fat)
    header += b"".join(struct.pack("<I", x) for x in difat)
    assert len(header) == SECTOR
    return bytes(header) + b"".join(sectors)


# --------------------------------------------------------------------------
# MS-OVBA streams
# --------------------------------------------------------------------------
_DOC_BASE = {
    "workbook": "0{00020819-0000-0000-C000-000000000046}",
    "sheet": "0{00020820-0000-0000-C000-000000000046}",
}


def module_header(name: str, kind: str, is_workbook: bool = False) -> str:
    """The 'Attribute VB_...' lines the VBE stores in front of a module."""
    lines = [f'Attribute VB_Name = "{name}"']
    if kind == "document":
        base = _DOC_BASE["workbook" if is_workbook else "sheet"]
        lines += [f'Attribute VB_Base = "{base}"',
                  "Attribute VB_GlobalNameSpace = False",
                  "Attribute VB_Creatable = False",
                  "Attribute VB_PredeclaredId = True",
                  "Attribute VB_Exposed = True",
                  "Attribute VB_TemplateDerived = False",
                  "Attribute VB_Customizable = True"]
    elif kind == "form":
        lines += ['Attribute VB_Base = "0{C62A69F0-16DC-11CE-9E98-00AA00574A4F}'
                  '{12345678-1234-1234-1234-123456789ABC}"',
                  "Attribute VB_GlobalNameSpace = False",
                  "Attribute VB_Creatable = False",
                  "Attribute VB_PredeclaredId = True",
                  "Attribute VB_Exposed = False",
                  "Attribute VB_TemplateDerived = False",
                  "Attribute VB_Customizable = False"]
    elif kind == "class":
        lines += ["Attribute VB_GlobalNameSpace = False",
                  "Attribute VB_Creatable = False",
                  "Attribute VB_PredeclaredId = False",
                  "Attribute VB_Exposed = False"]
    return "\r\n".join(lines) + "\r\n"


def _rec(rid: int, payload: bytes) -> bytes:
    return struct.pack("<HI", rid, len(payload)) + payload


def _dir_stream(project_name: str, modules, codepage: int = 1252) -> bytes:
    enc = f"cp{codepage}"
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 1))           # SYSKIND win32
    d += _rec(0x0002, struct.pack("<I", 0x409))       # LCID
    d += _rec(0x0014, struct.pack("<I", 0x409))       # LCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", codepage))    # CODEPAGE
    d += _rec(0x0004, project_name.encode(enc))       # NAME
    d += _rec(0x0005, b"") + _rec(0x0040, b"")        # DOCSTRING
    d += _rec(0x0006, b"") + _rec(0x003D, b"")        # HELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))           # HELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))           # LIBFLAGS
    # VERSION: Id, Reserved(=4), VersionMajor u32, VersionMinor u16
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)
    d += _rec(0x000C, b"") + _rec(0x003C, b"")        # CONSTANTS
    # (no references)
    d += _rec(0x000F, struct.pack("<H", len(modules)))  # MODULES count
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))        # PROJECTCOOKIE
    for name, kind, _code in modules:
        d += _rec(0x0019, name.encode(enc))
        d += _rec(0x0047, name.encode("utf-16-le"))
        d += _rec(0x001A, name.encode(enc)) + _rec(0x0032, name.encode("utf-16-le"))
        d += _rec(0x001C, b"") + _rec(0x0048, b"")
        d += _rec(0x0031, struct.pack("<I", 0))          # MODULEOFFSET: source at 0
        d += _rec(0x001E, struct.pack("<I", 0))
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))
        d += _rec(0x0021 if kind == "standard" else 0x0022, b"")
        d += _rec(0x002B, b"")                            # module terminator
    d += _rec(0x0010, b"")                                # dir terminator
    return bytes(d)


def _project_stream(project_name: str, modules) -> bytes:
    lines = ['ID="{%s}"' % str(uuid.UUID(int=0x1234)).upper()]
    for name, kind, _ in modules:
        if kind == "document":
            lines.append(f"Document={name}/&H00000000")
        elif kind == "standard":
            lines.append(f"Module={name}")
        elif kind == "class":
            lines.append(f"Class={name}")
        elif kind == "form":
            lines.append(f"BaseClass={name}")
            lines.append("Package={AC9F2F90-E877-11CE-9F68-00AA00574A4F}")
    lines += [f'Name="{project_name}"', 'HelpContextID="0"',
              'VersionCompatible32="393222000"', "",
              "[Host Extender Info]",
              "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
              "", "[Workspace]"]
    for name, _, _ in modules:
        lines.append(f"{name}=0, 0, 0, 0, C")
    return ("\r\n".join(lines) + "\r\n").encode("cp1252")


def build_vba_project(modules, project_name: str = "VBAProject", forms=None) -> bytes:
    """modules: list of (name, kind, code) with kind in standard|document|class|form.

    `code` is stored as-is (converted to CRLF); include the Attribute header
    (see module_header()) to look like a real VBE-written project.
    forms: optional {form module name: FORM_DICT-like dict} (see
    build_form_designer); each gets a designer storage next to the VBA storage.
    """
    vba = {
        # Reserved1 0x61CC, Version 0xFFFF (no p-code => recompile), Reserved2 0x00, Reserved3
        "_VBA_PROJECT": struct.pack("<HHBH", 0x61CC, 0xFFFF, 0x00, 0x0000),
        "dir": compress(_dir_stream(project_name, modules)),
    }
    for name, _kind, code in modules:
        text = code.replace("\r\n", "\n").replace("\n", "\r\n")
        vba[name] = compress(text.encode("cp1252"))
    tree = {
        "VBA": vba,
        "PROJECT": _project_stream(project_name, modules),
        "PROJECTwm": b"".join(n.encode("cp1252") + b"\x00" + n.encode("utf-16-le") + b"\x00\x00"
                              for n, _, _ in modules) + b"\x00\x00",
    }
    for name, form in (forms or {}).items():
        # a dict that already has an "f" stream is a prebuilt designer storage
        tree[name] = form if "f" in form else build_form_designer(name, form)
    return write_cfb(tree)


# --------------------------------------------------------------------------
# MS-OFORMS designer writer (UserForm storages)
# --------------------------------------------------------------------------
# Written independently of portable_excel.fileio from [MS-OFORMS]: every
# record is MinorVersion(1)=0, MajorVersion(1), cb(2), PropMask, DataBlock
# (fixed-size values in bit order, each aligned to its size, block padded to
# 4), ExtraDataBlock (strings / 8-byte pairs / string arrays, each padded to 4).
_HM = 2540.0 / 72.0
_GUID_STDFONT = bytes.fromhex("0352e30b918fce119de300aa004bb851")
_FMT = {"u8": "<B", "u16": "<H", "i16": "<h", "u32": "<I", "i32": "<i", "pic": "<H"}
_CLSID = {"Label": 21, "TextBox": 23, "ListBox": 24, "ComboBox": 25, "CheckBox": 26,
          "OptionButton": 27, "ToggleButton": 28, "CommandButton": 17, "Frame": 14,
          "SpinButton": 16, "ScrollBar": 47, "Image": 12, "TabStrip": 18, "MultiPage": 57}
_DISPLAY = {"TextBox": 1, "ListBox": 2, "ComboBox": 3, "CheckBox": 4, "OptionButton": 5,
            "ToggleButton": 6}


def _hm(points) -> int:
    return int(round(float(points) * _HM))


def _pad4(b: bytes) -> bytes:
    return b + b"\x00" * (-len(b) % 4)


def _ms_string(text: str):
    """-> (CountOfBytesWithCompressionFlag, bytes)"""
    if all(ord(ch) < 256 for ch in text):
        raw = text.encode("latin-1")
        return 0x80000000 | len(raw), raw
    raw = text.encode("utf-16-le")
    return len(raw), raw


def _props(items) -> tuple:
    """items: [(bit, kind, value)] -> (mask, DataBlock + ExtraDataBlock)."""
    mask = 0
    data = bytearray()
    extra = bytearray()
    for bit, kind, value in sorted(items, key=lambda t: t[0]):
        mask |= 1 << bit
        if kind == "flag":
            continue
        if kind == "pair":
            extra += struct.pack("<ii", int(value[0]), int(value[1]))
            continue
        if kind == "str":
            count, raw = _ms_string(value)
            data += b"\x00" * (-len(data) % 4)
            data += struct.pack("<I", count)
            extra += _pad4(raw)
            continue
        if kind == "arr":
            arr = bytearray()
            for text in value:
                count, raw = _ms_string(text)
                arr += struct.pack("<I", count) + _pad4(raw)
            data += b"\x00" * (-len(data) % 4)
            data += struct.pack("<I", len(arr))
            extra += arr
            continue
        fmt = _FMT[kind]
        size = struct.calcsize(fmt)
        data += b"\x00" * (-len(data) % size)
        data += struct.pack(fmt, value)
    data += b"\x00" * (-len(data) % 4)
    return mask, bytes(data) + bytes(extra)


def _record(major: int, items, mask_fmt: str = "<I") -> bytes:
    mask, body = _props(items)
    if mask_fmt == "<Q":
        mask |= 1 << 31  # MorphDataPropMask.Reserved
    payload = struct.pack(mask_fmt, mask) + body
    return struct.pack("<BBH", 0, major, len(payload)) + payload


def _text_props(font_name="Tahoma", height_twips=165, align=None, bold=False) -> bytes:
    items = [(0, "str", font_name), (2, "u32", height_twips), (5, "u8", 2)]
    if bold:
        items.append((1, "u32", 1))
    if align is not None:
        items.append((6, "u8", align))
    return _record(2, items)


def _std_font(name="Tahoma", size_pt=8.25) -> bytes:
    face = name.encode("latin-1")
    return (_GUID_STDFONT + struct.pack("<BHBHI", 1, 0, 0, 400, int(round(size_pt * 10000)))
            + struct.pack("<B", len(face)) + face)


def _ole_color(hex_rgb: str) -> int:
    h = hex_rgb.lstrip("#")
    r, g, b = int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)
    return r | (g << 8) | (b << 16)


def _control_object(c: dict) -> bytes:
    """The control's record in the "o" stream (Frame / MultiPage: none)."""
    t = c["type"]
    size = (_hm(c.get("width", 72)), _hm(c.get("height", 18)))
    align = {1: 1, 2: 3, 3: 2}.get(c.get("text_align")) if c.get("text_align") else None
    if t in _DISPLAY:
        flags = 0x2C80081B
        if c.get("multi_line"):
            flags |= 0x80000000
        if c.get("enabled") is False:
            flags &= ~0x2
        items = [(0, "u32", flags), (8, "pair", size)]
        style = 7 if (t == "ComboBox" and c.get("style") == 2) else _DISPLAY[t]
        items.append((6, "u8", style))
        if c.get("back_color"):
            items.append((1, "u32", _ole_color(c["back_color"])))
        if c.get("fore_color"):
            items.append((2, "u32", _ole_color(c["fore_color"])))
        if c.get("max_length"):
            items.append((3, "u32", c["max_length"]))
        if c.get("scroll_bars"):
            items.append((5, "u8", c["scroll_bars"]))
        if c.get("password_char"):
            items.append((9, "u16", ord(c["password_char"])))
        if c.get("column_count", 1) != 1:
            items.append((13, "i16", c["column_count"]))
        if c.get("list_style"):
            items.append((17, "u8", c["list_style"]))
        if c.get("multi_select"):
            items.append((21, "u8", c["multi_select"]))
        value = c.get("value")
        if isinstance(value, bool):
            value = "1" if value else "0"
        if value not in (None, ""):
            items.append((22, "str", str(value)))
        if c.get("caption"):
            items.append((23, "str", c["caption"]))
        if c.get("group_name"):
            items.append((32, "str", c["group_name"]))
        return _record(2, items, "<Q") + _text_props(align=align)
    if t == "Label":
        items = [(3, "str", c.get("caption", "")), (5, "pair", size)]
        if c.get("fore_color"):
            items.append((0, "u32", _ole_color(c["fore_color"])))
        if c.get("back_color"):
            items.append((1, "u32", _ole_color(c["back_color"])))
        return _record(2, items) + _text_props(align=align, bold=c.get("bold", False))
    if t == "CommandButton":
        items = [(3, "str", c.get("caption", "")), (5, "pair", size)]
        if c.get("back_color"):
            items.append((1, "u32", _ole_color(c["back_color"])))
        return _record(2, items) + _text_props(align=3)
    if t in ("SpinButton", "ScrollBar"):
        items = [(3, "pair", size), (5, "i32", c.get("min", 0)), (6, "i32", c.get("max", 100)),
                 (7, "i32", c.get("value", 0))]
        if t == "SpinButton":
            items.append((10, "i32", c.get("small_change", 1)))
        else:
            items.append((11, "i32", c.get("small_change", 1)))
            items.append((12, "i32", c.get("large_change", 1)))
        return _record(2, items)
    if t == "TabStrip":
        tabs = c.get("tabs", [])
        items = [(0, "i32", c.get("value", 0)), (4, "pair", size),
                 (5, "arr", [tb["caption"] for tb in tabs]),
                 (17, "arr", [tb["name"] for tb in tabs]), (22, "u32", len(tabs))]
        return _record(2, items) + _text_props(align=3) + b"\x03\x00\x00\x00" * len(tabs)
    if t == "Image":
        return _record(2, [(9, "pair", size)])
    raise ValueError(f"fixture writer does not support {t}")


def _site(c: dict, site_id: int, tab_index: int, obj_size: int, streamed: bool) -> bytes:
    flags = 0x2 if c.get("visible", True) else 0
    if c.get("tab_stop", c["type"] != "Label"):
        flags |= 0x1
    if c.get("default"):
        flags |= 0x4
    if c.get("cancel"):
        flags |= 0x8
    if streamed:
        flags |= 0x10
    items = [(0, "str", c["name"]), (2, "u32", site_id), (4, "u32", flags),
             (6, "i16", tab_index), (7, "u16", _CLSID[c["type"]]),
             (8, "pair", (_hm(c.get("left", 0)), _hm(c.get("top", 0))))]
    if obj_size:
        items.append((5, "u32", obj_size))
    if c.get("tag"):
        items.append((1, "str", c["tag"]))
    if c.get("control_tip_text"):
        items.append((11, "str", c["control_tip_text"]))
    return _record(0, items)  # OleSiteConcrete: Version(2)=0 -> bytes 00 00


def _form_control(sites, caption=None, size=None, back_color=None, nested=False,
                  next_id=1) -> bytes:
    items = [(3, "u32", next_id), (20, "pic", 0xFFFF)]  # fFont: 2-byte marker
    if nested:
        items.append((6, "u32", 0x8004))  # enabled + DONTSAVECLASSTABLE
    if size is not None:
        items += [(10, "pair", size), (11, "pair", (0, 0))]
    if caption is not None:
        items.append((19, "str", caption))
    if back_color:
        items.append((1, "u32", _ole_color(back_color)))
    rec = _record(4, items)
    out = bytearray(rec) + _std_font()
    if not nested:
        out += struct.pack("<H", 0)  # CountOfSiteClassInfo
    depth = bytearray()
    n = len(sites)
    left = n
    while left > 0:
        k = min(left, 127)
        depth += bytes([0, 0x80 | k, 1]) if k > 1 else bytes([0, 1])
        left -= k
    depth += b"\x00" * (-len(depth) % 4)
    body = bytes(depth) + b"".join(sites)
    out += struct.pack("<II", n, len(body)) + body
    return bytes(out)


def _container(controls, parent, ids, nested, caption=None, size=None, back_color=None):
    """-> storage dict with f, o and nested i<ID> storages."""
    storage: dict = {}
    sites, objects = [], bytearray()
    mine = [c for c in controls if c.get("parent") == parent]
    for tab_index, c in enumerate(mine):
        site_id = ids[0]
        ids[0] += 1
        if c["type"] == "Frame":
            sites.append(_site(c, site_id, tab_index, 0, False))
            storage[f"i{site_id:02d}"] = _container(
                controls, c["name"], ids, True, caption=c.get("caption", c["name"]),
                size=(_hm(c.get("width", 144)), _hm(c.get("height", 72))),
                back_color=c.get("back_color"))
            continue
        obj = _control_object(c)
        sites.append(_site(c, site_id, tab_index, len(obj), True))
        objects += obj
    storage["f"] = _form_control(sites, caption=caption, size=size, back_color=back_color,
                                 nested=nested, next_id=ids[0])
    storage["o"] = bytes(objects)
    return storage


def build_form_designer(name: str, form: dict) -> dict:
    """FORM_DICT-like dict (points; controls with "parent") -> designer storage."""
    width = float(form.get("width", 240))
    height = float(form.get("height", 180))
    client_w, client_h = width - 12.0, height - 29.25
    vbframe = (
        "VERSION 5.00\r\n"
        f"Begin {{C62A69F0-16DC-11CE-9E98-00AA00574A4F}} {name} \r\n"
        f'   Caption         =   "{form.get("caption", name).replace(chr(34), chr(34) * 2)}"\r\n'
        f"   ClientHeight    =   {int(round(client_h * 20))}\r\n"
        "   ClientLeft      =   120\r\n"
        "   ClientTop       =   465\r\n"
        f"   ClientWidth     =   {int(round(client_w * 20))}\r\n"
        f"   StartUpPosition =   {form.get('start_up_position', 1)}  'CenterOwner\r\n"
        "   TypeInfoVer     =   1\r\n"
        "End\r\n").encode("cp1252")
    ids = [1]
    storage = _container(form.get("controls", []), None, ids, False,
                         size=(_hm(client_w), _hm(client_h)), back_color=form.get("back_color"))
    storage["\x03VBFrame"] = vbframe
    storage["\x01CompObj"] = (b"\x01\x00\xfe\xff\x03\x0a\x00\x00\xff\xff\xff\xff"
                              + bytes.fromhex("f0692ac6dc16ce119e9800aa00574a4f")
                              + struct.pack("<I", 0) + struct.pack("<I", 0)
                              + struct.pack("<I", 0x71B239F4))
    return storage
