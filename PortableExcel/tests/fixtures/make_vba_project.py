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
    lines += [f'Name="{project_name}"', 'HelpContextID="0"',
              'VersionCompatible32="393222000"', "",
              "[Host Extender Info]",
              "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
              "", "[Workspace]"]
    for name, _, _ in modules:
        lines.append(f"{name}=0, 0, 0, 0, C")
    return ("\r\n".join(lines) + "\r\n").encode("cp1252")


def build_vba_project(modules, project_name: str = "VBAProject") -> bytes:
    """modules: list of (name, kind, code) with kind in standard|document|class|form.

    `code` is stored as-is (converted to CRLF); include the Attribute header
    (see module_header()) to look like a real VBE-written project.
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
    return write_cfb(tree)
