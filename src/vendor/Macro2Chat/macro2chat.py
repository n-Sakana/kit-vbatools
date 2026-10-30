#!/usr/bin/env python3
"""Extract and replace workbook VBA without Excel; Python 3.10+, stdlib only.

Port of src/05_Ole2.cs through src/08_BookIO.cs (normal, complete reads).
CLI: extract BOOK DIRECTORY | rebuild BOOK DIRECTORY NEW_BOOK
Module files contain UTF-8 visible source; leading Attribute lines are preserved.
"""

import argparse
import codecs
import hashlib
import io
import json
from pathlib import Path
import re
import struct
import sys
import xml.etree.ElementTree as ET
import zipfile

OLE_SIGNATURE = bytes.fromhex("d0cf11e0a1b11ae1")
INVALIDATED_CACHE = bytes.fromhex("cc61ffff000000")
FREE, END, FAT_SECTOR, DIFAT_SECTOR = -1, -2, -3, -4
NO_STREAM = 0xFFFFFFFF
FORMATS = {".xlsm", ".xlam", ".xlsb", ".xls"}


class FormatError(ValueError):
    """The input is incomplete or cannot be rebuilt faithfully."""


def unpack(fmt, data, offset=0):
    """Read a bounded little-endian binary field at the file boundary."""
    size = struct.calcsize("<" + fmt)
    if offset < 0 or offset + size > len(data):
        raise FormatError(f"Truncated binary field at byte {offset}")
    return struct.unpack_from("<" + fmt, data, offset)


def decompress(data, offset=0):
    """MS-OVBA compressed container, with the C# strict chunk checks."""
    if not 0 <= offset < len(data) or data[offset] != 1:
        raise FormatError("Invalid VBA compressed container signature/offset")
    out = bytearray()
    pos = offset + 1
    while pos < len(data):
        header, = unpack("H", data, pos)
        size = (header & 0xFFF) + 3
        end = pos + size
        pos += 2
        if (header >> 12) & 7 != 3 or end > len(data):
            raise FormatError("Invalid or truncated VBA compressed chunk")
        start = len(out)
        if not header & 0x8000:
            if size != 4098:
                raise FormatError("Invalid raw VBA chunk size")
            out.extend(data[pos:end])
            pos = end
            continue
        while pos < end:
            flags = data[pos]
            pos += 1
            for bit in range(8):
                if pos >= end:
                    break
                count = len(out) - start
                if flags & (1 << bit):
                    if pos + 2 > end or count == 0:
                        raise FormatError("Invalid VBA copy token")
                    token, = unpack("H", data, pos)
                    pos += 2
                    bits = max(4, (count - 1).bit_length())
                    length = (token & (0xFFFF >> bits)) + 3
                    distance = (token >> (16 - bits)) + 1
                    if distance > count or count + length > 4096:
                        raise FormatError("VBA copy token exceeds chunk bounds")
                    for _ in range(length):
                        out.append(out[-distance])
                else:
                    if count >= 4096:
                        raise FormatError("VBA chunk exceeds 4096 bytes")
                    out.append(data[pos])
                    pos += 1
    return bytes(out)


def compress_chunk(data):
    """Greedy nearest-offset matching, including overlapping copy tokens."""
    out = bytearray()
    pos = 0
    while pos < len(data):
        flag_pos = len(out)
        out.append(0)
        for bit in range(8):
            if pos >= len(data):
                break
            bits = max(4, (pos - 1).bit_length())
            limit = min((0xFFFF >> bits) + 3, len(data) - pos)
            best, distance = 0, 0
            if limit >= 3 and pos:
                # rfind runs the candidate search in C; nearest match wins ties.
                candidate = data.rfind(data[pos:pos + 3], 0, pos + 2)
                while candidate >= 0:
                    length = 3
                    while length < limit and data[candidate + length] == data[pos + length]:
                        length += 1
                    if length > best:
                        best, distance = length, pos - candidate
                        if best == limit:
                            break
                    candidate = data.rfind(data[pos:pos + 3], 0, candidate + 2)
            if best >= 3:
                out[flag_pos] |= 1 << bit
                out.extend(struct.pack("<H", ((distance - 1) << (16 - bits)) | (best - 3)))
                pos += best
            else:
                out.append(data[pos])
                pos += 1
    return out


def compress(data):
    out = bytearray(b"\x01")
    pos = 0
    while pos < len(data):
        chunk = data[pos:pos + 4096]
        packed = compress_chunk(chunk)
        if len(packed) > 4096 and len(chunk) < 4096:
            chunk = chunk[:3640]
            packed = compress_chunk(chunk)
        if len(packed) > 4096:
            out.extend(b"\xff\x3f")
            out.extend(chunk)
        else:
            out.extend(struct.pack("<H", 0xB000 | (len(packed) - 1)))
            out.extend(packed)
        pos += len(chunk)
    return bytes(out)


class OleFile:
    """Compound File reader and v3 writer; retains all logical streams/metadata."""

    def __init__(self, data):
        if len(data) < 512 or data[:8] != OLE_SIGNATURE:
            raise FormatError("Not an OLE2 compound file")
        self.data = data
        minor, self.version, order, shift, mini_shift = unpack("5H", data, 24)
        if (self.version, shift) not in {(3, 9), (4, 12)} or order != 0xFFFE or mini_shift != 6:
            raise FormatError("Invalid OLE2 header")
        self.sector_size = 1 << shift
        self.sector_count = len(data) // self.sector_size - 1
        if len(data) % self.sector_size or self.sector_count < 0:
            raise FormatError("Truncated OLE2 file")
        dir_count, fat_count, dir_start, _, cutoff, mini_start, mini_count, difat_start, difat_count = unpack("IIiIIiIiI", data, 40)
        if cutoff != 4096 or (self.version == 3 and dir_count):
            raise FormatError("Invalid OLE2 allocation header")
        fat_ids = [n for n in unpack("109i", data, 76) if n != FREE]
        seen = set()
        current = difat_start
        for _ in range(difat_count):
            if current in seen:
                raise FormatError("Cyclic OLE2 DIFAT")
            seen.add(current)
            values = unpack(f"{self.sector_size // 4}i", self.sector(current))
            fat_ids.extend(n for n in values[:-1] if n != FREE)
            current = values[-1]
        if len(fat_ids) != fat_count or len(set(fat_ids)) != len(fat_ids):
            raise FormatError("Invalid OLE2 FAT count")
        if difat_count and current != END:
            raise FormatError("Unterminated OLE2 DIFAT")
        self.fat = []
        for n in fat_ids:
            self.fat.extend(unpack(f"{self.sector_size // 4}i", self.sector(n)))
        for ids, marker in ((fat_ids, FAT_SECTOR), (seen, DIFAT_SECTOR)):
            if any(n >= len(self.fat) or self.fat[n] != marker for n in ids):
                raise FormatError("Invalid OLE2 FAT sector marker")
        directory = b"".join(self.sector(n) for n in self.chain(dir_start, self.fat, self.sector_count))
        if self.version == 4 and len(directory) != dir_count * self.sector_size:
            raise FormatError("Invalid OLE2 directory count")
        self.entries = []
        for pos in range(0, len(directory), 128):
            raw = directory[pos:pos + 128]
            name_len, kind, color, left, right, child = unpack("HBBIII", raw, 64)
            if kind not in (0, 1, 2, 5) or (kind and (name_len < 2 or name_len > 64 or name_len % 2)):
                raise FormatError("Invalid OLE2 directory entry")
            name = raw[:max(0, name_len - 2)].decode("utf-16le") if kind else ""
            start, size = unpack("iQ", raw, 116)
            if self.version == 3:
                size &= 0xFFFFFFFF
            self.entries.append(dict(raw=raw, name=name, kind=kind, left=left,
                                     right=right, child=child, parent=-1, start=start, size=size))
        if not self.entries or self.entries[0]["kind"] != 5:
            raise FormatError("Missing OLE2 root entry")
        for parent, entry in enumerate(self.entries):
            if entry["kind"] not in (1, 5):
                continue
            pending = [entry["child"]]
            visited = set()
            while pending:
                n = pending.pop()
                if n == NO_STREAM:
                    continue
                if n >= len(self.entries) or n in visited:
                    raise FormatError("Invalid OLE2 directory tree")
                visited.add(n)
                e = self.entries[n]
                if e["parent"] != -1 or e["kind"] not in (1, 2):
                    raise FormatError("Duplicate or invalid OLE2 directory parent")
                e["parent"] = parent
                pending.extend((e["left"], e["right"]))
        if any(e["kind"] in (1, 2) and e["parent"] < 0 for e in self.entries):
            raise FormatError("Unreachable OLE2 directory entry")
        mini_sectors = self.chain(mini_start, self.fat, self.sector_count)
        if len(mini_sectors) != mini_count:
            raise FormatError("Invalid OLE2 mini FAT count")
        self.mini_fat = []
        for n in mini_sectors:
            self.mini_fat.extend(unpack(f"{self.sector_size // 4}i", self.sector(n)))
        root = self.entries[0]
        self.mini_data = self.regular_stream(root["start"], root["size"])
        self.streams = {i: self.read_stream(e) for i, e in enumerate(self.entries) if e["kind"] == 2}

    def sector(self, n):
        if not 0 <= n < self.sector_count:
            raise FormatError(f"OLE2 sector outside file: {n}")
        return self.data[(n + 1) * self.sector_size:(n + 2) * self.sector_size]

    @staticmethod
    def chain(start, table, capacity):
        ids = []
        seen = set()
        current = start
        while current not in (END, FREE):
            if not 0 <= current < min(len(table), capacity) or current in seen:
                raise FormatError("Invalid or cyclic OLE2 sector chain")
            seen.add(current)
            ids.append(current)
            current = table[current]
        return ids

    def regular_stream(self, start, size):
        ids = self.chain(start, self.fat, self.sector_count)
        if len(ids) != (size + self.sector_size - 1) // self.sector_size:
            raise FormatError("OLE2 regular stream size/chain mismatch")
        return b"".join(self.sector(n) for n in ids)[:size]

    def read_stream(self, e):
        size = e["size"]
        if size == 0:
            return b""
        if size >= 4096:
            return self.regular_stream(e["start"], size)
        ids = self.chain(e["start"], self.mini_fat, len(self.mini_data) // 64)
        if len(ids) != (size + 63) // 64:
            raise FormatError("OLE2 mini stream size/chain mismatch")
        return b"".join(self.mini_data[n * 64:(n + 1) * 64] for n in ids)[:size]

    def find(self, name, kind, parent=None):
        matches = [i for i, e in enumerate(self.entries)
                   if e["kind"] == kind and e["name"].casefold() == name.casefold()
                   and (parent is None or e["parent"] == parent)]
        if len(matches) != 1:
            raise FormatError(f"Missing or ambiguous OLE2 entry: {name}")
        return matches[0]

    def rebuild(self, changes, removed_streams=()):
        if self.version != 3:
            raise FormatError("OLE2 version 4 reconstruction is not supported (same as C#)")
        entries = [dict(e) for e in self.entries]
        streams = {i: data for i, data in (self.streams | changes).items() if i not in removed_streams}
        for i in removed_streams:
            entries[i] = dict(raw=bytes(128), name="", kind=0, parent=-1, start=0, size=0)
        blocks = []
        chains = []

        def allocate(data):
            count = (len(data) + 511) // 512
            if not count:
                return END
            start = len(blocks)
            padded = data.ljust(count * 512, b"\0")
            blocks.extend(padded[n:n + 512] for n in range(0, len(padded), 512))
            chains.append((start, count))
            return start

        mini_data = bytearray()
        mini_fat = []
        for i, data in streams.items():
            e = entries[i]
            e["size"] = len(data)
            if 0 < len(data) < 4096:
                start = len(mini_fat)
                count = (len(data) + 63) // 64
                mini_fat.extend(range(start + 1, start + count + 1))
                mini_fat[-1] = END
                mini_data.extend(data.ljust(count * 64, b"\0"))
                e["start"] = start
            else:
                e["start"] = allocate(data)
        entries[0]["start"] = allocate(mini_data)
        entries[0]["size"] = len(mini_data)
        mini_fat.extend([FREE] * (-len(mini_fat) % 128))
        mini_start = allocate(struct.pack(f"<{len(mini_fat)}i", *mini_fat))
        children = {}
        for i, e in enumerate(entries):
            e["left"] = e["right"] = e["child"] = NO_STREAM
            if e["kind"] in (1, 2):
                children.setdefault(e["parent"], []).append(i)

        def balanced(ids):
            if not ids:
                return NO_STREAM
            middle = len(ids) // 2
            n = ids[middle]
            entries[n]["left"] = balanced(ids[:middle])
            entries[n]["right"] = balanced(ids[middle + 1:])
            return n

        for parent, ids in children.items():
            ids.sort(key=lambda n: (len(entries[n]["name"].encode("utf-16le")), entries[n]["name"].upper()))
            names = [entries[n]["name"].upper() for n in ids]
            if len(names) != len(set(names)):
                raise FormatError("Duplicate OLE2 child names")
            entries[parent]["child"] = balanced(ids)
        directory = bytearray()
        for e in entries:
            if not e["kind"]:
                directory.extend(bytes(128))
                continue
            raw = bytearray(e["raw"])
            struct.pack_into("<BIII", raw, 67, 1 if e["kind"] else 0,
                             e["left"], e["right"], e["child"])
            struct.pack_into("<iQ", raw, 116, e["start"], e["size"])
            directory.extend(raw)
        dir_start = allocate(directory)
        base = len(blocks)
        fat_count, difat_count = 0, 0
        while True:
            new_fat = (base + fat_count + difat_count + 127) // 128
            new_difat = (max(0, new_fat - 109) + 126) // 127
            if (new_fat, new_difat) == (fat_count, difat_count):
                break
            fat_count, difat_count = new_fat, new_difat
        fat_ids = list(range(base, base + fat_count))
        difat_ids = list(range(base + fat_count, base + fat_count + difat_count))
        fat = [FREE] * (fat_count * 128)
        for start, count in chains:
            fat[start:start + count] = list(range(start + 1, start + count)) + [END]
        for n in fat_ids:
            fat[n] = FAT_SECTOR
        for n in difat_ids:
            fat[n] = DIFAT_SECTOR
        blocks.extend(struct.pack("<128i", *fat[n:n + 128]) for n in range(0, len(fat), 128))
        for i, _ in enumerate(difat_ids):
            ids = fat_ids[109 + i * 127:109 + (i + 1) * 127]
            ids += [FREE] * (127 - len(ids))
            ids.append(difat_ids[i + 1] if i + 1 < len(difat_ids) else END)
            blocks.append(struct.pack("<128i", *ids))
        header = bytearray(512)
        header[:8] = OLE_SIGNATURE
        struct.pack_into("<5H", header, 24, 0x3E, 3, 0xFFFE, 9, 6)
        struct.pack_into("<IIiIIiIiI", header, 40, 0, fat_count, dir_start, 0, 4096,
                         mini_start, len(mini_fat) // 128,
                         difat_ids[0] if difat_ids else END, difat_count)
        struct.pack_into("<109i", header, 76, *(fat_ids[:109] + [FREE] * max(0, 109 - fat_count)))
        return bytes(header) + b"".join(blocks)


class DirReader:
    """Cursor for dir records; this boundary hides their variable byte lengths."""

    def __init__(self, data, pos):
        self.data, self.pos = data, pos

    def record(self, expected, expected_size=None):
        record_id, size = unpack("HI", self.data, self.pos)
        if (record_id != expected or self.pos + 6 + size > len(self.data)
                or (expected_size is not None and size != expected_size)):
            raise FormatError(f"Invalid dir record 0x{expected:04x} at {self.pos}")
        self.pos += 6
        value = self.data[self.pos:self.pos + size]
        self.pos += size
        return value

    def reserved(self, expected):
        record_id, _ = unpack("HI", self.data, self.pos)
        if record_id != expected:
            raise FormatError(f"Invalid dir reserved record 0x{expected:04x}")
        self.pos += 6


def parse_dir(data):
    """Find a unique valid PROJECTMODULES/code-page pair, as in C#."""
    successes = []
    for match in re.finditer(re.escape(bytes.fromhex("0f0002000000")), data):
        offset = match.start()
        if data[offset + 8:offset + 14] != bytes.fromhex("130002000000"):
            continue
        pages = [int.from_bytes(m.group(1), "little") for m in
                 re.finditer(b"\x03\x00\x02\x00\x00\x00(..)", data[:offset], re.DOTALL)] or [932]
        for page in pages:
            # Alternative candidates can be embedded in reference strings.
            try:
                encoding = {65001: "utf-8", 1200: "utf-16le", 1201: "utf-16be", 10000: "mac_roman"}.get(page, f"cp{page}")
                try:
                    codecs.lookup(encoding)
                except LookupError:
                    if len(pages) != 1:
                        continue
                    page, encoding = 932, "cp932"
                reader = DirReader(data, offset)
                count, = unpack("H", reader.record(0xF, 2))
                reader.record(0x13, 2)
                modules = []
                for _ in range(count):
                    name = reader.record(0x19).decode(encoding)
                    if data[reader.pos:reader.pos + 2] == b"\x47\0":
                        unicode_name = reader.record(0x47).decode("utf-16le")
                        if unicode_name and unicode_name != name:
                            raise FormatError("MODULENAME encodings disagree")
                    stream = reader.record(0x1A).decode(encoding)
                    unicode_stream = reader.record(0x32).decode("utf-16le")
                    if not name or not stream or (unicode_stream and unicode_stream != stream):
                        raise FormatError("Invalid MODULESTREAMNAME")
                    reader.record(0x1C)
                    reader.record(0x48).decode("utf-16le")
                    offset_position = reader.pos + 6
                    source_offset, = unpack("I", reader.record(0x31, 4))
                    reader.record(0x1E, 4)
                    reader.record(0x2C, 2)
                    type_id, = unpack("H", data, reader.pos)
                    if type_id not in (0x21, 0x22):
                        raise FormatError("Invalid MODULETYPE")
                    reader.reserved(type_id)
                    while data[reader.pos:reader.pos + 2] in (b"\x25\0", b"\x28\0"):
                        flag, = unpack("H", data, reader.pos)
                        reader.reserved(flag)
                    reader.reserved(0x2B)
                    modules.append(dict(name=name, stream=stream, offset=source_offset,
                                        offset_position=offset_position, type_id=type_id))
                reader.reserved(0x10)
                if reader.pos != len(data):
                    raise FormatError("Invalid PROJECTTERMINATOR")
                successes.append((page, encoding, modules))
            except (FormatError, UnicodeError):
                continue
    if len(successes) != 1:
        raise FormatError(f"Expected one valid VBA dir candidate; found {len(successes)}")
    return successes[0]


def normalize_newlines(text):
    return text.replace("\r\n", "\n").replace("\r", "\n").replace("\n", "\r\n")


class Book:
    def __init__(self, data):
        self.data = data
        self.entry_name = None
        self.zip_entries = []
        self.zip_comment = b""
        if data[:8] == OLE_SIGNATURE:
            project_bytes = data
        else:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                self.zip_entries = [(entry, archive.read(entry)) for entry in archive.infolist()]
                self.zip_comment = archive.comment
            names = [entry.filename for entry, _ in self.zip_entries]
            if len(set(names)) != len(names):
                raise FormatError("Duplicate ZIP entry names")
            candidates = [(entry.filename, payload) for entry, payload in self.zip_entries
                          if entry.filename.replace("\\", "/").rsplit("/", 1)[-1].casefold() == "vbaproject.bin"]
            if len(candidates) != 1:
                raise FormatError("Expected exactly one vbaProject.bin in workbook")
            self.entry_name, project_bytes = candidates[0]
        self.ole = OleFile(project_bytes)
        if any(e["name"] == "EncryptedPackage" for e in self.ole.entries):
            raise FormatError("Password-to-open encrypted workbooks are unsupported")
        self.vba_id = self.ole.find("VBA", 1)
        self.dir_id = self.ole.find("dir", 2, self.vba_id)
        self.dir_bytes = decompress(self.ole.streams[self.dir_id])
        self.codepage, self.encoding, records = parse_dir(self.dir_bytes)
        project_id = self.ole.find("PROJECT", 2)
        project_text = self.ole.streams[project_id].decode(self.encoding)
        types = {}
        for line in re.split(r"\r\n|\n|\r", project_text):
            for prefix, kind in (("Document=", "Document"), ("BaseClass=", "Form"), ("Module=", "Standard"), ("Class=", "Class")):
                if line.startswith(prefix):
                    name = line[len(prefix):].split("/", 1)[0] if kind == "Document" else line[len(prefix):]
                    if not name or name.casefold() in types:
                        raise FormatError("Empty or duplicate PROJECT module")
                    types[name.casefold()] = kind
        self.modules = []
        used_names, used_streams = set(), set()
        for record in records:
            name = record["name"]
            if name.casefold() not in types or name.casefold() in used_names:
                raise FormatError(f"PROJECT/dir module mismatch: {name}")
            kind = types[name.casefold()]
            if record["type_id"] != (0x21 if kind == "Standard" else 0x22):
                raise FormatError(f"PROJECT/dir type mismatch: {name}")
            stream_id = self.ole.find(record["stream"], 2, self.vba_id)
            if stream_id in used_streams:
                raise FormatError("Multiple modules refer to one stream")
            full_bytes = decompress(self.ole.streams[stream_id], record["offset"])
            full = full_bytes.decode(self.encoding)
            header = re.match(r"(?:[ \t]*Attribute VB_[^\r\n]*(?:\r\n|\r|\n|$))*", full, re.IGNORECASE)[0]
            if not re.search(r'^\s*Attribute VB_Name\s*=\s*"' + re.escape(name) + r'"\s*$', header, re.IGNORECASE | re.MULTILINE):
                raise FormatError(f"Missing or inconsistent VB_Name attribute: {name}")
            self.modules.append(record | dict(kind=kind, stream_id=stream_id, full=full,
                                               header=header, code=full[len(header):], source_bytes=full_bytes))
            used_names.add(name.casefold())
            used_streams.add(stream_id)
        if len(types) != len(records) or not records:
            raise FormatError("Incomplete or empty VBA module inventory")
        reserved = {"dir", "_vba_project", "project", "projectwm"}
        unlisted = [e["name"] for i, e in enumerate(self.ole.entries)
                    if e["parent"] == self.vba_id and e["kind"] == 2 and i not in used_streams
                    and e["name"].casefold() not in reserved and not e["name"].casefold().startswith("__srp_")]
        if unlisted:
            raise FormatError(f"VBA streams missing from module inventory: {unlisted}")
        order = {"Document": 0, "Form": 1, "Standard": 2, "Class": 3}
        self.modules.sort(key=lambda m: (order[m["kind"]], m["name"].casefold()))

    def manifest(self):
        modules = []
        for i, module in enumerate(self.modules, 1):
            name = re.sub(r"[^\w-]", "_", module["name"])
            extension = {"Standard": "bas", "Form": "frm"}.get(module["kind"], "cls")
            modules.append(dict(name=module["name"], kind=module["kind"], stream=module["stream"],
                                file=f"{i:03d}_{name}.{extension}"))
        return dict(schema=1, source_sha256=hashlib.sha256(self.data).hexdigest(),
                    codepage=self.codepage, encoding="utf-8", modules=modules)

    def replace(self, changes):
        streams = {}
        removed_streams = set()
        directory = bytearray(self.dir_bytes)
        for module in self.modules:
            if module["name"] in changes:
                source = changes[module["name"]].encode(self.encoding)
            else:
                source = module["source_bytes"]
            # All modules share project caches. Dropping only _VBA_PROJECT's
            # cache while retaining these prefixes crashes VBE7 on open.
            streams[module["stream_id"]] = compress(source)
            struct.pack_into("<I", directory, module["offset_position"], 0)
        streams[self.dir_id] = compress(directory)
        streams[self.ole.find("_VBA_PROJECT", 2, self.vba_id)] = INVALIDATED_CACHE
        for i, entry in enumerate(self.ole.entries):
            if (entry["parent"] == self.vba_id and entry["kind"] == 2
                    and entry["name"].casefold().startswith("__srp_")):
                removed_streams.add(i)
        project = self.ole.rebuild(streams, removed_streams)
        if self.entry_name is None:
            return project, streams, [], removed_streams
        output = io.BytesIO()
        removed = []
        marker = "vbaprojectsignature"
        signature_parts = [entry.filename for entry, _ in self.zip_entries if marker in entry.filename.casefold()]
        with zipfile.ZipFile(output, "w") as archive:
            archive.comment = self.zip_comment
            for entry, payload in self.zip_entries:
                name = entry.filename
                if name in signature_parts:
                    removed.append(name)
                    continue
                if name == self.entry_name:
                    payload = project
                elif signature_parts and (name.casefold().endswith(".rels") or name.casefold() == "[content_types].xml"):
                    root = ET.fromstring(payload)
                    tag = "Relationship" if name.casefold().endswith(".rels") else "Override"
                    matches = [child for child in root if child.tag.rsplit("}", 1)[-1] == tag
                               and any(marker in value.casefold() for value in child.attrib.values())]
                    for child in matches:
                        root.remove(child)
                    if matches:
                        if tag == "Relationship" and len(root) == 0:
                            removed.append(name)
                            continue
                        payload = ET.tostring(root, encoding="utf-8", xml_declaration=True)
                archive.writestr(entry, payload)
        return output.getvalue(), streams, removed, removed_streams


def extract(workbook, directory):
    """Export every module plus a source snapshot manifest into a new directory."""
    workbook, directory = Path(workbook).expanduser().resolve(), Path(directory).expanduser().resolve()
    if not workbook.is_file():
        raise ValueError(f"Workbook does not exist: {workbook}")
    if directory.exists():
        raise FileExistsError(f"Extraction directory already exists: {directory}")
    book = Book(workbook.read_bytes())
    manifest = book.manifest()
    directory.mkdir(parents=True)
    for module, item in zip(book.modules, manifest["modules"]):
        (directory / item["file"]).write_bytes(module["code"].encode("utf-8"))
    (directory / "manifest.json").write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    return dict(action="extract", directory=str(directory), modules=len(book.modules), codepage=book.codepage)


def rebuild(workbook, directory, output):
    """Create a separate workbook and verify its source and untouched streams."""
    workbook, directory, output = (Path(p).expanduser().resolve() for p in (workbook, directory, output))
    if not workbook.is_file() or not directory.is_dir():
        raise ValueError("The original workbook and extracted directory must exist")
    if output == workbook or output.exists():
        raise FileExistsError(f"Output must be a new file, separate from the original: {output}")
    if output.suffix.lower() not in FORMATS or output.suffix.lower() != workbook.suffix.lower():
        raise ValueError("Output must use the original .xlsm/.xlam/.xlsb/.xls extension")
    if not output.parent.is_dir():
        raise ValueError("Output parent directory does not exist")
    book = Book(workbook.read_bytes())
    expected_manifest = book.manifest()
    manifest = json.loads((directory / "manifest.json").read_text(encoding="utf-8-sig"))
    if manifest != expected_manifest:
        raise ValueError("Manifest does not match the original workbook; extract it again")
    expected_files = {item["file"] for item in manifest["modules"]} | {"manifest.json"}
    extras = [p.name for p in directory.iterdir() if p.is_file() and p.name not in expected_files]
    if extras:
        raise ValueError(f"Unexpected module files (adding/renaming modules is unsupported): {extras}")
    changes = {}
    for module, item in zip(book.modules, manifest["modules"]):
        code = (directory / item["file"]).read_bytes().decode("utf-8-sig")
        if normalize_newlines(code) != normalize_newlines(module["code"]):
            full = module["header"] + normalize_newlines(code)
            full.encode(book.encoding)  # Refuse unrepresentable characters before writing.
            changes[module["name"]] = full
    if changes:
        data, streams, removed, removed_streams = book.replace(changes)
    else:
        data, streams, removed, removed_streams = book.data, {}, [], set()
    verified = False
    # Exclusive creation protects both the source and any prior output.
    with output.open("xb") as handle:
        try:
            handle.write(data)
            handle.flush()
            # Reopen the actual file, not the in-memory result.
            result = Book(output.read_bytes())
            if result.manifest()["modules"] != expected_manifest["modules"]:
                raise FormatError("Output module inventory changed")
            for before, after in zip(book.modules, result.modules):
                expected = changes.get(before["name"], before["full"])
                if after["full"] != expected:
                    raise FormatError(f"Output source mismatch: {before['name']}")
                if after["offset"] != (0 if changes else before["offset"]):
                    raise FormatError("Output MODULEOFFSET is inconsistent with its cache")
            if len(result.ole.entries) != len(book.ole.entries):
                raise FormatError("Output OLE2 entry count changed")
            for i, (before, after) in enumerate(zip(book.ole.entries, result.ole.entries)):
                if i in removed_streams:
                    if after["kind"] != 0:
                        raise FormatError(f"Output kept compiled cache: {before['name']}")
                    continue
                if (before["name"], before["kind"], before["parent"], before["raw"][80:116]) != (after["name"], after["kind"], after["parent"], after["raw"][80:116]):
                    raise FormatError(f"Output OLE2 metadata changed: {before['name']}")
                if before["kind"] == 2 and result.ole.streams[i] != streams.get(i, book.ole.streams[i]):
                    raise FormatError(f"Output OLE2 stream mismatch: {before['name']}")
            if changes and result.ole.streams[result.ole.find("_VBA_PROJECT", 2, result.vba_id)] != INVALIDATED_CACHE:
                raise FormatError("VBA execution cache was not invalidated")
            if changes and any(e["parent"] == result.vba_id and e["kind"] == 2
                               and e["name"].casefold().startswith("__srp_") for e in result.ole.entries):
                raise FormatError("Output retained an SRP execution cache")
            if book.entry_name:
                original_parts = {entry.filename: payload for entry, payload in book.zip_entries}
                result_parts = {entry.filename: payload for entry, payload in result.zip_entries}
                if set(result_parts) != set(original_parts) - set(removed):
                    raise FormatError("Output ZIP inventory mismatch")
                for name, payload in result_parts.items():
                    if name != book.entry_name and name != "[Content_Types].xml" and not name.lower().endswith(".rels"):
                        if payload != original_parts[name]:
                            raise FormatError(f"Output ZIP part changed: {name}")
            verified = True
        finally:
            # Close before unlinking, including on Windows.
            if not verified:
                handle.close()
                output.unlink()
    return dict(action="rebuild", output=str(output), verified=True, modules=len(book.modules),
                changed_modules=list(changes), removed_signature_parts=removed)


def main():
    parser = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    commands = parser.add_subparsers(dest="command", required=True)
    export = commands.add_parser("extract", help="Export all module source files into a new directory")
    export.add_argument("workbook")
    export.add_argument("directory")
    build = commands.add_parser("rebuild", help="Write edited module files to a new workbook and verify")
    build.add_argument("workbook")
    build.add_argument("directory")
    build.add_argument("output")
    args = parser.parse_args()
    try:
        if args.command == "extract":
            result = extract(args.workbook, args.directory)
        else:
            result = rebuild(args.workbook, args.directory, args.output)
    except (OSError, ValueError, LookupError, zipfile.BadZipFile, ET.ParseError) as error:
        print(json.dumps(dict(error=type(error).__name__, message=str(error)), ensure_ascii=False), file=sys.stderr)
        return 1
    print(json.dumps(result, ensure_ascii=False))
    return 0


if __name__ == "__main__":
    sys.exit(main())
