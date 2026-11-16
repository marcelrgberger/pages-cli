"""Read text out of a .pages file without opening Pages.

A Pages document keeps its text in ``Index/Document.iwa``: a sequence of
Snappy-compressed chunks holding Protocol Buffers messages. This module
decodes just enough of that to return the text, and nothing is ever written.

Layouts handled: a single zip file, or a bundle directory containing either
``Index/Document.iwa`` or ``Index.zip``. Pages '09 (``index.xml``) is not
supported.

Text lives in ``TSWP.StorageArchive`` messages (type 2001), field 3. Field 1
is the storage kind: 0 is the document body. Text boxes, shapes, headers,
footnotes and similar are other storages. Table cell text is held elsewhere
and is not returned.
"""

import re
import zipfile
from pathlib import Path

_STORAGE_ARCHIVE = 2001
_KIND_BODY = 0
_DOCUMENT_IWA = "Index/Document.iwa"


def _read_varint(buf: bytes, pos: int) -> tuple[int, int]:
    value = 0
    shift = 0
    while True:
        byte = buf[pos]
        pos += 1
        value |= (byte & 0x7F) << shift
        if byte < 0x80:
            return value, pos
        shift += 7
        if shift > 63:
            raise ValueError("Invalid varint.")


def _snappy_decompress(buf: bytes) -> bytes:
    """Decompress a raw Snappy block (no stream framing, as IWA uses)."""
    try:
        expected, pos = _read_varint(buf, 0)
        out = bytearray()
        size = len(buf)
        while pos < size:
            tag = buf[pos]
            pos += 1
            kind = tag & 3
            if kind == 0:
                length = tag >> 2
                if length >= 60:
                    extra = length - 59
                    length = int.from_bytes(buf[pos:pos + extra], "little")
                    pos += extra
                length += 1
                chunk = buf[pos:pos + length]
                if len(chunk) != length:
                    raise ValueError("Truncated Snappy literal.")
                out += chunk
                pos += length
                continue
            if kind == 1:
                length = ((tag >> 2) & 7) + 4
                offset = ((tag >> 5) << 8) | buf[pos]
                pos += 1
            elif kind == 2:
                length = (tag >> 2) + 1
                offset = int.from_bytes(buf[pos:pos + 2], "little")
                pos += 2
            else:
                length = (tag >> 2) + 1
                offset = int.from_bytes(buf[pos:pos + 4], "little")
                pos += 4
            if offset == 0 or offset > len(out):
                raise ValueError("Invalid Snappy back-reference.")
            start = len(out) - offset
            if offset >= length:
                out += out[start:start + length]
            else:
                for i in range(length):
                    out.append(out[start + i])
    except IndexError:
        raise ValueError("Truncated Snappy data.") from None
    if len(out) != expected:
        raise ValueError("Snappy length mismatch.")
    return bytes(out)


def _decompress_iwa(data: bytes) -> bytes:
    """Join the chunks of an .iwa file (each: 0x00, 3-byte length, Snappy block)."""
    out = bytearray()
    pos = 0
    while pos < len(data):
        if data[pos] != 0 or pos + 4 > len(data):
            raise ValueError("Not an IWA chunk stream.")
        length = int.from_bytes(data[pos + 1:pos + 4], "little")
        pos += 4
        chunk = data[pos:pos + length]
        if len(chunk) != length:
            raise ValueError("Truncated IWA chunk.")
        out += _snappy_decompress(chunk)
        pos += length
    return bytes(out)


def _fields(buf: bytes) -> list[tuple[int, int, int | bytes]]:
    """Parse one Protocol Buffers message into (field, wire type, value)."""
    result = []
    pos = 0
    try:
        while pos < len(buf):
            key, pos = _read_varint(buf, pos)
            field, wire = key >> 3, key & 7
            if wire == 0:
                value, pos = _read_varint(buf, pos)
            elif wire == 2:
                length, pos = _read_varint(buf, pos)
                value = buf[pos:pos + length]
                if len(value) != length:
                    raise ValueError("Truncated message field.")
                pos += length
            elif wire == 1:
                value = buf[pos:pos + 8]
                pos += 8
            elif wire == 5:
                value = buf[pos:pos + 4]
                pos += 4
            else:
                raise ValueError(f"Unsupported wire type {wire}.")
            result.append((field, wire, value))
    except IndexError:
        raise ValueError("Truncated message.") from None
    return result


def _first(fields, number: int, default=None):
    for field, _, value in fields:
        if field == number:
            return value
    return default


def _messages(raw: bytes):
    """Yield (type, payload) for every archived message in a decompressed IWA.

    The stream is: varint length, ArchiveInfo, then the payload of each of
    its message_infos in turn (lengths in message_info field 3).
    """
    pos = 0
    while pos < len(raw):
        length, pos = _read_varint(raw, pos)
        info = _fields(raw[pos:pos + length])
        pos += length
        for field, _, value in info:
            if field != 2:
                continue
            message_info = _fields(value)
            size = _first(message_info, 3, 0)
            payload = raw[pos:pos + size]
            if len(payload) != size:
                raise ValueError("Truncated archive payload.")
            pos += size
            yield _first(message_info, 1), payload


def _read_document_iwa(path: Path) -> bytes:
    """Return the bytes of Index/Document.iwa from any supported layout."""
    if path.is_dir():
        direct = path / _DOCUMENT_IWA
        if direct.is_file():
            return direct.read_bytes()
        inner = path / "Index.zip"
        if inner.is_file():
            with zipfile.ZipFile(inner) as archive:
                return archive.read(_DOCUMENT_IWA)
    elif path.is_file() and zipfile.is_zipfile(path):
        with zipfile.ZipFile(path) as archive:
            if _DOCUMENT_IWA in archive.namelist():
                return archive.read(_DOCUMENT_IWA)
    raise ValueError(
        f"{path} is not a recognised iWork-format .pages file "
        f"(no Index/Document.iwa). Pages '09 files are not supported."
    )


# What Pages' scripting counts as a paragraph boundary. The boundary character
# is not part of any paragraph. U+0004 and U+0005 (stored in documents, and not
# insertable by script) were checked against a real 632-paragraph document: the
# split matched Pages' own list exactly. U+000B, U+000C, U+000D and U+2029 were
# checked by inserting them. U+2028 (soft line break) and U+FFFC (object anchor)
# stay inside a paragraph. Other stored control characters (e.g. U+000E) are
# left inline, which has not been checked.
_PARAGRAPH_BREAKS = re.compile("[\n\r\x0b\x0c\x04\x05\u2029]")


def split_paragraphs(body: str) -> list[str]:
    """Split body text into paragraphs the way Pages numbers them."""
    return _PARAGRAPH_BREAKS.split(body)


def read_text(path: str) -> dict:
    """Read the text of a .pages file without opening Pages.

    Returns ``{"body": str, "other": [{"kind": int | None, "text": str}]}``.
    ``body`` is the text exactly as stored; use split_paragraphs for Pages'
    paragraph numbering (a line feed is not the only paragraph break). ``other``
    lists non-empty text from other storages (text boxes, shapes, headers,
    footnotes); ``kind`` is Pages' storage kind, or None when absent (text
    boxes and shapes). Pages also stores some in-text objects and breaks as
    special characters, returned unchanged (seen in real documents: U+FFFC,
    U+2028, and control characters U+0004, U+0005, U+000C, U+000E).
    """
    source = Path(path).expanduser()
    if not source.exists():
        raise FileNotFoundError(f"File not found: {source}")
    raw = _decompress_iwa(_read_document_iwa(source))

    body: list[str] = []
    other: list[dict] = []
    for message_type, payload in _messages(raw):
        if message_type != _STORAGE_ARCHIVE:
            continue
        fields = _fields(payload)
        kind = _first(fields, 1)
        text = "".join(
            value.decode("utf-8", "replace")
            for field, wire, value in fields
            if field == 3 and wire == 2
        )
        if kind == _KIND_BODY:
            body.append(text)
        elif text:
            other.append({"kind": kind, "text": text})
    if not body:
        raise ValueError(f"No body text found in {source}.")
    return {"body": "\n".join(body), "other": other}
