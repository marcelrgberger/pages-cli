"""Unit tests for the offline .pages text reader (no Pages needed)."""
import json
import zipfile

import pytest
from click.testing import CliRunner

from pages_cli.core import offline
from pages_cli.pages_cli import cli


# --- helpers that build synthetic IWA data -------------------------------

def varint(n: int) -> bytes:
    out = bytearray()
    while True:
        byte = n & 0x7F
        n >>= 7
        out.append(byte | (0x80 if n else 0))
        if not n:
            return bytes(out)


def field_varint(number: int, value: int) -> bytes:
    return varint(number << 3) + varint(value)


def field_bytes(number: int, value: bytes) -> bytes:
    return varint(number << 3 | 2) + varint(len(value)) + value


def storage(text: str | None, kind: int | None) -> bytes:
    msg = b""
    if kind is not None:
        msg += field_varint(1, kind)
    if text is not None:
        msg += field_bytes(3, text.encode("utf-8"))
    return msg


def archive(*messages: tuple[int, bytes]) -> bytes:
    """Decompressed IWA content: one ArchiveInfo per message."""
    raw = b""
    for ident, (mtype, payload) in enumerate(messages, start=1):
        message_info = field_varint(1, mtype) + field_varint(3, len(payload))
        info = field_varint(1, ident) + field_bytes(2, message_info)
        raw += varint(len(info)) + info + payload
    return raw


def snappy_literal(data: bytes) -> bytes:
    """A valid Snappy block made only of literals."""
    out = varint(len(data))
    pos = 0
    while pos < len(data):
        chunk = data[pos:pos + 60]
        out += bytes([(len(chunk) - 1) << 2]) + chunk
        pos += len(chunk)
    return out


def iwa(raw: bytes) -> bytes:
    block = snappy_literal(raw)
    return b"\x00" + len(block).to_bytes(3, "little") + block


BODY = (2001, storage("First paragraph\nSecond  paragraph \n\nLast café 😀", 0))
TEXTBOX = (2001, storage("Inside a box", None))
HEADER = (2001, storage("Page header", 1))
EMPTY_PLACEHOLDER = (2001, storage(None, 1))
NOISE = (6247, b"\x08\x01 not text " * 5)


def zipped(path, files: dict):
    with zipfile.ZipFile(path, "w") as z:
        for name, data in files.items():
            z.writestr(name, data)


@pytest.fixture
def document_iwa():
    return iwa(archive(NOISE, BODY, EMPTY_PLACEHOLDER, TEXTBOX, HEADER))


# --- decompression --------------------------------------------------------

class TestSnappy:
    def test_literals(self):
        data = bytes(range(200)) * 2
        assert offline._snappy_decompress(snappy_literal(data)) == data

    def test_copy_with_overlap_repeats_pattern(self):
        # literal "ab", then copy length 6 offset 2 (tag type 1) -> "abababab"
        block = varint(8) + bytes([(2 - 1) << 2]) + b"ab" + bytes([((6 - 4) << 2) | 1, 2])
        assert offline._snappy_decompress(block) == b"abababab"

    def test_two_byte_offset_copy(self):
        block = varint(8) + bytes([(4 - 1) << 2]) + b"abcd" + bytes([(4 - 1) << 2 | 2, 4, 0])
        assert offline._snappy_decompress(block) == b"abcdabcd"

    def test_long_literal_uses_length_bytes(self):
        data = b"x" * 100
        block = varint(100) + bytes([60 << 2, 99]) + data
        assert offline._snappy_decompress(block) == data

    @pytest.mark.parametrize("block", [
        varint(5) + bytes([(3 - 1) << 2]) + b"ab",            # truncated literal
        varint(4) + bytes([0b00000101, 9]),                  # back-reference before start
        varint(9) + bytes([(2 - 1) << 2]) + b"ab",            # length mismatch
        b"",                                                  # no preamble
    ])
    def test_corrupt_data_raises_value_error(self, block):
        with pytest.raises(ValueError):
            offline._snappy_decompress(block)


# --- reading --------------------------------------------------------------

class TestReadText:
    def test_zip_file(self, tmp_path, document_iwa):
        path = tmp_path / "a.pages"
        zipped(path, {"Index/Document.iwa": document_iwa, "Index/Other.iwa": b"junk"})
        result = offline.read_text(str(path))
        assert result["body"] == "First paragraph\nSecond  paragraph \n\nLast café 😀"
        assert result["other"] == [
            {"kind": None, "text": "Inside a box"},
            {"kind": 1, "text": "Page header"},
        ]

    def test_bundle_with_index_directory(self, tmp_path, document_iwa):
        bundle = tmp_path / "b.pages"
        (bundle / "Index").mkdir(parents=True)
        (bundle / "Index" / "Document.iwa").write_bytes(document_iwa)
        assert offline.read_text(str(bundle))["body"].startswith("First paragraph")

    def test_bundle_with_index_zip(self, tmp_path, document_iwa):
        bundle = tmp_path / "c.pages"
        bundle.mkdir()
        zipped(bundle / "Index.zip", {"Index/Document.iwa": document_iwa})
        assert offline.read_text(str(bundle))["body"].startswith("First paragraph")

    def test_expands_user_home(self, tmp_path, document_iwa, monkeypatch):
        zipped(tmp_path / "h.pages", {"Index/Document.iwa": document_iwa})
        monkeypatch.setenv("HOME", str(tmp_path))
        assert offline.read_text("~/h.pages")["body"]

    def test_multi_chunk_iwa(self, tmp_path):
        first, second = archive(BODY), archive(TEXTBOX)
        path = tmp_path / "m.pages"
        zipped(path, {"Index/Document.iwa": iwa(first) + iwa(second)})
        result = offline.read_text(str(path))
        assert result["other"] == [{"kind": None, "text": "Inside a box"}]

    def test_missing_file(self, tmp_path):
        with pytest.raises(FileNotFoundError):
            offline.read_text(str(tmp_path / "nope.pages"))

    def test_not_a_pages_file(self, tmp_path):
        path = tmp_path / "x.pages"
        path.write_text("hello")
        with pytest.raises(ValueError, match="not a recognised"):
            offline.read_text(str(path))

    def test_zip_without_document_iwa(self, tmp_path):
        path = tmp_path / "old.pages"
        zipped(path, {"index.xml": b"<x/>"})
        with pytest.raises(ValueError, match="Pages '09"):
            offline.read_text(str(path))

    def test_no_body_storage(self, tmp_path):
        path = tmp_path / "nb.pages"
        zipped(path, {"Index/Document.iwa": iwa(archive(TEXTBOX))})
        with pytest.raises(ValueError, match="No body text"):
            offline.read_text(str(path))

    def test_truncated_chunk(self, tmp_path, document_iwa):
        path = tmp_path / "t.pages"
        zipped(path, {"Index/Document.iwa": document_iwa[:-10]})
        with pytest.raises(ValueError):
            offline.read_text(str(path))

    def test_empty_body_is_a_valid_empty_document(self, tmp_path):
        path = tmp_path / "e.pages"
        zipped(path, {"Index/Document.iwa": iwa(archive((2001, storage("", 0))))})
        assert offline.read_text(str(path)) == {"body": "", "other": []}


# --- CLI ------------------------------------------------------------------

class TestReadFileCommand:
    @pytest.fixture
    def pages_file(self, tmp_path, document_iwa):
        path = tmp_path / "doc.pages"
        zipped(path, {"Index/Document.iwa": document_iwa})
        return str(path)

    def test_plain_output_needs_no_pages(self, pages_file):
        result = CliRunner().invoke(cli, ["document", "read-file", pages_file])
        assert result.exit_code == 0
        assert result.output.startswith("First paragraph\nSecond  paragraph \n\nLast café")

    def test_numbered_matches_pages_paragraph_numbers(self, pages_file):
        result = CliRunner().invoke(cli, ["document", "read-file", "--numbered", pages_file])
        assert "1: First paragraph\n2: Second  paragraph \n3: \n4: Last" in result.output

    def test_other_storages_listed_on_request(self, pages_file):
        plain = CliRunner().invoke(cli, ["document", "read-file", pages_file])
        assert "Inside a box" not in plain.output
        more = CliRunner().invoke(cli, ["document", "read-file", "--other", pages_file])
        assert "Inside a box" in more.output and "Page header" in more.output

    def test_json_output(self, pages_file):
        result = CliRunner().invoke(cli, ["--json", "document", "read-file", "--other", pages_file])
        data = json.loads(result.output)
        assert data["paragraphs"] == ["First paragraph", "Second  paragraph ", "", "Last café 😀"]
        assert data["other"][0]["text"] == "Inside a box"

    def test_error_is_reported(self, tmp_path):
        result = CliRunner().invoke(cli, ["document", "read-file", str(tmp_path / "x.pages")])
        assert "File not found" in result.output
