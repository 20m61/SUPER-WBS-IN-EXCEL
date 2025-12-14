#!/usr/bin/env python3
"""
vbaProject.bin ジェネレーター

OLE複合ドキュメント形式でVBAプロジェクトバイナリを生成する。
Excel VBAのバイナリ形式に準拠したファイルを作成する。

参考:
- MS-OVBA: Office VBA File Format Structure
- MS-CFB: Compound File Binary Format
"""
import io
import struct
import zlib
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import olefile


# VBA圧縮アルゴリズム（MS-OVBA 2.4.1に準拠）
def compress_vba(data: bytes) -> bytes:
    """VBA圧縮を適用する（RLE風の圧縮）"""
    if not data:
        return b'\x01\xb0'  # 空データ用の最小圧縮

    output = io.BytesIO()
    # シグネチャバイト（圧縮済みを示す）
    output.write(b'\x01')

    pos = 0
    while pos < len(data):
        # チャンクを処理（最大4096バイト）
        chunk_start = pos
        chunk_end = min(pos + 4096, len(data))
        chunk_data = data[chunk_start:chunk_end]

        # 圧縮せずにそのまま格納（簡略化）
        compressed_chunk = io.BytesIO()

        i = 0
        while i < len(chunk_data):
            # フラグバイトの位置を記録
            flag_pos = compressed_chunk.tell()
            compressed_chunk.write(b'\x00')  # プレースホルダー

            flag_byte = 0
            tokens_written = 0

            for bit in range(8):
                if i >= len(chunk_data):
                    break
                # リテラルトークンとして書き込み
                compressed_chunk.write(bytes([chunk_data[i]]))
                i += 1
                tokens_written += 1

            if tokens_written > 0:
                # フラグバイトを更新（全てリテラル = 0x00）
                current_pos = compressed_chunk.tell()
                compressed_chunk.seek(flag_pos)
                compressed_chunk.write(bytes([flag_byte]))
                compressed_chunk.seek(current_pos)

        chunk_bytes = compressed_chunk.getvalue()

        # チャンクヘッダー（2バイト）
        chunk_size = len(chunk_bytes)
        # ヘッダー: サイズ(12ビット) | 圧縮タイプ(3ビット) | シグネチャ(1ビット)
        header = ((chunk_size - 1) & 0x0FFF) | 0xB000
        output.write(struct.pack('<H', header))
        output.write(chunk_bytes)

        pos = chunk_end

    return output.getvalue()


def decompress_vba(data: bytes) -> bytes:
    """VBA圧縮を解凍する"""
    if not data or data[0] != 0x01:
        return data

    output = io.BytesIO()
    pos = 1

    while pos < len(data):
        if pos + 2 > len(data):
            break

        header = struct.unpack('<H', data[pos:pos+2])[0]
        pos += 2

        chunk_size = (header & 0x0FFF) + 1
        is_compressed = (header & 0x8000) != 0

        if pos + chunk_size > len(data):
            chunk_size = len(data) - pos

        chunk_data = data[pos:pos+chunk_size]
        pos += chunk_size

        if not is_compressed:
            output.write(chunk_data)
        else:
            # 圧縮チャンクを解凍
            chunk_pos = 0
            decompressed_start = output.tell()

            while chunk_pos < len(chunk_data):
                if chunk_pos >= len(chunk_data):
                    break

                flag_byte = chunk_data[chunk_pos]
                chunk_pos += 1

                for bit in range(8):
                    if chunk_pos >= len(chunk_data):
                        break

                    if (flag_byte >> bit) & 1:
                        # コピートークン
                        if chunk_pos + 2 > len(chunk_data):
                            break
                        token = struct.unpack('<H', chunk_data[chunk_pos:chunk_pos+2])[0]
                        chunk_pos += 2

                        # オフセットと長さを計算
                        decompressed_current = output.tell() - decompressed_start
                        bit_count = max(4, (decompressed_current - 1).bit_length())
                        length_mask = (1 << (16 - bit_count)) - 1
                        offset_mask = ~length_mask & 0xFFFF

                        length = (token & length_mask) + 3
                        offset = ((token & offset_mask) >> (16 - bit_count)) + 1

                        # コピー
                        copy_start = output.tell() - offset
                        for _ in range(length):
                            output.seek(copy_start)
                            byte = output.read(1)
                            output.seek(0, 2)
                            output.write(byte)
                            copy_start += 1
                    else:
                        # リテラルトークン
                        output.write(bytes([chunk_data[chunk_pos]]))
                        chunk_pos += 1

    return output.getvalue()


@dataclass
class VBAModule:
    """VBAモジュール定義"""
    name: str
    code: str
    module_type: str = "Module"  # Module, Class, ThisWorkbook, Sheet

    @property
    def stream_name(self) -> str:
        """OLEストリーム名を返す"""
        return self.name


class VBAProjectGenerator:
    """VBAプロジェクトバイナリ生成器"""

    # プロジェクトプロパティ
    PROJECT_NAME = "VBAProject"
    PROJECT_VERSION = "1.1"
    PROJECT_CODEPAGE = 932  # Shift_JIS

    def __init__(self):
        self.modules: List[VBAModule] = []
        self._ole_data = io.BytesIO()

    def add_module(self, name: str, code: str, module_type: str = "Module"):
        """モジュールを追加する"""
        self.modules.append(VBAModule(name=name, code=code, module_type=module_type))

    def _create_dir_stream(self) -> bytes:
        """VBA DIRストリームを作成する"""
        stream = io.BytesIO()

        # PROJECTSYSKIND (0x0001)
        stream.write(struct.pack('<HI', 0x0001, 4))
        stream.write(struct.pack('<I', 0x00000001))  # Windows

        # PROJECTLCID (0x0002)
        stream.write(struct.pack('<HI', 0x0002, 4))
        stream.write(struct.pack('<I', 0x0411))  # Japanese

        # PROJECTLCIDINVOKE (0x0014)
        stream.write(struct.pack('<HI', 0x0014, 4))
        stream.write(struct.pack('<I', 0x0411))

        # PROJECTCODEPAGE (0x0003)
        stream.write(struct.pack('<HI', 0x0003, 2))
        stream.write(struct.pack('<H', self.PROJECT_CODEPAGE))

        # PROJECTNAME (0x0004)
        name_bytes = self.PROJECT_NAME.encode('ascii')
        stream.write(struct.pack('<HI', 0x0004, len(name_bytes)))
        stream.write(name_bytes)

        # PROJECTDOCSTRING (0x0005)
        stream.write(struct.pack('<HI', 0x0005, 0))
        stream.write(struct.pack('<HI', 0x0040, 0))  # Unicode version

        # PROJECTHELPFILEPATH (0x0006)
        stream.write(struct.pack('<HI', 0x0006, 0))
        stream.write(struct.pack('<HI', 0x003D, 0))

        # PROJECTHELPCONTEXT (0x0007)
        stream.write(struct.pack('<HI', 0x0007, 4))
        stream.write(struct.pack('<I', 0))

        # PROJECTLIBFLAGS (0x0008)
        stream.write(struct.pack('<HI', 0x0008, 4))
        stream.write(struct.pack('<I', 0))

        # PROJECTVERSION (0x0009)
        stream.write(struct.pack('<HI', 0x0009, 4))
        stream.write(struct.pack('<HH', 1, 1))  # Major.Minor

        # PROJECTCONSTANTS (0x000C)
        stream.write(struct.pack('<HI', 0x000C, 0))
        stream.write(struct.pack('<HI', 0x003C, 0))  # Unicode version

        # REFERENCE records would go here (external references)
        # For now, we'll add a minimal reference to stdole

        # PROJECTMODULES (0x000F) - number of modules
        stream.write(struct.pack('<HI', 0x000F, 2))
        stream.write(struct.pack('<H', len(self.modules)))

        # PROJECTCOOKIE (0x0013)
        stream.write(struct.pack('<HI', 0x0013, 2))
        stream.write(struct.pack('<H', 0xFFFF))

        # Module records
        for module in self.modules:
            self._write_module_record(stream, module)

        # TERMINATOR (0x0010)
        stream.write(struct.pack('<H', 0x0010))
        stream.write(struct.pack('<I', 0))

        # 圧縮
        return compress_vba(stream.getvalue())

    def _write_module_record(self, stream: io.BytesIO, module: VBAModule):
        """モジュールレコードを書き込む"""
        name_bytes = module.name.encode('ascii')

        # MODULENAME (0x0019)
        stream.write(struct.pack('<HI', 0x0019, len(name_bytes)))
        stream.write(name_bytes)

        # MODULENAMEUNICODE (0x0047)
        name_unicode = module.name.encode('utf-16-le')
        stream.write(struct.pack('<HI', 0x0047, len(name_unicode)))
        stream.write(name_unicode)

        # MODULESTREAMNAME (0x001A)
        stream.write(struct.pack('<HI', 0x001A, len(name_bytes)))
        stream.write(name_bytes)
        stream.write(struct.pack('<HI', 0x0032, len(name_unicode)))
        stream.write(name_unicode)

        # MODULEDOCSTRING (0x001C)
        stream.write(struct.pack('<HI', 0x001C, 0))
        stream.write(struct.pack('<HI', 0x0048, 0))

        # MODULEOFFSET (0x0031)
        stream.write(struct.pack('<HI', 0x0031, 4))
        stream.write(struct.pack('<I', 0))  # Offset will be 0 for our simple case

        # MODULEHELPCONTEXT (0x001E)
        stream.write(struct.pack('<HI', 0x001E, 4))
        stream.write(struct.pack('<I', 0))

        # MODULECOOKIE (0x002C)
        stream.write(struct.pack('<HI', 0x002C, 2))
        stream.write(struct.pack('<H', 0xFFFF))

        # MODULETYPE (0x0021 for procedural, 0x0022 for document)
        if module.module_type in ("ThisWorkbook", "Sheet"):
            stream.write(struct.pack('<HI', 0x0022, 0))  # Document module
        else:
            stream.write(struct.pack('<HI', 0x0021, 0))  # Procedural module

        # MODULEREADONLY (optional) - skip
        # MODULEPRIVATE (optional) - skip

        # TERMINATOR (0x002B)
        stream.write(struct.pack('<HI', 0x002B, 0))

    def _create_module_stream(self, module: VBAModule) -> bytes:
        """モジュールストリームを作成する"""
        # Attribute VB_Name を追加
        code = module.code
        if not code.startswith('Attribute VB_Name'):
            code = f'Attribute VB_Name = "{module.name}"\r\n' + code

        # Shift_JISエンコード
        code_bytes = code.encode('cp932', errors='replace')

        # 圧縮
        return compress_vba(code_bytes)

    def _create_project_stream(self) -> bytes:
        """PROJECTストリームを作成する（テキスト形式）"""
        lines = [
            f'ID="{{{self._generate_guid()}}}"',
            f'Document=ThisWorkbook/&H00000000',
        ]

        # モジュール参照を追加
        for module in self.modules:
            if module.module_type == "Module":
                lines.append(f'Module={module.name}')
            elif module.module_type == "Class":
                lines.append(f'Class={module.name}')

        lines.extend([
            f'Name="{self.PROJECT_NAME}"',
            'HelpContextID="0"',
            'VersionCompatible32="393222000"',
            'CMG="0705030503050305"',
            'DPB="0E0CD11CD11CD1"',
            'GC="1517131713171317"',
            '',
            '[Host Extender Info]',
            '&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000',
            '',
            '[Workspace]',
        ])

        return '\r\n'.join(lines).encode('cp932')

    def _create_projectwm_stream(self) -> bytes:
        """PROJECTwmストリームを作成する（モジュール名のUnicodeマッピング）"""
        stream = io.BytesIO()

        for module in self.modules:
            # モジュール名（null終端）
            name_bytes = module.name.encode('ascii') + b'\x00'
            stream.write(name_bytes)
            # Unicode名（null終端、UTF-16LE）
            name_unicode = module.name.encode('utf-16-le') + b'\x00\x00'
            stream.write(name_unicode)

        # 終端
        stream.write(b'\x00\x00')

        return stream.getvalue()

    def _generate_guid(self) -> str:
        """GUIDを生成する"""
        import uuid
        return str(uuid.uuid4()).upper()

    def generate(self) -> bytes:
        """vbaProject.binを生成する"""
        # OLE複合ドキュメントを作成
        # olefile は読み取り専用なので、直接バイナリを構築する必要がある

        # シンプルなアプローチ：テンプレートベースの生成
        return self._generate_ole_compound()

    def _generate_ole_compound(self) -> bytes:
        """OLE複合ドキュメントを生成する"""
        # MS-CFB形式のバイナリを構築

        # ストリームデータを準備
        streams: Dict[str, bytes] = {}

        # VBA/dirストリーム
        streams['VBA/dir'] = self._create_dir_stream()

        # VBA/_VBA_PROJECTストリーム
        streams['VBA/_VBA_PROJECT'] = self._create_vba_project_stream()

        # モジュールストリーム
        for module in self.modules:
            streams[f'VBA/{module.stream_name}'] = self._create_module_stream(module)

        # PROJECTストリーム
        streams['PROJECT'] = self._create_project_stream()

        # PROJECTwmストリーム
        streams['PROJECTwm'] = self._create_projectwm_stream()

        # OLE複合ドキュメントとして組み立て
        return self._build_cfb(streams)

    def _create_vba_project_stream(self) -> bytes:
        """_VBA_PROJECTストリームを作成する"""
        # 固定ヘッダー
        header = bytes([
            0xCC, 0x61,  # Reserved1
            0x00, 0x00,  # Version
            0x00,        # Reserved2
            0x00,        # Reserved3
        ])

        # パフォーマンスキャッシュ（空）
        return header + b'\x00' * 4

    def _build_cfb(self, streams: Dict[str, bytes]) -> bytes:
        """CFB（複合ファイルバイナリ）を構築する"""
        # セクターサイズ
        sector_size = 512
        mini_sector_size = 64
        mini_stream_cutoff = 4096

        # ヘッダーを構築
        output = io.BytesIO()

        # CFBヘッダー（512バイト）
        header = self._build_cfb_header(len(streams), sector_size)
        output.write(header)

        # FATセクター、ディレクトリ、ストリームデータを構築
        # これは複雑な処理なので、簡略化版を使用

        # ディレクトリエントリを構築
        dir_entries = self._build_directory_entries(streams)

        # ストリームデータを配置
        stream_data = self._build_stream_sectors(streams, sector_size)

        # FATを構築
        fat = self._build_fat(len(dir_entries) // sector_size + 1,
                              len(stream_data) // sector_size + 1)

        # 全てを結合
        output.write(fat)
        output.write(dir_entries)
        output.write(stream_data)

        # セクター境界にパディング
        data = output.getvalue()
        padding = sector_size - (len(data) % sector_size)
        if padding < sector_size:
            data += b'\x00' * padding

        return data

    def _build_cfb_header(self, num_streams: int, sector_size: int) -> bytes:
        """CFBヘッダーを構築する"""
        header = io.BytesIO()

        # シグネチャ
        header.write(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1')

        # CLSID (16バイト、ゼロ)
        header.write(b'\x00' * 16)

        # Minor Version
        header.write(struct.pack('<H', 0x003E))

        # Major Version (3 = 512バイトセクター)
        header.write(struct.pack('<H', 0x0003))

        # Byte Order (little endian)
        header.write(struct.pack('<H', 0xFFFE))

        # Sector Shift (9 = 512バイト)
        header.write(struct.pack('<H', 0x0009))

        # Mini Sector Shift (6 = 64バイト)
        header.write(struct.pack('<H', 0x0006))

        # Reserved (6バイト)
        header.write(b'\x00' * 6)

        # Total Sectors in FAT (Directory Sectors for V3)
        header.write(struct.pack('<I', 0))

        # Number of FAT Sectors
        header.write(struct.pack('<I', 1))

        # First Directory Sector
        header.write(struct.pack('<I', 0))

        # Transaction Signature Number
        header.write(struct.pack('<I', 0))

        # Mini Stream Cutoff Size
        header.write(struct.pack('<I', 0x00001000))

        # First Mini FAT Sector
        header.write(struct.pack('<I', 0xFFFFFFFE))

        # Number of Mini FAT Sectors
        header.write(struct.pack('<I', 0))

        # First DIFAT Sector
        header.write(struct.pack('<I', 0xFFFFFFFE))

        # Number of DIFAT Sectors
        header.write(struct.pack('<I', 0))

        # DIFAT (109エントリ、残りはセクター0を指す)
        header.write(struct.pack('<I', 0))  # FAT at sector 0
        header.write(b'\xFF\xFF\xFF\xFF' * 108)  # 残りは未使用

        # 512バイトにパディング
        data = header.getvalue()
        data += b'\x00' * (512 - len(data))

        return data

    def _build_directory_entries(self, streams: Dict[str, bytes]) -> bytes:
        """ディレクトリエントリを構築する"""
        entries = io.BytesIO()

        # ルートエントリ
        entries.write(self._create_dir_entry("Root Entry", 0x05, -1, -1, 1))

        # VBAストレージ
        entries.write(self._create_dir_entry("VBA", 0x01, -1, 2, -1))

        # その他のエントリ
        entry_id = 2
        for name, data in streams.items():
            entry_type = 0x02  # Stream
            entries.write(self._create_dir_entry(name.split('/')[-1], entry_type, -1, -1, -1, len(data)))
            entry_id += 1

        # セクター境界にパディング
        data = entries.getvalue()
        padding = 512 - (len(data) % 512)
        if padding < 512:
            data += b'\x00' * padding

        return data

    def _create_dir_entry(self, name: str, entry_type: int, left: int, right: int, child: int, size: int = 0) -> bytes:
        """ディレクトリエントリ（128バイト）を作成する"""
        entry = io.BytesIO()

        # Name (64バイト、UTF-16LE)
        name_bytes = name.encode('utf-16-le')[:62]
        entry.write(name_bytes)
        entry.write(b'\x00' * (64 - len(name_bytes)))

        # Name Length (バイト単位、null終端含む)
        entry.write(struct.pack('<H', len(name_bytes) + 2))

        # Object Type
        entry.write(bytes([entry_type]))

        # Color Flag (1 = black)
        entry.write(bytes([0x01]))

        # Left Sibling ID
        entry.write(struct.pack('<i', left))

        # Right Sibling ID
        entry.write(struct.pack('<i', right))

        # Child ID
        entry.write(struct.pack('<i', child))

        # CLSID (16バイト)
        entry.write(b'\x00' * 16)

        # State Bits
        entry.write(struct.pack('<I', 0))

        # Creation Time (8バイト)
        entry.write(b'\x00' * 8)

        # Modification Time (8バイト)
        entry.write(b'\x00' * 8)

        # Starting Sector
        entry.write(struct.pack('<I', 0))

        # Stream Size (8バイト for v4, 4バイト for v3)
        entry.write(struct.pack('<Q', size))

        return entry.getvalue()

    def _build_stream_sectors(self, streams: Dict[str, bytes], sector_size: int) -> bytes:
        """ストリームセクターを構築する"""
        output = io.BytesIO()

        for name, data in streams.items():
            output.write(data)
            # セクター境界にパディング
            padding = sector_size - (len(data) % sector_size)
            if padding < sector_size:
                output.write(b'\x00' * padding)

        return output.getvalue()

    def _build_fat(self, dir_sectors: int, stream_sectors: int) -> bytes:
        """FATセクターを構築する"""
        fat = io.BytesIO()

        total_sectors = 1 + dir_sectors + stream_sectors  # FAT + Dir + Streams

        # FATエントリ（各4バイト）
        for i in range(128):  # 512バイト / 4バイト = 128エントリ
            if i == 0:
                fat.write(struct.pack('<I', 0xFFFFFFFD))  # FAT sector
            elif i < total_sectors:
                if i == total_sectors - 1:
                    fat.write(struct.pack('<I', 0xFFFFFFFE))  # End of chain
                else:
                    fat.write(struct.pack('<I', i + 1))  # Next sector
            else:
                fat.write(struct.pack('<I', 0xFFFFFFFF))  # Free

        return fat.getvalue()


def generate_vba_project(modules: Dict[str, str]) -> bytes:
    """モジュール辞書からvbaProject.binを生成する

    Args:
        modules: {モジュール名: VBAコード} の辞書

    Returns:
        vbaProject.binのバイナリデータ
    """
    generator = VBAProjectGenerator()

    for name, code in modules.items():
        # モジュールタイプを判定
        if name == "ThisWorkbook":
            module_type = "ThisWorkbook"
        elif name.startswith("Sheet") or name.endswith("_View"):
            module_type = "Sheet"
        elif "Class" in name or name[0].isupper() and "_" not in name:
            module_type = "Class"
        else:
            module_type = "Module"

        generator.add_module(name, code, module_type)

    return generator.generate()


def main():
    """コマンドラインからの実行"""
    import argparse

    parser = argparse.ArgumentParser(description="VBAプロジェクトバイナリを生成する")
    parser.add_argument("--output", "-o", type=Path, default=Path("vbaProject.bin"),
                        help="出力ファイルパス")
    parser.add_argument("--vba-dir", type=Path, default=None,
                        help="VBAソースディレクトリ（.bas, .cls ファイル）")

    args = parser.parse_args()

    # VBAモジュールを読み込む
    modules = {}

    if args.vba_dir and args.vba_dir.exists():
        for path in sorted(args.vba_dir.glob("*.bas")):
            code = path.read_text(encoding="utf-8")
            # Attribute VB_Name を解析
            name = path.stem
            modules[name] = code

        for path in sorted(args.vba_dir.glob("*.cls")):
            code = path.read_text(encoding="utf-8")
            name = path.stem
            modules[name] = code
    else:
        # デフォルトのサンプルモジュール
        modules["Module1"] = """
Option Explicit

Public Sub HelloWorld()
    MsgBox "Hello, World!"
End Sub
"""

    # 生成
    vba_binary = generate_vba_project(modules)

    # 保存
    args.output.write_bytes(vba_binary)
    print(f"Generated: {args.output} ({len(vba_binary)} bytes)")


if __name__ == "__main__":
    main()
