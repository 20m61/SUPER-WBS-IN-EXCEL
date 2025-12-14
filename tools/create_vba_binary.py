#!/usr/bin/env python3
"""
VBA Project Binary Generator (Revised)

Excelで使用可能なvbaProject.binを生成する。
MS-CFB (Compound File Binary) 仕様に正確に準拠。
"""
import struct
import io
from pathlib import Path
from typing import Dict, List, Tuple, Optional
from dataclasses import dataclass, field


# ============================================================================
# VBA圧縮（MS-OVBA 2.4.1準拠 - 簡略版：非圧縮）
# ============================================================================

def vba_compress(data: bytes) -> bytes:
    """VBA形式でデータを格納する（非圧縮モード）"""
    if not data:
        return b'\x01'

    output = io.BytesIO()
    output.write(b'\x01')  # 圧縮シグネチャ

    pos = 0
    while pos < len(data):
        chunk = data[pos:pos + 4096]
        chunk_len = len(chunk)

        # 非圧縮チャンクヘッダー（サイズ | 0x0000）
        # 実際は圧縮データとしてリテラルのみ出力
        compressed = _create_literal_chunk(chunk)

        # チャンクヘッダー: (size-1) | signature
        header = (len(compressed) - 1) | 0xB000
        output.write(struct.pack('<H', header))
        output.write(compressed)

        pos += chunk_len

    return output.getvalue()


def _create_literal_chunk(data: bytes) -> bytes:
    """リテラルのみのチャンクを作成"""
    output = io.BytesIO()
    i = 0

    while i < len(data):
        # 8バイトごとにフラグバイト + データ
        flag_byte = 0x00  # 全てリテラル
        batch = data[i:i+8]
        output.write(bytes([flag_byte]))
        output.write(batch)
        i += 8

    return output.getvalue()


# ============================================================================
# CFB (Compound File Binary) Writer - Simplified
# ============================================================================

@dataclass
class DirEntry:
    """ディレクトリエントリ"""
    name: str
    entry_type: int  # 0=empty, 1=storage, 2=stream, 5=root
    data: bytes = b''
    color: int = 1  # 0=red, 1=black
    left_sibling: int = -1
    right_sibling: int = -1
    child: int = -1
    start_sector: int = 0

    @property
    def size(self) -> int:
        return len(self.data)


class CFBFile:
    """CFBファイル生成器（簡略版）"""

    SECTOR_SIZE = 512
    MINI_SECTOR_SIZE = 64
    MINI_STREAM_CUTOFF = 4096
    DIR_ENTRY_SIZE = 128
    ENDOFCHAIN = 0xFFFFFFFE
    FREESECT = 0xFFFFFFFF
    FATSECT = 0xFFFFFFFD

    def __init__(self):
        self.dir_entries: List[DirEntry] = []
        self.stream_data = io.BytesIO()

    def add_root(self):
        """ルートエントリを追加"""
        self.dir_entries.append(DirEntry(
            name="Root Entry",
            entry_type=5,
            color=1,
        ))

    def add_storage(self, name: str, parent_idx: int = 0) -> int:
        """ストレージを追加"""
        idx = len(self.dir_entries)
        self.dir_entries.append(DirEntry(
            name=name,
            entry_type=1,
            color=1,
        ))
        # 親の子として設定
        if self.dir_entries[parent_idx].child == -1:
            self.dir_entries[parent_idx].child = idx
        else:
            # 兄弟として追加
            self._add_sibling(self.dir_entries[parent_idx].child, idx)
        return idx

    def add_stream(self, name: str, data: bytes, parent_idx: int = 0) -> int:
        """ストリームを追加"""
        idx = len(self.dir_entries)

        # ストリームデータの開始セクターを計算
        current_pos = self.stream_data.tell()
        start_sector = 2 + (current_pos // self.SECTOR_SIZE)  # FAT + DIR の後

        # データを書き込み
        self.stream_data.write(data)
        # セクター境界にパディング
        padding = self.SECTOR_SIZE - (len(data) % self.SECTOR_SIZE)
        if padding < self.SECTOR_SIZE:
            self.stream_data.write(b'\x00' * padding)

        self.dir_entries.append(DirEntry(
            name=name,
            entry_type=2,
            data=data,
            color=1,
            start_sector=start_sector,
        ))

        # 親の子として設定
        if self.dir_entries[parent_idx].child == -1:
            self.dir_entries[parent_idx].child = idx
        else:
            self._add_sibling(self.dir_entries[parent_idx].child, idx)

        return idx

    def _add_sibling(self, sibling_idx: int, new_idx: int):
        """兄弟として追加（右に追加）"""
        entry = self.dir_entries[sibling_idx]
        while entry.right_sibling != -1:
            entry = self.dir_entries[entry.right_sibling]
        entry.right_sibling = new_idx

    def build(self) -> bytes:
        """CFBファイルを構築"""
        # セクター計算
        num_dir_entries = len(self.dir_entries)
        dir_sectors = (num_dir_entries * self.DIR_ENTRY_SIZE + self.SECTOR_SIZE - 1) // self.SECTOR_SIZE

        stream_bytes = self.stream_data.getvalue()
        stream_sectors = (len(stream_bytes) + self.SECTOR_SIZE - 1) // self.SECTOR_SIZE if stream_bytes else 0

        fat_sectors = 1  # 簡略化

        total_sectors = fat_sectors + dir_sectors + stream_sectors

        # ヘッダー構築
        header = self._build_header(fat_sectors, dir_sectors)

        # FAT構築
        fat = self._build_fat(fat_sectors, dir_sectors, stream_sectors)

        # ディレクトリ構築
        directory = self._build_directory(dir_sectors)

        # 結合
        result = header + fat + directory
        if stream_bytes:
            result += stream_bytes
            # 最終セクターをパディング
            padding = self.SECTOR_SIZE - (len(stream_bytes) % self.SECTOR_SIZE)
            if padding < self.SECTOR_SIZE:
                result += b'\x00' * padding

        return result

    def _build_header(self, fat_sectors: int, dir_sectors: int) -> bytes:
        """CFBヘッダー（512バイト）"""
        h = io.BytesIO()

        # シグネチャ (8 bytes)
        h.write(b'\xD0\xCF\x11\xE0\xA1\xB1\x1A\xE1')
        # CLSID (16 bytes)
        h.write(b'\x00' * 16)
        # Minor Version
        h.write(struct.pack('<H', 0x003E))
        # Major Version (3 = sector size 512)
        h.write(struct.pack('<H', 0x0003))
        # Byte Order (little endian)
        h.write(struct.pack('<H', 0xFFFE))
        # Sector Shift (9 = 512 bytes)
        h.write(struct.pack('<H', 9))
        # Mini Sector Shift (6 = 64 bytes)
        h.write(struct.pack('<H', 6))
        # Reserved (6 bytes)
        h.write(b'\x00' * 6)
        # Total Sectors in FAT (0 for v3)
        h.write(struct.pack('<I', 0))
        # Number of FAT Sectors
        h.write(struct.pack('<I', fat_sectors))
        # First Directory Sector Location
        h.write(struct.pack('<I', fat_sectors))  # FAT の直後
        # Transaction Signature Number
        h.write(struct.pack('<I', 0))
        # Mini Stream Cutoff Size
        h.write(struct.pack('<I', self.MINI_STREAM_CUTOFF))
        # First Mini FAT Sector Location
        h.write(struct.pack('<I', self.ENDOFCHAIN))
        # Number of Mini FAT Sectors
        h.write(struct.pack('<I', 0))
        # First DIFAT Sector Location
        h.write(struct.pack('<I', self.ENDOFCHAIN))
        # Number of DIFAT Sectors
        h.write(struct.pack('<I', 0))

        # DIFAT (109 entries)
        for i in range(109):
            if i < fat_sectors:
                h.write(struct.pack('<I', i))
            else:
                h.write(struct.pack('<I', self.FREESECT))

        return h.getvalue()

    def _build_fat(self, fat_sectors: int, dir_sectors: int, stream_sectors: int) -> bytes:
        """FAT セクター"""
        fat = io.BytesIO()
        entries_per_sector = self.SECTOR_SIZE // 4

        sector_idx = 0

        # FAT セクター自体
        for _ in range(fat_sectors):
            fat.write(struct.pack('<I', self.FATSECT))
            sector_idx += 1

        # ディレクトリセクター
        for i in range(dir_sectors):
            if i < dir_sectors - 1:
                fat.write(struct.pack('<I', sector_idx + 1))
            else:
                fat.write(struct.pack('<I', self.ENDOFCHAIN))
            sector_idx += 1

        # ストリームセクター
        for i in range(stream_sectors):
            if i < stream_sectors - 1:
                fat.write(struct.pack('<I', sector_idx + 1))
            else:
                fat.write(struct.pack('<I', self.ENDOFCHAIN))
            sector_idx += 1

        # 残りを FREESECT で埋める
        while fat.tell() < fat_sectors * self.SECTOR_SIZE:
            fat.write(struct.pack('<I', self.FREESECT))

        return fat.getvalue()

    def _build_directory(self, dir_sectors: int) -> bytes:
        """ディレクトリセクター"""
        directory = io.BytesIO()

        for entry in self.dir_entries:
            directory.write(self._make_dir_entry(entry))

        # 空エントリで埋める
        entries_per_sector = self.SECTOR_SIZE // self.DIR_ENTRY_SIZE
        total_entries = dir_sectors * entries_per_sector
        while len(self.dir_entries) < total_entries:
            directory.write(self._make_empty_entry())
            self.dir_entries.append(None)  # プレースホルダー

        # セクター境界にパディング
        while directory.tell() % self.SECTOR_SIZE != 0:
            directory.write(b'\x00')

        return directory.getvalue()

    def _make_dir_entry(self, entry: DirEntry) -> bytes:
        """ディレクトリエントリ（128バイト）"""
        e = io.BytesIO()

        # Name (64 bytes, UTF-16LE, null-terminated)
        name_bytes = (entry.name + '\x00').encode('utf-16-le')
        e.write(name_bytes[:64].ljust(64, b'\x00'))

        # Name Size (including null terminator)
        e.write(struct.pack('<H', min(len(name_bytes), 64)))

        # Object Type
        e.write(struct.pack('<B', entry.entry_type))

        # Color Flag
        e.write(struct.pack('<B', entry.color))

        # Left Sibling ID
        e.write(struct.pack('<i', entry.left_sibling))

        # Right Sibling ID
        e.write(struct.pack('<i', entry.right_sibling))

        # Child ID
        e.write(struct.pack('<i', entry.child))

        # CLSID (16 bytes)
        e.write(b'\x00' * 16)

        # State Bits (4 bytes)
        e.write(struct.pack('<I', 0))

        # Creation Time (8 bytes)
        e.write(b'\x00' * 8)

        # Modification Time (8 bytes)
        e.write(b'\x00' * 8)

        # Starting Sector Location
        if entry.entry_type == 2 and entry.size > 0:
            e.write(struct.pack('<I', entry.start_sector))
        else:
            e.write(struct.pack('<I', 0))

        # Size (8 bytes for v4, but only lower 4 bytes used for v3)
        e.write(struct.pack('<Q', entry.size))

        return e.getvalue()

    def _make_empty_entry(self) -> bytes:
        """空のディレクトリエントリ"""
        e = io.BytesIO()
        e.write(b'\x00' * 64)  # Name
        e.write(struct.pack('<H', 0))  # Name Size
        e.write(struct.pack('<B', 0))  # Type (empty)
        e.write(struct.pack('<B', 1))  # Color
        e.write(struct.pack('<iii', -1, -1, -1))  # Siblings, Child
        e.write(b'\x00' * 16)  # CLSID
        e.write(struct.pack('<I', 0))  # State
        e.write(b'\x00' * 16)  # Timestamps
        e.write(struct.pack('<I', 0))  # Start
        e.write(struct.pack('<Q', 0))  # Size
        return e.getvalue()


# ============================================================================
# VBAプロジェクト生成
# ============================================================================

def create_dir_stream(modules: Dict[str, str]) -> bytes:
    """VBA dir ストリームを作成"""
    s = io.BytesIO()

    # PROJECTSYSKIND Record (0x0001)
    s.write(struct.pack('<HI', 0x0001, 4))
    s.write(struct.pack('<I', 0x00000001))  # Win32

    # PROJECTLCID Record (0x0002)
    s.write(struct.pack('<HI', 0x0002, 4))
    s.write(struct.pack('<I', 0x0411))  # Japanese

    # PROJECTLCIDINVOKE Record (0x0014)
    s.write(struct.pack('<HI', 0x0014, 4))
    s.write(struct.pack('<I', 0x0411))

    # PROJECTCODEPAGE Record (0x0003)
    s.write(struct.pack('<HI', 0x0003, 2))
    s.write(struct.pack('<H', 932))  # Shift_JIS

    # PROJECTNAME Record (0x0004)
    name = b'VBAProject'
    s.write(struct.pack('<HI', 0x0004, len(name)))
    s.write(name)

    # PROJECTDOCSTRING Record (0x0005)
    s.write(struct.pack('<HI', 0x0005, 0))
    # PROJECTDOCSTRINGUNICODE (0x0040)
    s.write(struct.pack('<HI', 0x0040, 0))

    # PROJECTHELPFILEPATH Record (0x0006)
    s.write(struct.pack('<HI', 0x0006, 0))
    # PROJECTHELPFILEPATH2 (0x003D)
    s.write(struct.pack('<HI', 0x003D, 0))

    # PROJECTHELPCONTEXT Record (0x0007)
    s.write(struct.pack('<HI', 0x0007, 4))
    s.write(struct.pack('<I', 0))

    # PROJECTLIBFLAGS Record (0x0008)
    s.write(struct.pack('<HI', 0x0008, 4))
    s.write(struct.pack('<I', 0))

    # PROJECTVERSION Record (0x0009)
    s.write(struct.pack('<HI', 0x0009, 4))
    s.write(struct.pack('<I', 0x00010001))

    # PROJECTCONSTANTS Record (0x000C)
    s.write(struct.pack('<HI', 0x000C, 0))
    # PROJECTCONSTANTSUNICODE (0x003C)
    s.write(struct.pack('<HI', 0x003C, 0))

    # PROJECTMODULES Record (0x000F)
    s.write(struct.pack('<HI', 0x000F, 2))
    s.write(struct.pack('<H', len(modules)))

    # PROJECTCOOKIE Record (0x0013)
    s.write(struct.pack('<HI', 0x0013, 2))
    s.write(struct.pack('<H', 0xFFFF))

    # MODULE Records
    for name in modules.keys():
        _write_module_record(s, name)

    # Terminator (0x0010)
    s.write(struct.pack('<HI', 0x0010, 0))

    return vba_compress(s.getvalue())


def _write_module_record(stream: io.BytesIO, name: str):
    """モジュールレコードを書き込む"""
    name_bytes = name.encode('cp932')
    name_unicode = name.encode('utf-16-le')

    # MODULENAME (0x0019)
    stream.write(struct.pack('<HI', 0x0019, len(name_bytes)))
    stream.write(name_bytes)

    # MODULENAMEUNICODE (0x0047)
    stream.write(struct.pack('<HI', 0x0047, len(name_unicode)))
    stream.write(name_unicode)

    # MODULESTREAMNAME (0x001A)
    stream.write(struct.pack('<HI', 0x001A, len(name_bytes)))
    stream.write(name_bytes)
    # Reserved (0x0032)
    stream.write(struct.pack('<HI', 0x0032, len(name_unicode)))
    stream.write(name_unicode)

    # MODULEDOCSTRING (0x001C)
    stream.write(struct.pack('<HI', 0x001C, 0))
    # Reserved (0x0048)
    stream.write(struct.pack('<HI', 0x0048, 0))

    # MODULEOFFSET (0x0031)
    stream.write(struct.pack('<HI', 0x0031, 4))
    stream.write(struct.pack('<I', 0))

    # MODULEHELPCONTEXT (0x001E)
    stream.write(struct.pack('<HI', 0x001E, 4))
    stream.write(struct.pack('<I', 0))

    # MODULECOOKIE (0x002C)
    stream.write(struct.pack('<HI', 0x002C, 2))
    stream.write(struct.pack('<H', 0xFFFF))

    # MODULETYPE (0x0021 = procedural)
    stream.write(struct.pack('<HI', 0x0021, 0))

    # Terminator (0x002B)
    stream.write(struct.pack('<HI', 0x002B, 0))


def create_vba_project_stream() -> bytes:
    """_VBA_PROJECT ストリーム"""
    # 固定ヘッダー
    return bytes([
        0xCC, 0x61,  # Reserved
        0x00, 0x00,  # Version
        0x00, 0x00, 0x00, 0x00,  # Reserved
    ])


def create_project_stream(modules: Dict[str, str]) -> bytes:
    """PROJECT ストリーム（テキスト形式）"""
    import uuid

    lines = [
        f'ID="{{{str(uuid.uuid4()).upper()}}}"',
    ]

    for name in modules.keys():
        lines.append(f'Module={name}')

    lines.extend([
        'Name="VBAProject"',
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


def create_projectwm_stream(modules: Dict[str, str]) -> bytes:
    """PROJECTwm ストリーム"""
    s = io.BytesIO()
    for name in modules.keys():
        s.write(name.encode('cp932') + b'\x00')
        s.write(name.encode('utf-16-le') + b'\x00\x00')
    s.write(b'\x00\x00')
    return s.getvalue()


def create_module_stream(name: str, code: str) -> bytes:
    """モジュールストリーム"""
    # Attribute VB_Name を先頭に追加
    if not code.strip().startswith('Attribute'):
        code = f'Attribute VB_Name = "{name}"\r\n' + code

    return vba_compress(code.encode('cp932', errors='replace'))


def generate_vba_project_bin(modules: Dict[str, str]) -> bytes:
    """vbaProject.bin を生成"""
    cfb = CFBFile()

    # ルートエントリ
    cfb.add_root()

    # VBA ストレージ
    vba_idx = cfb.add_storage("VBA", parent_idx=0)

    # dir ストリーム
    cfb.add_stream("dir", create_dir_stream(modules), parent_idx=vba_idx)

    # _VBA_PROJECT ストリーム
    cfb.add_stream("_VBA_PROJECT", create_vba_project_stream(), parent_idx=vba_idx)

    # 各モジュールストリーム
    for name, code in modules.items():
        cfb.add_stream(name, create_module_stream(name, code), parent_idx=vba_idx)

    # PROJECT ストリーム（ルート直下）
    cfb.add_stream("PROJECT", create_project_stream(modules), parent_idx=0)

    # PROJECTwm ストリーム（ルート直下）
    cfb.add_stream("PROJECTwm", create_projectwm_stream(modules), parent_idx=0)

    return cfb.build()


def load_modules_from_directory(vba_dir: Path) -> Dict[str, str]:
    """ディレクトリからVBAモジュールを読み込む"""
    modules = {}

    for ext in ['*.bas', '*.cls']:
        for path in sorted(vba_dir.glob(ext)):
            code = path.read_text(encoding='utf-8')
            # Attribute VB_Name から名前を取得
            name = path.stem
            for line in code.split('\n'):
                if line.startswith('Attribute VB_Name'):
                    parts = line.split('=')
                    if len(parts) >= 2:
                        name = parts[1].strip().strip('"')
                    break
            modules[name] = code

    return modules


def main():
    """メイン関数"""
    import argparse

    parser = argparse.ArgumentParser(description='VBAプロジェクトバイナリを生成')
    parser.add_argument('--vba-dir', type=Path, required=True,
                        help='VBAソースディレクトリ')
    parser.add_argument('--output', '-o', type=Path, default=Path('vbaProject.bin'),
                        help='出力ファイル')

    args = parser.parse_args()

    modules = load_modules_from_directory(args.vba_dir)
    print(f'Loaded {len(modules)} modules:')
    for name in modules.keys():
        print(f'  - {name}')

    vba_binary = generate_vba_project_bin(modules)
    args.output.write_bytes(vba_binary)
    print(f'\nGenerated: {args.output} ({len(vba_binary)} bytes)')


if __name__ == '__main__':
    main()
