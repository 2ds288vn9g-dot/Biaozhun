from io import BytesIO
from pathlib import Path, PurePosixPath
from zipfile import ZipFile, ZIP_DEFLATED
import tarfile
import tempfile
import shutil
import subprocess
import os
import re
import unicodedata
import copy
from dataclasses import dataclass
try:
    import tkinter as tk
    from tkinter import filedialog
except ImportError:  # Render/Linux 服务器没有桌面文件选择窗口。
    tk = None
    filedialog = None

import pandas as pd
import streamlit as st
from lxml import etree

from engine import DEFAULT_ASSOCIATION, DEFAULT_PRIMARY_ASSOCIATION, detect, merge

W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS = {'w': W}
q = lambda name: f'{{{W}}}{name}'

ROOT = Path(__file__).resolve().parent
TEMPLATE = ROOT / 'assets' / '固定标准封面模板.docx'
MIME_DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
DEFAULT_PROPOSAL_LEADER = '刘伟朝'
LABELS = {
    'ics': 'ICS', 'ccs': 'CCS', 'standard_number': '标准编号（同时同步正文页眉）',
    'chinese_title': '中文标准名称', 'english_title': '英文标准名称',
    'international_alignment': '国际标准一致性标识', 'publication_date': '发布日期',
    'implementation_date': '实施日期', 'publisher_top': '第一协会名称',
    'publisher_bottom': '第二协会名称',
    'drafting_organizations': '前言：本文件起草单位',
    'drafting_authors': '前言：本文件主要起草人',
}


def template_bytes() -> bytes:
    if not TEMPLATE.exists():
        raise RuntimeError('内置封面模板缺失，请联系管理员。')
    return TEMPLATE.read_bytes()


def choose_save_folder(title: str) -> str:
    if tk is None or filedialog is None:
        raise RuntimeError('在线版不支持选择本机保存位置，请直接使用下方下载按钮。')
    root = tk.Tk()
    root.withdraw()
    root.attributes('-topmost', True)
    try:
        return filedialog.askdirectory(title=title)
    finally:
        root.destroy()


def save_to_folder(folder: str, name: str, data: bytes) -> str:
    path = Path(folder) / name
    path.write_bytes(data)
    return str(path)


def normalize_name(value: object) -> str:
    # 仅消除同一标准名称的展示格式差异，不进行模糊匹配。
    text = unicodedata.normalize('NFKC', str(value or ''))
    text = text.replace('\u200b', '').replace('\ufeff', '')
    text = re.sub(r'^\s*\d+\s*[.．、]\s*', '', text)
    # 母文件夹常带“（6个项目）”之类的数量说明，不属于标准题目本身。
    text = re.sub(r'\s*[（(]\s*\d+\s*个?项目?\s*[）)]\s*$', '', text)
    return re.sub(r'\s+', '', text).strip()


def verify_standard_document(content: bytes, standard: str) -> tuple[bool, str]:
    """以 Word 封面内容复核，防止“编制说明”等同题目文件被误改。"""
    try:
        with ZipFile(BytesIO(content)) as archive:
            document = archive.read('word/document.xml')
        root = etree.fromstring(document)
        text = ''.join(root.xpath('.//w:body//w:t/text()', namespaces={
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }))
    except Exception:
        return False, '无法读取 Word 正文'
    normalized = normalize_name(text)
    if standard not in normalized:
        return False, '正文封面未找到对应标准题目'
    has_ics = bool(re.search(r'\bICS\b', text, re.I))
    has_ccs = bool(re.search(r'\bCCS\b', text, re.I))
    has_number = bool(re.search(r'(?:T|Q)/[^\s]+(?:\s+[^\s—-]+){0,3}\s*[—-]\s*(?:\d{4}|X{2,})', text, re.I))
    has_standard_cover = '团体标准' in text
    if (has_ics and has_ccs) or (has_number and has_standard_cover):
        return True, ''
    return False, '未识别到 ICS/CCS 或标准编号与“团体标准”封面特征'


def target_priority(filename: str, standard: str) -> int:
    """2=明确标注标准文本；1=文件名就是标准题目；0=不作为候选。"""
    stem = Path(filename).stem
    if '标准文本' in stem:
        return 2
    return 1 if normalize_name(stem) == standard else 0


@dataclass
class ArchiveItem:
    filename: str
    data: bytes = b''
    is_dir: bool = False


def archive_stem(filename: str) -> str:
    lower = filename.lower()
    for suffix in ('.tar.gz', '.tar.bz2', '.tar.xz', '.tgz', '.tbz', '.txz', '.zip', '.rar', '.7z', '.tar'):
        if lower.endswith(suffix):
            return filename[:-len(suffix)]
    return Path(filename).stem


def archive_suffix(filename: str) -> str:
    lower = filename.lower()
    for suffix in ('.tar.gz', '.tar.bz2', '.tar.xz', '.tgz', '.tbz', '.txz', '.zip', '.rar', '.7z', '.tar'):
        if lower.endswith(suffix):
            return filename[-len(suffix):]
    raise RuntimeError('无法确定压缩包格式。')


def safe_archive_name(name: str) -> str:
    pure = PurePosixPath(name.replace('\\', '/'))
    if pure.is_absolute() or '..' in pure.parts:
        raise RuntimeError('压缩包包含不安全的文件路径，已停止处理。')
    return str(pure)


def items_from_folder(folder: Path) -> list[ArchiveItem]:
    items = []
    for path in sorted(folder.rglob('*')):
        relative = safe_archive_name(path.relative_to(folder).as_posix())
        if path.is_symlink():
            raise RuntimeError('压缩包包含链接文件，已停止处理。')
        items.append(ArchiveItem(relative + ('/' if path.is_dir() else ''), b'' if path.is_dir() else path.read_bytes(), path.is_dir()))
    return items


def read_archive(data: bytes, filename: str) -> list[ArchiveItem]:
    """读取常见压缩格式；目录和其他文件保持不变。"""
    lower = filename.lower()
    if lower.endswith('.zip'):
        with ZipFile(BytesIO(data)) as archive:
            return [ArchiveItem(safe_archive_name(info.filename), b'' if info.is_dir() else archive.read(info.filename), info.is_dir()) for info in archive.infolist()]
    if lower.endswith(('.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz', '.tar.xz', '.txz')):
        with tarfile.open(fileobj=BytesIO(data), mode='r:*') as archive:
            items = []
            for member in archive.getmembers():
                name = safe_archive_name(member.name)
                if member.isdir():
                    items.append(ArchiveItem(name + '/', b'', True))
                elif member.isfile():
                    source = archive.extractfile(member)
                    items.append(ArchiveItem(name, source.read() if source else b'', False))
                else:
                    raise RuntimeError('压缩包包含不支持的链接或设备文件，已停止处理。')
            return items
    if lower.endswith('.7z'):
        try:
            import py7zr
        except ImportError as error:
            raise RuntimeError('缺少 7Z 支持组件，请重新安装资料包依赖。') from error
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / 'source.7z'
            destination = root / 'extracted'
            destination.mkdir()
            source.write_bytes(data)
            with py7zr.SevenZipFile(source, mode='r') as archive:
                for name in archive.getnames():
                    safe_archive_name(name)
                archive.extractall(path=destination)
            return items_from_folder(destination)
    if lower.endswith('.rar'):
        try:
            import rarfile
        except ImportError as error:
            raise RuntimeError('缺少 RAR 支持组件，请重新安装资料包依赖。') from error
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            source = root / 'source.rar'
            destination = root / 'extracted'
            destination.mkdir()
            source.write_bytes(data)
            try:
                with rarfile.RarFile(source) as archive:
                    for info in archive.infolist():
                        safe_archive_name(info.filename)
                    archive.extractall(path=destination)
            except rarfile.RarCannotExec as error:
                raise RuntimeError('读取 RAR 需要电脑已安装 7-Zip、WinRAR 或 unrar 工具。') from error
            return items_from_folder(destination)
    raise RuntimeError('不支持的压缩包格式。支持 ZIP、RAR、7Z、TAR、TAR.GZ/TGZ、TAR.BZ2/TBZ、TAR.XZ/TXZ。')


def write_archive(items: list[ArchiveItem], source_filename: str) -> bytes:
    """按输入压缩格式重新打包，绝不将其他格式悄悄转换为 ZIP。"""
    # 有些压缩软件没有为目录写入“目录标记”，而是写了一个与目录同名的
    # 普通空白项；资源管理器便会显示一批没有扩展名的白色文件。只要某项
    # 存在下级路径，就一律把该项正规化为目录，删除其伪文件内容。
    names = {item.filename.rstrip('/') for item in items}
    implicit_dirs = {
        name for name in names
        if name and any(other.startswith(name + '/') for other in names)
    }
    normalized_items = []
    emitted_dirs = set()
    for item in items:
        name = item.filename.rstrip('/')
        if item.is_dir or name in implicit_dirs:
            if name and name not in emitted_dirs:
                normalized_items.append(ArchiveItem(name + '/', b'', True))
                emitted_dirs.add(name)
            continue
        normalized_items.append(item)
    items = normalized_items
    lower = source_filename.lower()
    if lower.endswith('.zip'):
        result = BytesIO()
        with ZipFile(result, 'w', ZIP_DEFLATED) as archive:
            for item in items:
                archive.writestr(item.filename, item.data)
        return result.getvalue()
    if lower.endswith(('.tar', '.tar.gz', '.tgz', '.tar.bz2', '.tbz', '.tar.xz', '.txz')):
        mode = 'w'
        if lower.endswith(('.tar.gz', '.tgz')):
            mode = 'w:gz'
        elif lower.endswith(('.tar.bz2', '.tbz')):
            mode = 'w:bz2'
        elif lower.endswith(('.tar.xz', '.txz')):
            mode = 'w:xz'
        result = BytesIO()
        with tarfile.open(fileobj=result, mode=mode) as archive:
            for item in items:
                info = tarfile.TarInfo(item.filename.rstrip('/'))
                if item.is_dir:
                    info.type = tarfile.DIRTYPE
                    archive.addfile(info)
                else:
                    info.size = len(item.data)
                    archive.addfile(info, BytesIO(item.data))
        return result.getvalue()
    if lower.endswith('.7z'):
        try:
            import py7zr
        except ImportError as error:
            raise RuntimeError('缺少 7Z 支持组件，请重新安装资料包依赖。') from error
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / 'contents'
            root.mkdir()
            for item in items:
                path = root / safe_archive_name(item.filename.rstrip('/'))
                if item.is_dir:
                    path.mkdir(parents=True, exist_ok=True)
                else:
                    path.parent.mkdir(parents=True, exist_ok=True)
                    path.write_bytes(item.data)
            target = Path(temp) / 'result.7z'
            with py7zr.SevenZipFile(target, mode='w') as archive:
                archive.writeall(root, arcname='')
            return target.read_bytes()
    if lower.endswith('.rar'):
        rar = shutil.which('rar')
        if not rar:
            raise RuntimeError('RAR 同格式输出需要电脑已安装 WinRAR，并将 rar 命令加入系统 PATH。为避免改变格式，本次未生成文件。')
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp) / 'contents'
            root.mkdir()
            for item in items:
                path = root / safe_archive_name(item.filename.rstrip('/'))
                if item.is_dir:
                    path.mkdir(parents=True, exist_ok=True)
                else:
                    path.parent.mkdir(parents=True, exist_ok=True)
                    path.write_bytes(item.data)
            target = Path(temp) / 'result.rar'
            completed = subprocess.run([rar, 'a', '-r', str(target), '.'], cwd=root, capture_output=True, text=True)
            if completed.returncode not in (0, 1) or not target.exists():
                raise RuntimeError('RAR 同格式输出失败，请确认 WinRAR 已正确安装。')
            return target.read_bytes()
    raise RuntimeError('不支持的压缩包格式。')


def read_mapping(data: bytes) -> tuple[dict, list[str]]:
    book = pd.ExcelFile(BytesIO(data))
    frame = None
    for sheet in book.sheet_names:
        raw = pd.read_excel(book, sheet_name=sheet, header=None, dtype=str).fillna('')
        for row_index in range(len(raw)):
            headers = [str(value).strip().replace('\n', '') for value in raw.iloc[row_index].tolist()]
            # 只接受同一行内 B/P/Q/R 固定表头，避免误读其他工作表中的“标准名称”。
            if len(headers) >= 18 and headers[1] == '标准名称' and headers[15] == '起草单位' and headers[16] == '起草人':
                frame = raw.iloc[row_index + 1:, :18].copy().fillna('')
                break
        if frame is not None:
            break
    if frame is None:
        raise RuntimeError('未找到包含 B列“标准名称”、P列“起草单位”、Q列“起草人”、R列“起草人与起草单位对应情况”的 Excel 数据表。')
    mapping, duplicates = {}, []
    for _, row in frame.iterrows():
        name = normalize_name(row.iloc[1])       # B 列：标准名称
        if not name:
            continue
        item = {'organizations': str(row.iloc[15]).strip(), 'authors': str(row.iloc[16]).strip(), 'relations': str(row.iloc[17]).strip()}
        if name in mapping:
            duplicates.append(name)
            mapping.pop(name, None)
        elif name not in duplicates:
            mapping[name] = item
    return mapping, sorted(set(duplicates))


def values_editor(source: bytes, key: str) -> dict:
    found = detect(source)
    st.subheader('识别结果（可直接修改）')
    left, right = st.columns(2)
    values = {}
    for i, (name, label) in enumerate(LABELS.items()):
        with (left if i % 2 == 0 else right):
            values[name] = st.text_input(label, found.get(name, ''), key=f'{key}_{name}')
    st.caption('第二协会默认“青海省标准化协会”；标准编号下方窗体域将自动填写相同编号。')
    values['publisher_top'] = values.get('publisher_top') or DEFAULT_PRIMARY_ASSOCIATION
    values['publisher_bottom'] = values.get('publisher_bottom') or DEFAULT_ASSOCIATION
    values['replaces_standard'] = values['standard_number']
    return values


def first_organization(value: str) -> str:
    return next((part.strip() for part in re.split(r'[、；;\n]+', value or '') if part.strip()), '')


def standard_with_prefix(old: str, prefix: str) -> str:
    prefix = (prefix or '').strip()
    if not prefix:
        return old
    match = re.match(r'^(?:T|Q)/[^\s]+(\s+.+)$', old or '', re.I)
    if not match:
        raise RuntimeError(f'无法识别原标准编号：{old or "空"}')
    return prefix.rstrip() + match.group(1)


def cell_text(cell) -> str:
    return ''.join(cell.xpath('.//w:t/text()', namespaces=NS)).strip()


def set_text(nodes, value: str) -> bool:
    if not nodes:
        return False
    # Word/WPS 常把同一单元格拆成多段 run，且每段颜色或字体不完全相同。
    # 替换时写入原内容最长的那段 run（通常是正文值本身），并清空其余段，
    # 让新内容继承一个完整的原有格式，避免出现深浅不一致。
    target = max(nodes, key=lambda node: len(node.text or ''))
    target.text = value
    for node in nodes:
        if node is not target:
            node.text = ''
    return True


def set_cell(cell, value: str) -> bool:
    return set_text(cell.xpath('.//w:t', namespaces=NS), value)


def set_cell_with_label_color(label_cell, value_cell, value: str) -> bool:
    """写入表格值，并只继承左侧标签的文字颜色（不改变字号/加粗）。"""
    label_nodes = label_cell.xpath('.//w:t', namespaces=NS)
    value_nodes = value_cell.xpath('.//w:t', namespaces=NS)
    if not label_nodes or not value_nodes:
        return False
    label_node = max(label_nodes, key=lambda node: len(node.text or ''))
    value_node = max(value_nodes, key=lambda node: len(node.text or ''))
    label_run = label_node.getparent()
    value_run = value_node.getparent()
    label_rpr = label_run.find(q('rPr'))
    source_color = label_rpr.find(q('color')) if label_rpr is not None else None
    value_rpr = value_run.find(q('rPr'))
    if value_rpr is None:
        value_rpr = etree.Element(q('rPr'))
        value_run.insert(0, value_rpr)
    # 先移除目标原来的蓝灰等显式颜色；标签没有显式颜色时使用 Word 默认黑色。
    for color in value_rpr.findall(q('color')):
        value_rpr.remove(color)
    if source_color is not None:
        value_rpr.append(copy.deepcopy(source_color))
    value_node.text = value
    for node in value_nodes:
        if node is not value_node:
            node.text = ''
    return True


def set_paragraph_label_value_preserving_layout(paragraph, label: str, value: str) -> bool:
    """只替换同段落“标签：值”中的值，保留其他 run、制表符与换行。"""
    nodes = paragraph.xpath('.//w:t', namespaces=NS)
    current = ''.join(node.text or '' for node in nodes)
    match = re.search(re.escape(label) + r'\s*[：:]?\s*', current)
    if not match or not nodes:
        return False
    value_start = match.end()
    position = 0
    for index, node in enumerate(nodes):
        text = node.text or ''
        end = position + len(text)
        if end <= value_start:
            position = end
            continue
        # 值与标签在同一个 run 时，拆出一个克隆 run；否则只改原值所在 run。
        offset = max(0, value_start - position)
        if offset:
            prefix = text[:offset]
            node.text = prefix
            run = node.getparent()
            new_run = copy.deepcopy(run)
            new_texts = new_run.xpath('.//w:t', namespaces=NS)
            if not new_texts:
                return False
            new_texts[0].text = value
            for extra in new_texts[1:]:
                extra.text = ''
            run.addnext(new_run)
        else:
            node.text = value
        # 标签后的旧值不再保留，但标签前的内容、制表符、换行节点不动。
        for later in nodes[index + 1:]:
            later.text = ''
        return True
    return False


def replace_exact_in_paragraphs(root, old: str, new: str) -> int:
    """替换正文、表格、文本框和页眉中与旧编号完全相同的文本。"""
    changed = 0
    for paragraph in root.xpath('.//w:p', namespaces=NS):
        texts = paragraph.xpath('.//w:t', namespaces=NS)
        joined = ''.join(node.text or '' for node in texts)
        if old and old in joined:
            # 标准编号常在换行前后拆成多个 run。优先只替换前缀，避免重新
            # 分配文本而改变原有换行位置和版式。
            old_prefix = re.match(r'^(?:T|Q)/[^\s]+', old, re.I)
            new_prefix = re.match(r'^(?:T|Q)/[^\s]+', new, re.I)
            prefix_changed = False
            if old_prefix and new_prefix:
                for node in texts:
                    if old_prefix.group(0) in (node.text or ''):
                        node.text = (node.text or '').replace(old_prefix.group(0), new_prefix.group(0), 1)
                        prefix_changed = True
                        break
            if not prefix_changed:
                # 非前缀替换时先尝试限定在同一文字段中，保持原 run 属性。
                for node in texts:
                    if old in (node.text or ''):
                        node.text = (node.text or '').replace(old, new, 1)
                        prefix_changed = True
                        break
            if not prefix_changed:
                # 极少数跨 run 的普通文本才退回原方法。
                set_text(texts, joined.replace(old, new, 1))
            changed += 1
    return changed


def replace_label_value(root, labels: tuple[str, ...], value: str) -> bool:
    target = {normalize_name(label) for label in labels}
    for row in root.xpath('.//w:tr', namespaces=NS):
        cells = row.xpath('./w:tc', namespaces=NS)
        for index, cell in enumerate(cells[:-1]):
            if normalize_name(cell_text(cell)) in target:
                return set_cell_with_label_color(cell, cells[index + 1], value)
    for paragraph in root.xpath('.//w:p', namespaces=NS):
        current = ''.join(paragraph.xpath('.//w:t/text()', namespaces=NS))
        for label in labels:
            if label in current:
                return set_paragraph_label_value_preserving_layout(paragraph, label, value)
    return False


def parse_relations(value: str) -> list[tuple[str, str]]:
    pairs = []
    for unit, names in re.findall(r'([^：:；;\n]+)[：:]\s*([^；;\n]+)', value or ''):
        for name in re.split(r'[、，,\s]+', names.strip()):
            if name:
                pairs.append((name, unit.strip()))
    return pairs


def table_header_index(rows, start: int = 0) -> int:
    """返回“序号/姓名/单位/职务或职称”四列表头所在行。"""
    for index, row in enumerate(rows[start:], start):
        text = normalize_name(''.join(cell_text(cell) for cell in row.xpath('./w:tc', namespaces=NS)))
        if '姓名' in text and '单位' in text and ('职务' in text or '职称' in text) and '序号' in text:
            return index
    return -1


def update_drafter_table(root, relation_text: str) -> bool:
    people = parse_relations(relation_text)
    if not people:
        return False
    tables = root.xpath('.//w:tbl', namespaces=NS)
    for table_position, labelled_table in enumerate(tables):
        rows = labelled_table.xpath('./w:tr', namespaces=NS)
        # 不能只按“姓名/单位”找表：编制说明常同时有“参与起草单位”表。
        # 只定位“标准起草人”标题之后的表格；标题与表格有时是相邻的两张表。
        drafter_label_index = next((
            i for i, row in enumerate(rows)
            if '标准起草人' in ''.join(cell_text(cell) for cell in row.xpath('./w:tc', namespaces=NS))
        ), -1)
        if drafter_label_index < 0:
            continue
        target_table = labelled_table
        header_index = table_header_index(rows, drafter_label_index + 1)
        if header_index < 0:
            # 某些模板把“标准起草人”作为上一张表的最后一行，列标题在下一张表。
            for candidate in tables[table_position + 1:]:
                candidate_rows = candidate.xpath('./w:tr', namespaces=NS)
                candidate_header = table_header_index(candidate_rows)
                if candidate_header >= 0:
                    target_table, rows, header_index = candidate, candidate_rows, candidate_header
                    break
        if header_index < 0:
            continue
        header_cells = rows[header_index].xpath('./w:tc', namespaces=NS)
        heads = [normalize_name(cell_text(cell)) for cell in header_cells]
        name_index = next((i for i, h in enumerate(heads) if h == '姓名'), -1)
        unit_index = next((i for i, h in enumerate(heads) if h == '单位'), -1)
        title_index = next((i for i, h in enumerate(heads) if '职务' in h or '职称' in h), -1)
        serial_index = next((i for i, h in enumerate(heads) if h == '序号'), -1)
        if min(name_index, unit_index, title_index, serial_index) < 0:
            continue
        data_rows = []
        for row in rows[header_index + 1:]:
            cells = row.xpath('./w:tc', namespaces=NS)
            if len(cells) <= max(name_index, unit_index, title_index, serial_index):
                break
            if not cell_text(cells[serial_index]).strip().isdigit():
                break
            data_rows.append(row)
        existing = {}
        for row in data_rows:
            cells = row.xpath('./w:tc', namespaces=NS)
            existing[cell_text(cells[name_index]).strip()] = row
        # 空白表没有示例数据行时，复制表头结构新增数据行，保留边框和列宽。
        template = data_rows[0] if data_rows else rows[header_index]
        for row in data_rows:
            target_table.remove(row)
        # 不能按 table.insert(行号) 插入：Word 表格在数据行之前还有 tblPr、
        # tblGrid 等 XML 节点，绝对索引会把人员行误插到“标准起草人”标题上方。
        # 以表头行作为锚点，逐行紧接在“序号/姓名/单位/职务”之后插入。
        insert_after = rows[header_index]
        for serial, (name, unit) in enumerate(people, 1):
            row = copy.deepcopy(existing.get(name, template))
            cells = row.xpath('./w:tc', namespaces=NS)
            set_cell(cells[serial_index], str(serial))
            set_cell(cells[name_index], name)
            set_cell(cells[unit_index], unit)
            # 同名人员保留原职务/职称；新增人员清空该栏。
            if name not in existing:
                set_cell(cells[title_index], '')
            insert_after.addnext(row)
            insert_after = row
        return True
    return False


def update_vote_title(root, association: str) -> bool:
    suffix = '团体标准审查投票单'
    for paragraph in root.xpath('.//w:p', namespaces=NS):
        texts = paragraph.xpath('.//w:t', namespaces=NS)
        current = ''.join(node.text or '' for node in texts).strip()
        if suffix in current:
            return set_text(texts, association.strip() + suffix)
    return False


def find_office_executable() -> str | None:
    """兼容 Windows 默认安装位置，无需手动把 soffice 写入 PATH。"""
    found = shutil.which('soffice') or shutil.which('libreoffice')
    if found:
        return found
    candidates = []
    for variable in ('ProgramFiles', 'ProgramFiles(x86)', 'LOCALAPPDATA'):
        base = os.environ.get(variable)
        if base:
            candidates.extend([
                Path(base) / 'LibreOffice' / 'program' / 'soffice.exe',
                Path(base) / 'LibreOffice' / 'program' / 'soffice.com',
            ])
    return next((str(path) for path in candidates if path.exists()), None)


def windows_office_convert(data: bytes, filename: str, target_extension: str) -> bytes | None:
    """LibreOffice 未安装时，使用已安装的 Microsoft Word 或 WPS 转换。"""
    if os.name != 'nt':
        return None
    try:
        import win32com.client
    except ImportError:
        return None
    # Microsoft Word 与 WPS 的 COM 接口均兼容 Documents.Open / SaveAs(2)。
    # DOCX=16，DOC=0；在 Word/WPS 中保持原有页面版式后再继续处理。
    file_format = 16 if target_extension.lower() == 'docx' else 0
    with tempfile.TemporaryDirectory() as temp:
        source = Path(temp) / Path(filename).name
        target = Path(temp) / (source.stem + '.' + target_extension)
        source.write_bytes(data)
        for prog_id in ('Word.Application', 'Kwps.Application'):
            app = document = None
            try:
                app = win32com.client.DispatchEx(prog_id)
                app.Visible = False
                if hasattr(app, 'DisplayAlerts'):
                    app.DisplayAlerts = 0
                document = app.Documents.Open(str(source), ReadOnly=True)
                if hasattr(document, 'SaveAs2'):
                    document.SaveAs2(str(target), FileFormat=file_format)
                else:
                    document.SaveAs(str(target), FileFormat=file_format)
                if target.exists():
                    return target.read_bytes()
            except Exception:
                pass
            finally:
                if document is not None:
                    try:
                        document.Close(False)
                    except Exception:
                        pass
                if app is not None:
                    try:
                        app.Quit()
                    except Exception:
                        pass
    return None


def com_replace_all(document, old: str, new: str) -> bool:
    if not old or old == new:
        return False
    finder = document.Content.Duplicate.Find
    finder.ClearFormatting()
    finder.Replacement.ClearFormatting()
    # wdReplaceAll=2, wdFindContinue=1；Word 与 WPS 均兼容此调用。
    return bool(finder.Execute(FindText=old, ReplaceWith=new, Replace=2, Forward=True, Wrap=1, Format=False))


def com_replace_standard_number(document, old_number: str, new_number: str) -> bool:
    """按“标准编号”表头定位其数据单元格，完整写入标准文本中的编号。"""
    # 不能只替换 T/SPCH 等前缀：原文件的序号或年份也可能不同。先依据
    # 表头定位编号数据格，覆盖为标准文本的完整标准编号，并保留该格原字体。
    for table in document.Tables:
        for row_index in range(1, table.Rows.Count + 1):
            row = table.Rows(row_index)
            for column_index in range(1, row.Cells.Count + 1):
                header_cell = row.Cells(column_index)
                header_text = normalize_name(str(header_cell.Range.Text).rstrip('\r\x07'))
                if header_text != normalize_name('标准编号') or row_index >= table.Rows.Count:
                    continue
                try:
                    number_cell = table.Cell(row_index + 1, column_index)
                except Exception:
                    # 合并单元格表格使用下一行的同序单元格作为备用。
                    number_cell = table.Rows(row_index + 1).Cells(column_index)
                original_range = number_cell.Range
                start, end = original_range.Start, original_range.End - 1
                original_font = com_font_snapshot(original_range)
                document.Range(start, end).Text = new_number
                new_range = document.Range(start, start + len(new_number))
                com_apply_font(new_range, original_font)
                return str(new_range.Text).rstrip('\r\x07') == new_number

    # 未识别为标准编号列时，才退回精确全文/前缀替换；不通过时会由调用方报错。
    changed = com_replace_all(document, old_number, new_number)
    old_prefix = re.match(r'^(?:T|Q)/[^\s]+', old_number or '', re.I)
    new_prefix = re.match(r'^(?:T|Q)/[^\s]+', new_number or '', re.I)
    if old_prefix and new_prefix:
        changed = com_replace_all(document, old_prefix.group(0), new_prefix.group(0)) or changed
    return changed


FONT_ATTRIBUTES = (
        'Name', 'NameAscii', 'NameFarEast', 'NameOther', 'Size', 'Bold', 'Italic',
        'Underline', 'StrikeThrough', 'DoubleStrikeThrough', 'Subscript', 'Superscript',
        'Color', 'ColorIndex', 'Shadow', 'Outline', 'Emboss', 'Engrave', 'AllCaps',
        'SmallCaps', 'Spacing', 'Scaling', 'Position', 'Kerning', 'Hidden',
)


def com_font_snapshot(source_range) -> dict[str, object]:
    snapshot = {}
    for attribute in FONT_ATTRIBUTES:
        try:
            snapshot[attribute] = getattr(source_range.Font, attribute)
        except Exception:
            pass
    return snapshot


def com_apply_font(target_range, snapshot: dict[str, object]) -> None:
    for attribute, value in snapshot.items():
        try:
            setattr(target_range.Font, attribute, value)
        except Exception:
            pass


def com_copy_font(source_range, target_range) -> None:
    """把标签的字符格式复制给新值，避免 Word/WPS 按旧值的浅色样式继承。"""
    com_apply_font(target_range, com_font_snapshot(source_range))


def com_replace_label_value(document, labels: tuple[str, ...], value: str) -> bool:
    """在 Word/WPS 原生 DOC 中只替换标签后的值，保留段落与表格格式。"""
    for table in document.Tables:
        for row in table.Rows:
            for index in range(1, row.Cells.Count):
                label_range = row.Cells(index).Range
                label_text = str(label_range.Text).rstrip('\r\x07').strip()
                if normalize_name(label_text) in {normalize_name(label) for label in labels}:
                    value_range = row.Cells(index + 1).Range
                    # 去掉单元格结束标记，不重建单元格、边框或行高。
                    document.Range(value_range.Start, value_range.End - 1).Text = value
                    com_copy_font(label_range, document.Range(value_range.Start, value_range.Start + len(value)))
                    return True
    for paragraph in document.Paragraphs:
        paragraph_range = paragraph.Range
        text = str(paragraph_range.Text).rstrip('\r\x07')
        for label in labels:
            match = re.search(re.escape(label) + r'\s*[：:]?\s*', text)
            if match:
                value_start = paragraph_range.Start + match.end()
                document.Range(value_start, paragraph_range.Start + len(text)).Text = value
                # 原意见表的标签和旧值处于同一原生段落。显式使用标签的字体，
                # 令新值与“负责起草单位”完全一致，而不沿用旧值可能存在的浅色属性。
                com_copy_font(
                    document.Range(paragraph_range.Start, value_start),
                    document.Range(value_start, value_start + len(value)),
                )
                return True
    return False


def windows_direct_edit_doc(data: bytes, filename: str, role: str, old_number: str, new_number: str, record: dict, routing_association: str, proposal_leader: str) -> bytes | None:
    """直接修改原始 DOC，避免 DOC→DOCX→DOC 造成的重新排版。"""
    if os.name != 'nt' or role == 'compilation':
        return None
    try:
        import win32com.client
    except ImportError:
        return None
    failures = []
    with tempfile.TemporaryDirectory() as temp:
        source = Path(temp) / Path(filename).name
        source.write_bytes(data)
        for prog_id in ('Word.Application', 'Kwps.Application'):
            app = document = None
            try:
                app = win32com.client.DispatchEx(prog_id)
                app.Visible = False
                if hasattr(app, 'DisplayAlerts'):
                    app.DisplayAlerts = 0
                document = app.Documents.Open(str(source), ReadOnly=False)
                number_changed = com_replace_standard_number(document, old_number, new_number)
                label_changed = True
                if role == 'proposal':
                    com_replace_label_value(document, ('项目申请单位',), first_organization(record['organizations']))
                    com_replace_label_value(document, ('归口单位', '第一起草单位'), routing_association)
                    com_replace_label_value(document, ('共同发起单位', '参加起草单位'), record['organizations'])
                    com_replace_label_value(document, ('负责人',), proposal_leader)
                    com_replace_label_value(document, ('联系电话',), '')
                    com_replace_label_value(document, ('E-mail', 'Email', '邮箱'), '')
                elif role == 'opinion':
                    label_changed = com_replace_label_value(document, ('负责起草单位',), first_organization(record['organizations']))
                elif role == 'vote':
                    for paragraph in document.Paragraphs:
                        text = str(paragraph.Range.Text).rstrip('\r\x07')
                        if '团体标准审查投票单' in text:
                            paragraph.Range.Text = routing_association + '团体标准审查投票单\r'
                            break
                # 意见汇总处理表绝不能把“没有真正替换”的原文件作为成功结果返回。
                if role == 'opinion' and not number_changed:
                    raise RuntimeError(f'未在原 DOC 中找到可替换的标准编号/前缀：{old_number}')
                if role == 'opinion' and not label_changed:
                    raise RuntimeError('未在原 DOC 中找到“负责起草单位”字段')
                document.Save()
                saved = source.read_bytes()
                if role == 'opinion' and saved == data:
                    raise RuntimeError('Word/WPS 保存后文件字节未变化，未产生有效修改')
                return saved
            except Exception as error:
                failures.append(f'{prog_id}：{error}')
            finally:
                if document is not None:
                    try:
                        document.Close(False)
                    except Exception:
                        pass
                if app is not None:
                    try:
                        app.Quit()
                    except Exception:
                        pass
    if failures:
        raise RuntimeError('原 DOC 直改失败；' + '；'.join(failures))
    return None


def office_convert(data: bytes, filename: str, target_extension: str) -> bytes:
    """通过 LibreOffice、Microsoft Word 或 WPS 保持 .doc 输入仍输出 .doc。"""
    executable = find_office_executable()
    if not executable:
        converted = windows_office_convert(data, filename, target_extension)
        if converted is not None:
            return converted
        raise RuntimeError('处理 DOC 需要 LibreOffice、Microsoft Word 或 WPS 其中之一。程序已自动检测并尝试转换；请安装任一办公软件后重试。')
    with tempfile.TemporaryDirectory() as temp:
        root = Path(temp)
        source = root / Path(filename).name
        output = root / 'output'
        output.mkdir()
        source.write_bytes(data)
        command = [executable, '--headless', '--convert-to', target_extension, '--outdir', str(output), str(source)]
        completed = subprocess.run(command, capture_output=True, text=True)
        target = output / (source.stem + '.' + target_extension)
        if completed.returncode != 0 or not target.exists():
            raise RuntimeError(f'无法转换 {Path(filename).suffix} 文件，请确认 LibreOffice 已正确安装。')
        return target.read_bytes()


def as_docx(data: bytes, filename: str) -> bytes:
    suffix = Path(filename).suffix.lower()
    if suffix == '.docx':
        return data
    if suffix == '.doc':
        return office_convert(data, filename, 'docx')
    raise RuntimeError('当前文件不是可编辑的 Word DOC/DOCX。')


def restore_word_format(data: bytes, original_filename: str) -> bytes:
    return office_convert(data, Path(original_filename).with_suffix('.docx').name, 'doc') if Path(original_filename).suffix.lower() == '.doc' else data


def edit_docx(data: bytes, role: str, old_number: str, new_number: str, record: dict, routing_association: str, proposal_leader: str) -> bytes:
    """保留 DOCX 内部关系与版式，只定向覆盖已确认字段。"""
    source = ZipFile(BytesIO(data))
    document = etree.fromstring(source.read('word/document.xml'))
    replace_exact_in_paragraphs(document, old_number, new_number)
    if role == 'proposal':
        # 兼容两类立项建议书：传统表使用“归口单位”，新版表使用“第一起草单位”。
        # 两个位置统一取界面中指定的归口单位；新版的参加起草单位取 Excel P 列全文。
        replace_label_value(document, ('项目申请单位',), first_organization(record['organizations']))
        replace_label_value(document, ('归口单位',), routing_association)
        replace_label_value(document, ('共同发起单位',), record['organizations'])
        replace_label_value(document, ('第一起草单位',), routing_association)
        replace_label_value(document, ('参加起草单位',), record['organizations'])
        replace_label_value(document, ('负责人',), proposal_leader)
        replace_label_value(document, ('联系电话',), '')
        replace_label_value(document, ('E-mail', 'Email', '邮箱'), '')
    elif role == 'compilation':
        replace_label_value(document, ('负责起草单位',), record['organizations'])
        if not update_drafter_table(document, record.get('relations', '')):
            raise RuntimeError('未找到“标准起草人”姓名/单位/职务表格，或 Excel R 列为空')
    elif role == 'opinion':
        replace_label_value(document, ('负责起草单位',), first_organization(record['organizations']))
    elif role == 'vote':
        if not update_vote_title(document, routing_association):
            raise RuntimeError('未找到“团体标准审查投票单”标题')
    output = BytesIO()
    with ZipFile(output, 'w', ZIP_DEFLATED) as archive:
        for info in source.infolist():
            payload = source.read(info.filename)
            if info.filename == 'word/document.xml':
                payload = etree.tostring(document, xml_declaration=True, encoding='UTF-8', standalone=True)
            elif info.filename.startswith('word/header') and info.filename.endswith('.xml'):
                header = etree.fromstring(payload)
                if replace_exact_in_paragraphs(header, old_number, new_number):
                    payload = etree.tostring(header, xml_declaration=True, encoding='UTF-8', standalone=True)
            archive.writestr(info, payload)
    return output.getvalue()


def file_role(filename: str) -> str:
    stem = Path(filename).stem
    if '立项建议书' in stem:
        return 'proposal'
    if '编制说明' in stem:
        return 'compilation'
    if '意见汇总处理表' in stem:
        return 'opinion'
    if '投票单' in stem:
        return 'vote'
    if '审查专家签字表决结果汇总表' in stem:
        return 'expert_summary'
    return ''


def is_office_temporary(filename: str) -> bool:
    """Word/WPS/LibreOffice 产生的锁定文件不属于资料包内容。"""
    name = Path(filename).name
    # Word/WPS 常见 ~$xxx.docx；部分 WPS/LibreOffice 还会生成 .~xxx 或 .~lock.xxx#。
    return name.startswith(('~$', '.~'))


def batch_generate(archive_data: bytes, archive_name: str, excel_data: bytes, publisher_top: str, publisher_bottom: str, standard_prefix: str, proposal_leader: str) -> tuple[bytes, list[str], int, list[str]]:
    mapping, duplicates = read_mapping(excel_data)
    original = read_archive(archive_data, archive_name)
    unchanged: dict[str, str] = {}
    changed: set[str] = set()
    matched_folders: dict[str, str] = {}
    for info in original:
        if info.is_dir:
            continue
        parts = Path(info.filename).parts[:-1]
        for index in range(1, len(parts) + 1):
            folder = '/'.join(parts[:index])
            standard = normalize_name(parts[index - 1])
            if standard in mapping:
                matched_folders[folder] = standard
    # 先对每个已匹配母文件夹收集候选项，再决定唯一目标，绝不边扫描边修改。
    candidates: dict[str, list[tuple[object, int]]] = {}
    rejected: dict[str, list[str]] = {}
    for info in original:
        if info.is_dir or is_office_temporary(info.filename) or Path(info.filename).suffix.lower() not in {'.docx', '.doc'}:
            continue
        parts = [part for part in Path(info.filename).parts[:-1] if part not in {'.', '/'}]
        parent_paths = ['/'.join(parts[:index]) for index in range(1, len(parts) + 1)]
        matched = [path for path in reversed(parent_paths) if path in matched_folders]
        if not matched:
            continue
        mother_path = matched[0]
        standard = matched_folders[mother_path]
        priority = target_priority(Path(info.filename).name, standard)
        if not priority:
            continue
        try:
            passed, reason = verify_standard_document(as_docx(info.data, info.filename), standard)
        except Exception as error:
            passed, reason = False, str(error)
        if passed:
            candidates.setdefault(mother_path, []).append((info, priority))
        else:
            rejected.setdefault(mother_path, []).append(f'{Path(info.filename).name}（{reason}）')

    selected: dict[str, str] = {}
    for mother_path, options in candidates.items():
        highest = max(priority for _, priority in options)
        winners = [info for info, priority in options if priority == highest]
        mother = Path(mother_path).name
        if len(winners) == 1:
            selected[winners[0].filename] = mother_path
        else:
            names = '、'.join(Path(info.filename).name for info in winners)
            unchanged[mother_path] = f'{mother}：检测到多个候选标准文本（{names}），为避免误改已跳过'

    configurations = {}
    for filename, mother_path in selected.items():
        try:
            source_item = next(item for item in original if item.filename == filename)
            old_number = detect(as_docx(source_item.data, source_item.filename)).get('standard_number', '').strip()
            new_number = standard_with_prefix(old_number, standard_prefix)
            if not old_number:
                raise RuntimeError('标准文本未识别到标准编号')
            configurations[mother_path] = {'standard_file': filename, 'old_number': old_number, 'new_number': new_number, 'record': mapping[matched_folders[mother_path]]}
        except Exception as error:
            unchanged[mother_path] = f'{Path(mother_path).name}：{error}'

    count = 0
    output_items = []
    for info in original:
        if is_office_temporary(info.filename):
            continue
        content = info.data
        parts = [part for part in Path(info.filename).parts[:-1] if part not in {'.', '/'}]
        parent_paths = ['/'.join(parts[:index]) for index in range(1, len(parts) + 1)]
        mothers = [path for path in reversed(parent_paths) if path in configurations]
        mother_path = mothers[0] if mothers else ''
        # 已匹配标准母文件夹内的 PDF 均为不需要交付的子文件，直接不写入输出包。
        if mother_path and not info.is_dir and Path(info.filename).suffix.lower() == '.pdf':
            continue
        if mother_path and not info.is_dir and Path(info.filename).suffix.lower() in {'.docx', '.doc'}:
            mother = Path(mother_path).name
            config = configurations[mother_path]
            try:
                direct_doc = None
                if Path(info.filename).suffix.lower() == '.doc' and info.filename != config['standard_file']:
                    direct_doc = windows_direct_edit_doc(
                        content,
                        info.filename,
                        file_role(info.filename),
                        config['old_number'],
                        config['new_number'],
                        config['record'],
                        publisher_bottom.strip() or DEFAULT_ASSOCIATION,
                        proposal_leader.strip() or DEFAULT_PROPOSAL_LEADER,
                    )
                    # 意见汇总处理表的原 DOC 必须由 Word/WPS 原格式直改；若 COM
                    # 不可用，绝不悄悄退回 DOCX 转换，以免生成已知会重排的文件。
                    if file_role(info.filename) == 'opinion' and direct_doc is None:
                        raise RuntimeError(
                            '意见汇总处理表为原始 DOC，未检测到可用的 Word/WPS 直改组件。'
                            '请用资料包内“启动.bat”启动（会安装 pywin32），并确保已安装且关闭 WPS 或 Microsoft Word。'
                        )
                if direct_doc is not None:
                    content = direct_doc
                else:
                    working = as_docx(content, info.filename)
                    if info.filename == config['standard_file']:
                        values = detect(working)
                        values['standard_number'] = config['new_number']
                        values['drafting_organizations'] = config['record']['organizations']
                        values['drafting_authors'] = config['record']['authors']
                        values['publisher_top'] = publisher_top.strip() or DEFAULT_PRIMARY_ASSOCIATION
                        values['publisher_bottom'] = publisher_bottom.strip() or DEFAULT_ASSOCIATION
                        values['replaces_standard'] = config['new_number']
                        content = merge(template_bytes(), working, values)
                    else:
                        role = file_role(info.filename)
                        content = edit_docx(working, role, config['old_number'], config['new_number'], config['record'], publisher_bottom.strip() or DEFAULT_ASSOCIATION, proposal_leader.strip() or DEFAULT_PROPOSAL_LEADER)
                    content = restore_word_format(content, info.filename)
                changed.add(mother_path)
                count += 1
            except Exception as error:
                unchanged.setdefault(mother_path, f'{mother}：{Path(info.filename).name} 处理失败：{error}')
        elif mother_path and not info.is_dir and file_role(info.filename):
            unchanged.setdefault(mother_path, f'{Path(mother_path).name}：{Path(info.filename).name} 不是可编辑的 DOC/DOCX，无法在不改变文件格式的前提下安全修改')
        output_items.append(ArchiveItem(info.filename, content, info.is_dir))
    for folder, standard in matched_folders.items():
        if folder not in changed and folder not in unchanged:
            details = '；'.join(rejected.get(folder, []))
            unchanged[folder] = f'{Path(folder).name}：未找到通过严格识别的标准文本' + (f'（{details}）' if details else '')
    for name in duplicates:
        unchanged.setdefault(name, 'Excel 中存在重复标准名称，已跳过以避免信息错位')
    return write_archive(output_items, archive_name), list(unchanged.values()), count, duplicates


st.set_page_config(page_title='标准封面修改平台', layout='wide')
st.title('标准封面修改平台')
st.caption('固定使用已确认的标准封面模板。仅上传待修改文件即可。')
single, batch = st.tabs(['单个生成', '批量生成'])

with single:
    source = st.file_uploader('拖放或点击上传待修改原文件（.docx）', type=['docx'], key='single_source')
    if source:
        values = values_editor(source.getvalue(), 'single')
        single_proposal_leader = st.text_input('立项建议书负责人', value=DEFAULT_PROPOSAL_LEADER, key='single_proposal_leader')
        if tk is not None and st.button('选择单个文件保存位置'):
            selected = choose_save_folder('选择单个生成文件保存位置')
            if selected: st.session_state.single_folder = selected
        st.caption(st.session_state.get('single_folder', '') or '未选择保存位置：生成后可通过浏览器下载。')
        if st.button('生成 DOCX', type='primary'):
            try:
                if file_role(source.name) == 'proposal':
                    record = {'organizations': values.get('drafting_organizations', ''), 'relations': ''}
                    st.session_state.single_result = edit_docx(
                        source.getvalue(), 'proposal', values.get('standard_number', ''), values.get('standard_number', ''),
                        record, values.get('publisher_bottom', DEFAULT_ASSOCIATION),
                        single_proposal_leader.strip() or DEFAULT_PROPOSAL_LEADER,
                    )
                    st.session_state.single_name = Path(source.name).stem + '_已修改.docx'
                else:
                    st.session_state.single_result = merge(template_bytes(), source.getvalue(), values)
                    st.session_state.single_name = Path(source.name).stem + '_封面已修改.docx'
                if st.session_state.get('single_folder'):
                    saved = save_to_folder(st.session_state.single_folder, st.session_state.single_name, st.session_state.single_result)
                    st.success('已保存到：' + saved)
            except Exception as error:
                st.error(str(error))
        if st.session_state.get('single_result'):
            st.download_button('下载 DOCX', st.session_state.single_result, file_name=st.session_state.single_name, mime=MIME_DOCX)

with batch:
    st.subheader('拖放批量文件')
    archive = st.file_uploader('将压缩包直接拖到此处，或点击选择文件', type=['zip', 'rar', '7z', 'tar', 'gz', 'tgz', 'bz2', 'tbz', 'xz', 'txz'], key='batch_source')
    excel = st.file_uploader('将 Excel 对照表直接拖到此处，或点击选择文件（.xlsx/.xls）', type=['xlsx', 'xls'], key='batch_excel')
    left, right = st.columns(2)
    with left:
        batch_publisher_top = st.text_input('第一协会名称（批量统一使用）', value=DEFAULT_PRIMARY_ASSOCIATION, key='batch_publisher_top')
    with right:
        batch_publisher_bottom = st.text_input('归口单位（第二协会、立项建议书及投票单统一使用）', value=DEFAULT_ASSOCIATION, key='batch_publisher_bottom')
    batch_proposal_leader = st.text_input('立项建议书负责人', value=DEFAULT_PROPOSAL_LEADER, key='batch_proposal_leader')
    standard_prefix = st.text_input('标准编号前缀（例如 T/QAS；留空则保留原前缀）', value='', key='batch_standard_prefix')
    st.info('严格识别标准文本后，按其标准编号批量同步封面、页眉、正文及相关子文件。归口单位会同步用于标准文本第二协会、立项建议书的归口单位/第一起草单位，以及投票单标题。Excel 固定读取 B列=标准名称、P列=起草单位、Q列=起草人、R列=起草人与起草单位对应情况。')
    if tk is not None and st.button('选择批量压缩包保存位置'):
        selected = choose_save_folder('选择批量压缩包保存位置')
        if selected: st.session_state.batch_folder = selected
    st.caption(st.session_state.get('batch_folder', '') or '未选择保存位置：生成后可通过浏览器下载。')
    if archive and excel and st.button('生成批量压缩包', type='primary'):
        try:
            result, unchanged, count, _ = batch_generate(archive.getvalue(), archive.name, excel.getvalue(), batch_publisher_top, batch_publisher_bottom, standard_prefix, batch_proposal_leader)
            st.session_state.batch_result = result
            st.session_state.batch_name = archive_stem(archive.name) + '_封面批量已修改' + archive_suffix(archive.name)
            st.session_state.batch_unchanged = unchanged
            st.session_state.batch_count = count
            if st.session_state.get('batch_folder'):
                saved = save_to_folder(st.session_state.batch_folder, st.session_state.batch_name, st.session_state.batch_result)
                st.success('已保存到：' + saved)
            if count:
                st.success(f'已按规则处理 {count} 个 DOCX。')
            else:
                st.warning('未找到可修改的标准文本：请确认目标文件名含“标准文本”，或与母文件夹标准题目完全一致，并具备标准封面特征。')
        except Exception as error:
            st.error(str(error))
    if st.session_state.get('batch_result'):
        st.download_button('下载批量压缩包', st.session_state.batch_result, file_name=st.session_state.batch_name, mime='application/octet-stream')
        unchanged = st.session_state.get('batch_unchanged', [])
        if unchanged:
            st.warning('以下母文件未修改：')
            st.code('\n'.join(unchanged), language=None)
        elif st.session_state.get('batch_result') and st.session_state.get('batch_count', 0):
            st.success('所有压缩包内母文件均已按规则处理。')

st.caption('第二页起正文保持不变；前言仅覆盖指定段落。')
