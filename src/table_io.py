#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
通用表格文件读取器
支持格式：
- Excel 工作簿 (.xlsx / .xlsm)      → openpyxl
- Excel 97-2003 (.xls)              → xlrd
- CSV / TSV / TXT（多编码自适应）    → pandas
- 伪Excel文件（内容为CSV/TSV/HTML表格
  但扩展名为 .xls/.xlsx，常见于各
  类业务系统导出）                   → 自动识别
"""
import io
import csv
import xml.etree.ElementTree as ET

import pandas as pd

SUPPORTED_FILE_FILTER = (
    "所有支持的格式 (*.xlsx *.xlsm *.xls *.csv *.tsv *.txt);;"
    "Excel 工作簿 (*.xlsx *.xlsm);;"
    "Excel 97-2003 (*.xls);;"
    "CSV/文本文件 (*.csv *.tsv *.txt);;"
    "所有文件 (*)"
)

_TEXT_ENCODINGS = ('utf-8-sig', 'utf-16', 'gbk', 'big5', 'latin1')
_DELIMITERS = (',', '\t', ';', '|',)


def _sniff_delimiter(sample: str) -> str:
    """从文本样本中嗅探分隔符，失败时选产生列数最多的候选"""
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=''.join(_DELIMITERS))
        return dialect.delimiter
    except Exception:
        pass
    best, best_cols = ',', 0
    first_line = sample.splitlines()[0] if sample.strip() else ''
    for d in _DELIMITERS:
        cols = len(first_line.split(d))
        if cols > best_cols:
            best, best_cols = d, cols
    return best


def _decode_text(raw: bytes):
    """多编码尝试解码文本文件"""
    for enc in _TEXT_ENCODINGS:
        try:
            return raw.decode(enc), enc
        except (UnicodeDecodeError, UnicodeError):
            continue
    return raw.decode('utf-8', errors='replace'), 'utf-8(replace)'


def _read_spreadsheet_ml(raw: bytes) -> pd.DataFrame:
    """解析 Excel 2003 XML (SpreadsheetML) 格式"""
    SS_NS = '{urn:schemas-microsoft-com:office:spreadsheet}'
    root = ET.fromstring(raw)
    worksheet = root.find(f'{SS_NS}Worksheet')
    if worksheet is None:
        raise ValueError('文件中不包含Excel工作表')
    table = worksheet.find(f'{SS_NS}Table')
    if table is None:
        raise ValueError('文件中不包含数据表')

    rows = []
    for row_el in table.findall(f'{SS_NS}Row'):
        values, expect_idx = [], 1
        for cell_el in row_el.findall(f'{SS_NS}Cell'):
            idx_attr = cell_el.get(f'{SS_NS}Index')
            if idx_attr is not None:
                idx = int(idx_attr)
                while expect_idx < idx:
                    values.append('')
                    expect_idx += 1
            data_el = cell_el.find(f'{SS_NS}Data')
            values.append(data_el.text if data_el is not None and
                          data_el.text is not None else '')
            expect_idx += 1
        rows.append(values)

    if not rows:
        return pd.DataFrame()
    header = [str(h).strip() for h in rows[0]]
    # 补齐缺失表头
    for i, h in enumerate(header):
        if not h:
            header[i] = f'列{i + 1}'
    data_rows = rows[1:]
    width = len(header)
    padded = [(r + [''] * width)[:width] for r in data_rows]
    return pd.DataFrame(padded, columns=header)


def _normalize(df: pd.DataFrame) -> pd.DataFrame:
    """统一表格形态：列名去空白、单元格转字符串、剔除全空行与全空列"""
    if df is None or df.empty:
        return pd.DataFrame()
    df.columns = [str(c).strip() for c in df.columns]
    # 单元格统一转字符串，避免 Excel(数值型) 与 CSV(字符型) 行为不一致
    df = df.astype(object).where(pd.notna(df), '')
    _to_str = lambda v: '' if v is None else str(v).strip()
    # pandas 2.1+ 用 DataFrame.map，旧版本回退 applymap
    df = df.map(_to_str) if hasattr(df, 'map') else df.applymap(_to_str)
    # 剔除全空行 / 全空列
    df = df.dropna(axis=1, how='all')
    df = df[~(df == '').all(axis=1)]
    return df.reset_index(drop=True)


def read_table_any(path: str) -> pd.DataFrame:
    """智能读取表格文件，返回 DataFrame；无法识别时抛出带中文说明的异常"""
    with open(path, 'rb') as f:
        raw = f.read()

    if not raw.strip():
        raise ValueError('文件为空')

    # ---- 1. 按文件魔数识别 ----
    if raw[:4] == b'PK\x03\x04':                      # zip容器 → xlsx/xlsm
        return _normalize(pd.read_excel(io.BytesIO(raw), engine='openpyxl'))
    if raw[:8] == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1':  # OLE容器 → xls
        return _normalize(pd.read_excel(io.BytesIO(raw), engine='xlrd'))

    # ---- 2. 文本类内容（CSV/TSV/HTML/XML，含各种伪Excel）----
    text, enc = _decode_text(raw)
    head = text.lstrip('\ufeff ').lstrip()[:512].lower()

    if head.startswith('<!doctype html') or head.startswith('<html') \
            or '<table' in head:
        tables = pd.read_html(io.StringIO(text))
        if not tables:
            raise ValueError('HTML文件中未找到数据表格')
        # 取最大的表（通常业务数据表最大）
        return _normalize(max(tables, key=lambda t: t.shape[0] * t.shape[1]))

    if head.startswith('<?xml'):
        if 'spreadsheet' in head or 'workbook' in head \
                or 'urn:schemas-microsoft-com:office:spreadsheet' in text[:2048].lower():
            return _normalize(_read_spreadsheet_ml(raw))
        raise ValueError('无法识别的XML格式，请另存为 .xlsx 或 .csv 后重试')

    # 纯文本表格：嗅探分隔符
    sample = text[:64 * 1024]
    sep = _sniff_delimiter(sample)
    df = pd.read_csv(io.StringIO(text), sep=sep, dtype=str,
                     encoding=enc, keep_default_na=False,
                     engine='python')
    # 全空列清理；仅剩一列时可能是分隔符识别错误，尝试其它分隔符
    if df.shape[1] <= 1:
        for alt in _DELIMITERS:
            if alt == sep:
                continue
            df2 = pd.read_csv(io.StringIO(text), sep=alt, dtype=str,
                              encoding=enc, keep_default_na=False,
                              engine='python')
            if df2.shape[1] > df.shape[1]:
                df = df2
    return _normalize(df)
