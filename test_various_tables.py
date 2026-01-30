# -*- coding: utf-8 -*-
"""
다양한 테이블 구조 테스트

테스트 케이스:
1. 기본 중첩 rowspan (Head1 4행, Head2 2행, Head3 2행)
2. 3단계 중첩 rowspan (Head1 > Head2 > Head3 계층)
3. colspan + rowspan 혼합
4. 단일 헤더 열 (헤더 1개만)
5. 복잡한 불규칙 구조
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from io import BytesIO
import sys
import os

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent / "merge"))

from merge.table_merger import (
    TableMerger, TableParser, HeaderConfig,
    print_table_structure, analyze_rowspan_structure
)


# 네임스페이스
NS_SECTION = "http://www.hancom.co.kr/hwpml/2011/section"
NS_PARA = "http://www.hancom.co.kr/hwpml/2011/paragraph"
NS_CORE = "http://www.hancom.co.kr/hwpml/2011/core"

ET.register_namespace('hs', NS_SECTION)
ET.register_namespace('hp', NS_PARA)
ET.register_namespace('hc', NS_CORE)


def make_tag(ns, name):
    return f"{{{ns}}}{name}"


def create_cell(row, col, text, rowspan=1, colspan=1, field_name=None):
    """셀 요소 생성"""
    tc = ET.Element(make_tag(NS_PARA, 'tc'))

    addr = ET.SubElement(tc, make_tag(NS_PARA, 'cellAddr'))
    addr.set('colAddr', str(col))
    addr.set('rowAddr', str(row))

    span = ET.SubElement(tc, make_tag(NS_PARA, 'cellSpan'))
    span.set('colSpan', str(colspan))
    span.set('rowSpan', str(rowspan))

    sz = ET.SubElement(tc, make_tag(NS_PARA, 'cellSz'))
    sz.set('width', '5000')
    sz.set('height', '1000')

    sublist = ET.SubElement(tc, make_tag(NS_PARA, 'subList'))
    p = ET.SubElement(sublist, make_tag(NS_PARA, 'p'))
    run = ET.SubElement(p, make_tag(NS_PARA, 'run'))

    if field_name:
        field = ET.SubElement(run, make_tag(NS_PARA, 'fieldBegin'))
        field.set('name', field_name)

    t = ET.SubElement(run, make_tag(NS_PARA, 't'))
    t.text = text

    return tc


def create_table_case1():
    """
    케이스 1: 기본 중첩 rowspan
    +-------+-------+-------+-------+
    | Head1 | Head2 | Col1  | Col2  |
    | (4행) | (2행) +-------+-------+
    |       |       | A     | B     |
    |       +-------+-------+-------+
    |       | Head3 | C     | D     |
    |       | (2행) +-------+-------+
    |       |       | E     | F     |
    +-------+-------+-------+-------+
    """
    tbl = ET.Element(make_tag(NS_PARA, 'tbl'))
    tbl.set('id', 'case1_basic_nested')

    tr0 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr0.append(create_cell(0, 0, 'Head1', rowspan=4))
    tr0.append(create_cell(0, 1, 'Head2', rowspan=2))
    tr0.append(create_cell(0, 2, 'Col1', field_name='Col1'))
    tr0.append(create_cell(0, 3, 'Col2', field_name='Col2'))

    tr1 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr1.append(create_cell(1, 2, 'A'))
    tr1.append(create_cell(1, 3, 'B'))

    tr2 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr2.append(create_cell(2, 1, 'Head3', rowspan=2))
    tr2.append(create_cell(2, 2, 'C'))
    tr2.append(create_cell(2, 3, 'D'))

    tr3 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr3.append(create_cell(3, 2, 'E'))
    tr3.append(create_cell(3, 3, 'F'))

    return tbl


def create_table_case6():
    """
    케이스 6: 빈 셀 있는 테이블 (fill_empty_first 테스트)
    +-------+-------+-------+-------+
    | Head1 | Head2 | Col1  | Col2  |
    | (6행) | (2행) +-------+-------+
    |       |       |       |       |  ← 빈 셀
    |       +-------+-------+-------+
    |       | Head3 |       |       |  ← 빈 셀
    |       | (4행) +-------+-------+
    |       |       |       |       |  ← 빈 셀
    |       |       +-------+-------+
    |       |       |       |       |  ← 빈 셀
    +-------+-------+-------+-------+
    """
    tbl = ET.Element(make_tag(NS_PARA, 'tbl'))
    tbl.set('id', 'case6_empty_cells')

    tr0 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr0.append(create_cell(0, 0, 'Head1', rowspan=6))
    tr0.append(create_cell(0, 1, 'Head2', rowspan=2))
    tr0.append(create_cell(0, 2, 'Col1', field_name='Col1'))
    tr0.append(create_cell(0, 3, 'Col2', field_name='Col2'))

    tr1 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr1.append(create_cell(1, 2, ''))  # 빈 셀
    tr1.append(create_cell(1, 3, ''))  # 빈 셀

    tr2 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr2.append(create_cell(2, 1, 'Head3', rowspan=4))
    tr2.append(create_cell(2, 2, ''))  # 빈 셀
    tr2.append(create_cell(2, 3, ''))  # 빈 셀

    tr3 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr3.append(create_cell(3, 2, ''))  # 빈 셀
    tr3.append(create_cell(3, 3, ''))  # 빈 셀

    tr4 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr4.append(create_cell(4, 2, ''))  # 빈 셀
    tr4.append(create_cell(4, 3, ''))  # 빈 셀

    tr5 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr5.append(create_cell(5, 2, ''))  # 빈 셀
    tr5.append(create_cell(5, 3, ''))  # 빈 셀

    return tbl


def create_table_case7():
    """
    케이스 7: add_rows_smart 테스트 (header_, data_ 접두사)
    +-------+---------------+-------+-------+
    |       | header_group  | data_ | data_ |
    | (4행) | (2행)         | col1  | col2  |
    |       |               +-------+-------+
    |       |               | A     | B     |
    |       +---------------+-------+-------+
    |       | GroupB        | C     | D     |
    |       | (2행)         +-------+-------+
    |       |               | E     | F     |
    +-------+---------------+-------+-------+

    필드명:
    - Col 0: 없음 (항상 확장)
    - Col 1: "header_group" (헤더 병합)
    - Col 2: "data_col1" (데이터)
    - Col 3: "data_col2" (데이터)
    """
    tbl = ET.Element(make_tag(NS_PARA, 'tbl'))
    tbl.set('id', 'case7_smart_header')

    tr0 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr0.append(create_cell(0, 0, 'Main', rowspan=4))
    tr0.append(create_cell(0, 1, 'GroupA', rowspan=2, field_name='header_group'))  # header_ 접두사
    tr0.append(create_cell(0, 2, 'data_col1', field_name='data_col1'))  # data_ 접두사
    tr0.append(create_cell(0, 3, 'data_col2', field_name='data_col2'))  # data_ 접두사

    tr1 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr1.append(create_cell(1, 2, 'A'))
    tr1.append(create_cell(1, 3, 'B'))

    tr2 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr2.append(create_cell(2, 1, 'GroupB', rowspan=2))
    tr2.append(create_cell(2, 2, 'C'))
    tr2.append(create_cell(2, 3, 'D'))

    tr3 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr3.append(create_cell(3, 2, 'E'))
    tr3.append(create_cell(3, 3, 'F'))

    return tbl


def create_table_case2():
    """
    케이스 2: 3단계 중첩 rowspan
    +-------+-------+-------+-------+-------+
    | H1    | H2    | H3    | Col1  | Col2  |
    | (6행) | (4행) | (2행) +-------+-------+
    |       |       |       | A     | B     |
    |       |       +-------+-------+-------+
    |       |       | H4    | C     | D     |
    |       |       | (2행) +-------+-------+
    |       |       |       | E     | F     |
    |       +-------+-------+-------+-------+
    |       | H5    | H6    | G     | H     |
    |       | (2행) | (2행) +-------+-------+
    |       |       |       | I     | J     |
    +-------+-------+-------+-------+-------+
    """
    tbl = ET.Element(make_tag(NS_PARA, 'tbl'))
    tbl.set('id', 'case2_three_level')

    tr0 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr0.append(create_cell(0, 0, 'H1', rowspan=6))
    tr0.append(create_cell(0, 1, 'H2', rowspan=4))
    tr0.append(create_cell(0, 2, 'H3', rowspan=2))
    tr0.append(create_cell(0, 3, 'Col1', field_name='Col1'))
    tr0.append(create_cell(0, 4, 'Col2', field_name='Col2'))

    tr1 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr1.append(create_cell(1, 3, 'A'))
    tr1.append(create_cell(1, 4, 'B'))

    tr2 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr2.append(create_cell(2, 2, 'H4', rowspan=2))
    tr2.append(create_cell(2, 3, 'C'))
    tr2.append(create_cell(2, 4, 'D'))

    tr3 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr3.append(create_cell(3, 3, 'E'))
    tr3.append(create_cell(3, 4, 'F'))

    tr4 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr4.append(create_cell(4, 1, 'H5', rowspan=2))
    tr4.append(create_cell(4, 2, 'H6', rowspan=2))
    tr4.append(create_cell(4, 3, 'G'))
    tr4.append(create_cell(4, 4, 'H'))

    tr5 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr5.append(create_cell(5, 3, 'I'))
    tr5.append(create_cell(5, 4, 'J'))

    return tbl


def create_table_case3():
    """
    케이스 3: colspan + rowspan 혼합
    +---------------+-------+-------+
    |    Header     | Col1  | Col2  |
    |   (colspan=2) +-------+-------+
    |   (rowspan=2) | A     | B     |
    +-------+-------+-------+-------+
    | Sub1  | Sub2  | C     | D     |
    | (2행) | (2행) +-------+-------+
    |       |       | E     | F     |
    +-------+-------+-------+-------+
    """
    tbl = ET.Element(make_tag(NS_PARA, 'tbl'))
    tbl.set('id', 'case3_colspan_rowspan')

    tr0 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr0.append(create_cell(0, 0, 'Header', rowspan=2, colspan=2))
    tr0.append(create_cell(0, 2, 'Col1', field_name='Col1'))
    tr0.append(create_cell(0, 3, 'Col2', field_name='Col2'))

    tr1 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr1.append(create_cell(1, 2, 'A'))
    tr1.append(create_cell(1, 3, 'B'))

    tr2 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr2.append(create_cell(2, 0, 'Sub1', rowspan=2))
    tr2.append(create_cell(2, 1, 'Sub2', rowspan=2))
    tr2.append(create_cell(2, 2, 'C'))
    tr2.append(create_cell(2, 3, 'D'))

    tr3 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr3.append(create_cell(3, 2, 'E'))
    tr3.append(create_cell(3, 3, 'F'))

    return tbl


def create_table_case4():
    """
    케이스 4: 단일 헤더 열
    +-------+-------+-------+
    | Head  | Col1  | Col2  |
    | (4행) +-------+-------+
    |       | A     | B     |
    |       +-------+-------+
    |       | C     | D     |
    |       +-------+-------+
    |       | E     | F     |
    +-------+-------+-------+
    """
    tbl = ET.Element(make_tag(NS_PARA, 'tbl'))
    tbl.set('id', 'case4_single_header')

    tr0 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr0.append(create_cell(0, 0, 'Head', rowspan=4))
    tr0.append(create_cell(0, 1, 'Col1', field_name='Col1'))
    tr0.append(create_cell(0, 2, 'Col2', field_name='Col2'))

    tr1 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr1.append(create_cell(1, 1, 'A'))
    tr1.append(create_cell(1, 2, 'B'))

    tr2 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr2.append(create_cell(2, 1, 'C'))
    tr2.append(create_cell(2, 2, 'D'))

    tr3 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr3.append(create_cell(3, 1, 'E'))
    tr3.append(create_cell(3, 2, 'F'))

    return tbl


def create_table_case5():
    """
    케이스 5: 불규칙 rowspan
    +-------+-------+-------+-------+
    | H1    | H2    | Col1  | Col2  |
    | (3행) | (2행) +-------+-------+
    |       |       | A     | B     |
    |       +-------+-------+-------+
    |       | H3    | C     | D     |
    +-------+-------+-------+-------+
    | H4    | H5    | E     | F     |
    | (2행) | (1행) +-------+-------+
    |       +-------+-------+-------+
    |       | H6    | G     | H     |
    +-------+-------+-------+-------+
    """
    tbl = ET.Element(make_tag(NS_PARA, 'tbl'))
    tbl.set('id', 'case5_irregular')

    tr0 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr0.append(create_cell(0, 0, 'H1', rowspan=3))
    tr0.append(create_cell(0, 1, 'H2', rowspan=2))
    tr0.append(create_cell(0, 2, 'Col1', field_name='Col1'))
    tr0.append(create_cell(0, 3, 'Col2', field_name='Col2'))

    tr1 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr1.append(create_cell(1, 2, 'A'))
    tr1.append(create_cell(1, 3, 'B'))

    tr2 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr2.append(create_cell(2, 1, 'H3'))
    tr2.append(create_cell(2, 2, 'C'))
    tr2.append(create_cell(2, 3, 'D'))

    tr3 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr3.append(create_cell(3, 0, 'H4', rowspan=2))
    tr3.append(create_cell(3, 1, 'H5'))
    tr3.append(create_cell(3, 2, 'E'))
    tr3.append(create_cell(3, 3, 'F'))

    tr4 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr4.append(create_cell(4, 1, 'H6'))
    tr4.append(create_cell(4, 2, 'G'))
    tr4.append(create_cell(4, 3, 'H'))

    return tbl


def create_hwpx_with_tables(output_path: str, tables: list):
    """여러 테이블을 포함한 HWPX 생성"""
    section = ET.Element(make_tag(NS_SECTION, 'sec'))

    for tbl in tables:
        section.append(tbl)
        # 테이블 사이에 빈 문단 추가
        p = ET.SubElement(section, make_tag(NS_PARA, 'p'))

    section_xml = ET.tostring(section, encoding='UTF-8', xml_declaration=True)

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr('mimetype', 'application/hwp+zip', compress_type=zipfile.ZIP_STORED)
        zf.writestr('Contents/section0.xml', section_xml)
        zf.writestr('META-INF/manifest.xml', '<?xml version="1.0"?><manifest/>')
        zf.writestr('version.xml', '<?xml version="1.0"?><version/>')
        zf.writestr('Contents/content.hpf', '<?xml version="1.0"?><content/>')

    return output_path


def test_case1(hwpx_path: str):
    """케이스 1 테스트: 기본 중첩 rowspan"""
    print("\n" + "="*70)
    print("[케이스 1] 기본 중첩 rowspan")
    print("="*70)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 0)

    print("\n[원본]")
    print_table_structure(table)

    data_list = [
        {"_header": "Head3", "Col1": "G", "Col2": "H"},
        {"_header": "Head3", "Col1": "I", "Col2": "J"},
        {"_header": "NewH", "Col1": "K", "Col2": "L"},
        {"_header": "NewH", "Col1": "M", "Col2": "N"},
    ]

    print("\n[추가 데이터]")
    for d in data_list:
        print(f"  {d}")

    merger.add_rows_auto(data_list, header_col=1, data_cols=[2, 3], extend_cols=[0])

    print("\n[결과]")
    print_table_structure(table)

    return merger


def test_case2(hwpx_path: str):
    """케이스 2 테스트: 3단계 중첩"""
    print("\n" + "="*70)
    print("[케이스 2] 3단계 중첩 rowspan")
    print("="*70)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 1)

    print("\n[원본]")
    print_table_structure(table)

    data_list = [
        {"_header": "H6", "Col1": "K", "Col2": "L"},
        {"_header": "H7", "Col1": "M", "Col2": "N"},
        {"_header": "H7", "Col1": "O", "Col2": "P"},
    ]

    print("\n[추가 데이터]")
    for d in data_list:
        print(f"  {d}")

    merger.add_rows_auto(data_list, header_col=2, data_cols=[3, 4], extend_cols=[0, 1])

    print("\n[결과]")
    print_table_structure(table)

    return merger


def test_case3(hwpx_path: str):
    """케이스 3 테스트: colspan + rowspan 혼합"""
    print("\n" + "="*70)
    print("[케이스 3] colspan + rowspan 혼합")
    print("="*70)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 2)

    print("\n[원본]")
    print_table_structure(table)

    data_list = [
        {"_header": "Sub2", "Col1": "G", "Col2": "H"},
        {"_header": "Sub3", "Col1": "I", "Col2": "J"},
        {"_header": "Sub3", "Col1": "K", "Col2": "L"},
    ]

    print("\n[추가 데이터]")
    for d in data_list:
        print(f"  {d}")

    merger.add_rows_auto(data_list, header_col=1, data_cols=[2, 3], extend_cols=[0])

    print("\n[결과]")
    print_table_structure(table)

    return merger


def test_case4(hwpx_path: str):
    """케이스 4 테스트: 단일 헤더 열"""
    print("\n" + "="*70)
    print("[케이스 4] 단일 헤더 열 (헤더 변경 없이 확장)")
    print("="*70)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 3)

    print("\n[원본]")
    print_table_structure(table)

    # 단일 헤더 열이므로 항상 확장만
    data_list = [
        {"_header": "Head", "Col1": "G", "Col2": "H"},
        {"_header": "Head", "Col1": "I", "Col2": "J"},
        {"_header": "Head", "Col1": "K", "Col2": "L"},
    ]

    print("\n[추가 데이터]")
    for d in data_list:
        print(f"  {d}")

    merger.add_rows_auto(data_list, header_col=0, data_cols=[1, 2], extend_cols=[])

    print("\n[결과]")
    print_table_structure(table)

    return merger


def test_case5(hwpx_path: str):
    """케이스 5 테스트: 불규칙 rowspan"""
    print("\n" + "="*70)
    print("[케이스 5] 불규칙 rowspan")
    print("="*70)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 4)

    print("\n[원본]")
    print_table_structure(table)

    data_list = [
        {"_header": "H6", "Col1": "I", "Col2": "J"},
        {"_header": "H7", "Col1": "K", "Col2": "L"},
        {"_header": "H7", "Col1": "M", "Col2": "N"},
    ]

    print("\n[추가 데이터]")
    for d in data_list:
        print(f"  {d}")

    merger.add_rows_auto(data_list, header_col=1, data_cols=[2, 3], extend_cols=[0])

    print("\n[결과]")
    print_table_structure(table)

    return merger


def test_case6(hwpx_path: str):
    """케이스 6 테스트: 빈 셀 먼저 채우기"""
    print("\n" + "="*70)
    print("[케이스 6] 빈 셀 먼저 채우기 (fill_empty_first=True)")
    print("="*70)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 5)

    print("\n[원본] - 빈 셀 5개")
    print_table_structure(table)

    # 5개 빈 셀 + 2개 추가 = 총 7개 데이터
    data_list = [
        {"_header": "Head2", "Col1": "A", "Col2": "B"},   # Head2 빈 셀 채움
        {"_header": "Head3", "Col1": "C", "Col2": "D"},   # Head3 첫 번째 빈 셀
        {"_header": "Head3", "Col1": "E", "Col2": "F"},   # Head3 두 번째 빈 셀
        {"_header": "Head3", "Col1": "G", "Col2": "H"},   # Head3 세 번째 빈 셀
        {"_header": "Head3", "Col1": "I", "Col2": "J"},   # Head3 네 번째 빈 셀
        {"_header": "Head3", "Col1": "K", "Col2": "L"},   # 빈 셀 없음 → 새 행 추가
        {"_header": "Head4", "Col1": "M", "Col2": "N"},   # Head4 새로 생성
    ]

    print("\n[추가 데이터] (7개: 빈 셀 5개 + 새 행 2개)")
    for i, d in enumerate(data_list):
        print(f"  {i+1}. {d}")

    merger.add_rows_auto(
        data_list,
        header_col=1,
        data_cols=[2, 3],
        extend_cols=[0],
        fill_empty_first=True
    )

    print("\n[결과] - 빈 셀 채우고 새 행 추가")
    print_table_structure(table)

    return merger


def test_case7(hwpx_path: str):
    """케이스 7 테스트: add_rows_smart (header_, data_ 접두사)"""
    print("\n" + "="*70)
    print("[케이스 7] add_rows_smart - header_, data_ 접두사 자동 분석")
    print("="*70)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 6)

    print("\n[원본]")
    print_table_structure(table)

    # 필드명 정보 출력
    print("\n[필드명 분석]")
    for (row, col), cell in sorted(table.cells.items()):
        if cell.field_name:
            prefix = ""
            if cell.field_name.startswith('header_'):
                prefix = " → 헤더 병합"
            elif cell.field_name.startswith('data_'):
                prefix = " → 데이터"
            print(f"  Col {col}: '{cell.field_name}'{prefix}")

    # 데이터 - 필드명을 키로 사용 (header_group, data_col1, data_col2)
    data_list = [
        {"header_group": "GroupB", "data_col1": "G", "data_col2": "H"},   # GroupB 확장
        {"header_group": "GroupB", "data_col1": "I", "data_col2": "J"},   # GroupB 계속 확장
        {"header_group": "GroupC", "data_col1": "K", "data_col2": "L"},   # GroupC 새로 생성
        {"header_group": "GroupC", "data_col1": "M", "data_col2": "N"},   # GroupC 확장
    ]

    print("\n[추가 데이터] - 필드명을 키로 사용")
    for i, d in enumerate(data_list):
        print(f"  {i+1}. {d}")

    # add_rows_smart 호출 (자동으로 header_col, data_cols 분석)
    merger.add_rows_smart(data_list)

    print("\n[결과]")
    print_table_structure(table)

    return merger


def run_all_tests():
    """모든 테스트 실행"""
    os.makedirs("data", exist_ok=True)

    # 테스트 파일 생성
    tables = [
        create_table_case1(),
        create_table_case2(),
        create_table_case3(),
        create_table_case4(),
        create_table_case5(),
        create_table_case6(),
        create_table_case7(),
    ]

    input_path = "data/test_various_tables.hwpx"
    create_hwpx_with_tables(input_path, tables)
    print(f"테스트 파일 생성: {input_path}")

    print("\n" + "#"*70)
    print("# 다양한 테이블 구조 테스트")
    print("#"*70)

    # 각 케이스 테스트
    mergers = []
    mergers.append(test_case1(input_path))
    mergers.append(test_case2(input_path))
    mergers.append(test_case3(input_path))
    mergers.append(test_case4(input_path))
    mergers.append(test_case5(input_path))
    mergers.append(test_case6(input_path))
    mergers.append(test_case7(input_path))

    # 결과 저장
    print("\n" + "#"*70)
    print("# 결과 파일 저장")
    print("#"*70)

    for i, merger in enumerate(mergers, 1):
        output_path = f"data/test_case{i}_result.hwpx"
        merger.save(output_path)
        print(f"  케이스 {i}: {output_path}")

    print("\n모든 테스트 완료!")


if __name__ == "__main__":
    run_all_tests()
