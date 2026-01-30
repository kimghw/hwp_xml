# -*- coding: utf-8 -*-
"""
중첩 rowspan 테스트

테스트 테이블 구조:
+---------------+---------------+-------+-------+
|               |               | Col1  | Col2  |
|    Head1      |    Head2      +-------+-------+
|   (4행)       |   (2행)       | A     | B     |
|               +---------------+-------+-------+
|               |    Head3      | C     | D     |
|               |   (2행)       +-------+-------+
|               |               | E     | F     |
+---------------+---------------+-------+-------+

Row 0: Head1(4x1), Head2(2x1), Col1(1x1), Col2(1x1)
Row 1: ↓,          ↓,          A,         B
Row 2: ↓,          Head3(2x1), C,         D
Row 3: ↓,          ↓,          E,         F
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from io import BytesIO
import sys
import os

# 경로 설정
sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).parent / "merge"))

from merge.table_merger import (
    TableMerger, TableParser, HeaderConfig,
    create_header_config_from_table, print_table_structure,
    analyze_rowspan_structure, get_row_template_info
)


def create_test_hwpx_with_nested_rowspan(output_path: str):
    """
    중첩 rowspan을 가진 테스트 HWPX 파일 생성

    구조:
    +---------------+---------------+-------+-------+
    |               |               | Col1  | Col2  |
    |    Head1      |    Head2      +-------+-------+
    |   (4행)       |   (2행)       | A     | B     |
    |               +---------------+-------+-------+
    |               |    Head3      | C     | D     |
    |               |   (2행)       +-------+-------+
    |               |               | E     | F     |
    +---------------+---------------+-------+-------+
    """
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

        # cellAddr
        addr = ET.SubElement(tc, make_tag(NS_PARA, 'cellAddr'))
        addr.set('colAddr', str(col))
        addr.set('rowAddr', str(row))

        # cellSpan
        span = ET.SubElement(tc, make_tag(NS_PARA, 'cellSpan'))
        span.set('colSpan', str(colspan))
        span.set('rowSpan', str(rowspan))

        # cellSz
        sz = ET.SubElement(tc, make_tag(NS_PARA, 'cellSz'))
        sz.set('width', '5000')
        sz.set('height', '1000')

        # subList > p > run > t
        sublist = ET.SubElement(tc, make_tag(NS_PARA, 'subList'))
        p = ET.SubElement(sublist, make_tag(NS_PARA, 'p'))
        run = ET.SubElement(p, make_tag(NS_PARA, 'run'))

        # 필드명 추가
        if field_name:
            field = ET.SubElement(run, make_tag(NS_PARA, 'fieldBegin'))
            field.set('name', field_name)

        t = ET.SubElement(run, make_tag(NS_PARA, 't'))
        t.text = text

        return tc

    # section 루트
    section = ET.Element(make_tag(NS_SECTION, 'sec'))

    # 테이블 생성
    tbl = ET.SubElement(section, make_tag(NS_PARA, 'tbl'))
    tbl.set('id', 'test_nested_rowspan')

    # Row 0: Head1(4x1), Head2(2x1), Col1, Col2
    tr0 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr0.append(create_cell(0, 0, 'Head1', rowspan=4))
    tr0.append(create_cell(0, 1, 'Head2', rowspan=2))
    tr0.append(create_cell(0, 2, 'Col1', field_name='Col1'))
    tr0.append(create_cell(0, 3, 'Col2', field_name='Col2'))

    # Row 1: (Head1 확장), (Head2 확장), A, B
    tr1 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr1.append(create_cell(1, 2, 'A'))
    tr1.append(create_cell(1, 3, 'B'))

    # Row 2: (Head1 확장), Head3(2x1), C, D
    tr2 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr2.append(create_cell(2, 1, 'Head3', rowspan=2))
    tr2.append(create_cell(2, 2, 'C'))
    tr2.append(create_cell(2, 3, 'D'))

    # Row 3: (Head1 확장), (Head3 확장), E, F
    tr3 = ET.SubElement(tbl, make_tag(NS_PARA, 'tr'))
    tr3.append(create_cell(3, 2, 'E'))
    tr3.append(create_cell(3, 3, 'F'))

    # XML 문자열 생성
    section_xml = ET.tostring(section, encoding='UTF-8', xml_declaration=True)

    # HWPX 파일 생성
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # mimetype (압축 안함)
        zf.writestr('mimetype', 'application/hwp+zip', compress_type=zipfile.ZIP_STORED)

        # section0.xml
        zf.writestr('Contents/section0.xml', section_xml)

        # 필수 메타데이터 (빈 파일로 대체)
        zf.writestr('META-INF/manifest.xml', '<?xml version="1.0"?><manifest/>')
        zf.writestr('version.xml', '<?xml version="1.0"?><version/>')
        zf.writestr('Contents/content.hpf', '<?xml version="1.0"?><content/>')

    print(f"테스트 파일 생성: {output_path}")
    return output_path


def test_parse_nested_rowspan(hwpx_path: str):
    """중첩 rowspan 파싱 테스트"""
    print("\n" + "="*60)
    print("[1] 테이블 파싱 테스트")
    print("="*60)

    parser = TableParser()
    tables = parser.parse_tables(hwpx_path)

    print(f"테이블 개수: {len(tables)}")

    if tables:
        table = tables[0]
        print(f"\n테이블 크기: {table.row_count}행 x {table.col_count}열")
        print(f"셀 개수: {len(table.cells)}")

        print("\n[셀 목록]")
        for (row, col), cell in sorted(table.cells.items()):
            span = f"(rowspan={cell.row_span}, colspan={cell.col_span})" if cell.row_span > 1 or cell.col_span > 1 else ""
            field = f" [field:{cell.field_name}]" if cell.field_name else ""
            print(f"  ({row}, {col}): '{cell.text}' {span}{field}")

        # rowspan 구조 분석
        analyze_rowspan_structure(table)

        return table
    return None


def test_add_row_basic(hwpx_path: str):
    """기본 행 추가 테스트"""
    print("\n" + "="*60)
    print("[2] 기본 행 추가 테스트")
    print("="*60)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 0)

    print("\n[원본 테이블]")
    print_table_structure(table)

    # 마지막 행 분석
    print("\n[마지막 행 분석]")
    template = get_row_template_info(table)
    for cell_info in template['cells']:
        status = cell_info['status']
        col = cell_info['col']
        text = cell_info['text'] or "(empty)"
        print(f"  Col {col}: {status}, text='{text}'")

    # 데이터 추가
    print("\n[데이터 추가 시도]")
    data = {"Col1": "NEW1", "Col2": "NEW2"}
    print(f"  데이터: {data}")

    initial_row_count = table.row_count
    merger._add_row_with_data(data)

    print(f"\n[결과]")
    print(f"  행 수: {initial_row_count} -> {table.row_count}")

    # 추가된 행 확인
    new_row = table.row_count - 1
    print(f"\n[새 행 (Row {new_row}) 셀 정보]")
    for col in range(table.col_count):
        cell = table.get_cell(new_row, col)
        if cell:
            if cell.row == new_row:
                print(f"  Col {col}: '{cell.text}' (row={cell.row}, rowspan={cell.row_span})")
            else:
                print(f"  Col {col}: ↓ (rowspan from Row {cell.row})")
        else:
            print(f"  Col {col}: - (없음)")

    return merger


def test_add_row_with_new_header(hwpx_path: str):
    """새 헤더와 함께 행 추가 테스트"""
    print("\n" + "="*60)
    print("[3] 새 헤더와 함께 행 추가 테스트 (add_row_with_headers)")
    print("="*60)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 0)

    print("\n[원본 테이블]")
    print_table_structure(table)

    # HeaderConfig 수동 설정
    print("\n[HeaderConfig 설정]")
    config = [
        HeaderConfig(col=0, action='extend'),                # Head1 확장
        HeaderConfig(col=1, action='new', text='Head4', rowspan=2),  # 새 헤더
        HeaderConfig(col=2, action='data'),                  # 데이터 열
        HeaderConfig(col=3, action='data'),                  # 데이터 열
    ]
    for hc in config:
        print(f"  Col {hc.col}: action={hc.action}, text='{hc.text}', rowspan={hc.rowspan}")

    # 첫 번째 행 추가
    print("\n[첫 번째 행 추가]")
    data1 = {"Col1": "X1", "Col2": "Y1"}
    print(f"  데이터: {data1}")

    initial_row_count = table.row_count
    merger.add_row_with_headers(data1, config)

    print(f"\n[결과]")
    print(f"  행 수: {initial_row_count} -> {table.row_count}")

    # 두 번째 행 추가 (Head4 rowspan 아래)
    print("\n[두 번째 행 추가 (Head4가 커버)]")
    config2 = [
        HeaderConfig(col=0, action='extend'),   # Head1 계속 확장
        HeaderConfig(col=1, action='extend'),   # Head4 커버
        HeaderConfig(col=2, action='data'),
        HeaderConfig(col=3, action='data'),
    ]
    data2 = {"Col1": "X2", "Col2": "Y2"}
    print(f"  데이터: {data2}")

    merger.add_row_with_headers(data2, config2)
    print(f"  행 수: {table.row_count}")

    # 결과 확인
    print("\n[최종 테이블]")
    print_table_structure(table)

    return merger


def test_auto_header_config(hwpx_path: str):
    """자동 HeaderConfig 생성 테스트"""
    print("\n" + "="*60)
    print("[4] 자동 HeaderConfig 생성 테스트 (create_header_config_from_table)")
    print("="*60)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 0)

    print("\n[원본 테이블]")
    print_table_structure(table)

    # 자동 설정 생성
    print("\n[자동 생성된 HeaderConfig]")
    config = create_header_config_from_table(
        table,
        data_cols=[2, 3],                    # Col1, Col2가 데이터 열
        new_headers={1: ("NewHead", 2)}      # Col 1에 새 헤더 생성
    )
    for hc in config:
        print(f"  Col {hc.col}: action={hc.action}, text='{hc.text}', rowspan={hc.rowspan}, colspan={hc.col_span}")

    # 행 추가
    print("\n[행 추가]")
    data = {"Col1": "AUTO1", "Col2": "AUTO2"}
    print(f"  데이터: {data}")

    merger.add_row_with_headers(data, config)

    print("\n[결과 테이블]")
    print_table_structure(table)

    return merger


def test_multiple_rows_with_nested_rowspan(hwpx_path: str):
    """여러 행 연속 추가 테스트"""
    print("\n" + "="*60)
    print("[5] 여러 행 연속 추가 테스트")
    print("="*60)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 0)

    print("\n[원본 테이블]")
    print_table_structure(table)

    # 3개 행 연속 추가
    data_list = [
        {"Col1": "Row1-A", "Col2": "Row1-B"},
        {"Col1": "Row2-A", "Col2": "Row2-B"},
        {"Col1": "Row3-A", "Col2": "Row3-B"},
    ]

    print("\n[3개 행 연속 추가]")
    for i, data in enumerate(data_list):
        print(f"\n--- 추가 {i+1}: {data} ---")
        initial = table.row_count
        merger._add_row_with_data(data)
        print(f"  행 수: {initial} -> {table.row_count}")

        # 마지막 행 상태 출력
        last_row = table.row_count - 1
        print(f"  Row {last_row}:")
        for col in range(table.col_count):
            cell = table.get_cell(last_row, col)
            if cell:
                if cell.row == last_row:
                    print(f"    Col {col}: '{cell.text}'")
                else:
                    print(f"    Col {col}: ↓ (from Row {cell.row}, rowspan={cell.row_span})")

    print("\n[최종 테이블]")
    print_table_structure(table)

    return merger


def test_add_rows_auto(hwpx_path: str):
    """자동 행 추가 테스트 (헤더 이름 기준)"""
    print("\n" + "="*60)
    print("[6] 자동 행 추가 테스트 (add_rows_auto)")
    print("="*60)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 0)

    print("\n[원본 테이블]")
    print_table_structure(table)

    # 데이터 리스트 - _header로 헤더 이름 지정
    data_list = [
        {"_header": "Head3", "Col1": "G", "Col2": "H"},   # Head3 확장
        {"_header": "Head3", "Col1": "I", "Col2": "J"},   # Head3 계속 확장
        {"_header": "Head4", "Col1": "K", "Col2": "L"},   # Head4 새로 생성
        {"_header": "Head4", "Col1": "M", "Col2": "N"},   # Head4 확장
        {"_header": "Head5", "Col1": "O", "Col2": "P"},   # Head5 새로 생성
    ]

    print("\n[추가할 데이터]")
    for i, d in enumerate(data_list):
        print(f"  {i+1}. {d}")

    # 자동 행 추가
    # header_col=1: Head2/Head3가 바뀌는 열
    # data_cols=[2, 3]: Col1, Col2
    # extend_cols=[0]: Head1은 항상 확장
    merger.add_rows_auto(
        data_list,
        header_col=1,        # Head2/Head3 열
        data_cols=[2, 3],    # Col1, Col2 열
        extend_cols=[0]      # Head1은 항상 확장
    )

    print("\n[결과 테이블]")
    print_table_structure(table)

    # 저장
    output_path = "data/test_add_rows_auto_result.hwpx"
    merger.save(output_path)
    print(f"\n결과 파일: {output_path}")

    return merger


def test_save_result(hwpx_path: str):
    """결과 파일 저장 테스트"""
    print("\n" + "="*60)
    print("[7] 결과 HWPX 파일 저장 테스트")
    print("="*60)

    merger = TableMerger()
    table = merger.load_base_table(hwpx_path, 0)

    print("\n[원본 테이블]")
    print_table_structure(table)

    # 새 헤더와 함께 행 추가
    config = [
        HeaderConfig(col=0, action='extend'),
        HeaderConfig(col=1, action='new', text='추가헤더', rowspan=2),
        HeaderConfig(col=2, action='data'),
        HeaderConfig(col=3, action='data'),
    ]

    merger.add_row_with_headers({"Col1": "데이터1", "Col2": "데이터2"}, config)

    # 두 번째 행 (추가헤더가 커버)
    config2 = [
        HeaderConfig(col=0, action='extend'),
        HeaderConfig(col=1, action='extend'),
        HeaderConfig(col=2, action='data'),
        HeaderConfig(col=3, action='data'),
    ]
    merger.add_row_with_headers({"Col1": "데이터3", "Col2": "데이터4"}, config2)

    print("\n[수정된 테이블]")
    print_table_structure(table)

    # 저장
    output_path = "data/test_nested_rowspan_result.hwpx"
    merger.save(output_path)
    print(f"\n결과 파일 저장: {output_path}")

    return output_path


def run_all_tests():
    """모든 테스트 실행"""
    # 테스트 파일 생성
    test_hwpx = "data/test_nested_rowspan.hwpx"
    os.makedirs("data", exist_ok=True)
    create_test_hwpx_with_nested_rowspan(test_hwpx)

    print("\n" + "#"*60)
    print("# 중첩 rowspan 테스트 시작")
    print("#"*60)

    # 테스트 실행
    table = test_parse_nested_rowspan(test_hwpx)
    if table is None:
        print("ERROR: 테이블 파싱 실패")
        return

    test_add_row_basic(test_hwpx)
    test_add_row_with_new_header(test_hwpx)
    test_auto_header_config(test_hwpx)
    test_multiple_rows_with_nested_rowspan(test_hwpx)
    test_add_rows_auto(test_hwpx)

    # 결과 파일 저장
    output_path = test_save_result(test_hwpx)

    print("\n" + "#"*60)
    print("# 모든 테스트 완료")
    print("#"*60)
    print(f"\n결과 파일: {output_path}")


if __name__ == "__main__":
    run_all_tests()
