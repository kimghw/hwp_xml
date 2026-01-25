# -*- coding: utf-8 -*-
"""
HWPX XML에서 paragraph 정보 추출하여 COM API 결과와 비교

## 분석 결과

### XML hp:p 요소 구조
- id: 문단 ID (대부분 2147483648로 재사용됨 - 위치 식별 불가)
- paraPrIDRef: 문단 스타일 속성 참조 ID (header.xml 참조)
- styleIDRef: 스타일 ID 참조

### COM API para_id
- SetPos(list_id, para_id, char_id)에서 사용
- list_id 내에서 0부터 순차 증가
- list_id는 테이블 셀, 각주, 머리글 등 텍스트 컨테이너 단위

### 결론: XML id ≠ COM API para_id

| 구분 | XML id 속성 | COM API para_id |
|------|------------|-----------------|
| 값 범위 | 2147483648 대량 재사용 | list_id 내 0부터 증가 |
| 용도 | 편집 히스토리 추적 | 커서 위치 지정 |
| 위치 식별 | 불가능 | 가능 |

### 매핑 방법
1. 문서 순서(doc_order) 매칭: 양쪽 결과를 순서대로 1:1 매칭
2. 텍스트 매칭: 동일 텍스트로 매칭
3. 셀 위치 매칭: (table_index, row, col) → list_id
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Dict, Tuple
from dataclasses import dataclass


@dataclass
class XmlPara:
    """XML에서 추출한 문단 정보"""
    # XML 속성
    xml_id: str = ""  # hp:p의 id 속성
    para_pr_id: str = ""  # paraPrIDRef
    style_id: str = ""  # styleIDRef

    # 위치 정보
    location: str = ""  # "body" 또는 "table_R_C" (행_열)
    table_index: int = -1  # 테이블 인덱스 (-1이면 본문)
    cell_row: int = -1
    cell_col: int = -1
    para_index_in_cell: int = 0  # 셀 내 문단 순서

    # 문서 순서
    doc_order: int = 0  # 문서 내 전체 순서

    # 텍스트
    text: str = ""


def extract_paragraphs_from_hwpx(hwpx_path: str) -> List[XmlPara]:
    """HWPX 파일에서 모든 문단 정보 추출"""
    paragraphs = []
    doc_order = 0

    with zipfile.ZipFile(hwpx_path, 'r') as zf:
        # section 파일 찾기
        section_files = [f for f in zf.namelist() if f.startswith('Contents/section') and f.endswith('.xml')]
        section_files.sort()

        for section_file in section_files:
            with zf.open(section_file) as f:
                tree = ET.parse(f)
                root = tree.getroot()

                # 네임스페이스 처리
                ns = {
                    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
                    'hs': 'http://www.hancom.co.kr/hwpml/2011/section'
                }

                # 본문 직접 문단 (테이블 외부)
                table_index = 0

                for child in root:
                    tag_local = child.tag.split('}')[-1]

                    if tag_local == 'p':
                        # 본문 문단
                        para = _parse_paragraph(child, ns)
                        para.location = "body"
                        para.doc_order = doc_order
                        doc_order += 1
                        paragraphs.append(para)

                    elif tag_local == 'tbl':
                        # 테이블 내 문단
                        table_paras = _parse_table_paragraphs(child, ns, table_index, doc_order)
                        for p in table_paras:
                            p.doc_order = doc_order
                            doc_order += 1
                        paragraphs.extend(table_paras)
                        table_index += 1

    return paragraphs


def _parse_paragraph(p_elem, ns: dict) -> XmlPara:
    """hp:p 요소에서 문단 정보 추출"""
    para = XmlPara()
    para.xml_id = p_elem.get('id', '')
    para.para_pr_id = p_elem.get('paraPrIDRef', '')
    para.style_id = p_elem.get('styleIDRef', '')

    # 텍스트 추출 (hp:run > hp:t)
    texts = []
    for run in p_elem:
        if run.tag.endswith('}run'):
            for t in run:
                if t.tag.endswith('}t') and t.text:
                    texts.append(t.text)
    para.text = ''.join(texts)

    return para


def _parse_table_paragraphs(tbl_elem, ns: dict, table_index: int, start_order: int) -> List[XmlPara]:
    """테이블 내 모든 문단 추출"""
    paragraphs = []

    for tr in tbl_elem:
        if not tr.tag.endswith('}tr'):
            continue

        for tc in tr:
            if not tc.tag.endswith('}tc'):
                continue

            # 셀 주소
            cell_addr = tc.find('.//hp:cellAddr', {'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph'})
            if cell_addr is None:
                for elem in tc:
                    if elem.tag.endswith('}cellAddr'):
                        cell_addr = elem
                        break

            row = int(cell_addr.get('rowAddr', 0)) if cell_addr is not None else 0
            col = int(cell_addr.get('colAddr', 0)) if cell_addr is not None else 0

            # subList 내 문단
            para_index = 0
            for sublist in tc:
                if not sublist.tag.endswith('}subList'):
                    continue

                for p in sublist:
                    if not p.tag.endswith('}p'):
                        continue

                    para = _parse_paragraph(p, ns)
                    para.location = f"table{table_index}_R{row}_C{col}"
                    para.table_index = table_index
                    para.cell_row = row
                    para.cell_col = col
                    para.para_index_in_cell = para_index
                    para_index += 1
                    paragraphs.append(para)

    return paragraphs


def main():
    """메인 함수"""
    import sys
    import os

    # 기본 HWPX 파일
    hwpx_path = r"C:\hwp_xml\extract_table_field_xml.hwpx"

    if len(sys.argv) > 1:
        hwpx_path = sys.argv[1]

    if not os.path.exists(hwpx_path):
        print(f"파일 없음: {hwpx_path}")
        return

    print("=" * 70)
    print("HWPX XML 문단 추출 결과")
    print("=" * 70)
    print(f"파일: {hwpx_path}")
    print()

    paragraphs = extract_paragraphs_from_hwpx(hwpx_path)

    print(f"총 {len(paragraphs)}개 문단 발견")
    print()

    # XML id 통계
    id_counts = {}
    for p in paragraphs:
        id_counts[p.xml_id] = id_counts.get(p.xml_id, 0) + 1

    print("XML id 분포:")
    for xml_id, count in sorted(id_counts.items(), key=lambda x: -x[1])[:10]:
        print(f"  id='{xml_id}': {count}개")
    print()

    # 본문 vs 테이블
    body_count = sum(1 for p in paragraphs if p.location == "body")
    table_count = len(paragraphs) - body_count
    print(f"본문 문단: {body_count}개")
    print(f"테이블 문단: {table_count}개")
    print()

    # 처음 20개 문단 출력
    print("-" * 70)
    print("처음 20개 문단:")
    print("-" * 70)
    for i, p in enumerate(paragraphs[:20]):
        text_preview = p.text[:30] + "..." if len(p.text) > 30 else p.text
        print(f"[{i:3d}] {p.location:20s} | xml_id={p.xml_id:15s} | paraPrID={p.para_pr_id:3s} | {text_preview}")

    print()
    print("-" * 70)
    print("결론: XML id ≠ COM API para_id")
    print("-" * 70)
    print("""
XML id 속성:
  - 2147483648 대량 재사용 → 위치 식별 불가
  - 편집 히스토리 추적용

COM API para_id:
  - list_id 내에서 0부터 순차 증가
  - 위치 식별 가능 (list_id + para_id)

매핑 방법:
  1. 문서 순서(doc_order) 매칭
  2. 텍스트 매칭
  3. 셀 위치 (table, row, col) → list_id 매핑
""")


if __name__ == "__main__":
    main()
