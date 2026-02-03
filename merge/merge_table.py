# -*- coding: utf-8 -*-
"""
HWPX 테이블 병합 모듈

템플릿 파일의 테이블에 추가 파일의 데이터를 병합합니다.

주요 기능:
- 템플릿 테이블 필드명 수집
- 추가 파일에서 테이블 데이터 추출
- 필드명 매칭으로 테이블 병합
"""

from typing import Dict, List, Optional, Set, Any
from pathlib import Path

from .table import TableParser, TableMerger
from .models import HwpxData


class TableMergeHandler:
    """테이블 병합 처리기"""

    def __init__(self, format_content: bool = True, use_sdk_for_levels: bool = True):
        """
        Args:
            format_content: add_ 필드에 글머리 기호 포맷팅 적용 여부
            use_sdk_for_levels: SDK로 레벨 분석 여부
        """
        self.table_parser = TableParser()
        self.format_content = format_content
        self.use_sdk_for_levels = use_sdk_for_levels

        # 템플릿 파일의 테이블 필드명 캐시: {field_name: [table_index, ...]}
        self._template_table_fields: Dict[str, List[int]] = {}

    def collect_template_fields(self, template_data: HwpxData):
        """
        템플릿 파일의 테이블 필드명 수집

        Args:
            template_data: 템플릿 파일 데이터
        """
        self._template_table_fields.clear()

        try:
            tables = self.table_parser.parse_tables(template_data.path)
            for table_idx, table in enumerate(tables):
                for cell in table.cells.values():
                    if cell.field_name:
                        # 동일 필드명이 여러 테이블에 있을 수 있으므로 리스트로 저장
                        if cell.field_name not in self._template_table_fields:
                            self._template_table_fields[cell.field_name] = []
                        if table_idx not in self._template_table_fields[cell.field_name]:
                            self._template_table_fields[cell.field_name].append(table_idx)
        except Exception:
            # 테이블 파싱 실패 시 빈 상태로 유지
            pass

    def get_fields_from_element(self, tbl_elem) -> Set[str]:
        """
        테이블 요소에서 필드명 추출

        Args:
            tbl_elem: 테이블 XML 요소

        Returns:
            필드명 집합
        """
        fields = set()

        for tc in tbl_elem.iter():
            if tc.tag.endswith('}tc'):
                # tc의 name 속성
                tc_name = tc.get('name', '')
                if tc_name:
                    fields.add(tc_name)

                # subList 내 fieldBegin의 name
                for child in tc.iter():
                    if child.tag.endswith('}fieldBegin'):
                        name = child.get('name', '')
                        if name:
                            fields.add(name)

        return fields

    def find_matching_table(self, addition_fields: Set[str]) -> Optional[int]:
        """
        추가 테이블의 필드명과 일치하는 템플릿 테이블 인덱스 반환

        Args:
            addition_fields: 추가 파일 테이블의 필드명 집합

        Returns:
            일치하는 테이블 인덱스, 없으면 None
        """
        if not addition_fields:
            return None

        # 필드명이 템플릿 테이블에 있는지 확인
        matching_tables = {}  # table_idx -> match_count

        for field_name in addition_fields:
            if field_name in self._template_table_fields:
                # 필드명이 여러 테이블에 있을 수 있음
                for table_idx in self._template_table_fields[field_name]:
                    matching_tables[table_idx] = matching_tables.get(table_idx, 0) + 1

        if not matching_tables:
            return None

        # 가장 많이 일치하는 테이블 반환
        return max(matching_tables, key=matching_tables.get)

    def extract_table_data(self, tbl_elem, fields: Set[str]) -> List[Dict[str, str]]:
        """
        추가(addition) 테이블에서 필드명-값 데이터를 행별로 추출

        Args:
            tbl_elem: 테이블 XML 요소
            fields: 추출할 필드명 집합

        Returns:
            행별 데이터 리스트 [{field_name: value}, ...]
        """
        # 행별 데이터 수집: {row_idx: {field_name: text}}
        row_data: Dict[int, Dict[str, str]] = {}

        # gstub/stub 셀 정보 수집: (start_row, end_row, field_name, text)
        gstub_cells = []

        for tc in tbl_elem.iter():
            if not tc.tag.endswith('}tc'):
                continue

            field_name = tc.get('name', '')
            if not field_name:
                continue

            # 셀 주소와 span 정보 추출
            row_idx = 0
            row_span = 1
            for child in tc:
                if child.tag.endswith('}cellAddr'):
                    row_idx = int(child.get('rowAddr', 0))
                elif child.tag.endswith('}cellSpan'):
                    row_span = int(child.get('rowSpan', 1))

            # subList에서 텍스트 추출 (여러 문단은 줄바꿈으로 구분)
            paragraphs_text = []
            for sublist in tc:
                if sublist.tag.endswith('}subList'):
                    for p in sublist:
                        if p.tag.endswith('}p'):
                            p_text = ""
                            for run in p:
                                if run.tag.endswith('}run'):
                                    for t in run:
                                        if t.tag.endswith('}t') and t.text:
                                            p_text += t.text
                            paragraphs_text.append(p_text)
            text = '\n'.join(paragraphs_text)

            # gstub/stub 셀은 rowspan 정보와 함께 저장
            if field_name.startswith('gstub_') or field_name.startswith('stub_'):
                end_row = row_idx + row_span - 1
                gstub_cells.append((row_idx, end_row, field_name, text))

            # 매칭되는 필드만 저장 (input_, gstub_, stub_, add_ 등)
            if field_name in fields or field_name.startswith('gstub_') or field_name.startswith('stub_') or field_name.startswith('add_'):
                if row_idx not in row_data:
                    row_data[row_idx] = {}
                row_data[row_idx][field_name] = text

        # gstub/stub 값을 해당 rowspan 범위의 모든 행에 전파
        for start_row, end_row, field_name, text in gstub_cells:
            for r in range(start_row, end_row + 1):
                if r not in row_data:
                    row_data[r] = {}
                if field_name not in row_data[r]:
                    row_data[r][field_name] = text

        # 행 순서대로 리스트 반환 (헤더 행, data_ 행, 빈 input 행 제외)
        # 단, add_ 필드는 헤더 행에 있어도 추출함
        result = []

        # add_ 필드는 모든 행에서 추출 (헤더 행 포함)
        add_fields_data = {}
        for row_idx in sorted(row_data.keys()):
            data = row_data[row_idx]
            for field_name, value in data.items():
                if field_name.startswith('add_') and value:
                    if field_name not in add_fields_data:
                        add_fields_data[field_name] = value

        # add_ 필드가 있으면 별도 행으로 추가
        if add_fields_data:
            result.append(add_fields_data)

        for row_idx in sorted(row_data.keys()):
            if row_idx == 0:  # 헤더 행 스킵 (add_ 필드는 이미 처리됨)
                continue

            data = row_data[row_idx]
            if not data:  # 빈 행 스킵
                continue

            # add_ 필드 제외 (이미 처리됨)
            data_without_add = {k: v for k, v in data.items() if not k.startswith('add_')}
            if not data_without_add:  # add_ 필드만 있는 행은 스킵
                continue

            # data_ 필드만 있는 행 스킵 (데이터 행)
            non_data_fields = [k for k in data_without_add.keys() if not k.startswith('data_')]
            if not non_data_fields:
                continue

            # input_ 값이 모두 비어있으면 스킵 (gstub/stub만 있는 빈 행)
            input_values = [v for k, v in data_without_add.items() if k.startswith('input_')]
            if input_values and all(not v for v in input_values):
                continue

            result.append(data_without_add)

        return result

    def apply_merges(self, root, table_merge_data: Dict[int, List[Dict[str, str]]]):
        """
        수집된 추가 데이터를 템플릿 테이블에 머지

        Args:
            root: section XML 루트 요소
            table_merge_data: {table_idx: [행 데이터 리스트]}
        """
        if not table_merge_data:
            return

        # root에서 테이블 요소 찾기
        table_elements = []
        for elem in root.iter():
            if elem.tag.endswith('}tbl'):
                table_elements.append(elem)

        # 각 테이블에 머지 적용
        for table_idx, addition_data_list in table_merge_data.items():
            if table_idx >= len(table_elements):
                continue

            tbl_elem = table_elements[table_idx]

            # tbl_elem을 직접 파싱하여 TableInfo 생성 (element 참조 일치)
            table_info = self.table_parser._parse_table(tbl_elem)

            # TableMerger를 사용하여 머지
            merger = TableMerger(
                format_add_content=self.format_content,
                use_sdk_for_levels=self.use_sdk_for_levels,
            )
            merger.base_table = table_info

            # stub/gstub/input 기반 머지
            merger.merge_with_stub(addition_data_list)


def merge_table_data(
    template_path: str,
    addition_paths: List[str],
    output_path: str,
    table_index: int = 0,
    format_content: bool = True,
    use_sdk_for_levels: bool = True
) -> str:
    """
    단독 테이블 병합 함수

    템플릿 테이블에 여러 추가 파일의 데이터를 병합합니다.

    Args:
        template_path: 템플릿 HWPX 파일 경로
        addition_paths: 추가 데이터 HWPX 파일 경로 리스트
        output_path: 출력 HWPX 파일 경로
        table_index: 병합할 테이블 인덱스
        format_content: add_ 필드 글머리 포맷팅 적용 여부
        use_sdk_for_levels: SDK로 레벨 분석 여부

    Returns:
        출력 파일 경로
    """
    merger = TableMerger(
        format_add_content=format_content,
        use_sdk_for_levels=use_sdk_for_levels
    )
    merger.load_base_table(template_path, table_index)

    parser = TableParser()

    for addition_path in addition_paths:
        tables = parser.parse_tables(addition_path)
        if tables:
            # 첫 번째 테이블에서 데이터 추출
            # TODO: 필드명 매칭으로 테이블 찾기
            pass

    merger.save(output_path)
    return output_path
