# -*- coding: utf-8 -*-
"""
HWPX 테이블 병합 모듈

템플릿 파일의 테이블에 추가 파일의 데이터를 병합합니다.

주요 기능:
- 템플릿 테이블 필드명 수집
- 추가 파일에서 테이블 데이터 추출
- 필드명 매칭으로 테이블 병합
"""

from typing import Dict, List, Optional, Set, Any, TYPE_CHECKING
from pathlib import Path
from dataclasses import dataclass, field

from .table import TableParser, TableMerger
from .table.row_extractor import RowExtractor
from .models import HwpxData
from .formatters import BaseFormatter

if TYPE_CHECKING:
    from .models import OutlineNode


@dataclass
class TableMergePlan:
    """테이블 병합 계획"""
    table_idx: int = 0  # 템플릿 테이블 인덱스
    template_fields: List[str] = field(default_factory=list)  # 템플릿 필드명
    addition_data: List[Dict[str, str]] = field(default_factory=list)  # 추가할 데이터
    source_file: str = ""  # 데이터 출처 파일


class TableMergeHandler:
    """테이블 병합 처리기"""

    def __init__(
        self,
        format_content: bool = True,
        use_sdk_for_levels: bool = True,
        add_formatter: Optional[BaseFormatter] = None,
    ):
        """
        Args:
            format_content: add_ 필드에 글머리 기호 포맷팅 적용 여부
            use_sdk_for_levels: SDK로 레벨 분석 여부
            add_formatter: add_ 필드용 포맷터 (BaseFormatter 상속)
        """
        self.table_parser = TableParser()
        self.row_extractor = RowExtractor()
        self.format_content = format_content
        self.use_sdk_for_levels = use_sdk_for_levels
        self.add_formatter = add_formatter

        # 기준 파일의 테이블 필드명 캐시: {field_name: [table_index, ...]}
        self._base_table_fields: Dict[str, List[int]] = {}

    def get_fields_from_file(self, hwpx_data: HwpxData):
        """
        HWPX 파일의 테이블 필드명 수집

        Args:
            hwpx_data: HWPX 파일 데이터
        """
        self._base_table_fields.clear()

        try:
            tables = self.table_parser.parse_tables(hwpx_data.path)
            for table_idx, table in enumerate(tables):
                # 테이블 요소에서 필드명 추출
                fields = self.get_fields_from_element(table.element)
                for field_name in fields:
                    if field_name not in self._base_table_fields:
                        self._base_table_fields[field_name] = []
                    if table_idx not in self._base_table_fields[field_name]:
                        self._base_table_fields[field_name].append(table_idx)
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

    def find_matching_table(self, fields: Set[str]) -> Optional[int]:
        """
        필드명과 일치하는 기준 테이블 인덱스 반환

        Args:
            fields: 테이블의 필드명 집합

        Returns:
            일치하는 테이블 인덱스, 없으면 None
        """
        if not fields:
            return None

        # 필드명이 기준 테이블에 있는지 확인
        matching_tables = {}  # table_idx -> match_count

        for field_name in fields:
            if field_name in self._base_table_fields:
                for table_idx in self._base_table_fields[field_name]:
                    matching_tables[table_idx] = matching_tables.get(table_idx, 0) + 1

        if not matching_tables:
            return None

        # 가장 많이 일치하는 테이블 반환
        return max(matching_tables, key=matching_tables.get)

    def collect_table_data(
        self,
        hwpx_data_list: List[HwpxData],
        merged_tree: List['OutlineNode']
    ) -> List[TableMergePlan]:
        """
        병합할 테이블 데이터 수집

        템플릿 테이블의 필드와 Addition 테이블의 필드를 매칭하여
        병합할 데이터를 수집합니다.

        처리 방식:
        - 필드 없는 테이블: 본문처럼 그대로 복사 (이 메서드에서 처리 안 함)
        - 필드 있는 테이블: 데이터만 수집하여 반환

        Args:
            hwpx_data_list: 파싱된 HWPX 데이터 리스트
            merged_tree: 병합된 개요 트리

        Returns:
            테이블 병합 계획 리스트 (필드 매칭된 테이블만)
        """
        from .outline import flatten_outline_tree

        plans: List[TableMergePlan] = []

        if len(hwpx_data_list) < 2:
            return plans

        # 템플릿(첫 번째 파일)의 테이블 필드 수집
        template_data = hwpx_data_list[0]
        self.get_fields_from_file(template_data)

        # 템플릿 테이블 필드 정보 저장
        template_table_fields: Dict[int, List[str]] = {}
        for field_name, table_indices in self._base_table_fields.items():
            for table_idx in table_indices:
                if table_idx not in template_table_fields:
                    template_table_fields[table_idx] = []
                template_table_fields[table_idx].append(field_name)

        # 병합된 트리에서 테이블 문단 추출
        merged_paragraphs = flatten_outline_tree(merged_tree)

        # Addition 파일들의 테이블 데이터 수집
        template_path = str(Path(template_data.path).resolve())

        for para in merged_paragraphs:
            if not para.has_table:
                continue

            # 템플릿 파일은 건너뜀
            if str(Path(para.source_file).resolve()) == template_path:
                continue

            # 테이블 요소에서 필드 추출
            for tbl in para.element.iter():
                if not tbl.tag.endswith('}tbl'):
                    continue

                fields = self.get_fields_from_element(tbl)
                matching_idx = self.find_matching_table(fields)

                if matching_idx is not None:
                    # 테이블 데이터 추출
                    table_data = self.extract_table_data(tbl, fields)

                    # 기존 계획에 추가하거나 새 계획 생성
                    existing_plan = None
                    for plan in plans:
                        if plan.table_idx == matching_idx:
                            existing_plan = plan
                            break

                    if existing_plan:
                        existing_plan.addition_data.extend(table_data)
                    else:
                        plan = TableMergePlan(
                            table_idx=matching_idx,
                            template_fields=template_table_fields.get(matching_idx, []),
                            addition_data=table_data,
                            source_file=para.source_file,
                        )
                        plans.append(plan)

        return plans

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
                add_formatter=self.add_formatter,
            )
            merger.base_table = table_info

            # stub/gstub/input 기반 머지
            merger.merge_with_stub(addition_data_list)
