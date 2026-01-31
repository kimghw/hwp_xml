# -*- coding: utf-8 -*-
"""
HWPX 파일 병합 모듈

개요(Outline) 기준으로 여러 HWPX 파일을 병합합니다.

규칙:
- 같은 level + 같은 이름 → 내용 이어붙이기
- 같은 level + 다른 이름 → 개요 추가
- 파일 순서 = 내용 순서
- 일반 문단은 위 개요에 종속
"""

import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Optional, Union, Set, Tuple, Any
from pathlib import Path
from io import BytesIO
import copy
import re

# 로컬 모듈
from .models import Paragraph, OutlineNode, HwpxData
from .parser import HwpxParser, NAMESPACES
from .outline import (
    build_outline_tree,
    merge_outline_trees,
    flatten_outline_tree,
    filter_outline_tree,
    get_all_outline_names,
    print_outline_tree,
)
from .table import TableParser, TableMerger
from .content_formatter import ContentFormatter

# 네임스페이스 등록
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class HwpxMerger:
    """HWPX 파일 병합"""

    def __init__(self, format_content: bool = True, use_sdk_for_levels: bool = True):
        """
        Args:
            format_content: 내용 문단에 글머리 기호 양식 적용 여부
            use_sdk_for_levels: SDK로 계층 레벨 분석 여부 (True: SDK 사용, False: 정규식만)
        """
        self.hwpx_data_list: List[HwpxData] = []
        self.parser = HwpxParser()
        self.table_parser = TableParser()
        self.exclude_outlines: Set[str] = set()
        # 템플릿 파일의 테이블 필드명 캐시: {field_name: table_index}
        self._template_table_fields: Dict[str, int] = {}
        self._template_path: Optional[Path] = None  # 템플릿 파일 경로 (정규화)
        self.format_content = format_content
        self.use_sdk_for_levels = use_sdk_for_levels
        self.content_formatter = ContentFormatter(style="default", use_sdk=use_sdk_for_levels) if format_content else None
        # 템플릿의 기존 글머리 포맷 예시 (SDK 참고용)
        self._existing_format: Optional[str] = None

    def add_file(self, hwpx_path: Union[str, Path]):
        """병합할 파일 추가"""
        data = self.parser.parse(hwpx_path)
        self.hwpx_data_list.append(data)

    def set_exclude_outlines(self, outlines: Union[Set[str], List[str]]):
        """
        제외할 개요 설정

        Args:
            outlines: 제외할 개요 이름 집합/리스트
                      예: {"1. 서론", "3."}  # "3."으로 시작하는 모든 개요 제외
        """
        self.exclude_outlines = set(outlines) if isinstance(outlines, list) else outlines

    def get_outline_list(self) -> List[Tuple[str, str]]:
        """
        모든 파일의 개요 목록 반환 (선택 UI용)

        Returns:
            (표시용 문자열, 실제 이름) 튜플 리스트
        """
        all_names = []
        seen = set()

        for data in self.hwpx_data_list:
            names = get_all_outline_names(data.outline_tree)
            for display, name in names:
                if name not in seen:
                    all_names.append((display, name))
                    seen.add(name)

        return all_names

    def merge(self, output_path: Optional[Union[str, Path]] = None, exclude_outlines: Optional[Set[str]] = None):
        """
        파일 병합 후 저장

        Args:
            output_path: 출력 파일 경로 (None이면 merge/output/merged.hwpx)
            exclude_outlines: 제외할 개요 (None이면 self.exclude_outlines 사용)
        """
        if not self.hwpx_data_list:
            raise ValueError("병합할 파일이 없습니다.")

        # 출력 경로 설정
        output_dir = Path(__file__).parent / "output"
        if output_path is None:
            output_path = output_dir / "merged.hwpx"
        else:
            output_path = Path(output_path)
            # 파일명만 주어진 경우 기본 출력 디렉토리에 저장
            if not output_path.is_absolute() and not output_path.parent.exists():
                output_path = output_dir / output_path.name

        # 출력 디렉토리 생성
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # 제외 개요 설정
        exclude = exclude_outlines if exclude_outlines is not None else self.exclude_outlines

        # 1. 개요 트리 병합
        trees = [data.outline_tree for data in self.hwpx_data_list]
        merged_tree = merge_outline_trees(trees, exclude)

        # 2. 병합된 문단 리스트 생성
        merged_paragraphs = flatten_outline_tree(merged_tree)

        # 내부 메서드 호출
        return self._merge_with_paragraphs(output_path, merged_paragraphs)

    def merge_with_tree(self, output_path: Union[str, Path], merged_tree: List) -> Path:
        """
        수정된 개요 트리로 병합 파일 생성

        merge_and_review에서 element를 수정한 후 호출합니다.

        Args:
            output_path: 출력 파일 경로
            merged_tree: 수정된 개요 트리 (OutlineNode 리스트)

        Returns:
            출력 파일 경로
        """
        if not self.hwpx_data_list:
            raise ValueError("병합할 파일이 없습니다.")

        output_path = Path(output_path)
        output_path.parent.mkdir(parents=True, exist_ok=True)

        # 병합된 문단 리스트 생성 (수정된 element 사용)
        merged_paragraphs = flatten_outline_tree(merged_tree)

        return self._merge_with_paragraphs(output_path, merged_paragraphs)

    def _merge_with_paragraphs(self, output_path: Path, merged_paragraphs: List[Paragraph]) -> Path:
        """문단 리스트로 병합 파일 생성 (내부 메서드)"""
        # 3. 템플릿 파일 (첫 번째 파일)
        template_data = self.hwpx_data_list[0]
        self._template_path = Path(template_data.path).resolve()

        # 4. BinData 병합 (이미지 ID 재매핑)
        merged_bin_data, bin_id_map = self._merge_bin_data()

        # 4.5 템플릿 테이블 필드명 수집
        self._collect_template_table_fields(template_data)

        # 4.6 템플릿의 기존 글머리 포맷 수집 (SDK 참고용)
        if self.format_content and self.use_sdk_for_levels:
            self._collect_existing_format(merged_paragraphs)

        # 5. section XML 생성 (테이블 머지 포함)
        merged_section_xml = self._create_merged_section(merged_paragraphs, template_data, bin_id_map)

        # 6. header.xml 병합 (스타일 병합)
        merged_header_xml = self._merge_headers()

        # 7. HWPX 파일 생성
        self._write_hwpx(output_path, template_data, merged_section_xml, merged_header_xml, merged_bin_data)

        return output_path

    def _merge_bin_data(self) -> Tuple[Dict[str, bytes], Dict[str, Dict[str, str]]]:
        """BinData 병합 및 ID 재매핑"""
        merged = {}
        bin_id_map = {}  # {source_file: {old_id: new_id}}

        next_id = 1

        for data in self.hwpx_data_list:
            bin_id_map[data.path] = {}

            for name, content in data.bin_data.items():
                # 기존 파일명에서 ID 추출 (예: BinData/image1.png -> 1)
                match = re.search(r'image(\d+)', name)
                old_id = match.group(1) if match else str(next_id)

                # 새 파일명 생성
                ext = Path(name).suffix
                new_name = f"BinData/image{next_id}{ext}"

                # 중복 체크 (동일 내용이면 재사용)
                content_exists = False
                for existing_name, existing_content in merged.items():
                    if existing_content == content:
                        # 동일 이미지 발견
                        existing_match = re.search(r'image(\d+)', existing_name)
                        if existing_match:
                            bin_id_map[data.path][old_id] = existing_match.group(1)
                            content_exists = True
                            break

                if not content_exists:
                    merged[new_name] = content
                    bin_id_map[data.path][old_id] = str(next_id)
                    next_id += 1

        return merged, bin_id_map

    def _merge_headers(self) -> bytes:
        """header.xml 병합 (첫 번째 파일 기준, 필요 시 스타일 추가)"""
        # 현재는 첫 번째 파일의 header 사용
        # TODO: 스타일 충돌 시 ID 재매핑 필요
        return self.hwpx_data_list[0].header_xml

    def _is_from_template(self, source_file: str) -> bool:
        """문단이 템플릿 파일에서 왔는지 확인"""
        if self._template_path is None:
            return False
        return Path(source_file).resolve() == self._template_path

    def _collect_template_table_fields(self, template_data: HwpxData):
        """템플릿 파일의 테이블 필드명 수집"""
        self._template_table_fields.clear()

        try:
            tables = self.table_parser.parse_tables(template_data.path)
            for table_idx, table in enumerate(tables):
                for cell in table.cells.values():
                    if cell.field_name:
                        self._template_table_fields[cell.field_name] = table_idx
        except Exception:
            # 테이블 파싱 실패 시 빈 상태로 유지
            pass

    def _collect_existing_format(self, paragraphs: List[Paragraph]):
        """
        병합 대상 파일들에서 기존 글머리 포맷 수집 (SDK 참고용)

        이미 글머리 기호가 적용된 문단을 찾아 예시로 사용
        템플릿 파일 우선, 없으면 다른 파일에서 수집
        """
        self._existing_format = None

        # 글머리 기호 패턴
        bullet_chars = '□■◆◇○●◎•-–—·∙'

        format_examples = []

        # 1단계: 템플릿 파일에서 먼저 수집
        for para in paragraphs:
            if not self._is_from_template(para.source_file):
                continue

            if para.has_table or para.has_image or para.is_outline:
                continue

            text = para.text.strip() if para.text else ''
            if not text:
                continue

            first_char = text.lstrip()[0] if text.lstrip() else ''
            if first_char in bullet_chars or text.startswith((' □', '   ○', '    -')):
                format_examples.append(text)

            if len(format_examples) >= 10:
                break

        # 2단계: 템플릿에 없으면 다른 파일에서 수집
        if not format_examples:
            for para in paragraphs:
                if self._is_from_template(para.source_file):
                    continue  # 템플릿 이미 확인함

                if para.has_table or para.has_image or para.is_outline:
                    continue

                text = para.text.strip() if para.text else ''
                if not text:
                    continue

                first_char = text.lstrip()[0] if text.lstrip() else ''
                if first_char in bullet_chars or text.startswith((' □', '   ○', '    -')):
                    format_examples.append(text)

                if len(format_examples) >= 10:
                    break

        if format_examples:
            self._existing_format = '\n'.join(format_examples)

    def _get_table_fields_from_element(self, tbl_elem) -> Set[str]:
        """테이블 요소에서 필드명 추출"""
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

    def _find_matching_template_table(self, addition_table_fields: Set[str]) -> Optional[int]:
        """
        추가 테이블의 필드명과 일치하는 템플릿 테이블 인덱스 반환

        Args:
            addition_table_fields: 추가 파일 테이블의 필드명 집합

        Returns:
            일치하는 테이블 인덱스, 없으면 None
        """
        if not addition_table_fields:
            return None

        # 필드명이 템플릿 테이블에 있는지 확인
        matching_tables = {}  # table_idx -> match_count

        for field_name in addition_table_fields:
            if field_name in self._template_table_fields:
                table_idx = self._template_table_fields[field_name]
                matching_tables[table_idx] = matching_tables.get(table_idx, 0) + 1

        if not matching_tables:
            return None

        # 가장 많이 일치하는 테이블 반환
        return max(matching_tables, key=matching_tables.get)

    def _create_merged_section(self, paragraphs: List[Paragraph], template_data: HwpxData,
                                bin_id_map: Dict[str, Dict[str, str]]) -> bytes:
        """병합된 section XML 생성"""
        # 템플릿 section XML 파싱
        template_section_name = sorted(template_data.section_xmls.keys())[0]
        template_section = template_data.section_xmls[template_section_name]

        root = ET.parse(BytesIO(template_section)).getroot()

        # 기존 문단 제거 (sec > p 요소들)
        p_elements = []
        for child in list(root):
            if child.tag.endswith('}p'):
                p_elements.append(child)

        for p in p_elements:
            root.remove(p)

        # 테이블 머지 상태 추적: {table_idx: 머지할 데이터}
        table_merge_data: Dict[int, List[Dict[str, str]]] = {}

        # 템플릿 테이블 요소 위치 추적: {table_idx: paragraph_elem}
        template_table_paragraphs: Dict[int, Any] = {}

        # 문단 순서 인덱스 → 요소 매핑 (글머리 기호 적용 후 참조용)
        # para.index는 원본 파일 내 인덱스라 중복될 수 있으므로 순차 인덱스 사용
        para_elements: Dict[int, Any] = {}
        para_seq_map: Dict[int, int] = {}  # para.id(객체 ID) → seq_idx

        # 병합된 문단 추가
        for seq_idx, para in enumerate(paragraphs):
            elem = copy.deepcopy(para.element)
            para_elements[seq_idx] = elem
            para_seq_map[id(para)] = seq_idx

            # Addition 문단의 스타일을 Template 스타일로 변경
            if not self._is_from_template(para.source_file) and para.para_pr_id:
                current_style = elem.get('paraPrIDRef', '')
                if current_style != para.para_pr_id:
                    elem.set('paraPrIDRef', para.para_pr_id)

            # 이미지 ID 재매핑
            if para.source_file in bin_id_map:
                self._remap_bin_ids(elem, bin_id_map[para.source_file])

            # 테이블이 있는 문단 처리
            if para.has_table:
                # 템플릿 파일의 문단인 경우 테이블 위치 기록
                if self._is_from_template(para.source_file):
                    for tbl in elem.iter():
                        if tbl.tag.endswith('}tbl'):
                            fields = self._get_table_fields_from_element(tbl)
                            if fields:
                                table_idx = self._find_matching_template_table(fields)
                                if table_idx is not None:
                                    template_table_paragraphs[table_idx] = elem
                    root.append(elem)

                # 추가(addition) 파일의 문단인 경우
                else:
                    tables_to_merge = []
                    tables_to_copy = []

                    for tbl in list(elem.iter()):
                        if not tbl.tag.endswith('}tbl'):
                            continue

                        fields = self._get_table_fields_from_element(tbl)
                        matching_table_idx = self._find_matching_template_table(fields)

                        if matching_table_idx is not None:
                            # 필드가 일치하는 테이블 → 머지
                            tables_to_merge.append((tbl, matching_table_idx, fields))
                        else:
                            # 필드 없음 → 복사
                            tables_to_copy.append(tbl)

                    # 테이블 머지 처리
                    for tbl, table_idx, fields in tables_to_merge:
                        table_data = self._extract_addition_table_data(tbl, fields)
                        if table_idx not in table_merge_data:
                            table_merge_data[table_idx] = []
                        table_merge_data[table_idx].extend(table_data)

                    # 테이블 복사 처리 (필드 없는 테이블은 문단 그대로 추가)
                    if tables_to_copy and not tables_to_merge:
                        root.append(elem)
                    elif tables_to_copy:
                        # 머지할 테이블과 복사할 테이블이 섞여있는 경우
                        # 복사할 테이블만 있는 문단 새로 생성
                        for tbl in tables_to_copy:
                            new_para = self._create_paragraph_with_table(elem, tbl)
                            root.append(new_para)
            else:
                root.append(elem)

        # 테이블 머지 적용
        self._apply_table_merges(root, template_data, table_merge_data)

        # 글머리 기호 양식 적용 (개요 단위로 내용 문단 모아서 처리)
        if self.format_content and self.content_formatter:
            self._apply_bullet_format_by_outline(root, paragraphs, para_elements, para_seq_map)

        # XML 직렬화
        return ET.tostring(root, encoding='UTF-8', xml_declaration=True)

    def _extract_addition_table_data(self, tbl_elem, fields: Set[str]) -> List[Dict[str, str]]:
        """추가(addition) 테이블에서 필드명-값 데이터 추출"""
        data = {}

        for tc in tbl_elem.iter():
            if not tc.tag.endswith('}tc'):
                continue

            field_name = tc.get('name', '')

            # subList에서 텍스트 추출
            text = ""
            for sublist in tc:
                if sublist.tag.endswith('}subList'):
                    for p in sublist:
                        if p.tag.endswith('}p'):
                            for run in p:
                                if run.tag.endswith('}run'):
                                    for t in run:
                                        if t.tag.endswith('}t') and t.text:
                                            text += t.text

            if field_name and field_name in fields:
                data[field_name] = text

        return [data] if data else []

    def _create_paragraph_with_table(self, original_para, tbl_elem) -> Any:
        """테이블을 포함하는 새 문단 요소 생성"""
        new_para = copy.deepcopy(original_para)

        # 기존 테이블 제거
        for run in list(new_para):
            if run.tag.endswith('}run'):
                for child in list(run):
                    if child.tag.endswith('}tbl'):
                        run.remove(child)

        # 새 테이블 추가 (첫 번째 run에)
        for run in new_para:
            if run.tag.endswith('}run'):
                run.append(copy.deepcopy(tbl_elem))
                break

        return new_para

    def _apply_table_merges(self, root, template_data: HwpxData, table_merge_data: Dict[int, List[Dict[str, str]]]):
        """수집된 추가 데이터를 템플릿 테이블에 머지"""
        if not table_merge_data:
            return

        # 템플릿 파일에서 테이블 파싱
        try:
            template_tables = self.table_parser.parse_tables(template_data.path)
        except Exception:
            return

        # root에서 테이블 요소 찾기
        table_elements = []
        for elem in root.iter():
            if elem.tag.endswith('}tbl'):
                table_elements.append(elem)

        # 각 테이블에 머지 적용
        for table_idx, addition_data_list in table_merge_data.items():
            if table_idx >= len(template_tables) or table_idx >= len(table_elements):
                continue

            template_table = template_tables[table_idx]
            tbl_elem = table_elements[table_idx]

            # TableMerger를 사용하여 머지
            merger = TableMerger()
            merger.base_table = template_table
            merger.base_table.element = tbl_elem

            # stub/gstub/input 기반 머지
            merger.merge_with_stub(addition_data_list)

    def _apply_bullet_format_by_outline(self, root, paragraphs: List[Paragraph], para_elements: Dict[int, Any], para_seq_map: Dict[int, int]):
        """
        개요 단위로 내용 문단들을 모아서 글머리 기호 양식 적용

        use_sdk_for_levels=True: SDK로 계층 구조 분석 후 정규식으로 글머리 적용
        use_sdk_for_levels=False: 정규식만으로 레벨 감지 및 글머리 적용

        같은 개요 아래의 여러 내용 문단을 한 번에 분석해서
        계층 구조를 파악하고 적절한 레벨(□/○/-)을 지정
        """
        # 개요별 내용 문단 그룹화
        content_groups: List[Tuple[int, List[Paragraph]]] = []  # [(outline_index, [content_paras])]
        current_outline_idx = -1
        current_content_paras = []

        for para in paragraphs:
            if para.is_outline:
                # 이전 그룹 저장
                if current_content_paras:
                    content_groups.append((current_outline_idx, current_content_paras))
                current_outline_idx = para.index
                current_content_paras = []
            else:
                # 내용 문단 (테이블/이미지 제외)
                if not para.has_table and not para.has_image and para.text and para.text.strip():
                    current_content_paras.append(para)

        # 마지막 그룹 저장
        if current_content_paras:
            content_groups.append((current_outline_idx, current_content_paras))

        # 각 그룹에 대해 레벨 분석 후 글머리 적용
        for outline_idx, content_paras in content_groups:
            if not content_paras:
                continue

            # 여러 문단의 텍스트를 줄바꿈으로 합침
            combined_text = '\n'.join(para.text for para in content_paras)

            if self.use_sdk_for_levels:
                # SDK로 레벨 분석 후 정규식으로 글머리 적용 (기존 포맷 참고)
                result = self.content_formatter.format_with_analyzed_levels(
                    combined_text,
                    existing_format=self._existing_format
                )
            else:
                # 정규식만으로 레벨 감지 및 글머리 적용
                result = self.content_formatter.auto_format(combined_text, use_sdk_for_levels=False)

            if result.success and result.formatted_text != combined_text:
                # 변환된 텍스트를 줄 단위로 분리 (빈 줄 제외)
                formatted_lines = [line for line in result.formatted_text.split('\n') if line.strip()]

                # 줄 수가 일치하는 경우에만 적용
                if len(formatted_lines) == len(content_paras):
                    for i, para in enumerate(content_paras):
                        new_text = formatted_lines[i]
                        seq_idx = para_seq_map.get(id(para))
                        if seq_idx is not None:
                            elem = para_elements.get(seq_idx)
                            if elem is not None:
                                self._update_paragraph_text(elem, new_text)
                else:
                    # 줄 수가 맞지 않으면 개별 문단씩 처리
                    for para in content_paras:
                        if self.use_sdk_for_levels:
                            result = self.content_formatter.format_with_analyzed_levels(
                                para.text,
                                existing_format=self._existing_format
                            )
                        else:
                            result = self.content_formatter.auto_format(para.text, use_sdk_for_levels=False)

                        if result.success and result.formatted_text != para.text:
                            # 첫 줄만 사용 (빈 줄 제외)
                            lines = [l for l in result.formatted_text.split('\n') if l.strip()]
                            if lines:
                                seq_idx = para_seq_map.get(id(para))
                                if seq_idx is not None:
                                    elem = para_elements.get(seq_idx)
                                    if elem is not None:
                                        self._update_paragraph_text(elem, lines[0])

    def _has_bullet(self, text: str) -> bool:
        """텍스트에 글머리 기호가 있는지 확인"""
        text = text.strip()
        if not text:
            return False

        # 글머리 기호 패턴
        bullet_chars = '□■◆◇○●◎•-–—·∙'
        first_char = text[0] if text else ''
        if first_char in bullet_chars:
            return True

        # 숫자 글머리 확인 (1. 또는 1) 형식)
        if re.match(r'^\d+[.)]\s', text):
            return True

        return False

    def _update_paragraph_text(self, elem, new_text: str):
        """문단 요소의 텍스트를 새 텍스트로 교체"""
        # 첫 번째 run의 t 요소 찾기
        for run in elem:
            if run.tag.endswith('}run'):
                for t in run:
                    if t.tag.endswith('}t'):
                        t.text = new_text
                        # 나머지 t 요소 제거 (텍스트를 하나로 합침)
                        for other_t in list(run)[list(run).index(t)+1:]:
                            if other_t.tag.endswith('}t'):
                                run.remove(other_t)
                        return

    def _remap_bin_ids(self, elem, id_map: Dict[str, str]):
        """요소 내 BinData ID 재매핑"""
        # binDataIDRef 속성 찾아서 변경
        for child in elem.iter():
            bin_ref = child.get('binDataIDRef')
            if bin_ref and bin_ref in id_map:
                child.set('binDataIDRef', id_map[bin_ref])

            # imgID도 처리
            img_id = child.get('imgID')
            if img_id and img_id in id_map:
                child.set('imgID', id_map[img_id])

    def _update_content_hpf(self, content_hpf: bytes, bin_data: Dict[str, bytes]) -> bytes:
        """content.hpf에 이미지 항목 추가"""
        root = ET.parse(BytesIO(content_hpf)).getroot()

        # manifest 찾기
        manifest = None
        for elem in root.iter():
            if elem.tag.endswith('}manifest'):
                manifest = elem
                break

        if manifest is None:
            return content_hpf

        # 기존 이미지 항목 ID 수집
        existing_ids = set()
        for item in manifest:
            if item.tag.endswith('}item'):
                item_id = item.get('id', '')
                if item_id.startswith('image'):
                    existing_ids.add(item_id)

        # 새 이미지 항목 추가
        ns = '{http://www.idpf.org/2007/opf/}'
        for name in sorted(bin_data.keys()):
            # BinData/image1.jpeg -> image1
            match = re.search(r'(image\d+)', name)
            if not match:
                continue

            image_id = match.group(1)
            if image_id in existing_ids:
                continue

            # 확장자로 media-type 결정
            ext = Path(name).suffix.lower()
            media_types = {
                '.jpeg': 'image/jpeg',
                '.jpg': 'image/jpeg',
                '.png': 'image/png',
                '.gif': 'image/gif',
                '.bmp': 'image/bmp',
            }
            media_type = media_types.get(ext, 'application/octet-stream')

            # item 요소 생성
            item = ET.Element(f'{ns}item')
            item.set('id', image_id)
            item.set('href', name)
            item.set('media-type', media_type)
            item.set('isEmbeded', '1')
            manifest.append(item)

        return ET.tostring(root, encoding='UTF-8', xml_declaration=True)

    def _write_hwpx(self, output_path: Path, template_data: HwpxData,
                    section_xml: bytes, header_xml: bytes, bin_data: Dict[str, bytes]):
        """HWPX 파일 생성"""
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            # mimetype (압축 안 함)
            if 'mimetype' in template_data.other_files:
                zf.writestr('mimetype', template_data.other_files['mimetype'], compress_type=zipfile.ZIP_STORED)

            # header.xml
            zf.writestr('Contents/header.xml', header_xml)

            # section0.xml (병합된 것)
            zf.writestr('Contents/section0.xml', section_xml)

            # BinData
            for name, content in bin_data.items():
                zf.writestr(name, content)

            # 기타 파일 (mimetype 제외, content.hpf는 별도 처리)
            for name, content in template_data.other_files.items():
                if name == 'mimetype' or name.startswith('BinData/'):
                    continue

                if name == 'Contents/content.hpf':
                    # content.hpf에 이미지 항목 추가
                    updated_content = self._update_content_hpf(content, bin_data)
                    zf.writestr(name, updated_content)
                else:
                    zf.writestr(name, content)


def get_outline_structure(hwpx_path: Union[str, Path]) -> List[OutlineNode]:
    """HWPX 파일에서 개요 구조 추출"""
    parser = HwpxParser()
    data = parser.parse(hwpx_path)
    return data.outline_tree


# 기본 출력 디렉토리
DEFAULT_OUTPUT_DIR = Path(__file__).parent / "output"


def merge_hwpx_files(
    hwpx_paths: List[Union[str, Path]],
    output_path: Optional[Union[str, Path]] = None,
    exclude_outlines: Optional[Set[str]] = None,
    format_content: bool = True,
    use_sdk_for_levels: bool = True
) -> Path:
    """
    여러 HWPX 파일을 개요 기준으로 병합

    Args:
        hwpx_paths: 병합할 파일 경로들
        output_path: 출력 파일 경로 (None이면 merge/output/merged.hwpx)
        exclude_outlines: 제외할 개요 이름 집합
        format_content: 글머리 기호 양식 적용 여부
        use_sdk_for_levels: SDK로 계층 레벨 분석 여부 (True: SDK 사용, False: 정규식만)

    Returns:
        출력 파일 경로
    """
    # 출력 경로 설정
    if output_path is None:
        output_path = DEFAULT_OUTPUT_DIR / "merged.hwpx"
    else:
        output_path = Path(output_path)
        # 파일명만 주어진 경우 기본 출력 디렉토리에 저장
        if not output_path.parent.exists() and output_path.parent == Path('.'):
            output_path = DEFAULT_OUTPUT_DIR / output_path.name

    # 출력 디렉토리 생성
    output_path.parent.mkdir(parents=True, exist_ok=True)

    merger = HwpxMerger(format_content=format_content, use_sdk_for_levels=use_sdk_for_levels)

    for path in hwpx_paths:
        merger.add_file(path)

    if exclude_outlines:
        merger.set_exclude_outlines(exclude_outlines)

    return merger.merge(output_path)


if __name__ == "__main__":
    import sys
    import argparse

    parser = argparse.ArgumentParser(
        description="HWPX 파일 개요 기준 병합",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
예제:
  # 파일 구조 확인
  python -m merge.merge_hwpx file1.hwpx file2.hwpx

  # 개요 목록 출력 (제외할 개요 선택용)
  python -m merge.merge_hwpx --list-outlines file1.hwpx file2.hwpx

  # 기본 병합
  python -m merge.merge_hwpx -o output.hwpx file1.hwpx file2.hwpx

  # 특정 개요 제외하여 병합
  python -m merge.merge_hwpx -o output.hwpx --exclude "1. 서론" "3. 결론" file1.hwpx file2.hwpx

  # 접두사로 제외 (3으로 시작하는 모든 개요 제외)
  python -m merge.merge_hwpx -o output.hwpx --exclude "3." file1.hwpx file2.hwpx
"""
    )

    parser.add_argument("files", nargs="+", help="병합할 HWPX 파일들")
    parser.add_argument("-o", "--output", help="출력 파일 경로")
    parser.add_argument("--exclude", nargs="*", help="제외할 개요 이름 (복수 가능)")
    parser.add_argument("--list-outlines", action="store_true", help="개요 목록만 출력")

    args = parser.parse_args()

    # 개요 목록 출력
    if args.list_outlines:
        print("=" * 60)
        print("개요 목록 (제외할 개요 선택 참고)")
        print("=" * 60)

        merger = HwpxMerger()
        for path in args.files:
            merger.add_file(path)

        outlines = merger.get_outline_list()
        for i, (display, name) in enumerate(outlines, 1):
            print(f"  {i:3d}. {display}")
            print(f"       → 제외 시: --exclude \"{name}\"")

        print("\n접두사로 제외하려면:")
        print("  --exclude \"1.\"  → '1.'로 시작하는 모든 개요 제외")
        print("  --exclude \"2\"   → '2.'로 시작하는 모든 개요 제외")
        sys.exit(0)

    # 병합
    if args.output:
        exclude_set = set(args.exclude) if args.exclude else None

        print(f"병합 파일: {args.files}")
        print(f"출력 파일: {args.output}")
        if exclude_set:
            print(f"제외 개요: {exclude_set}")
        print("=" * 60)

        # 각 파일 구조 출력
        for i, path in enumerate(args.files):
            print(f"\n[파일 {i+1}] {path}")
            print("-" * 40)
            tree = get_outline_structure(path)
            print_outline_tree(tree)

        # 병합
        result = merge_hwpx_files(args.files, args.output, exclude_set)
        print("\n" + "=" * 60)
        print(f"병합 완료: {result}")

        # 결과 구조 출력
        print("\n[병합 결과]")
        print("-" * 40)
        merged_tree = get_outline_structure(result)
        print_outline_tree(merged_tree)

    else:
        # 구조 확인만
        print(f"파일 수: {len(args.files)}")
        print("=" * 60)

        for i, path in enumerate(args.files):
            print(f"\n[파일 {i+1}] {path}")
            print("-" * 40)
            tree = get_outline_structure(path)
            print_outline_tree(tree)
