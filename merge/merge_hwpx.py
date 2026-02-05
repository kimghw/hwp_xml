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
from .merge_table import TableMergeHandler
from .content_formatter import ContentFormatter
from .formatters import BaseFormatter

# 네임스페이스 등록
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class HwpxMerger:
    """HWPX 파일 병합"""

    def __init__(
        self,
        format_content: bool = True,
        use_sdk_for_levels: bool = True,
        add_formatter: Optional[BaseFormatter] = None,
    ):
        """
        Args:
            format_content: 내용 문단에 글머리 기호 양식 적용 여부
            use_sdk_for_levels: SDK로 계층 레벨 분석 여부 (True: SDK 사용, False: 정규식만)
            add_formatter: add_ 필드용 포맷터 (BaseFormatter 상속)
        """
        self.hwpx_data_list: List[HwpxData] = []
        self.parser = HwpxParser()
        self.exclude_outlines: Set[str] = set()
        self._template_path: Optional[Path] = None  # 템플릿 파일 경로 (정규화)
        self.format_content = format_content
        self.use_sdk_for_levels = use_sdk_for_levels
        self.add_formatter = add_formatter
        self.content_formatter = ContentFormatter(style="default", use_sdk=use_sdk_for_levels) if format_content else None
        # 템플릿의 기존 글머리 포맷 예시 (SDK 참고용)
        self._existing_format: Optional[str] = None
        # 테이블 병합 처리기 (add_formatter 전달)
        self.table_handler = TableMergeHandler(
            format_content, use_sdk_for_levels, add_formatter=add_formatter
        )

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

    def merge_with_tree(
        self,
        output_path: Union[str, Path],
        merged_tree: List,
    ) -> Path:
        """
        수정된 개요 트리로 병합 파일 생성

        테이블은 이미 병합 완료된 상태로 전달됩니다.

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

        merged_paragraphs = flatten_outline_tree(merged_tree)
        return self._merge_with_paragraphs(output_path, merged_paragraphs)

    def _merge_with_paragraphs(self, output_path: Path, merged_paragraphs: List[Paragraph]) -> Path:
        """문단 리스트로 병합 파일 생성 (내부 메서드)"""
        # 3. 템플릿 파일 (첫 번째 파일)
        template_data = self.hwpx_data_list[0]
        self._template_path = Path(template_data.path).resolve()

        # 4. BinData 병합 (이미지 ID 재매핑)
        merged_bin_data, bin_id_map = self._merge_bin_data()

        # 4.5 템플릿의 기존 글머리 포맷 수집 (SDK 참고용)
        if self.format_content and self.use_sdk_for_levels:
            self._collect_existing_format(merged_paragraphs)

        # 5. section XML 생성 (테이블은 이미 병합 완료)
        merged_section_xml = self._create_merged_section(merged_paragraphs, template_data, bin_id_map)

        # 6. header.xml 병합 (스타일 병합)
        merged_header_xml = self._merge_headers(merged_section_xml)

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

    def _merge_headers(self, merged_section_xml: bytes) -> bytes:
        """
        header.xml 병합

        - 템플릿(첫 번째 파일)을 기본으로 사용
        - 병합된 section에서 참조하는 스타일 ID가 템플릿에 없으면
          다른 입력 파일의 정의를 가져와 추가한다.
        - charPr은 템플릿 스타일만 사용 (Addition 파일의 charPr는 복사하지 않음)
        """
        template_header_xml = self.hwpx_data_list[0].header_xml
        header_root = ET.fromstring(template_header_xml)

        ns = {'hh': NAMESPACES['hh']}

        # section에서 필요한 ID 수집
        sec_root = ET.fromstring(merged_section_xml)
        needed = {
            'charPr': set(),
            'paraPr': set(),
            'tabPr': set(),
            'borderFill': set(),
        }
        for elem in sec_root.iter():
            for attr, key in (
                ('charPrIDRef', 'charPr'),
                ('paraPrIDRef', 'paraPr'),
                ('tabPrIDRef', 'tabPr'),
                ('borderFillIDRef', 'borderFill'),
            ):
                ref = elem.attrib.get(attr)
                if ref:
                    needed[key].add(ref)

        # 템플릿에 이미 있는 ID 집합
        existing = {
            'charPr': {e.get('id') for e in header_root.findall('.//hh:charPr', ns)},
            'paraPr': {e.get('id') for e in header_root.findall('.//hh:paraPr', ns)},
            'tabPr': {e.get('id') for e in header_root.findall('.//hh:tabPr', ns)},
            'borderFill': {e.get('id') for e in header_root.findall('.//hh:borderFill', ns)},
        }

        # 모든 입력 파일에서 정의 수집 (id -> Element)
        defs = {'charPr': {}, 'paraPr': {}, 'tabPr': {}, 'borderFill': {}}
        for data in self.hwpx_data_list:
            root = ET.fromstring(data.header_xml)
            for key, xpath in (
                ('charPr', './/hh:charPr'),
                ('paraPr', './/hh:paraPr'),
                ('tabPr', './/hh:tabPr'),
                ('borderFill', './/hh:borderFill'),
            ):
                for elem in root.findall(xpath, ns):
                    _id = elem.get('id')
                    if _id and _id not in defs[key]:
                        defs[key][_id] = copy.deepcopy(elem)

        # 필요한데 템플릿에 없는 정의 추가
        targets = {
            'charPr': header_root.find('.//hh:charProperties', ns),
            'paraPr': header_root.find('.//hh:paraProperties', ns),
            'tabPr': header_root.find('.//hh:tabProperties', ns),
            'borderFill': header_root.find('.//hh:borderFills', ns),
        }

        # charPr은 템플릿만 사용, 나머지 스타일만 병합
        for key in ['paraPr', 'tabPr', 'borderFill']:
            target = targets.get(key)
            if target is None:
                continue
            missing_ids = needed[key] - existing[key]
            for _id in sorted(missing_ids, key=lambda x: int(x) if x.isdigit() else x):
                if _id in defs[key]:
                    target.append(copy.deepcopy(defs[key][_id]))
                    existing[key].add(_id)

            # itemCnt 업데이트
            if 'itemCnt' in target.attrib:
                count = len(target.findall(f'hh:{key}', ns))
                target.set('itemCnt', str(count))

        return ET.tostring(header_root, encoding='utf-8', xml_declaration=True)

    def _is_from_template(self, source_file: str) -> bool:
        """문단이 템플릿 파일에서 왔는지 확인"""
        if self._template_path is None:
            return False
        return Path(source_file).resolve() == self._template_path

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

    def _create_merged_section(self, paragraphs: List[Paragraph], template_data: HwpxData,
                                bin_id_map: Dict[str, Dict[str, str]]) -> bytes:
        """병합된 section XML 생성"""
        # 템플릿 section XML 파싱
        template_section_name = sorted(template_data.section_xmls.keys())[0]
        template_section = template_data.section_xmls[template_section_name]

        root = ET.parse(BytesIO(template_section)).getroot()

        # 원본 루트 태그 저장 (네임스페이스 복원용)
        template_str = template_section.decode('utf-8')
        xml_decl_end = template_str.find('?>') + 2
        root_end = template_str.find('>', xml_decl_end) + 1
        self._original_root_tag = template_str[xml_decl_end:root_end].strip()

        # 기존 문단 제거 (sec > p 요소들)
        p_elements = []
        for child in list(root):
            if child.tag.endswith('}p'):
                p_elements.append(child)

        for p in p_elements:
            root.remove(p)

        # 문단 순서 인덱스 → 요소 매핑 (글머리 기호 적용 후 참조용)
        para_elements: Dict[int, Any] = {}
        para_seq_map: Dict[int, int] = {}  # para.id(객체 ID) → seq_idx

        # 병합된 문단 추가 (테이블은 이미 병합 완료됨)
        for seq_idx, para in enumerate(paragraphs):
            elem = copy.deepcopy(para.element)
            para_elements[seq_idx] = elem
            para_seq_map[id(para)] = seq_idx

            # Addition 문단의 스타일을 Template 스타일로 변경
            if not self._is_from_template(para.source_file):
                # 문단 스타일 변경
                if para.para_pr_id:
                    current_style = elem.get('paraPrIDRef', '')
                    if current_style != para.para_pr_id:
                        elem.set('paraPrIDRef', para.para_pr_id)

                # 문자 스타일을 템플릿 기본 스타일(ID=0)로 변경
                self._remap_char_styles_to_template(elem)

            # 이미지 ID 재매핑
            if para.source_file in bin_id_map:
                self._remap_bin_ids(elem, bin_id_map[para.source_file])

            # 테이블이 있는 문단 처리
            if para.has_table:
                # 템플릿 파일의 문단: 그대로 추가
                if self._is_from_template(para.source_file):
                    root.append(elem)
                # Addition 파일의 문단: 필드 없는 테이블만 복사
                else:
                    has_field_table = False
                    for tbl in elem.iter():
                        if not tbl.tag.endswith('}tbl'):
                            continue
                        fields = self.table_handler.get_fields_from_element(tbl)
                        if self.table_handler.find_matching_table(fields) is not None:
                            has_field_table = True
                            break
                    if not has_field_table:
                        root.append(elem)
            else:
                root.append(elem)

        # 글머리 기호 양식 적용 (개요 단위로 내용 문단 모아서 처리)
        if self.format_content and self.content_formatter:
            self._apply_bullet_format_by_outline(root, paragraphs, para_elements, para_seq_map)

        # XML 직렬화
        xml_str = ET.tostring(root, encoding='unicode')

        # ElementTree가 생성한 루트 태그를 원본으로 교체 (네임스페이스 유지)
        # 실제로 사용되는 네임스페이스만 원본 태그에 추가
        if hasattr(self, '_original_root_tag'):
            new_root_end = xml_str.find('>')
            body_str = xml_str[new_root_end + 1:]

            # 원본 루트 태그에 없는 네임스페이스 프리픽스 찾기
            orig_ns = set(re.findall(r'xmlns:(\w+)=', self._original_root_tag))

            # 본문에서 사용되는 프리픽스 찾기 (태그만, 속성 제외)
            used_prefixes = set(re.findall(r'<(\w+):', body_str))

            # 누락된 네임스페이스 추가
            missing_prefixes = used_prefixes - orig_ns
            extra_ns = ''
            for prefix in missing_prefixes:
                if prefix in NAMESPACES:
                    extra_ns += f' xmlns:{prefix}="{NAMESPACES[prefix]}"'

            if extra_ns:
                fixed_root = self._original_root_tag[:-1] + extra_ns + '>'
            else:
                fixed_root = self._original_root_tag

            xml_str = fixed_root + body_str

        xml_header = "<?xml version='1.0' encoding='utf-8'?>\n"
        return (xml_header + xml_str).encode('utf-8')

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

    def _remap_char_styles_to_template(self, elem):
        """
        Addition 파일의 문자 스타일을 템플릿 기본 스타일로 변경

        Addition 파일에서 가져온 문단의 모든 run 요소의 charPrIDRef를
        템플릿의 기본 문자 스타일(ID="0")로 변경합니다.

        이렇게 하면 Addition 파일의 글꼴, 색상, 크기가 템플릿의 기본 스타일을 따릅니다.
        """
        TEMPLATE_DEFAULT_CHAR_STYLE = "0"  # 템플릿 기본 문자 스타일 ID

        for child in elem.iter():
            # run 요소의 charPrIDRef 속성을 템플릿 기본값으로 변경
            if child.tag.endswith('}run'):
                char_ref = child.get('charPrIDRef')
                if char_ref:
                    child.set('charPrIDRef', TEMPLATE_DEFAULT_CHAR_STYLE)

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
        """HWPX 파일 생성 (원본 ZipInfo 속성 보존)"""
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            # 원본 ZipInfo를 기반으로 새 ZipInfo 생성
            def make_zip_info(name: str, compress_type: int = zipfile.ZIP_DEFLATED) -> zipfile.ZipInfo:
                if name in template_data.zip_infos:
                    orig = template_data.zip_infos[name]
                    info = zipfile.ZipInfo(name)
                    info.compress_type = compress_type
                    info.external_attr = orig.external_attr
                    info.date_time = orig.date_time
                    return info
                else:
                    info = zipfile.ZipInfo(name)
                    info.compress_type = compress_type
                    # 기본 external_attr (Windows에서 생성된 것처럼)
                    info.external_attr = 0x81A40000  # 2175008768
                    return info

            # mimetype (압축 안 함)
            if 'mimetype' in template_data.other_files:
                info = make_zip_info('mimetype', zipfile.ZIP_STORED)
                zf.writestr(info, template_data.other_files['mimetype'])

            # header.xml
            info = make_zip_info('Contents/header.xml')
            zf.writestr(info, header_xml)

            # section0.xml (병합된 것)
            info = make_zip_info('Contents/section0.xml')
            zf.writestr(info, section_xml)

            # 추가 섹션들 (section1.xml, section2.xml 등) - 템플릿에서 그대로 복사
            for section_name, section_content in template_data.section_xmls.items():
                if section_name != 'Contents/section0.xml':
                    info = make_zip_info(section_name)
                    zf.writestr(info, section_content)

            # BinData
            for name, content in bin_data.items():
                info = make_zip_info(name)
                zf.writestr(info, content)

            # 기타 파일 (mimetype 제외, content.hpf는 별도 처리)
            for name, content in template_data.other_files.items():
                if name == 'mimetype' or name.startswith('BinData/'):
                    continue

                info = make_zip_info(name)
                if name == 'Contents/content.hpf':
                    # content.hpf에 이미지 항목 추가
                    updated_content = self._update_content_hpf(content, bin_data)
                    zf.writestr(info, updated_content)
                else:
                    zf.writestr(info, content)


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

