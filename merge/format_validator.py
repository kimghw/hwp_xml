# -*- coding: utf-8 -*-
"""
HWPX 병합 데이터 형식 검증 및 수정 모듈

Claude Agent SDK를 이용한 Sub Agent로 병합 데이터의 형식을 검토하고 수정합니다.

검토 항목:
1. 그림/테이블 캡션 스타일
2. 개요 3단계의 글머리 기호 (네모 → 원 → -)
3. add_ 필드 텍스트 형식 검증 (병합 전 형식 맞춤)
4. input_ 필드 데이터 형식 검증 (선택적)

사용 예:
    # add_ 필드 검증
    validator = AddFieldValidator()
    result = validator.validate_add_content(
        "추가할 내용...",
        base_cell_style="bullet_list"
    )

    # 개요 본문 검증
    result = validator.validate_outline(
        "## 섹션 제목\\n내용...",
        target_level=2
    )

    # YAML 설정 기반 검증
    from merge.formatters import load_config
    config = load_config("formatter_config.yaml")
    fixer = FormatFixer.from_config(config)
"""

import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Union, Tuple, Callable, Any
from pathlib import Path
from io import BytesIO
import copy
import re

# formatters 모듈 import (YAML 설정 지원, 정규식 기반)
try:
    from .formatters import (
        BulletFormatter as RegexBulletFormatter,
        CaptionFormatter as RegexCaptionFormatter,
        load_config,
        FormatterConfig,
        BULLET_STYLES as FORMATTER_BULLET_STYLES,
    )
    HAS_FORMATTERS = True
except ImportError:
    HAS_FORMATTERS = False
    FORMATTER_BULLET_STYLES = None
    RegexBulletFormatter = None
    RegexCaptionFormatter = None

# SDK 기반 포맷터 import (핵심 - agent 모듈)
try:
    from agent import (
        BulletFormatter as SDKBulletFormatter,
        CaptionFormatter as SDKCaptionFormatter,
    )
    HAS_SDK_FORMATTERS = True
except ImportError:
    HAS_SDK_FORMATTERS = False
    SDKBulletFormatter = None
    SDKCaptionFormatter = None


# 글머리 기호 정의 (기본값)
BULLET_STYLES = {
    0: "■",  # 네모 (1단계)
    1: "●",  # 원 (2단계)
    2: "-",  # 대시 (3단계)
}

# 기본 캡션 스타일
DEFAULT_CAPTION_STYLE = {
    "table": "표 {num}. {title}",
    "figure": "그림 {num}. {title}",
}

# 셀 스타일 타입
CELL_STYLE_TYPES = {
    "plain": "일반 텍스트",
    "bullet_list": "글머리 기호 목록",
    "numbered_list": "번호 매기기 목록",
    "heading": "헤딩/제목",
}


@dataclass
class ValidationResult:
    """검증 결과"""
    is_valid: bool = True
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    fixes: List[Dict] = field(default_factory=list)  # 수정 내역


@dataclass
class AddFieldValidationResult:
    """add_ 필드 검증 결과"""
    success: bool = True
    original_text: str = ""
    validated_text: str = ""
    changes_made: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)


@dataclass
class CellStyleInfo:
    """셀 스타일 정보"""
    style_type: str = "plain"  # plain, bullet_list, numbered_list, heading
    indent_level: int = 0
    font_size: Optional[int] = None
    is_bold: bool = False
    line_spacing: Optional[float] = None


@dataclass
class CaptionInfo:
    """캡션 정보"""
    type: str = ""  # "table" or "figure"
    text: str = ""
    style_valid: bool = True
    element: any = None
    para_index: int = 0


@dataclass
class BulletInfo:
    """글머리 기호 정보"""
    level: int = 0  # 개요 레벨 (0, 1, 2)
    current_bullet: str = ""
    expected_bullet: str = ""
    is_valid: bool = True
    element: any = None
    para_index: int = 0
    text: str = ""


class FormatValidator:
    """HWPX 형식 검증기"""

    def __init__(self, caption_styles: Dict[str, str] = None, bullet_styles: Dict[int, str] = None):
        self.caption_styles = caption_styles or DEFAULT_CAPTION_STYLE
        self.bullet_styles = bullet_styles or BULLET_STYLES

    def validate(self, hwpx_path: Union[str, Path]) -> ValidationResult:
        """HWPX 파일 형식 검증"""
        hwpx_path = Path(hwpx_path)
        result = ValidationResult()

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            # section 파일들 검증
            section_files = sorted([
                f for f in zf.namelist()
                if f.startswith('Contents/section') and f.endswith('.xml')
            ])

            for section_file in section_files:
                section_content = zf.read(section_file)
                self._validate_section(section_content, result)

        return result

    def _validate_section(self, xml_content: bytes, result: ValidationResult):
        """section XML 검증"""
        root = ET.parse(BytesIO(xml_content)).getroot()

        para_idx = 0
        for elem in root.iter():
            if elem.tag.endswith('}p'):
                # 캡션 검증
                self._validate_caption(elem, para_idx, result)

                # 글머리 기호 검증
                self._validate_bullet(elem, para_idx, result)

                para_idx += 1

    def _validate_caption(self, p_elem, para_idx: int, result: ValidationResult):
        """캡션 스타일 검증"""
        text = self._extract_text(p_elem)

        # 테이블 캡션 패턴
        table_patterns = [
            r'^표\s*\d+[.\s]',
            r'^Table\s*\d+[.\s]',
            r'^\[표\s*\d+\]',
        ]

        # 그림 캡션 패턴
        figure_patterns = [
            r'^그림\s*\d+[.\s]',
            r'^Figure\s*\d+[.\s]',
            r'^\[그림\s*\d+\]',
        ]

        for pattern in table_patterns:
            if re.match(pattern, text, re.IGNORECASE):
                # 테이블 캡션 발견
                expected_pattern = r'^표\s+\d+\.\s+'
                if not re.match(expected_pattern, text):
                    result.warnings.append(
                        f"문단 {para_idx}: 테이블 캡션 형식이 '표 N. 제목' 형식이 아닙니다: '{text[:30]}...'"
                    )
                    result.fixes.append({
                        'type': 'caption',
                        'para_index': para_idx,
                        'original': text,
                        'expected_format': '표 N. 제목',
                    })
                break

        for pattern in figure_patterns:
            if re.match(pattern, text, re.IGNORECASE):
                # 그림 캡션 발견
                expected_pattern = r'^그림\s+\d+\.\s+'
                if not re.match(expected_pattern, text):
                    result.warnings.append(
                        f"문단 {para_idx}: 그림 캡션 형식이 '그림 N. 제목' 형식이 아닙니다: '{text[:30]}...'"
                    )
                    result.fixes.append({
                        'type': 'caption',
                        'para_index': para_idx,
                        'original': text,
                        'expected_format': '그림 N. 제목',
                    })
                break

    def _validate_bullet(self, p_elem, para_idx: int, result: ValidationResult):
        """글머리 기호 검증 (개요 3단계)"""
        text = self._extract_text(p_elem)
        if not text:
            return

        # 개요 3단계 글머리 기호 패턴 검사
        bullet_patterns = {
            '■': 0,  # 네모 → level 0
            '□': 0,
            '●': 1,  # 원 → level 1
            '○': 1,
            '-': 2,  # 대시 → level 2
            '‑': 2,
            '–': 2,
        }

        first_char = text[0] if text else ''

        if first_char in bullet_patterns:
            detected_level = bullet_patterns[first_char]

            # paraPrIDRef에서 개요 레벨 확인 (개요 스타일인 경우)
            para_pr_id = p_elem.get('paraPrIDRef', '')

            # 글머리 기호 순서 검증은 컨텍스트가 필요함
            # 여기서는 일단 기록만 함
            result.fixes.append({
                'type': 'bullet_check',
                'para_index': para_idx,
                'bullet': first_char,
                'detected_level': detected_level,
                'text': text[:50],
            })

    def _extract_text(self, p_elem) -> str:
        """문단에서 텍스트 추출"""
        texts = []
        for run in p_elem.iter():
            if run.tag.endswith('}t') and run.text:
                texts.append(run.text)
        return ''.join(texts)


class FormatFixer:
    """HWPX 형식 수정기 (SDK 기반)"""

    def __init__(
        self,
        bullet_styles: Dict[int, str] = None,
        caption_formatter = None,
        bullet_formatter = None,
        use_sdk: bool = True
    ):
        """
        Args:
            bullet_styles: 글머리 스타일 {level: symbol}
            caption_formatter: 캡션 포맷터 (SDK 또는 정규식)
            bullet_formatter: 글머리 포맷터 (SDK 또는 정규식)
            use_sdk: SDK 포맷터 사용 여부 (기본: True)
        """
        self.bullet_styles = bullet_styles or BULLET_STYLES
        self._bullet_counters = {}  # 개요 레벨별 카운터
        self.use_sdk = use_sdk and HAS_SDK_FORMATTERS

        # SDK 포맷터 우선, fallback으로 정규식 포맷터
        if caption_formatter:
            self._caption_formatter = caption_formatter
        elif self.use_sdk and SDKCaptionFormatter:
            self._caption_formatter = SDKCaptionFormatter()
        elif HAS_FORMATTERS and RegexCaptionFormatter:
            self._caption_formatter = RegexCaptionFormatter()
        else:
            self._caption_formatter = None

        if bullet_formatter:
            self._bullet_formatter = bullet_formatter
        elif self.use_sdk and SDKBulletFormatter:
            self._bullet_formatter = SDKBulletFormatter()
        elif HAS_FORMATTERS and RegexBulletFormatter:
            self._bullet_formatter = RegexBulletFormatter()
        else:
            self._bullet_formatter = None

        # SDK 사용 가능 여부 로깅
        if self.use_sdk and self._bullet_formatter:
            print(f"    [SDK] BulletFormatter 활성화")

    @classmethod
    def from_config(cls, config: 'FormatterConfig', use_sdk: bool = True) -> 'FormatFixer':
        """
        YAML 설정에서 FormatFixer 생성

        Args:
            config: FormatterConfig 객체 (load_config로 로드)
            use_sdk: SDK 포맷터 사용 여부 (기본: True)

        Returns:
            FormatFixer 인스턴스

        사용 예:
            from merge.formatters import load_config
            config = load_config("formatter_config.yaml")
            fixer = FormatFixer.from_config(config)
        """
        if not HAS_FORMATTERS:
            raise ImportError("formatters 모듈을 import할 수 없습니다.")

        # 스타일에서 글머리 기호 추출
        bullet_style = config.bullet.style
        if FORMATTER_BULLET_STYLES and bullet_style in FORMATTER_BULLET_STYLES:
            style_config = FORMATTER_BULLET_STYLES[bullet_style]
            # (symbol, indent) 튜플에서 symbol만 추출
            bullet_styles = {
                level: symbol_tuple[0].strip()
                for level, symbol_tuple in style_config.items()
            }
        else:
            bullet_styles = BULLET_STYLES

        # SDK 포맷터 우선 생성
        if use_sdk and HAS_SDK_FORMATTERS:
            caption_formatter = SDKCaptionFormatter()
            bullet_formatter = SDKBulletFormatter(style=bullet_style)
        else:
            caption_formatter = RegexCaptionFormatter() if RegexCaptionFormatter else None
            bullet_formatter = RegexBulletFormatter(style=bullet_style) if RegexBulletFormatter else None

        return cls(
            bullet_styles=bullet_styles,
            caption_formatter=caption_formatter,
            bullet_formatter=bullet_formatter,
            use_sdk=use_sdk
        )

    def fix_bullets_in_tree(self, outline_tree, parent_level: int = -1) -> List[Dict]:
        """
        개요 트리에서 글머리 기호 수정 (SDK 레벨 분석 사용)

        규칙:
        - 마지막 개요(자식 없는 노드)의 본문에만 글머리 적용
        - 개요 문단 자체에는 글머리 적용 안 함
        - add_ 필드는 테이블 병합에서 별도 처리
        """
        fixes = []

        for node in outline_tree:
            # 하위 개요가 있으면 먼저 재귀 처리
            if node.children:
                child_fixes = self.fix_bullets_in_tree(node.children, node.level)
                fixes.extend(child_fixes)
                # 자식이 있는 노드는 본문에 글머리 적용 안 함
                continue

            # 마지막 개요(leaf node)만 본문에 글머리 적용
            # 노드의 문단들 텍스트 수집 (개요 문단 제외)
            content_paras = [p for p in node.paragraphs if not p.is_outline and p.text]
            if not content_paras:
                continue

            para_texts = [p.text for p in content_paras]

            # SDK로 레벨 분석 + 기존 글머리/번호 제거
            if self.use_sdk and self._bullet_formatter and para_texts:
                combined_text = '\n'.join(para_texts)
                try:
                    # analyze_and_strip: 레벨 분석 + 기존 글머리/번호 제거
                    if hasattr(self._bullet_formatter, 'analyze_and_strip'):
                        analyzed_levels, stripped_texts = self._bullet_formatter.analyze_and_strip(combined_text)
                    else:
                        # fallback: 레벨만 분석
                        analyzed_levels = self._bullet_formatter.analyze_levels(combined_text)
                        stripped_texts = para_texts  # 원본 유지
                except Exception as e:
                    print(f"    [SDK 오류] 레벨 분석 실패: {e}")
                    analyzed_levels = [0] * len(para_texts)
                    stripped_texts = para_texts
            else:
                analyzed_levels = [0] * len(para_texts)
                stripped_texts = para_texts

            # 각 내용 문단에 분석된 레벨과 정제된 텍스트 적용
            for i, para in enumerate(content_paras):
                # SDK 분석 레벨 사용
                level = analyzed_levels[i] if i < len(analyzed_levels) else 0
                level = max(0, min(level, 2))  # 0~2 범위로 제한
                expected_bullet = self.bullet_styles.get(level, '-')

                # SDK에서 정제된 텍스트 사용 (기존 글머리/번호 제거됨)
                pure_text = stripped_texts[i].strip() if i < len(stripped_texts) else para.text.strip()

                # 새 글머리로 텍스트 생성
                new_text = expected_bullet + ' ' + pure_text

                if new_text != para.text:
                    fixes.append({
                        'type': 'bullet_fix',
                        'para_index': para.index,
                        'original_bullet': para.text[:10] + '...' if len(para.text) > 10 else para.text,
                        'new_bullet': expected_bullet,
                        'level': level,
                        'original_text': para.text,
                        'new_text': new_text,
                        'sdk_analyzed': self.use_sdk,
                    })
                    para.text = new_text
                    self._update_element_text(para.element, new_text)

        return fixes

    def fix_caption_format(self, paragraphs: list) -> List[Dict]:
        """캡션 형식 수정"""
        fixes = []
        table_num = 0
        figure_num = 0

        for para in paragraphs:
            text = para.text

            # 테이블 캡션 수정
            table_match = re.match(r'^(?:표|Table|\[표)\s*(\d+)[.\s\]]*(.*)$', text, re.IGNORECASE)
            if table_match:
                table_num = int(table_match.group(1))
                title = table_match.group(2).strip()
                new_text = f"표 {table_num}. {title}"

                if text != new_text:
                    fixes.append({
                        'type': 'caption_fix',
                        'para_index': para.index,
                        'caption_type': 'table',
                        'original': text,
                        'new': new_text,
                    })
                    # 텍스트 수정
                    para.text = new_text
                    # XML 요소도 수정
                    self._update_element_text(para.element, new_text)
                continue

            # 그림 캡션 수정
            figure_match = re.match(r'^(?:그림|Figure|\[그림)\s*(\d+)[.\s\]]*(.*)$', text, re.IGNORECASE)
            if figure_match:
                figure_num = int(figure_match.group(1))
                title = figure_match.group(2).strip()
                new_text = f"그림 {figure_num}. {title}"

                if text != new_text:
                    fixes.append({
                        'type': 'caption_fix',
                        'para_index': para.index,
                        'caption_type': 'figure',
                        'original': text,
                        'new': new_text,
                    })
                    # 텍스트 수정
                    para.text = new_text
                    # XML 요소도 수정
                    self._update_element_text(para.element, new_text)

        return fixes

    def _update_element_text(self, elem, new_text: str):
        """
        문단 요소의 텍스트를 새 텍스트로 교체

        Args:
            elem: 문단 XML 요소 (p)
            new_text: 새 텍스트
        """
        if elem is None:
            return

        # 첫 번째 run의 t 요소 찾아서 텍스트 교체
        for run in elem:
            if run.tag.endswith('}run'):
                t_elements = [t for t in run if t.tag.endswith('}t')]
                if t_elements:
                    # 첫 번째 t에 새 텍스트 설정
                    t_elements[0].text = new_text
                    # 나머지 t 요소 제거 (텍스트를 하나로 합침)
                    for t in t_elements[1:]:
                        run.remove(t)
                    return


class AddFieldValidator:
    """
    add_ 필드 및 개요 본문 형식 검증기

    Claude Code SDK를 사용하여 병합 전 데이터 형식을 검증하고 조정합니다.
    """

    def __init__(self, sdk_validator: Optional[Callable] = None):
        """
        Args:
            sdk_validator: Claude Code SDK 검증 함수 (외부 주입)
                          None이면 기본 규칙 기반 검증 사용
        """
        self.sdk_validator = sdk_validator

    def validate_add_content(
        self,
        text: str,
        base_cell_style: Optional[str] = None,
        base_cell_info: Optional[CellStyleInfo] = None,
        separator: str = " "
    ) -> AddFieldValidationResult:
        """
        add_ 필드 내용 검증

        기존 셀의 스타일에 맞게 텍스트를 검증하고 조정합니다.

        Args:
            text: 추가할 텍스트
            base_cell_style: 기존 셀 스타일 타입
                           ("plain", "bullet_list", "numbered_list", "heading")
            base_cell_info: 상세 셀 스타일 정보 (CellStyleInfo)
            separator: 같은 문단 구분자 (기본: 빈칸 1개)

        Returns:
            AddFieldValidationResult: 검증 결과
        """
        if not text:
            return AddFieldValidationResult(
                success=True,
                original_text=text,
                validated_text=text,
                changes_made=[]
            )

        original = text
        changes = []
        warnings = []

        # SDK 검증기가 있으면 사용
        if self.sdk_validator:
            try:
                validated = self._validate_with_sdk(text, "add", base_cell_style)
                if validated != text:
                    changes.append("SDK 검증으로 형식 조정됨")
                    text = validated
            except Exception as e:
                warnings.append(f"SDK 검증 실패, 기본 규칙 사용: {str(e)}")

        # 기본 규칙 기반 검증
        style = base_cell_style or (base_cell_info.style_type if base_cell_info else "plain")

        if style == "bullet_list":
            text, bullet_changes = self._format_as_bullet_list(text)
            changes.extend(bullet_changes)

        elif style == "numbered_list":
            text, num_changes = self._format_as_numbered_list(text)
            changes.extend(num_changes)

        elif style == "heading":
            text, heading_changes = self._format_as_heading(text, base_cell_info)
            changes.extend(heading_changes)

        else:
            # plain 스타일
            text, plain_changes = self._format_as_plain(text, separator)
            changes.extend(plain_changes)

        return AddFieldValidationResult(
            success=True,
            original_text=original,
            validated_text=text,
            changes_made=changes,
            warnings=warnings
        )

    def validate_outline(
        self,
        text: str,
        target_level: int = 1,
        max_level: int = 6
    ) -> AddFieldValidationResult:
        """
        개요 본문 검증

        마크다운 헤딩 레벨과 목록 스타일을 검증합니다.

        Args:
            text: 개요 텍스트
            target_level: 목표 헤딩 레벨 (1-6)
            max_level: 최대 허용 헤딩 레벨

        Returns:
            AddFieldValidationResult: 검증 결과
        """
        if not text:
            return AddFieldValidationResult(
                success=True,
                original_text=text,
                validated_text=text,
                changes_made=[]
            )

        original = text
        changes = []
        warnings = []

        # SDK 검증기가 있으면 사용
        if self.sdk_validator:
            try:
                validated = self._validate_with_sdk(text, "outline", str(target_level))
                if validated != text:
                    changes.append("SDK 검증으로 개요 형식 조정됨")
                    text = validated
            except Exception as e:
                warnings.append(f"SDK 검증 실패, 기본 규칙 사용: {str(e)}")

        # 기본 규칙: 헤딩 레벨 조정
        lines = text.split('\n')
        adjusted_lines = []

        for line in lines:
            # 마크다운 헤딩 패턴
            heading_match = re.match(r'^(#{1,6})\s+(.+)$', line)
            if heading_match:
                current_level = len(heading_match.group(1))
                content = heading_match.group(2)

                if current_level > max_level:
                    new_level = max_level
                    changes.append(f"헤딩 레벨 {current_level} → {new_level} 조정")
                    line = '#' * new_level + ' ' + content

            adjusted_lines.append(line)

        text = '\n'.join(adjusted_lines)

        return AddFieldValidationResult(
            success=True,
            original_text=original,
            validated_text=text,
            changes_made=changes,
            warnings=warnings
        )

    def validate_input_content(
        self,
        text: str,
        expected_format: Optional[str] = None
    ) -> AddFieldValidationResult:
        """
        input_ 필드 내용 검증 (선택적)

        Args:
            text: 입력 텍스트
            expected_format: 예상 형식 ("date", "number", "text", None)

        Returns:
            AddFieldValidationResult: 검증 결과
        """
        if not text:
            return AddFieldValidationResult(
                success=True,
                original_text=text,
                validated_text=text,
                changes_made=[]
            )

        original = text
        changes = []
        warnings = []

        if expected_format == "date":
            text, date_changes, date_warnings = self._validate_date_format(text)
            changes.extend(date_changes)
            warnings.extend(date_warnings)

        elif expected_format == "number":
            text, num_changes, num_warnings = self._validate_number_format(text)
            changes.extend(num_changes)
            warnings.extend(num_warnings)

        return AddFieldValidationResult(
            success=len(warnings) == 0,
            original_text=original,
            validated_text=text,
            changes_made=changes,
            warnings=warnings
        )

    def validate_batch(
        self,
        data_list: List[Dict[str, str]],
        field_styles: Optional[Dict[str, str]] = None
    ) -> List[Dict[str, AddFieldValidationResult]]:
        """
        여러 데이터 행을 일괄 검증

        Args:
            data_list: [{field_name: value}, ...] 형태의 데이터
            field_styles: {field_name: style_type} 필드별 스타일 지정

        Returns:
            [{field_name: AddFieldValidationResult}, ...] 검증 결과
        """
        results = []
        field_styles = field_styles or {}

        for data in data_list:
            row_result = {}
            for field_name, value in data.items():
                if field_name.startswith('add_'):
                    style = field_styles.get(field_name, "plain")
                    row_result[field_name] = self.validate_add_content(value, style)
                elif field_name.startswith('input_'):
                    row_result[field_name] = self.validate_input_content(value)
                else:
                    # gstub_, stub_ 등은 검증 없이 통과
                    row_result[field_name] = AddFieldValidationResult(
                        success=True,
                        original_text=value,
                        validated_text=value,
                        changes_made=[]
                    )
            results.append(row_result)

        return results

    def _validate_with_sdk(self, text: str, content_type: str, style: Optional[str]) -> str:
        """Claude Code SDK를 사용한 검증"""
        if self.sdk_validator is None:
            return text

        return self.sdk_validator(text, content_type, style)

    def _format_as_bullet_list(self, text: str) -> Tuple[str, List[str]]:
        """bullet list 형식으로 변환"""
        changes = []
        lines = text.split('\n')
        formatted_lines = []

        for line in lines:
            line = line.strip()
            if not line:
                formatted_lines.append(line)
                continue

            # 이미 bullet이 있는지 확인
            if line.startswith(('- ', '• ', '* ', '· ', '■ ', '● ')):
                # 형식 통일 (기존 bullet 유지)
                pass
            else:
                # bullet 추가
                line = '- ' + line
                changes.append("bullet 추가")

            formatted_lines.append(line)

        return '\n'.join(formatted_lines), changes

    def _format_as_numbered_list(self, text: str) -> Tuple[str, List[str]]:
        """numbered list 형식으로 변환"""
        changes = []
        lines = text.split('\n')
        formatted_lines = []
        num = 1

        for line in lines:
            line = line.strip()
            if not line:
                formatted_lines.append(line)
                continue

            # 이미 번호가 있는지 확인
            num_match = re.match(r'^(\d+)[.)]\s*(.+)$', line)
            if num_match:
                content = num_match.group(2)
                line = f"{num}. {content}"
                if num_match.group(1) != str(num):
                    changes.append(f"번호 재정렬: {num_match.group(1)} → {num}")
            else:
                line = f"{num}. {line}"
                changes.append(f"번호 {num} 추가")

            formatted_lines.append(line)
            num += 1

        return '\n'.join(formatted_lines), changes

    def _format_as_heading(self, text: str, style_info: Optional[CellStyleInfo]) -> Tuple[str, List[str]]:
        """heading 형식으로 변환"""
        changes = []

        # 이미 헤딩 형식인지 확인
        if text.startswith('#'):
            return text, changes

        # 스타일 정보에서 레벨 추정
        level = 2  # 기본값
        if style_info and style_info.font_size:
            if style_info.font_size >= 20:
                level = 1
            elif style_info.font_size >= 16:
                level = 2
            else:
                level = 3

        text = '#' * level + ' ' + text
        changes.append(f"헤딩 레벨 {level} 적용")

        return text, changes

    def _format_as_plain(self, text: str, separator: str = " ") -> Tuple[str, List[str]]:
        """plain 형식 정리 (같은 문단 추가 시 구분자 사용)"""
        changes = []

        # 불필요한 공백 제거
        original_len = len(text)
        text = re.sub(r'[ \t]+', ' ', text)  # 연속 공백 → 단일 공백
        text = re.sub(r'\n{3,}', '\n\n', text)  # 연속 줄바꿈 → 최대 2개

        if len(text) != original_len:
            changes.append("불필요한 공백/줄바꿈 정리")

        return text.strip(), changes

    def _validate_date_format(self, text: str) -> Tuple[str, List[str], List[str]]:
        """날짜 형식 검증"""
        changes = []
        warnings = []

        # 다양한 날짜 패턴 인식
        patterns = [
            (r'(\d{4})[./\-](\d{1,2})[./\-](\d{1,2})', 'YYYY-MM-DD'),
            (r'(\d{1,2})[./\-](\d{1,2})[./\-](\d{4})', 'DD-MM-YYYY'),
            (r'(\d{4})년\s*(\d{1,2})월\s*(\d{1,2})일', 'YYYY년 MM월 DD일'),
        ]

        for pattern, format_name in patterns:
            match = re.match(pattern, text.strip())
            if match:
                return text, changes, warnings

        warnings.append(f"날짜 형식을 인식할 수 없음: {text}")
        return text, changes, warnings

    def _validate_number_format(self, text: str) -> Tuple[str, List[str], List[str]]:
        """숫자 형식 검증"""
        changes = []
        warnings = []

        text = text.strip()
        cleaned = text.replace(',', '').replace(' ', '')

        try:
            float(cleaned)
            if ',' in text:
                changes.append("천단위 구분자 유지")
            return text, changes, warnings
        except ValueError:
            warnings.append(f"숫자 형식이 아님: {text}")
            return text, changes, warnings


def create_sdk_validator(api_client: Any = None) -> Optional[Callable]:
    """
    Claude Code SDK 검증 함수 생성

    실제 SDK 클라이언트가 있으면 사용하고, 없으면 None 반환

    Args:
        api_client: Claude Code SDK API 클라이언트

    Returns:
        검증 함수 또는 None
    """
    if api_client is None:
        return None

    def validate(text: str, content_type: str, style: Optional[str]) -> str:
        """SDK를 사용한 실제 검증"""
        prompt = f"""
다음 텍스트를 {content_type} 형식으로 검증하고 필요시 수정해주세요.
스타일: {style or 'plain'}

텍스트:
{text}

수정된 텍스트만 반환해주세요. 설명은 필요없습니다.
"""
        response = api_client.complete(prompt)
        return response.strip()

    return validate


class FormatReviewAgent:
    """
    형식 검토 Sub Agent

    Claude Agent SDK와 함께 사용하기 위한 검토 에이전트 정의
    """

    AGENT_DEFINITION = {
        "name": "format-reviewer",
        "description": "HWPX 병합 데이터의 형식을 검토합니다. 캡션 스타일과 글머리 기호를 확인합니다.",
        "prompt": """당신은 HWPX 문서 형식 검토 전문가입니다.

다음 규칙에 따라 문서 형식을 검토하세요:

1. 캡션 스타일:
   - 테이블: "표 N. 제목" 형식
   - 그림: "그림 N. 제목" 형식

2. 개요 글머리 기호 (3단계):
   - 1단계: ■ (네모)
   - 2단계: ● (원)
   - 3단계: - (대시)

3. 검토 결과 형식:
   - 오류와 경고를 구분
   - 수정 제안 포함
   - JSON 형식으로 결과 반환

검토 후 수정이 필요한 항목과 수정 방법을 제안하세요.""",
        "tools": ["Read", "Grep", "Glob"],
    }

    @staticmethod
    def create_review_prompt(hwpx_path: str) -> str:
        """검토 요청 프롬프트 생성"""
        return f"""
format-reviewer 에이전트를 사용하여 다음 HWPX 파일을 검토하세요:
파일: {hwpx_path}

검토 항목:
1. 모든 테이블/그림 캡션이 표준 형식인지 확인
2. 개요 3단계의 글머리 기호가 올바른지 확인 (■ → ● → -)
3. 수정이 필요한 항목 목록 작성

결과를 JSON 형식으로 반환하세요.
"""


def validate_and_fix(hwpx_path: Union[str, Path], auto_fix: bool = False) -> Tuple[ValidationResult, List[Dict]]:
    """
    HWPX 파일 검증 및 수정

    Args:
        hwpx_path: HWPX 파일 경로
        auto_fix: 자동 수정 여부

    Returns:
        (검증 결과, 수정 내역)
    """
    validator = FormatValidator()
    result = validator.validate(hwpx_path)

    fixes = []
    if auto_fix and (result.errors or result.warnings):
        fixer = FormatFixer()
        # 자동 수정 로직 (outline tree 필요)
        # fixes = fixer.fix_bullets_in_tree(...)

    return result, fixes


def print_validation_result(result: ValidationResult):
    """검증 결과 출력"""
    print("=" * 60)
    print("형식 검증 결과")
    print("=" * 60)

    if result.is_valid and not result.errors and not result.warnings:
        print("✓ 모든 형식이 올바릅니다.")
        return

    if result.errors:
        print(f"\n오류 ({len(result.errors)}개):")
        for error in result.errors:
            print(f"  ✗ {error}")

    if result.warnings:
        print(f"\n경고 ({len(result.warnings)}개):")
        for warning in result.warnings:
            print(f"  ⚠ {warning}")

    if result.fixes:
        print(f"\n수정 필요 항목 ({len(result.fixes)}개):")
        for fix in result.fixes[:10]:  # 처음 10개만
            print(f"  - {fix}")


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("사용법: python format_validator.py <hwpx_file>")
        sys.exit(1)

    hwpx_path = sys.argv[1]
    print(f"파일: {hwpx_path}")

    result, fixes = validate_and_fix(hwpx_path)
    print_validation_result(result)
