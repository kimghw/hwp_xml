# -*- coding: utf-8 -*-
"""
HWPX 병합용 데이터 모델

개요:
- Paragraph: 문단 정보
- OutlineNode: 개요 노드 (트리 구조)
- HwpxData: HWPX 파일 데이터
"""

from dataclasses import dataclass, field
from typing import List, Dict, Set, Any


@dataclass
class Paragraph:
    """문단 정보"""
    index: int = 0
    is_outline: bool = False
    level: int = -1
    text: str = ""
    para_pr_id: str = ""
    element: Any = None
    source_file: str = ""  # 원본 파일

    # 테이블/그림 정보
    has_table: bool = False
    has_image: bool = False
    table_count: int = 0
    image_count: int = 0


@dataclass
class OutlineNode:
    """개요 노드 (트리 구조)"""
    level: int = 0
    name: str = ""
    paragraphs: List[Paragraph] = field(default_factory=list)
    children: List['OutlineNode'] = field(default_factory=list)

    def get_content_paragraphs(self) -> List[Paragraph]:
        """개요 자체를 제외한 내용 문단만 반환"""
        if not self.paragraphs:
            return []
        return self.paragraphs[1:] if len(self.paragraphs) > 1 else []

    def get_all_paragraphs(self) -> List[Paragraph]:
        """이 노드와 모든 하위 노드의 문단 반환 (순서대로)"""
        result = list(self.paragraphs)
        for child in self.children:
            result.extend(child.get_all_paragraphs())
        return result


@dataclass
class HwpxData:
    """HWPX 파일 데이터"""
    path: str
    paragraphs: List[Paragraph] = field(default_factory=list)
    outline_tree: List[OutlineNode] = field(default_factory=list)

    # XML 내용
    header_xml: bytes = b''
    section_xmls: Dict[str, bytes] = field(default_factory=dict)

    # 바이너리 데이터 (이미지 등)
    bin_data: Dict[str, bytes] = field(default_factory=dict)

    # 기타 파일
    other_files: Dict[str, bytes] = field(default_factory=dict)

    # 스타일 정의
    outline_para_ids: Set[str] = field(default_factory=set)
    outline_levels: Dict[str, int] = field(default_factory=dict)

    # ZIP 파일 정보 (external_attr 등 보존용)
    zip_infos: Dict[str, Any] = field(default_factory=dict)
