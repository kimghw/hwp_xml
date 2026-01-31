# -*- coding: utf-8 -*-
"""
HWPX 파일 파싱 모듈

개요:
- HwpxParser: HWPX 파일 전체 파싱
"""

import zipfile
import xml.etree.ElementTree as ET
from typing import List, Union, Tuple
from pathlib import Path
from io import BytesIO
import copy

from .models import Paragraph, HwpxData
from .outline import build_outline_tree

# XML 네임스페이스
NAMESPACES = {
    'hp': 'http://www.hancom.co.kr/hwpml/2011/paragraph',
    'hs': 'http://www.hancom.co.kr/hwpml/2011/section',
    'hc': 'http://www.hancom.co.kr/hwpml/2011/core',
    'hh': 'http://www.hancom.co.kr/hwpml/2011/head',
}

# 네임스페이스 등록
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)


class HwpxParser:
    """HWPX 파일 파싱"""

    def parse(self, hwpx_path: Union[str, Path]) -> HwpxData:
        """HWPX 파일 전체 파싱"""
        hwpx_path = Path(hwpx_path)
        data = HwpxData(path=str(hwpx_path))

        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            # 모든 파일 읽기
            for name in zf.namelist():
                content = zf.read(name)

                if name == 'Contents/header.xml':
                    data.header_xml = content
                    self._parse_header(content, data)
                elif name.startswith('Contents/section') and name.endswith('.xml'):
                    data.section_xmls[name] = content
                elif name.startswith('BinData/'):
                    data.bin_data[name] = content
                else:
                    data.other_files[name] = content

            # 문단 파싱
            para_idx = 0
            for section_name in sorted(data.section_xmls.keys()):
                section_content = data.section_xmls[section_name]
                section_paras = self._parse_section(section_content, para_idx, data, str(hwpx_path))
                para_idx += len(section_paras)
                data.paragraphs.extend(section_paras)

        # 개요 트리 생성
        data.outline_tree = build_outline_tree(data.paragraphs)

        return data

    def _parse_header(self, xml_content: bytes, data: HwpxData):
        """header.xml에서 개요 스타일 파싱"""
        root = ET.parse(BytesIO(xml_content)).getroot()

        for para_pr in root.iter():
            if para_pr.tag.endswith('}paraPr'):
                para_pr_id = para_pr.get('id', '')
                if not para_pr_id:
                    continue

                for child in para_pr:
                    if child.tag.endswith('}heading'):
                        if child.get('type') == 'OUTLINE':
                            level = int(child.get('level', '0'))
                            data.outline_para_ids.add(para_pr_id)
                            data.outline_levels[para_pr_id] = level
                        break

    def _parse_section(self, xml_content: bytes, start_idx: int, data: HwpxData, source_file: str) -> List[Paragraph]:
        """section XML에서 문단 파싱 (테이블 내부 문단 제외)"""
        paragraphs = []
        root = ET.parse(BytesIO(xml_content)).getroot()

        para_idx = start_idx

        # 섹션의 직접 자식 p 요소만 파싱 (테이블 내부 문단 제외)
        for elem in root:
            if not elem.tag.endswith('}p'):
                continue

            para_pr_id = elem.get('paraPrIDRef', '')
            is_outline = para_pr_id in data.outline_para_ids
            level = data.outline_levels.get(para_pr_id, -1)

            text = self._extract_text(elem)

            # 테이블/그림 정보 추출
            table_count, image_count = self._count_table_image(elem)

            para = Paragraph(
                index=para_idx,
                is_outline=is_outline,
                level=level,
                text=text,
                para_pr_id=para_pr_id,
                element=copy.deepcopy(elem),
                source_file=source_file,
                has_table=table_count > 0,
                has_image=image_count > 0,
                table_count=table_count,
                image_count=image_count,
            )
            paragraphs.append(para)
            para_idx += 1

        return paragraphs

    def _extract_text(self, p_elem) -> str:
        """문단에서 텍스트 추출"""
        texts = []
        for run in p_elem:
            if run.tag.endswith('}run'):
                for t in run:
                    if t.tag.endswith('}t') and t.text:
                        texts.append(t.text)
        return ''.join(texts)

    def _count_table_image(self, p_elem) -> Tuple[int, int]:
        """문단 내 테이블/그림 개수 카운트"""
        table_count = 0
        image_count = 0

        for child in p_elem.iter():
            if child.tag.endswith('}tbl'):
                table_count += 1
            elif child.tag.endswith('}pic'):
                image_count += 1

        return table_count, image_count
