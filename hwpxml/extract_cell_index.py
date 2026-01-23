"""
HWPX 테이블 셀에서 [index:##{숫자}] 패턴 추출

셀 내용에서 인덱스 마커를 찾아 매핑 정보를 생성합니다.
"""

import re
import json
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any, Union
from pathlib import Path

try:
    # 패키지로 import 시
    from .get_table_property import GetTableProperty, TableProperty, CellProperty
except ImportError:
    # 직접 실행 시
    from get_table_property import GetTableProperty, TableProperty, CellProperty


@dataclass
class CellIndexInfo:
    """셀 인덱스 정보"""
    document_name: str          # 문서 파일명
    table_id: str               # 테이블 ID
    row_index: int              # 셀 행 좌표
    col_index: int              # 셀 열 좌표
    index_number: int           # [index:##{숫자}]에서 추출한 숫자
    full_match: str             # 전체 매칭 문자열 (예: [index:##1])
    cell_text: str              # 셀 전체 텍스트

    def to_dict(self) -> Dict[str, Any]:
        return {
            'document_name': self.document_name,
            'table_id': self.table_id,
            'row_index': self.row_index,
            'col_index': self.col_index,
            'index_number': self.index_number,
            'full_match': self.full_match,
            'cell_text': self.cell_text,
        }


@dataclass
class DocumentIndexMap:
    """문서 전체의 인덱스 매핑 정보"""
    document_name: str
    document_path: str
    indexes: List[CellIndexInfo] = field(default_factory=list)

    # 빠른 조회용 딕셔너리
    _by_index: Dict[int, CellIndexInfo] = field(default_factory=dict, repr=False)
    _by_cell: Dict[tuple, CellIndexInfo] = field(default_factory=dict, repr=False)

    def add(self, info: CellIndexInfo):
        """인덱스 정보 추가"""
        self.indexes.append(info)
        self._by_index[info.index_number] = info
        self._by_cell[(info.table_id, info.row_index, info.col_index)] = info

    def get_by_index(self, index_number: int) -> Optional[CellIndexInfo]:
        """인덱스 번호로 조회"""
        return self._by_index.get(index_number)

    def get_by_cell(self, table_id: str, row: int, col: int) -> Optional[CellIndexInfo]:
        """셀 좌표로 조회"""
        return self._by_cell.get((table_id, row, col))

    def get_all_indexes(self) -> List[int]:
        """모든 인덱스 번호 반환 (정렬됨)"""
        return sorted(self._by_index.keys())

    def to_dict(self) -> Dict[str, Any]:
        return {
            'document_name': self.document_name,
            'document_path': self.document_path,
            'count': len(self.indexes),
            'indexes': [info.to_dict() for info in self.indexes],
        }

    def to_simple_dict(self) -> Dict[str, Dict[int, Dict[str, int]]]:
        """
        간단한 딕셔너리 형태로 변환

        Returns:
            {table_id: {index_number: {"row": row, "col": col}, ...}, ...}
        """
        output = {}
        for info in self.indexes:
            table_id = info.table_id
            if table_id not in output:
                output[table_id] = {}
            output[table_id][info.index_number] = {
                'row': info.row_index,
                'col': info.col_index
            }
        return output

    def to_json(self, indent: int = 2) -> str:
        """JSON 문자열로 변환"""
        return json.dumps(self.to_simple_dict(), indent=indent, ensure_ascii=False)

    def save_json(self, output_path: Union[str, Path] = None) -> Path:
        """
        JSON 파일로 저장

        Args:
            output_path: 저장 경로 (None이면 문서명_index.json)

        Returns:
            저장된 파일 경로
        """
        if output_path is None:
            # 기본 파일명: 문서명_index.json
            base_name = Path(self.document_name).stem
            output_path = Path(self.document_path).parent / f"{base_name}_index.json"
        else:
            output_path = Path(output_path)

        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(self.to_json())

        return output_path


class ExtractCellIndex:
    """
    HWPX 파일에서 [index:##{숫자}] 패턴 추출

    사용 예:
        extractor = ExtractCellIndex()
        result = extractor.from_hwpx("document.hwpx")

        # 인덱스 번호로 셀 찾기
        cell_info = result.get_by_index(1)

        # 전체 인덱스 목록
        all_indexes = result.get_all_indexes()
    """

    # [index:##{숫자}] 패턴 정규식
    INDEX_PATTERN = re.compile(r'\[index:##(\d+)\]')

    def __init__(self, pattern: str = None):
        """
        초기화

        Args:
            pattern: 사용자 정의 정규식 패턴 (기본: [index:##{숫자}])
        """
        if pattern:
            self.pattern = re.compile(pattern)
        else:
            self.pattern = self.INDEX_PATTERN

        self.table_parser = GetTableProperty()

    def extract_from_cell(self, cell: CellProperty, document_name: str,
                          table_id: str) -> List[CellIndexInfo]:
        """단일 셀에서 인덱스 추출"""
        results = []

        if not cell.text:
            return results

        matches = self.pattern.finditer(cell.text)
        for match in matches:
            index_number = int(match.group(1))
            info = CellIndexInfo(
                document_name=document_name,
                table_id=table_id,
                row_index=cell.row_index,
                col_index=cell.col_index,
                index_number=index_number,
                full_match=match.group(0),
                cell_text=cell.text,
            )
            results.append(info)

        return results

    def extract_from_table(self, table: TableProperty,
                           document_name: str) -> List[CellIndexInfo]:
        """테이블에서 모든 인덱스 추출"""
        results = []

        for row in table.cells:
            for cell in row:
                cell_indexes = self.extract_from_cell(
                    cell, document_name, table.id
                )
                results.extend(cell_indexes)

        return results

    def from_hwpx(self, hwpx_path: Union[str, Path],
                  save_json: bool = False,
                  json_path: Union[str, Path] = None) -> DocumentIndexMap:
        """
        HWPX 파일에서 모든 인덱스 추출

        Args:
            hwpx_path: HWPX 파일 경로
            save_json: True면 JSON 파일로 저장
            json_path: JSON 저장 경로 (None이면 자동 생성)

        Returns:
            DocumentIndexMap 객체
        """
        hwpx_path = Path(hwpx_path)
        document_name = hwpx_path.name

        result = DocumentIndexMap(
            document_name=document_name,
            document_path=str(hwpx_path),
        )

        # 테이블 파싱
        tables = self.table_parser.from_hwpx(hwpx_path)

        for table in tables:
            indexes = self.extract_from_table(table, document_name)
            for info in indexes:
                result.add(info)

        # JSON 저장
        if save_json:
            saved_path = result.save_json(json_path)
            print(f"JSON 저장: {saved_path}")

        return result

    def from_xml_string(self, xml_string: Union[str, bytes],
                        document_name: str = "unknown") -> DocumentIndexMap:
        """XML 문자열에서 인덱스 추출"""
        result = DocumentIndexMap(
            document_name=document_name,
            document_path="",
        )

        tables = self.table_parser.from_xml_string(xml_string)

        for table in tables:
            indexes = self.extract_from_table(table, document_name)
            for info in indexes:
                result.add(info)

        return result


# 편의 함수
def extract_indexes_from_hwpx(hwpx_path: Union[str, Path],
                               save_json: bool = False,
                               json_path: Union[str, Path] = None) -> DocumentIndexMap:
    """
    HWPX 파일에서 인덱스 추출 (편의 함수)

    Args:
        hwpx_path: HWPX 파일 경로
        save_json: True면 JSON 파일로 저장
        json_path: JSON 저장 경로

    Returns:
        DocumentIndexMap 객체
    """
    extractor = ExtractCellIndex()
    return extractor.from_hwpx(hwpx_path, save_json=save_json, json_path=json_path)


def get_index_mapping(hwpx_path: Union[str, Path],
                      save_json: bool = False,
                      json_path: Union[str, Path] = None) -> Dict[str, Dict]:
    """
    인덱스 번호 → 셀 좌표 딕셔너리 반환

    Args:
        hwpx_path: HWPX 파일 경로
        save_json: True면 JSON 파일로 저장
        json_path: JSON 저장 경로

    Returns:
        {table_id: {index: {"row": row, "col": col}, ...}, ...}
    """
    result = extract_indexes_from_hwpx(hwpx_path, save_json=save_json, json_path=json_path)
    return result.to_simple_dict()


# 테스트
if __name__ == "__main__":
    import sys

    if len(sys.argv) > 1:
        hwpx_path = sys.argv[1]
    else:
        hwpx_path = "/mnt/c/hwp_xml/table.hwpx"

    print(f"파일: {hwpx_path}")
    print("=" * 50)

    extractor = ExtractCellIndex()
    result = extractor.from_hwpx(hwpx_path)

    print(f"문서명: {result.document_name}")
    print(f"인덱스 개수: {len(result.indexes)}")

    if result.indexes:
        print("\n인덱스 목록:")
        for info in result.indexes:
            print(f"  [{info.index_number}] 테이블:{info.table_id} "
                  f"셀:({info.row_index},{info.col_index}) "
                  f"매칭:{info.full_match}")
    else:
        print("\n인덱스를 찾을 수 없습니다.")
