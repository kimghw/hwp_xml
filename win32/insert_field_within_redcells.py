# -*- coding: utf-8 -*-
"""빨간색 배경 셀에 필드 이름 설정 (위쪽 텍스트 기준)"""

import sys
import os
import tempfile
import zipfile
import shutil
import xml.etree.ElementTree as ET
from pathlib import Path

_project_root = Path(__file__).parent.parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

try:
    from config import WIN32HWP_DIR
    if str(WIN32HWP_DIR) not in sys.path:
        sys.path.insert(0, str(WIN32HWP_DIR))
except ImportError:
    win32hwp_dir = os.environ.get('WIN32HWP_DIR', r'C:\win32hwp')
    if win32hwp_dir not in sys.path:
        sys.path.insert(0, win32hwp_dir)

from hwpxml.get_cell_detail import GetCellDetail

try:
    from hwp_file_manager import create_hwp_instance
except ImportError:
    from win32.hwp_file_manager import create_hwp_instance


def is_red_color(color: str) -> bool:
    """빨간색 계열인지 확인"""
    if not color:
        return False

    color_lower = color.lower().strip().lstrip('#')

    red_colors = ['ff0000', 'cf2741', 'ff0000ff', 'cf2741ff']
    if color_lower in red_colors:
        return True

    if len(color_lower) >= 6:
        try:
            r = int(color_lower[0:2], 16)
            g = int(color_lower[2:4], 16)
            b = int(color_lower[4:6], 16)
            if r > 180 and g < 80 and b < 80:
                return True
        except:
            pass
    return False


def set_red_field(hwp_path: str, output_path: str = None):
    """빨간색 배경 셀에 필드 이름 설정"""
    hwp_path = Path(hwp_path)
    if not hwp_path.exists():
        print(f"파일이 없습니다: {hwp_path}")
        return

    if output_path is None:
        output_path = hwp_path

    print(f"입력: {hwp_path}")
    print(f"출력: {output_path}")
    print()

    # 한글 실행 (숨김)
    hwp = create_hwp_instance(visible=False)

    # 파일 열기 (HAction 방식 - 보안 팝업 방지)
    hwp.HAction.GetDefault("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = str(hwp_path)
    hwp.HParameterSet.HFileOpenSave.Format = "HWP"
    hwp.HAction.Execute("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)

    # 임시 HWPX로 저장
    temp_dir = tempfile.gettempdir()
    hwpx_path = os.path.join(temp_dir, "set_red_field_temp.hwpx")

    hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = hwpx_path
    hwp.HParameterSet.HFileOpenSave.Format = "HWPX"
    hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
    print(f"HWPX 변환 완료")

    # 테이블 파싱
    parser = GetCellDetail()
    all_tables = parser.from_hwpx_by_table(hwpx_path)

    if not all_tables:
        print("테이블이 없습니다")
        hwp.Quit()
        return

    # 테이블별 셀 맵 생성 (병합 정보 포함)
    table_cells = {}
    for tbl_idx, cells in enumerate(all_tables):
        table_cells[tbl_idx] = {}
        for cell in cells:
            text = ' '.join([p.text for p in cell.paragraphs]).strip()
            bg_color = cell.border.bg_color if cell.border else ''
            table_cells[tbl_idx][(cell.row, cell.col)] = {
                'text': text,
                'bg_color': bg_color,
                'row_span': cell.row_span,
                'col_span': cell.col_span
            }

    def find_cell_at(cell_map, row, col):
        """해당 위치의 셀 찾기 (병합된 셀 포함)"""
        # 정확히 해당 위치에 셀이 있으면 반환
        if (row, col) in cell_map:
            return cell_map[(row, col)]

        # 병합된 셀 찾기: 이전 행/열에서 span이 이 위치를 포함하는 셀 찾기
        for (r, c), info in cell_map.items():
            row_span = info.get('row_span', 1)
            col_span = info.get('col_span', 1)
            if r <= row < r + row_span and c <= col < c + col_span:
                return info
        return {}

    # HWPX 압축 해제
    extract_dir = tempfile.mkdtemp()
    set_count = 0

    try:
        with zipfile.ZipFile(hwpx_path, 'r') as zf:
            zf.extractall(extract_dir)

        contents_dir = os.path.join(extract_dir, 'Contents')
        section_files = sorted([
            f for f in os.listdir(contents_dir)
            if f.startswith('section') and f.endswith('.xml')
        ])

        current_tbl_idx = 0

        for section_file in section_files:
            section_path = os.path.join(contents_dir, section_file)
            tree = ET.parse(section_path)
            root = tree.getroot()

            modified = False

            # 테이블 찾기
            for tbl in root.iter():
                if tbl.tag.endswith('}tbl'):
                    if current_tbl_idx not in table_cells:
                        current_tbl_idx += 1
                        continue

                    cell_map = table_cells[current_tbl_idx]

                    # 이 테이블의 셀들 처리
                    for tr in tbl:
                        if not tr.tag.endswith('}tr'):
                            continue
                        for tc in tr:
                            if not tc.tag.endswith('}tc'):
                                continue

                            # 셀 주소 가져오기
                            row, col = -1, -1
                            for child in tc:
                                if child.tag.endswith('}cellAddr'):
                                    row = int(child.get('rowAddr', -1))
                                    col = int(child.get('colAddr', -1))
                                    break

                            if row < 0 or col < 0:
                                continue

                            # 셀 정보 가져오기
                            cell_info = cell_map.get((row, col), {})
                            bg_color = cell_info.get('bg_color', '')

                            # 빨간색 배경이 아니면 스킵
                            if not is_red_color(bg_color):
                                continue

                            # 텍스트가 있으면 스킵 (빈 셀에서만 필드 설정)
                            cell_text = cell_info.get('text', '').strip()
                            if cell_text:
                                continue

                            # 왼쪽으로 이동해서 최대 3개 텍스트 찾기 (병합 셀 고려)
                            left_texts = []
                            for c in range(col - 1, -1, -1):
                                info = find_cell_at(cell_map, row, c)
                                t = info.get('text', '').strip()
                                if t:
                                    left_texts.append(t)
                                    if len(left_texts) >= 3:
                                        break

                            # 위쪽으로 이동해서 최대 3개 텍스트 찾기 (병합 셀 고려)
                            top_texts = []
                            for r in range(row - 1, -1, -1):
                                info = find_cell_at(cell_map, r, col)
                                t = info.get('text', '').strip()
                                if t:
                                    top_texts.append(t)
                                    if len(top_texts) >= 3:
                                        break

                            # 필드명 생성: [좌1][좌2][좌3][위1][위2][위3]
                            parts = []
                            # 왼쪽: 가까운 순서대로 (역순 불필요)
                            for t in left_texts:
                                parts.append('[' + t + ']')
                            # 위쪽: 가까운 순서대로
                            for t in top_texts:
                                parts.append('[' + t + ']')

                            field_name = ''.join(parts)

                            if field_name:
                                tc.set('name', field_name)
                                set_count += 1
                                modified = True
                                print(f"  테이블{current_tbl_idx} ({row},{col}) -> [{field_name}]")

                    current_tbl_idx += 1

            # 수정된 XML 저장
            if modified:
                tree.write(section_path, encoding='utf-8', xml_declaration=True)

        # 수정된 HWPX 다시 압축
        with zipfile.ZipFile(hwpx_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, extract_dir)
                    zf.write(file_path, arcname)

        print()
        print(f"설정된 필드: {set_count}개")

        # HWP로 변환하여 저장 (HAction 방식)
        hwp.HAction.GetDefault("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = hwpx_path
        hwp.HParameterSet.HFileOpenSave.Format = "HWPX"
        hwp.HAction.Execute("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = str(output_path)
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        print(f"저장 완료: {output_path}")

    finally:
        shutil.rmtree(extract_dir, ignore_errors=True)
        hwp.Quit()


if __name__ == "__main__":
    hwp_path = r"C:\hwp_xml\test.hwp"
    output_path = r"C:\hwp_xml\test_field.hwp"

    if len(sys.argv) > 1:
        hwp_path = sys.argv[1]
    if len(sys.argv) > 2:
        output_path = sys.argv[2]

    set_red_field(hwp_path, output_path)
