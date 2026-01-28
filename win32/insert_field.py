# -*- coding: utf-8 -*-
"""색상 기반 셀 필드 자동 설정

사용법:
1. HWP 파일 선택 시:
   - 원본.hwp → 원본_insert_field.hwpx 변환
   - 한글에서 파일 열림 → 색상 블럭 작업
   - 문서 닫으면 자동으로 필드 설정 진행

2. _insert_field.hwpx 파일 선택 시:
   - tc.name에 필드명 설정
   - YAML 출력 (data/원본/)
   - 원본.hwp, 원본_insert_field.hwpx 유지
"""

import sys
import os
import time
import tempfile
import zipfile
import shutil
import yaml
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


def is_yellow_color(color: str) -> bool:
    """노란색 계열인지 확인"""
    if not color:
        return False

    color_lower = color.lower().strip().lstrip('#')

    yellow_colors = ['ffff00', 'ffff00ff', 'fff000', 'fff000ff']
    if color_lower in yellow_colors:
        return True

    if len(color_lower) >= 6:
        try:
            r = int(color_lower[0:2], 16)
            g = int(color_lower[2:4], 16)
            b = int(color_lower[4:6], 16)
            # 노란색: R과 G가 높고, B가 낮음
            if r > 200 and g > 200 and b < 100:
                return True
        except:
            pass
    return False


def clear_tc_names_in_hwpx(hwpx_path: str) -> int:
    """HWPX에서 모든 tc.name 속성 삭제

    Returns:
        삭제된 필드 수
    """
    hwpx_path = Path(hwpx_path)
    extract_dir = tempfile.mkdtemp()
    total_cleared = 0

    try:
        with zipfile.ZipFile(str(hwpx_path), 'r') as zf:
            zf.extractall(extract_dir)

        contents_dir = os.path.join(extract_dir, 'Contents')
        section_files = sorted([
            f for f in os.listdir(contents_dir)
            if f.startswith('section') and f.endswith('.xml')
        ])

        for section_file in section_files:
            section_path = os.path.join(contents_dir, section_file)
            tree = ET.parse(section_path)
            root = tree.getroot()

            for tc in root.iter():
                if tc.tag.endswith('}tc'):
                    if 'name' in tc.attrib:
                        del tc.attrib['name']
                        total_cleared += 1

            tree.write(section_path, encoding='utf-8', xml_declaration=True)

        # 수정된 HWPX 다시 압축
        with zipfile.ZipFile(str(hwpx_path), 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, extract_dir)
                    zf.write(file_path, arcname)

    finally:
        shutil.rmtree(extract_dir, ignore_errors=True)

    return total_cleared


def field_clear(file_path: str):
    """파일의 모든 tc.name(필드) 삭제

    - HWP: HWPX 변환 → tc.name 삭제 → HWP 덮어쓰기 → HWPX 삭제
    - HWPX: tc.name만 삭제
    """
    file_path = Path(file_path)
    ext = file_path.suffix.lower()

    print(f"입력: {file_path}")

    if ext == '.hwp':
        # HWP → HWPX 변환
        hwp = create_hwp_instance(visible=False)

        hwp.HAction.GetDefault("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = str(file_path)
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HAction.Execute("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)

        # 임시 HWPX
        temp_hwpx = file_path.parent / f"{file_path.stem}_temp_clear.hwpx"

        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = str(temp_hwpx)
        hwp.HParameterSet.HFileOpenSave.Format = "HWPX"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

        hwp.Clear(1)  # 문서 닫기
        print(f"HWPX 변환: {temp_hwpx}")

        # tc.name 삭제
        cleared = clear_tc_names_in_hwpx(str(temp_hwpx))
        print(f"삭제된 필드: {cleared}개")

        # HWPX → HWP 덮어쓰기
        hwp.HAction.GetDefault("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = str(temp_hwpx)
        hwp.HParameterSet.HFileOpenSave.Format = "HWPX"
        hwp.HAction.Execute("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)

        hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
        hwp.HParameterSet.HFileOpenSave.filename = str(file_path)
        hwp.HParameterSet.HFileOpenSave.Format = "HWP"
        hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

        hwp.Quit()
        print(f"HWP 저장: {file_path}")

        # 임시 HWPX 삭제
        temp_hwpx.unlink()
        print(f"임시 파일 삭제: {temp_hwpx}")

    elif ext == '.hwpx':
        # HWPX: tc.name만 삭제
        cleared = clear_tc_names_in_hwpx(str(file_path))
        print(f"삭제된 필드: {cleared}개")
        print(f"HWPX 저장: {file_path}")

    else:
        print(f"지원하지 않는 파일 형식: {ext}")
        return

    print("완료!")


def convert_hwp_to_hwpx(hwp_path: str) -> str:
    """HWP 파일을 _insert_field.hwpx로 변환 후 열기, 문서 닫힘 대기"""
    hwp_path = Path(hwp_path)
    hwpx_path = hwp_path.parent / f"{hwp_path.stem}_insert_field.hwpx"

    print(f"입력: {hwp_path}")
    print(f"출력: {hwpx_path}")

    hwp = create_hwp_instance(visible=True)

    # 파일 열기
    hwp.HAction.GetDefault("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = str(hwp_path)
    hwp.HParameterSet.HFileOpenSave.Format = "HWP"
    hwp.HAction.Execute("FileOpen", hwp.HParameterSet.HFileOpenSave.HSet)

    # HWPX로 저장
    hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
    hwp.HParameterSet.HFileOpenSave.filename = str(hwpx_path)
    hwp.HParameterSet.HFileOpenSave.Format = "HWPX"
    hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)

    print(f"HWPX 변환 완료: {hwpx_path}")
    print()
    print("색상 블럭 작업 후 저장하고 문서를 닫으세요...")
    print("(문서가 닫히면 자동으로 필드 설정이 진행됩니다)")

    # 문서 닫힘 대기 (폴링 방식)
    while True:
        try:
            doc_count = hwp.XHwpDocuments.Count
            if doc_count == 0:
                break
            time.sleep(1)
        except:
            # hwp 객체 접근 실패 = 한글 종료됨
            break

    print()
    print("문서 닫힘 감지, 필드 설정 진행...")
    return str(hwpx_path)


def process_hwpx_field(hwpx_path: str):
    """HWPX 파일에서 색상 기반 필드 설정"""
    hwpx_path = Path(hwpx_path)
    if not hwpx_path.exists():
        print(f"파일이 없습니다: {hwpx_path}")
        return

    # 원본 파일명 추출 (_insert_field 제거)
    stem = hwpx_path.stem
    if stem.endswith('_insert_field'):
        original_stem = stem[:-len('_insert_field')]
    else:
        original_stem = stem

    print(f"입력: {hwpx_path}")
    print()

    # 테이블 파싱
    parser = GetCellDetail()
    all_tables = parser.from_hwpx_by_table(str(hwpx_path))

    if not all_tables:
        print("테이블이 없습니다")
        return

    # 테이블별 셀 맵 생성 (병합 정보 포함)
    table_cells = {}
    table_info = {}  # 테이블 정보 (list_id, table_id)
    for tbl_idx, cells in enumerate(all_tables):
        table_cells[tbl_idx] = {}
        # 첫 번째 셀에서 테이블 정보 가져오기
        if cells:
            table_info[tbl_idx] = {
                'list_id': cells[0].list_id,
                'table_id': cells[0].table_id
            }
        for cell in cells:
            text = ' '.join([p.text for p in cell.paragraphs]).strip()
            bg_color = cell.border.bg_color if cell.border else ''
            table_cells[tbl_idx][(cell.row, cell.col)] = {
                'text': text,
                'bg_color': bg_color,
                'row_span': cell.row_span,
                'col_span': cell.col_span,
                'list_id': cell.list_id,
                'table_id': cell.table_id
            }

    def find_cell_at(cell_map, row, col):
        """해당 위치의 셀 찾기 (병합된 셀 포함)
        Returns: (info, start_row, start_col) - 셀 정보와 시작 위치
        """
        # 정확히 해당 위치에 셀이 있으면 반환
        if (row, col) in cell_map:
            return cell_map[(row, col)], row, col

        # 병합된 셀 찾기: 이전 행/열에서 span이 이 위치를 포함하는 셀 찾기
        for (r, c), info in cell_map.items():
            row_span = info.get('row_span', 1)
            col_span = info.get('col_span', 1)
            if r <= row < r + row_span and c <= col < c + col_span:
                return info, r, c
        return {}, -1, -1

    # HWPX 압축 해제
    extract_dir = tempfile.mkdtemp()
    set_count = 0
    field_results = []  # 필드 설정 결과 저장

    try:
        with zipfile.ZipFile(str(hwpx_path), 'r') as zf:
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
                            cell_info, _, _ = find_cell_at(cell_map, row, col)
                            if not cell_info:
                                cell_info = {}
                            bg_color = cell_info.get('bg_color', '')
                            cell_text = cell_info.get('text', '').strip()

                            # 노란색 셀: 셀 텍스트를 필드명으로 사용 (20자 제한)
                            if is_yellow_color(bg_color):
                                if cell_text:
                                    field_name = cell_text[:20]
                                    tc.set('name', field_name)
                                    set_count += 1
                                    modified = True
                                    print(f"  테이블{current_tbl_idx} ({row},{col}) -> [{field_name}]")
                                    # 결과 저장
                                    tbl_info = table_info.get(current_tbl_idx, {})
                                    field_results.append({
                                        'table_idx': current_tbl_idx,
                                        'list_id': tbl_info.get('list_id', ''),
                                        'table_id': tbl_info.get('table_id', ''),
                                        'row': row,
                                        'col': col,
                                        'field_name': field_name,
                                        'type': 'yellow'
                                    })
                                continue

                            # 빨간색 배경이 아니면 스킵
                            if not is_red_color(bg_color):
                                continue

                            # 텍스트가 있으면 스킵 (빈 셀에서만 필드 설정)
                            if cell_text:
                                continue

                            # 왼쪽으로 이동해서 최대 3개 텍스트 찾기 (빨간색 범위 내에서만)
                            left_texts = []
                            c = col - 1
                            while c >= 0 and len(left_texts) < 3:
                                info, start_r, start_c = find_cell_at(cell_map, row, c)
                                # 빨간색 셀이 아니면 탐색 중단
                                if not is_red_color(info.get('bg_color', '')):
                                    break
                                t = info.get('text', '').strip()
                                if t:
                                    left_texts.append(t)
                                # 병합 셀의 시작 열로 점프 (다음 반복에서 그 왼쪽으로)
                                c = start_c - 1 if start_c >= 0 else c - 1

                            # 위쪽으로 이동해서 최대 3개 텍스트 찾기 (빨간색 범위 내에서만)
                            top_texts = []
                            r = row - 1
                            while r >= 0 and len(top_texts) < 3:
                                info, start_r, start_c = find_cell_at(cell_map, r, col)
                                # 빨간색 셀이 아니면 탐색 중단
                                if not is_red_color(info.get('bg_color', '')):
                                    break
                                t = info.get('text', '').strip()
                                if t:
                                    top_texts.append(t)
                                # 병합 셀의 시작 행으로 점프 (다음 반복에서 그 위쪽으로)
                                r = start_r - 1 if start_r >= 0 else r - 1

                            # 필드명 생성: [L:좌1][L:좌2][T:위1][T:위2]
                            parts = []
                            # 왼쪽: L: 접두사
                            for t in left_texts:
                                parts.append('[L:' + t + ']')
                            # 위쪽: T: 접두사
                            for t in top_texts:
                                parts.append('[T:' + t + ']')

                            field_name = ''.join(parts)

                            if field_name:
                                tc.set('name', field_name)
                                set_count += 1
                                modified = True
                                print(f"  테이블{current_tbl_idx} ({row},{col}) -> [{field_name}]")
                                # 결과 저장
                                tbl_info = table_info.get(current_tbl_idx, {})
                                field_results.append({
                                    'table_idx': current_tbl_idx,
                                    'list_id': tbl_info.get('list_id', ''),
                                    'table_id': tbl_info.get('table_id', ''),
                                    'row': row,
                                    'col': col,
                                    'field_name': field_name,
                                    'type': 'red'
                                })

                    current_tbl_idx += 1

            # 수정된 XML 저장
            if modified:
                tree.write(section_path, encoding='utf-8', xml_declaration=True)

        # 수정된 HWPX 다시 압축
        with zipfile.ZipFile(str(hwpx_path), 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_dir, dirs, files in os.walk(extract_dir):
                for file in files:
                    file_path = os.path.join(root_dir, file)
                    arcname = os.path.relpath(file_path, extract_dir)
                    zf.write(file_path, arcname)

        print()
        print(f"설정된 필드: {set_count}개")
        print(f"HWPX 저장 완료: {hwpx_path}")

        # YAML 파일 출력 (data/원본파일명/ 폴더에 저장)
        if field_results:
            # data 폴더 생성 (원본 파일명 기준)
            data_dir = hwpx_path.parent / 'data' / original_stem
            data_dir.mkdir(parents=True, exist_ok=True)

            yaml_path = data_dir / f"{original_stem}_field.yaml"
            with open(yaml_path, 'w', encoding='utf-8') as f:
                yaml.dump(field_results, f, allow_unicode=True, default_flow_style=False, sort_keys=False)
            print(f"YAML 저장: {yaml_path}")

    finally:
        shutil.rmtree(extract_dir, ignore_errors=True)


if __name__ == "__main__":
    from hwp_file_manager import open_file_dialog

    if len(sys.argv) > 1:
        file_path = sys.argv[1]
    else:
        file_path = open_file_dialog("HWP/HWPX 파일 선택")
        if not file_path:
            print("파일을 선택하지 않았습니다.")
            sys.exit(1)

    file_path = Path(file_path)
    ext = file_path.suffix.lower()

    if ext == '.hwp':
        # HWP 파일: HWPX로 변환 → 문서 닫힘 대기 → 필드 설정
        hwpx_path = convert_hwp_to_hwpx(str(file_path))
        process_hwpx_field(hwpx_path)
    elif ext == '.hwpx':
        # HWPX 파일: 필드 설정 처리
        process_hwpx_field(str(file_path))
    else:
        print(f"지원하지 않는 파일 형식: {ext}")
        sys.exit(1)
