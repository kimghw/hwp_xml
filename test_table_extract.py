# -*- coding: utf-8 -*-
"""
테이블 추출 테스트 스크립트

파일 탐색기로 HWP/HWPX 파일을 선택하여 테이블 추출을 테스트합니다.
HWP 파일 선택 시 자동으로 HWPX로 변환합니다.
"""

import os
import subprocess
import tempfile
from pathlib import Path

from core import open_hwp_dialog, windows_to_wsl_path
from hwpxml.get_cell_detail import GetCellDetail


def convert_hwp_to_hwpx_via_cmd(hwp_path: str) -> str:
    """WSL에서 Windows Python을 통해 HWP를 HWPX로 변환"""
    hwpx_path = str(Path(hwp_path).with_suffix(".hwpx"))

    # 프로젝트 루트 경로 기준
    project_root = Path(__file__).parent
    script_path = project_root / "_temp_convert.py"
    win_project_root = subprocess.run(
        ["wslpath", "-w", str(project_root)],
        capture_output=True, text=True
    ).stdout.strip()

    # win32/__init__.py를 우회하여 직접 모듈 import
    script_content = f'''# -*- coding: utf-8 -*-
import sys
import importlib.util
spec = importlib.util.spec_from_file_location("convert_hwp", r"{win_project_root}\\win32\\convert_hwp.py")
module = importlib.util.module_from_spec(spec)
spec.loader.exec_module(module)
result = module.convert_hwp_to_hwpx(r"{hwp_path}", r"{hwpx_path}")
print(result)
'''
    script_path.write_text(script_content, encoding='utf-8')

    try:
        result = subprocess.run(
            ["cmd.exe", "/c", f"cd /d {win_project_root} && python _temp_convert.py"],
            capture_output=True,
            timeout=120
        )

        # 인코딩 처리
        try:
            stderr = result.stderr.decode('utf-8')
        except UnicodeDecodeError:
            stderr = result.stderr.decode('cp949', errors='ignore')

        if result.returncode != 0:
            raise RuntimeError(f"변환 실패: {stderr}")

        return hwpx_path
    finally:
        if script_path.exists():
            script_path.unlink()


def main():
    print("=" * 60)
    print("테이블 추출 테스트")
    print("=" * 60)

    # 파일 선택 (HWP/HWPX 모두 지원)
    print("\nHWP 또는 HWPX 파일을 선택하세요...")
    win_path = open_hwp_dialog()

    if not win_path:
        print("파일 선택이 취소되었습니다.")
        return

    print(f"\n선택된 파일: {win_path}")

    # HWP인 경우 HWPX로 변환
    if win_path.lower().endswith(".hwp"):
        print("\nHWP 파일을 HWPX로 변환 중...")
        try:
            hwpx_win_path = convert_hwp_to_hwpx_via_cmd(win_path)
            print(f"변환 완료: {hwpx_win_path}")
            win_path = hwpx_win_path
        except Exception as e:
            print(f"변환 실패: {e}")
            return

    # WSL 경로로 변환
    wsl_path = windows_to_wsl_path(win_path)
    print(f"WSL 경로: {wsl_path}")

    # 테이블 추출
    print("\n" + "-" * 60)
    print("테이블 추출 중...")
    print("-" * 60)

    try:
        parser = GetCellDetail()
        tables = parser.from_hwpx_by_table(wsl_path)

        print(f"\n발견된 테이블 수: {len(tables)}")

        for i, table_cells in enumerate(tables):
            print(f"\n{'='*40}")
            print(f"테이블 {i + 1}: {len(table_cells)}개 셀")
            print("=" * 40)

            if not table_cells:
                print("  (빈 테이블)")
                continue

            # 행/열 범위 계산
            max_row = max(c.row for c in table_cells)
            max_col = max(c.col for c in table_cells)
            print(f"  크기: {max_row + 1}행 x {max_col + 1}열")

            # 처음 몇 개 셀 정보 출력
            print(f"\n  처음 5개 셀:")
            for j, cell in enumerate(table_cells[:5]):
                text_preview = cell.text[:20] + "..." if len(cell.text) > 20 else cell.text
                text_preview = text_preview.replace("\n", "\\n")
                print(f"    [{cell.row},{cell.col}] "
                      f"span=({cell.row_span}x{cell.col_span}) "
                      f"text=\"{text_preview}\"")

    except Exception as e:
        print(f"\n오류 발생: {e}")
        import traceback
        traceback.print_exc()

    print("\n" + "=" * 60)
    print("테스트 완료")


if __name__ == "__main__":
    main()
