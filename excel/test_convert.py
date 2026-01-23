# -*- coding: utf-8 -*-
"""셀 정보 시트 포함 변환 테스트"""

import sys
from pathlib import Path

_project_root = Path(__file__).parent.parent
sys.path.insert(0, str(_project_root))

from hwpx_to_excel import convert_hwpx_to_excel

hwpx_path = r"C:\hwp_xml\table.hwpx"
output_path = r"C:\hwp_xml\excel\output_with_cellinfo.xlsx"

print(f"입력: {hwpx_path}")
print(f"출력: {output_path}")

result = convert_hwpx_to_excel(
    hwpx_path,
    output_path,
    include_cell_info=True,
    hide_para_rows=True
)

print(f"완료: {result}")
