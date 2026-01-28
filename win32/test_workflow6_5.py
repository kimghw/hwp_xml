import sys
import os
sys.path.insert(0, r'C:\hwp_xml')
sys.path.insert(0, r'C:\hwp_xml\win32')
os.chdir(r'C:\hwp_xml')

# Workflow 6: 색상 기반 필드 설정
from win32.insert_field import convert_hwp_to_hwpx, process_hwpx_field

hwp_path = r'C:\hwp_xml\test.hwp'
print("=" * 60)
print("Workflow 6 실행")
print("=" * 60)

# HWP -> HWPX 변환 (문서 닫힘 대기)
hwpx_path = convert_hwp_to_hwpx(hwp_path)

# 필드 설정
process_hwpx_field(hwpx_path)

print("\n")
print("=" * 60)
print("Workflow 5 실행")
print("=" * 60)

from workflow.workflow5_integrated import Workflow5

workflow5 = Workflow5()
workflow5.run(hwp_path)
