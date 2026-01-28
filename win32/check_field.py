import zipfile
import xml.etree.ElementTree as ET

# test_insert_field.hwpx 에서 tc.name 확인
hwpx_path = r'C:\hwp_xml\test_insert_field.hwpx'
count = 0
with zipfile.ZipFile(hwpx_path, 'r') as zf:
    for name in zf.namelist():
        if 'section' in name and name.endswith('.xml'):
            with zf.open(name) as f:
                tree = ET.parse(f)
                root = tree.getroot()
                for tc in root.iter():
                    if tc.tag.endswith('}tc') and 'name' in tc.attrib:
                        print(f'tc.name: {tc.attrib["name"][:60]}...')
                        count += 1

print(f'\n총 {count}개 필드 발견')
