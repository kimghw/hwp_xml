# -*- coding: utf-8 -*-
"""HWP 파일의 북마크(책갈피) 확인 - 다양한 방법 시도"""

import os
from hwp_utils import create_hwp_instance, open_file_dialog


def get_bookmarks_method1(hwp) -> list:
    """방법1: HeadCtrl 순회 - 북마크 개수만"""
    bookmarks = []
    ctrl = hwp.HeadCtrl
    while ctrl:
        try:
            if ctrl.CtrlID == 'bokm':
                bookmarks.append(ctrl)
        except:
            pass
        ctrl = ctrl.Next
    return bookmarks


def get_bookmarks_method2(hwp) -> list:
    """방법2: Bookmark 액션으로 목록 가져오기 시도"""
    try:
        # HBookMark 파라미터셋 사용
        hwp.HAction.GetDefault("Bookmark", hwp.HParameterSet.HBookMark.HSet)
        # Command=1: 목록 조회 모드?
        hwp.HParameterSet.HBookMark.Command = 1
        result = hwp.HAction.Execute("Bookmark", hwp.HParameterSet.HBookMark.HSet)
        print(f"  Bookmark Execute 결과: {result}")

        # HBookMark에서 이름 가져오기 시도
        for attr in ['Name', 'BookmarkName', 'Text', 'List']:
            try:
                val = getattr(hwp.HParameterSet.HBookMark, attr, None)
                if val:
                    print(f"  HBookMark.{attr}: {val}")
            except:
                pass
    except Exception as e:
        print(f"  방법2 오류: {e}")
    return []


def get_bookmarks_method3(hwp) -> list:
    """방법3: GetFieldList 사용"""
    try:
        # 북마크도 필드로 취급될 수 있음
        field_list = hwp.GetFieldList(0, 0)  # 모든 필드
        if field_list:
            print(f"  GetFieldList(0,0): {field_list}")

        field_list2 = hwp.GetFieldList(2, 0)  # 누름틀 필드만
        if field_list2:
            print(f"  GetFieldList(2,0): {field_list2}")
    except Exception as e:
        print(f"  방법3 오류: {e}")
    return []


def get_bookmarks_from_hwpx(hwpx_path) -> list:
    """HWPX XML에서 북마크 이름 추출"""
    import zipfile
    from xml.etree import ElementTree as ET

    bookmarks = []
    with zipfile.ZipFile(hwpx_path, 'r') as zf:
        for name in zf.namelist():
            if name.startswith('Contents/section') and name.endswith('.xml'):
                content = zf.read(name).decode('utf-8')
                root = ET.fromstring(content)
                for elem in root.iter():
                    if 'bookmark' in elem.tag.lower():
                        bm_name = elem.get('name', '')
                        if bm_name:
                            bookmarks.append(bm_name)
    return bookmarks


if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1:
        filepath = sys.argv[1]
    else:
        filepath = r'C:\hwp_xml\test_step5.hwp'

    if not os.path.exists(filepath):
        print(f'파일 없음: {filepath}')
        exit()

    print(f'파일: {filepath}')

    hwp = create_hwp_instance(visible=True)
    hwp.Open(filepath)

    print('\n=== 방법1: HeadCtrl 순회 ===')
    ctrls = get_bookmarks_method1(hwp)
    print(f'  북마크 컨트롤 개수: {len(ctrls)}개')

    print('\n=== 방법2: Bookmark 액션 ===')
    get_bookmarks_method2(hwp)

    print('\n=== 방법3: GetFieldList ===')
    get_bookmarks_method3(hwp)

    # HWPX로 변환
    print('\n=== HWPX 변환 후 XML에서 추출 ===')
    base = os.path.splitext(filepath)[0]
    hwpx_path = base + '_bm_test.hwpx'
    hwp.SaveAs(hwpx_path, "HWPX")

    bm_names = get_bookmarks_from_hwpx(hwpx_path)
    print(f'  북마크 {len(bm_names)}개:')
    for i, name in enumerate(bm_names, 1):
        print(f'    {i}. {name}')

    print('\n=== 결론 ===')
    print(f'  HWP COM API: 북마크 개수만 확인 가능 ({len(ctrls)}개)')
    print(f'  HWPX XML: 북마크 이름 추출 가능 ({len(bm_names)}개)')
    print(f'  -> HWPX 변환 시 북마크가 유지됨!')
