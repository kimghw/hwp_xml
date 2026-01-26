# -*- coding: utf-8 -*-
import win32com.client as win32
import logging

# 로깅 설정
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# 테스트할 모듈 목록
modules = [
    ("FilePathCheckerModuleExample", "파일 경로 접근만 허용"),
    ("SecurityModule", "모든 보안 경고 자동 허용"),
]

for module_name, desc in modules:
    logging.info(f"\n{'='*50}")
    logging.info(f"테스트: {module_name} ({desc})")
    logging.info(f"{'='*50}")

    try:
        hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        hwp.RegisterModule("FilePathCheckDLL", module_name)
        hwp.Open(r"C:\hwp_xml\test_step5.hwp", "HWP", "forceopen:true")
        hwp.XHwpWindows.Item(0).Visible = True

        # 북마크 찾기
        ctrl = hwp.HeadCtrl
        bookmark_count = 0
        bookmark_names = []

        while ctrl:
            try:
                if ctrl.CtrlID == 'bokm':
                    bookmark_count += 1
                    props = ctrl.Properties
                    name = props.Item('Name') if props else None
                    bookmark_names.append(name)
            except:
                pass
            ctrl = ctrl.Next

        logging.info(f"북마크 개수: {bookmark_count}")
        logging.info(f"북마크 이름들: {bookmark_names[:3]}...")  # 처음 3개만

        hwp.Clear(1)
        hwp.Quit()

    except Exception as e:
        logging.error(f"Error: {e}")
