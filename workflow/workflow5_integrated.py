# -*- coding: utf-8 -*-
"""
Workflow 5: 북마크별 Excel 생성

HWP 파일에서 북마크가 HWPX로 변환되지 않는 경우를 처리:
1. HWP에서 북마크 확인
2. 북마크가 없으면 → 일반 workflow4 진행
3. 북마크가 있으면 → 마커 삽입 후 HWPX 변환 → 북마크별 시트 분리

프로세스:
1. HWP 열기
2. 북마크 목록 확인
3. 북마크 위치에 마커 텍스트 삽입 (HWPX에서 인식용)
4. HWPX 임시 변환
5. workflow1 실행 → `{파일}_meta.yaml`
6. workflow2 실행 → `{파일}_para.yaml`
7. 북마크별 시트 분리 + 문단 분할 → `{파일}_by_bookmark.xlsx`
8. 원본 HWP에서 마커 삭제
9. 임시 파일 삭제

출력 파일:
- {base}_meta.yaml  : 테이블/셀 메타데이터
- {base}_para.yaml  : 문단 스타일 정보
- {base}_by_bookmark.xlsx : 북마크별 시트 분리 Excel
"""

# ============================================================
# 설정
# ============================================================
DEFAULT_HWP_PATH = None  # 스크립트와 동일한 이름의 .hwp 파일 (main에서 설정)
BOOKMARK_MARKER_PREFIX = "{{BOOKMARK:"
BOOKMARK_MARKER_SUFFIX = "}}"
# ============================================================

import sys
import os
import tempfile
from pathlib import Path

# 프로젝트 루트 경로 설정
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

from win32.hwp_utils import get_hwp_instance, create_hwp_instance, get_active_filepath, open_file_dialog, save_hwp


class Workflow5:
    """북마크별 Excel 생성 워크플로우"""

    def __init__(self):
        self.hwp = None
        self.hwp_created = False
        self.filepath = None
        self.temp_hwpx = None
        self.bookmarks = []  # 북마크 목록
        self.markers_inserted = False  # 마커 삽입 여부
        # 메타데이터 저장용
        self.cell_positions = None
        self.field_names = None
        self.existing_fields = []  # 기존 필드 목록

    def _get_hwp(self):
        """한글 인스턴스 가져오기"""
        self.hwp = get_hwp_instance()
        if self.hwp:
            print("기존 한글 인스턴스 연결")
            self.hwp_created = False
            try:
                self.hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
            except:
                pass
        else:
            print("새 한글 인스턴스 생성")
            self.hwp = create_hwp_instance(visible=True)
            self.hwp_created = True
        return self.hwp

    def _open_file(self, filepath: str = None) -> str:
        """HWP 파일 열기"""
        if filepath:
            self.filepath = filepath
        else:
            self.filepath = get_active_filepath(self.hwp)
            if self.filepath:
                print(f"열린 문서 사용: {self.filepath}")
                return self.filepath

            print("파일을 선택하세요...")
            self.filepath = open_file_dialog()
            if not self.filepath:
                raise ValueError("파일이 선택되지 않았습니다.")

        if not get_active_filepath(self.hwp) == self.filepath:
            self.hwp.Open(self.filepath)
        print(f"파일 열림: {self.filepath}")
        return self.filepath

    def _get_bookmarks(self) -> int:
        """HWP에서 북마크 개수 확인 (HeadCtrl 순회 방식)"""
        bookmark_count = 0

        # HeadCtrl 순회로 북마크 개수 확인
        ctrl = self.hwp.HeadCtrl
        while ctrl:
            if ctrl.CtrlID == 'bokm':
                bookmark_count += 1
            ctrl = ctrl.Next

        print(f"  발견된 북마크: {bookmark_count}개")
        self.bookmark_count = bookmark_count

        return bookmark_count

    def _insert_bookmark_markers(self):
        """북마크 위치에 마커 텍스트 삽입"""
        if not self.bookmarks:
            return

        print("\n북마크 마커 삽입 중...")
        inserted_count = 0

        for bm_name in self.bookmarks:
            try:
                # 북마크로 이동
                self.hwp.MoveToBookmark(bm_name, False, False)

                # 마커 텍스트 삽입
                marker = f"{BOOKMARK_MARKER_PREFIX}{bm_name}{BOOKMARK_MARKER_SUFFIX}"
                self.hwp.HAction.Run("BreakPara")  # 새 문단
                self.hwp.HAction.GetDefault("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                self.hwp.HParameterSet.HInsertText.Text = marker
                self.hwp.HAction.Execute("InsertText", self.hwp.HParameterSet.HInsertText.HSet)
                self.hwp.HAction.Run("BreakPara")  # 새 문단

                inserted_count += 1
            except Exception as e:
                print(f"  마커 삽입 실패 ({bm_name}): {e}")

        self.markers_inserted = inserted_count > 0
        print(f"  {inserted_count}개 마커 삽입됨")

    def _remove_bookmark_markers(self):
        """원본 HWP에서 마커 텍스트 삭제"""
        if not self.markers_inserted:
            return

        print("\n원본에서 마커 삭제 중...")

        # 원본 파일 다시 열기
        self.hwp.Open(self.filepath)

        # 마커 텍스트는 삽입하지 않았으므로 삭제 불필요
        # (HWPX 변환 전에만 마커 삽입하고, 원본은 건드리지 않음)

        print("  마커 삭제 완료")

    def _save_as_hwpx(self) -> str:
        """임시 HWPX로 저장"""
        temp_dir = tempfile.gettempdir()
        self.temp_hwpx = os.path.join(temp_dir, "workflow5_temp.hwpx")

        if os.path.exists(self.temp_hwpx):
            try:
                os.remove(self.temp_hwpx)
            except:
                pass

        self.hwp.HAction.GetDefault("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)
        self.hwp.HParameterSet.HFileOpenSave.filename = self.temp_hwpx
        self.hwp.HParameterSet.HFileOpenSave.Format = "HWPX"
        self.hwp.HAction.Execute("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)

        # 문서 닫고 파일 잠금 해제
        self.hwp.Clear(1)

        print(f"임시 HWPX 저장: {self.temp_hwpx}")
        return self.temp_hwpx

    def _extract_existing_fields(self, base_path: str) -> str:
        """기존 셀 필드 이름 추출 및 삭제"""
        print("\n" + "-" * 60)
        print("기존 셀 필드 이름 추출 중...")

        from win32.extract_field import ExtractField

        extractor = ExtractField(self.hwp)
        # 테이블 셀에 설정된 필드 이름 추출
        self.existing_fields = extractor.extract_cell_field_names()

        if not self.existing_fields:
            print("  기존 셀 필드 없음")
            return None

        # YAML 저장
        field_yaml = base_path + "_field.yaml"
        extractor.save_yaml(field_yaml)

        # 필드 삭제
        print("기존 필드 삭제 중...")
        extractor.delete_all_fields()

        return field_yaml

    def _run_workflow1(self, base_path: str) -> str:
        """Workflow 1: 테이블 메타데이터 추출"""
        print("\n" + "=" * 60)
        print("Workflow 1: 테이블 메타데이터 추출")
        print("=" * 60)

        from win32.insert_table_field import InsertTableField

        inserter = InsertTableField(self.hwp)

        print("필드명 삽입 중...")
        count = inserter.insert_field_to_xml(self.temp_hwpx)
        print(f"  {count}개 테이블 처리")

        if count == 0:
            print("  테이블 없음, 건너뜀")
            return None

        self.hwp.Open(self.temp_hwpx)

        print("캡션 삽입 중...")
        table_infos = inserter.collect_table_list_ids()
        caption_count = inserter.insert_caption_text(table_infos)
        print(f"  {caption_count}개 캡션 삽입")

        from win32.extract_cell_meta import ExtractCellMeta

        meta_yaml = base_path + "_meta.yaml"
        extractor = ExtractCellMeta(self.hwp)

        self.hwp.HAction.GetDefault("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)
        self.hwp.HParameterSet.HFileOpenSave.filename = self.temp_hwpx
        self.hwp.HParameterSet.HFileOpenSave.Format = "HWPX"
        self.hwp.HAction.Execute("FileSaveAs_S", self.hwp.HParameterSet.HFileOpenSave.HSet)

        self.cell_positions = extractor._extract_cell_positions()
        self.field_names = extractor._extract_field_names_from_hwpx(self.temp_hwpx)
        yaml_content = extractor._merge_to_yaml(self.cell_positions, self.field_names)

        with open(meta_yaml, 'w', encoding='utf-8') as f:
            f.write(yaml_content)

        print(f"메타데이터 저장: {meta_yaml}")

        self._clear_field_names_in_hwpx(self.temp_hwpx)

        return meta_yaml

    def _clear_field_names_in_hwpx(self, hwpx_path: str):
        """HWPX에서 tc.name 속성만 삭제"""
        import zipfile
        import shutil
        import xml.etree.ElementTree as ET

        extract_dir = tempfile.mkdtemp()
        total_cleared = 0

        try:
            with zipfile.ZipFile(hwpx_path, 'r') as zf:
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

            with zipfile.ZipFile(hwpx_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for root_dir, dirs, files in os.walk(extract_dir):
                    for file in files:
                        file_path = os.path.join(root_dir, file)
                        arcname = os.path.relpath(file_path, extract_dir)
                        zf.write(file_path, arcname)

            print(f"필드명 삭제: {total_cleared}개 셀")

        finally:
            shutil.rmtree(extract_dir, ignore_errors=True)

    def _run_bookmark_excel(self, base_path: str, split_by_para: bool = True) -> str:
        """북마크별 시트 분리 Excel 생성"""
        print("\n" + "=" * 60)
        print("북마크별 Excel 변환")
        print("=" * 60)

        from excel.hwpx_to_excel import HwpxToExcel

        excel_path = base_path + "_by_bookmark.xlsx"
        converter = HwpxToExcel()

        # 북마크별 시트 분리 변환
        converter.convert_all_by_bookmark(
            self.temp_hwpx,
            excel_path,
            include_body=True,
            split_by_para=split_by_para,
            include_cell_info=False
        )

        print(f"Excel 저장: {excel_path}")
        return excel_path

    def _cleanup(self):
        """임시 파일 삭제"""
        print("\n" + "-" * 60)
        print("정리 중...")

        if self.temp_hwpx and os.path.exists(self.temp_hwpx):
            try:
                os.remove(self.temp_hwpx)
                print(f"  삭제: {self.temp_hwpx}")
            except:
                pass

    def _should_close(self) -> bool:
        """종료 여부 판단"""
        return self.hwp_created

    def run(self, filepath: str = None, split_by_para: bool = True) -> dict:
        """
        워크플로우 실행

        Args:
            filepath: HWP 파일 경로
            split_by_para: 문단별 행 분할 여부

        Returns:
            생성된 파일 경로 dict
        """
        print("=" * 60)
        print("Workflow 5: 북마크별 Excel 생성")
        print("=" * 60)

        results = {
            'field_yaml': None,
            'meta_yaml': None,
            'excel': None,
            'has_bookmarks': False,
        }

        try:
            # 1. 한글 인스턴스
            self._get_hwp()

            # 2. 파일 열기
            self._open_file(filepath)

            base_path = os.path.splitext(self.filepath)[0]

            # 3. HWP에서 북마크 개수 확인 (HeadCtrl 순회)
            bookmark_count = self._get_bookmarks()
            results['has_bookmarks'] = bookmark_count > 0

            if bookmark_count == 0:
                print("\n북마크가 없습니다. 일반 변환으로 진행합니다.")
                return results

            # 4. 기존 필드 추출 및 삭제
            results['field_yaml'] = self._extract_existing_fields(base_path)

            # 5. HWPX 변환
            self._save_as_hwpx()

            # 6. 원본 HWP 복원
            self.hwp.Open(self.filepath)

            # 7. Workflow 1: 메타데이터 추출
            results['meta_yaml'] = self._run_workflow1(base_path)

            # 8. 북마크별 Excel 변환
            results['excel'] = self._run_bookmark_excel(base_path, split_by_para)

            # 11. 문서 닫고 정리
            self.hwp.Clear(1)  # temp HWPX 파일 잠금 해제
            self._cleanup()

            # 12. 완료
            print("\n" + "=" * 60)
            print("완료!")
            print("=" * 60)
            print(f"  북마크 수: {len(self.bookmarks)}개")
            if results['field_yaml']:
                print(f"  기존필드:   {results['field_yaml']}")
            print(f"  메타데이터: {results['meta_yaml']}")
            print(f"  Excel:     {results['excel']}")

            if self._should_close():
                try:
                    self.hwp.Quit()
                    print("한글 종료됨")
                except:
                    pass

        except Exception as e:
            print(f"\n오류: {e}")
            import traceback
            traceback.print_exc()

        return results


def main():
    """메인 함수"""
    filepath = None

    if len(sys.argv) > 1:
        filepath = sys.argv[1]
        if not os.path.isabs(filepath):
            filepath = os.path.abspath(filepath)
    else:
        # 스크립트와 동일한 이름의 .hwp 파일을 기본 경로로
        script_hwp = str(Path(__file__).with_suffix('.hwp'))
        if os.path.exists(script_hwp):
            filepath = script_hwp

    if filepath and not os.path.exists(filepath):
        print(f"파일을 찾을 수 없습니다: {filepath}")
        filepath = None

    workflow = Workflow5()
    workflow.run(filepath)


if __name__ == "__main__":
    main()
