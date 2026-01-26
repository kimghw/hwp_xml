# -*- coding: utf-8 -*-
"""
win32 모듈 - 한글 COM API 연동

Windows 환경에서 한글 프로그램과 COM API를 통해 상호작용
(Windows 전용)
"""

# Windows 환경에서만 import 가능
import platform

if platform.system() == "Windows":
    from .get_table_property import (
        GetTableProperty,
        TableProperty,
        CellInfo,
        CtrlType,
        get_tables_from_file,
        get_table_data_as_list,
    )
    from .get_table_info import get_all_table_info

    __all__ = [
        'GetTableProperty',
        'TableProperty',
        'CellInfo',
        'CtrlType',
        'get_tables_from_file',
        'get_table_data_as_list',
        'get_all_table_info',
    ]
else:
    __all__ = []
