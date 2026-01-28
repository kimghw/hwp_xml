' ============================================================
' Workflow 7: Excel → HWP 필드값 반영
' ============================================================
' meta 시트 구조:
' tbl_idx(1), table_id(2), row(3), col(4), row_span(5), col_span(6),
' list_id(7), para_idx(8), para_text(9), field_name(10), field_type(11)
'
' 처리 방식:
' 1. meta 시트에서 같은 셀(tbl_idx, row, col)의 문단들을 합침
' 2. 빈 문단은 빈 줄(vbLf)로 유지
' 3. field_name으로 HWP에 값 입력 (PutFieldText)
'
' 사용법:
' 1. workflow5로 생성된 Excel 파일 열기
' 2. meta 시트에서 para_text 수정
' 3. 이 매크로 실행
' ============================================================

Option Explicit

Sub Workflow7_ExcelToHwp()
    Dim hwp As Object
    Dim wsMeta As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim updateCount As Long
    Dim hwpPath As String

    ' meta 시트 확인
    On Error Resume Next
    Set wsMeta = ThisWorkbook.Sheets("meta")
    On Error GoTo 0

    If wsMeta Is Nothing Then
        MsgBox "meta 시트가 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 한글 연결 (열린 인스턴스 또는 새로 생성)
    On Error Resume Next
    Set hwp = GetObject(, "HWPFrame.HwpObject")
    On Error GoTo 0

    If hwp Is Nothing Then
        Set hwp = CreateObject("HWPFrame.HwpObject")
        hwp.SetMessageBoxMode &H7FFFFFFF
        hwp.RegisterModule "FilePathCheckerModuleExample", "FilePathCheckerModule"
        hwp.RegisterModule "FilePathCheckDLL", "SecurityModule"

        hwpPath = SelectHwpFile()
        If hwpPath = "" Then
            MsgBox "파일을 선택하지 않았습니다.", vbExclamation
            Exit Sub
        End If
        OpenHwpFile hwp, hwpPath
    End If

    ' ============================================================
    ' 1단계: meta 시트에서 셀별 문단 합치기
    ' ============================================================
    ' 같은 (tbl_idx, row, col)의 문단들을 para_idx 순서대로 합침
    ' 빈 문단(para_text = "")은 빈 줄로 유지

    Dim cellDict As Object   ' cellKey → 합쳐진 텍스트
    Dim fieldDict As Object  ' cellKey → field_name
    Set cellDict = CreateObject("Scripting.Dictionary")
    Set fieldDict = CreateObject("Scripting.Dictionary")

    lastRow = wsMeta.Cells(wsMeta.Rows.Count, 1).End(xlUp).Row

    ' meta 시트 컬럼 인덱스
    Const COL_TBL_IDX As Integer = 1
    Const COL_TABLE_ID As Integer = 2
    Const COL_ROW As Integer = 3
    Const COL_COL As Integer = 4
    Const COL_PARA_IDX As Integer = 8
    Const COL_PARA_TEXT As Integer = 9
    Const COL_FIELD_NAME As Integer = 10

    Dim tblIdx As Long, cellRow As Long, cellCol As Long
    Dim paraIdx As Long, paraText As String, fieldName As String
    Dim cellKey As String
    Dim existingText As String

    For i = 2 To lastRow
        tblIdx = wsMeta.Cells(i, COL_TBL_IDX).Value
        cellRow = wsMeta.Cells(i, COL_ROW).Value
        cellCol = wsMeta.Cells(i, COL_COL).Value
        paraIdx = wsMeta.Cells(i, COL_PARA_IDX).Value
        paraText = CStr(wsMeta.Cells(i, COL_PARA_TEXT).Value)
        fieldName = Trim(CStr(wsMeta.Cells(i, COL_FIELD_NAME).Value))

        cellKey = tblIdx & "_" & cellRow & "_" & cellCol

        ' 문단 합치기 (빈 문단도 빈 줄로 유지)
        If cellDict.Exists(cellKey) Then
            existingText = cellDict(cellKey)
            cellDict(cellKey) = existingText & vbLf & paraText
        Else
            cellDict.Add cellKey, paraText
        End If

        ' field_name 저장 (첫 번째만)
        If fieldName <> "" And Not fieldDict.Exists(cellKey) Then
            fieldDict.Add cellKey, fieldName
        End If
    Next i

    ' ============================================================
    ' 2단계: field_name으로 HWP에 값 입력
    ' ============================================================
    updateCount = 0

    Dim keys As Variant
    Dim combinedText As String

    keys = fieldDict.keys

    For i = 0 To fieldDict.Count - 1
        cellKey = keys(i)
        fieldName = fieldDict(cellKey)
        combinedText = cellDict(cellKey)

        ' 필드 존재 여부 확인 후 값 입력
        If hwp.FieldExist(fieldName) Then
            hwp.PutFieldText fieldName, combinedText
            updateCount = updateCount + 1
        End If
    Next i

    MsgBox updateCount & "개 필드 업데이트 완료 (총 " & fieldDict.Count & "개 필드)", vbInformation

End Sub

' ============================================================
' HWP 파일 선택 대화상자
' ============================================================
Function SelectHwpFile() As String
    Dim fd As Object
    Set fd = Application.FileDialog(msoFileDialogFilePicker)

    With fd
        .Title = "HWP 파일 선택"
        .Filters.Clear
        .Filters.Add "한글 파일", "*.hwp;*.hwpx"
        .AllowMultiSelect = False

        If .Show = -1 Then
            SelectHwpFile = .SelectedItems(1)
        Else
            SelectHwpFile = ""
        End If
    End With
End Function

' ============================================================
' HWP 파일 열기 (보안 팝업 방지)
' ============================================================
Sub OpenHwpFile(hwp As Object, filePath As String)
    Dim pset As Object
    Set pset = hwp.HParameterSet.HFileOpenSave

    hwp.HAction.GetDefault "FileOpen", pset.HSet
    pset.filename = filePath
    pset.Format = "AUTO"
    hwp.HAction.Execute "FileOpen", pset.HSet
End Sub

' ============================================================
' 필드 목록 확인 (디버그용)
' ============================================================
Sub CheckFieldList()
    Dim hwp As Object
    Dim fieldList As String

    On Error Resume Next
    Set hwp = GetObject(, "HWPFrame.HwpObject")
    On Error GoTo 0

    If hwp Is Nothing Then
        MsgBox "열린 한글 문서가 없습니다.", vbExclamation
        Exit Sub
    End If

    ' 필드 목록 조회
    fieldList = hwp.GetFieldList(0, 1)  ' 0: 현재 문서, 1: 중복 제거

    If fieldList = "" Then
        MsgBox "문서에 필드가 없습니다.", vbInformation
    Else
        MsgBox "필드 목록:" & vbCrLf & Replace(fieldList, Chr(2), vbCrLf), vbInformation
    End If
End Sub

' ============================================================
' 필드값 미리보기 (디버그용)
' ============================================================
Sub PreviewFieldValues()
    Dim hwp As Object
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim fieldName As String
    Dim currentValue As String
    Dim msg As String

    On Error Resume Next
    Set hwp = GetObject(, "HWPFrame.HwpObject")
    Set ws = ThisWorkbook.Sheets("meta")
    On Error GoTo 0

    If hwp Is Nothing Or ws Is Nothing Then
        MsgBox "한글 문서 또는 meta 시트가 없습니다.", vbExclamation
        Exit Sub
    End If

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    msg = "필드명 -> 현재값:" & vbCrLf & vbCrLf

    For i = 2 To Application.Min(lastRow, 20)  ' 최대 20개만
        fieldName = Trim(CStr(ws.Cells(i, 10).Value))
        If fieldName <> "" Then
            If hwp.FieldExist(fieldName) Then
                currentValue = hwp.GetFieldText(fieldName)
                msg = msg & fieldName & " -> " & currentValue & vbCrLf
            End If
        End If
    Next i

    MsgBox msg, vbInformation
End Sub
