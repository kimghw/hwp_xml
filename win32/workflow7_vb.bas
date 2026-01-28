Option Explicit

Const COL_TBL_IDX As Integer = 1
Const COL_TABLE_ID As Integer = 2
Const COL_ROW As Integer = 3
Const COL_COL As Integer = 4
Const COL_ROW_SPAN As Integer = 5
Const COL_COL_SPAN As Integer = 6
Const COL_LIST_ID As Integer = 7
Const COL_PARA_IDX As Integer = 8
Const COL_FIELD_NAME As Integer = 9
Const COL_FIELD_TYPE As Integer = 10
Const COL_PARA_TEXT As Integer = 11

Sub Workflow7_ByFieldName()
    Dim hwp As Object
    Dim wsMeta As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim updateCount As Long
    Dim cellDict As Object
    Dim fieldDict As Object
    Dim tblIdx As Long
    Dim cellRow As Long
    Dim cellCol As Long
    Dim paraText As String
    Dim fieldName As String
    Dim cellKey As String
    Dim existingText As String
    Dim keys As Variant
    Dim combinedText As String
    Dim hwpPath As String
    Dim openParam As Object

    On Error Resume Next
    Set wsMeta = ThisWorkbook.Sheets("meta")
    On Error GoTo 0
    If wsMeta Is Nothing Then
        MsgBox "meta sheet not found.", vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    Set hwp = GetObject(, "HWPFrame.HwpObject")
    If hwp Is Nothing Then
        Set hwp = GetObject(, "Hwp.HwpObject")
    End If
    On Error GoTo 0

    If hwp Is Nothing Then
        On Error Resume Next
        Set hwp = CreateObject("HWPFrame.HwpObject")
        If hwp Is Nothing Then
            Set hwp = CreateObject("Hwp.HwpObject")
        End If
        On Error GoTo 0

        If hwp Is Nothing Then
            MsgBox "Cannot create HWP object. Please open HWP file first.", vbExclamation
            Exit Sub
        End If

        hwp.SetMessageBoxMode &H7FFFFFFF
        hwp.RegisterModule "FilePathCheckDLL", "SecurityModule"

        hwpPath = Application.GetOpenFilename("HWP Files (*.hwp;*.hwpx),*.hwp;*.hwpx", , "Select HWP file")
        If hwpPath = "False" Then
            MsgBox "No file selected.", vbExclamation
            Exit Sub
        End If

        Set openParam = hwp.HParameterSet.HFileOpenSave
        hwp.HAction.GetDefault "FileOpen", openParam.HSet
        openParam.FileName = hwpPath
        openParam.Format = "AUTO"
        hwp.HAction.Execute "FileOpen", openParam.HSet
    End If

    Set cellDict = CreateObject("Scripting.Dictionary")
    Set fieldDict = CreateObject("Scripting.Dictionary")

    lastRow = wsMeta.Cells(wsMeta.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        tblIdx = wsMeta.Cells(i, COL_TBL_IDX).Value
        cellRow = wsMeta.Cells(i, COL_ROW).Value
        cellCol = wsMeta.Cells(i, COL_COL).Value
        paraText = CStr(wsMeta.Cells(i, COL_PARA_TEXT).Value)
        fieldName = Trim(CStr(wsMeta.Cells(i, COL_FIELD_NAME).Value))

        cellKey = tblIdx & "_" & cellRow & "_" & cellCol

        If cellDict.Exists(cellKey) Then
            existingText = cellDict(cellKey)
            cellDict(cellKey) = existingText & vbLf & paraText
        Else
            cellDict.Add cellKey, paraText
        End If

        If fieldName <> "" And Not fieldDict.Exists(cellKey) Then
            fieldDict.Add cellKey, fieldName
        End If
    Next i

    updateCount = 0
    keys = fieldDict.keys

    For i = 0 To fieldDict.Count - 1
        cellKey = keys(i)
        fieldName = fieldDict(cellKey)
        combinedText = cellDict(cellKey)

        If hwp.FieldExist(fieldName) Then
            hwp.PutFieldText fieldName, combinedText
            updateCount = updateCount + 1
        End If
    Next i

    MsgBox "Updated: " & updateCount & " fields", vbInformation
End Sub
