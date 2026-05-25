# Option Explicit

' ============================================================
'  MAIN: รันตัวนี้ตัวเดียว → ทำครบทุกขั้นตอนอัตโนมัติ
'  ลำดับ: 1) แยกชีท  2) ลบคอลัมน์ S และ T  3) ปรับความกว้างคอลัมน์ A
' ============================================================
Sub RunAll()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' ขั้นตอนที่ 1: แยกชีทตามค่าคอลัมน์ S
    Call SplitToSheets_ByColS

    ' ขั้นตอนที่ 2: ลบคอลัมน์ S และ T ออกจากทุกชีท
    Call DeleteColumnsST_AllSheets

    ' ขั้นตอนที่ 3: ปรับความกว้างคอลัมน์ A ทุกชีทเป็น 9
    Call SetColumnAWidthAllSheets

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "เสร็จสิ้นทุกขั้นตอนเรียบร้อยแล้วครับ!" & vbCrLf & _
           "  ✔ แยกชีทตามคอลัมน์ S" & vbCrLf & _
           "  ✔ ลบคอลัมน์ S และ T ออกทุกชีท" & vbCrLf & _
           "  ✔ ปรับความกว้างคอลัมน์ A เป็น 9", vbInformation
End Sub

' ============================================================
'  ขั้นตอนที่ 1: แยก Sheet ตามค่าในคอลัมน์ S
' ============================================================
Private Sub SplitToSheets_ByColS()
    Dim wsSrc As Worksheet, wsNew As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim nameCol As String, nm As String
    Dim dict As Object, k As Variant
    Dim hdrRow As Long: hdrRow = 1       ' แถวหัวตาราง (แก้ได้ถ้าหัวอยู่แถวอื่น)
    nameCol = "S"                        ' คอลัมน์อ้างอิงในการแยกชีท

    Set wsSrc = ActiveSheet

    lastRow = wsSrc.Cells(wsSrc.Rows.Count, nameCol).End(xlUp).Row
    If lastRow <= hdrRow Then
        MsgBox "ไม่พบข้อมูลใต้หัวตาราง", vbExclamation
        Exit Sub
    End If

    lastCol = GetLastColSmart(wsSrc, hdrRow, lastRow)

    ' รวบรวมชื่อที่ไม่ซ้ำจากคอลัมน์ S
    Set dict = CreateObject("Scripting.Dictionary")
    Dim r As Long
    For r = hdrRow + 1 To lastRow
        nm = Trim$(CStr(wsSrc.Cells(r, nameCol).Value))
        If Len(nm) > 0 Then
            If Not dict.Exists(nm) Then dict.Add nm, 1
        End If
    Next r

    If dict.Count = 0 Then
        MsgBox "คอลัมน์ " & nameCol & " ไม่มีข้อมูล", vbExclamation
        Exit Sub
    End If

    ' เปิด AutoFilter บนช่วงข้อมูลจริง
    Dim filtRange As Range
    Set filtRange = wsSrc.Range(wsSrc.Cells(hdrRow, 1), wsSrc.Cells(lastRow, lastCol))
    If Not wsSrc.AutoFilterMode Then filtRange.AutoFilter

    For Each k In dict.Keys
        Dim safeName As String
        safeName = SanitizeSheetName(CStr(k))

        ' สร้างหรือล้างชีทปลายทาง
        If SheetExists(safeName) Then
            Set wsNew = ThisWorkbook.Worksheets(safeName)
            wsNew.Cells.Clear
        Else
            Set wsNew = ThisWorkbook.Worksheets.Add(After:=Worksheets(Worksheets.Count))
            On Error Resume Next
            wsNew.Name = safeName
            If Err.Number <> 0 Then
                Err.Clear
                wsNew.Name = Left$("Sheet_" & Format(Now, "hhmmss") & "_" & _
                             WorksheetFunction.RandBetween(100, 999), 31)
            End If
            On Error GoTo 0
        End If

        ' คัดลอกแถวหัวตาราง (พร้อมรูปแบบและสูตร)
        wsSrc.Rows(hdrRow).Copy
        wsNew.Rows(1).PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False

        ' กรองข้อมูลตามค่าในคอลัมน์ S
        filtRange.AutoFilter Field:=wsSrc.Range(nameCol & hdrRow).Column, Criteria1:=k

        ' คัดลอกเฉพาะแถวที่มองเห็น (ข้ามหัวตาราง)
        Dim vis As Range
        On Error Resume Next
        Set vis = wsSrc.Range(wsSrc.Cells(hdrRow + 1, 1), _
                              wsSrc.Cells(lastRow, lastCol)).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0

        If Not vis Is Nothing Then
            vis.Copy
            wsNew.Cells(2, 1).PasteSpecial Paste:=xlPasteAll
            Application.CutCopyMode = False
        End If
        Set vis = Nothing

        ' คัดลอกความกว้างคอลัมน์จากต้นฉบับ แล้ว AutoFit
        Dim c As Long
        For c = 1 To lastCol
            wsNew.Columns(c).ColumnWidth = wsSrc.Columns(c).ColumnWidth
        Next c
        wsNew.Columns.AutoFit
    Next k

    ' ยกเลิกตัวกรอง
    On Error Resume Next
    wsSrc.AutoFilter.ShowAllData
    On Error GoTo 0
End Sub

' ============================================================
'  ขั้นตอนที่ 2: ลบคอลัมน์ S และ T ออกจากทุกชีท
'  หมายเหตุ: ต้องลบ T ก่อน แล้วค่อยลบ S
'           เพราะถ้าลบ S ก่อน คอลัมน์ T จะเลื่อนมาเป็น S
' ============================================================
Private Sub DeleteColumnsST_AllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        ws.Columns("T").Delete     ' ลบคอลัมน์ T ก่อน
        ws.Columns("S").Delete     ' แล้วค่อยลบคอลัมน์ S
        On Error GoTo 0
    Next ws
End Sub

' ============================================================
'  ขั้นตอนที่ 3: ปรับความกว้างคอลัมน์ A ทุกชีทเป็น 9
' ============================================================
Private Sub SetColumnAWidthAllSheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        ws.Columns("A").ColumnWidth = 9
    Next ws
End Sub

' ============================================================
'  Helper: หาคอลัมน์สุดท้ายจากแถวข้อมูลจริง (ไม่ใช่หัวตาราง)
' ============================================================
Private Function GetLastColSmart(ws As Worksheet, hdrRow As Long, lastRow As Long) As Long
    Dim r As Long, c As Long, maxC As Long
    For r = hdrRow + 1 To Application.Min(hdrRow + 50, lastRow)
        c = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If c > maxC Then maxC = c
    Next r
    c = ws.UsedRange.Columns(ws.UsedRange.Columns.Count).Column
    If c > maxC Then maxC = c
    GetLastColSmart = maxC
End Function

' ============================================================
'  Helper: ทำความสะอาดชื่อชีท (ลบอักขระต้องห้าม)
' ============================================================
Private Function SanitizeSheetName(ByVal s As String) As String
    Dim badChars As Variant, ch As Variant
    badChars = Array("\", "/", "?", "*", "[", "]", ":", "'")
    For Each ch In badChars: s = Replace(s, ch, "_"): Next ch
    s = Trim$(s)
    If Len(s) = 0 Then s = "Unnamed"
    If Left$(s, 1) = "=" Then s = "_" & s
    If Len(s) > 31 Then s = Left$(s, 31)
    SanitizeSheetName = s
End Function

' ============================================================
'  Helper: ตรวจสอบว่าชีทมีอยู่แล้วหรือไม่
' ============================================================
Private Function SheetExists(ByVal s As String) As Boolean
    On Error Resume Next
    SheetExists = Not ThisWorkbook.Worksheets(s) Is Nothing
    On Error GoTo 0
End Function