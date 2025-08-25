Attribute VB_Name = "modConst"

Option Explicit
Public Const PRODUCT_VERSION As String = "0.1.1"

' ========= Sheet & Column Names =========
Public Const SHEET_SRC   As String = "Compounding"                ' ????
Public Const SHEET_ALLOC As String = "CompoundingAllocation"      ' ?????
Public Const SHEET_SUM   As String = "CompoundingBatchSummary"    ' ??????
Public Const SHEET_CFG   As String = "CompoundingTab"             ' ???
Public Const APPLY_LOSS_FACTOR As Boolean = False   ' ???? 1.07;??????? => False

' ========= Config Cells (可按需调整落位) =========
Public Const CELL_HORIZON_DAYS As String = "I1"         ' 计划窗（天）
Public Const CELL_EFFCAP_PER_BATCH As String = "I2"     ' 每批可用能力（t）
Public Const CELL_WINDOW_DAYS As String = "I3"          ' 每批窗口天数

' ========= Defaults =========
Public Const DEF_HORIZON_DAYS As Long = 60
Public Const DEF_EFFCAP_PER_BATCH As Double = 37.5
Public Const DEF_WINDOW_DAYS As Long = 21

' ========= Conversion Factors =========
Public Const FACTOR_LOSS As Double = 1.07
Public Const FACTOR_UNIT As Double = 10.4 ' unit to kg factor inside formula chain

' ========= Other =========
' modConst.bas ????:

' ====== Config cells on CompoundingTab ======
Public Const CELL_RUN_DATE    As String = "I4"   ' ??:??????
Public Const CELL_SRC_SHEET   As String = "I5"   ' ??:?????
' ??????

Public Const CELL_LEAD_DAYS As String = "I6"   ' CompoundingTab!I6
Public Const DEF_LEAD_DAYS  As Long = 1       ' ???? 1 ?

Public Function ReadLeadDays() As Long
    On Error GoTo Fallback
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_CFG)
    If IsNumeric(ws.Range(CELL_LEAD_DAYS).Value) Then
        ReadLeadDays = Application.Max(1, CLng(ws.Range(CELL_LEAD_DAYS).Value))
        Exit Function
    End If
Fallback:
    ReadLeadDays = DEF_LEAD_DAYS
End Function



Public Function ReadRunDate() As Date
    On Error GoTo Fallback
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHEET_CFG)
    If IsDate(ws.Range(CELL_RUN_DATE).Value) Then
        ReadRunDate = CDate(ws.Range(CELL_RUN_DATE).Value)
        Exit Function
    End If
Fallback:
    ReadRunDate = Date
End Function

Public Function ReadSourceSheetName() As String
    On Error Resume Next
    Dim s$: s = Trim$(CStr(ThisWorkbook.Worksheets(SHEET_CFG).Range(CELL_SRC_SHEET).Value))
    If Len(s) = 0 Then s = SHEET_SRC   ' ?????????
    ReadSourceSheetName = s
End Function



' 读取配置：HorizonDays
Public Function ReadHorizonDays() As Long
    On Error GoTo Fallback
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CFG)
    If IsNumeric(ws.Range(CELL_HORIZON_DAYS).Value) Then
        ReadHorizonDays = CLng(ws.Range(CELL_HORIZON_DAYS).Value)
        If ReadHorizonDays <= 0 Then ReadHorizonDays = DEF_HORIZON_DAYS
        Exit Function
    End If
Fallback:
    ReadHorizonDays = DEF_HORIZON_DAYS
End Function

' 读取配置：每批能力 EffCap(t)
Public Function ReadEffCapPerBatch() As Double
    On Error GoTo Fallback
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CFG)
    If IsNumeric(ws.Range(CELL_EFFCAP_PER_BATCH).Value) Then
        ReadEffCapPerBatch = CDbl(ws.Range(CELL_EFFCAP_PER_BATCH).Value)
        If ReadEffCapPerBatch <= 0 Then ReadEffCapPerBatch = DEF_EFFCAP_PER_BATCH
        Exit Function
    End If
Fallback:
    ReadEffCapPerBatch = DEF_EFFCAP_PER_BATCH
End Function

' 读取配置：Window(d)
Public Function ReadWindowDays() As Long
    On Error GoTo Fallback
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHEET_CFG)
    If IsNumeric(ws.Range(CELL_WINDOW_DAYS).Value) Then
        ReadWindowDays = CLng(ws.Range(CELL_WINDOW_DAYS).Value)
        If ReadWindowDays <= 0 Then ReadWindowDays = DEF_WINDOW_DAYS
        Exit Function
    End If
Fallback:
    ReadWindowDays = DEF_WINDOW_DAYS
End Function



