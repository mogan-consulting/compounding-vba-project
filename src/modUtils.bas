Attribute VB_Name = "modUtils"
Option Explicit

Public Function GetSheetByNameSafe(ByVal name As String, Optional createIfMissing As Boolean = False) As Worksheet
    On Error Resume Next
    Set GetSheetByNameSafe = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If GetSheetByNameSafe Is Nothing And createIfMissing Then
        Set GetSheetByNameSafe = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetSheetByNameSafe.name = name
    End If
End Function

' 在表头中按别名数组查找列号
Public Function FindCol(ByVal ws As Worksheet, ByVal candidates As Variant) As Long
    Dim lastCol&, c&, head$, i&
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        head = ws.Cells(1, c).Text
        ' ??:???????/????/?????,?????
        head = WorksheetFunction.Clean(head)
        head = Replace(head, Chr(160), " ")
        Do While InStr(head, "  ") > 0: head = Replace(head, "  ", " "): Loop
        head = LCase$(Trim$(head))

        For i = LBound(candidates) To UBound(candidates)
            If head = LCase$(CStr(candidates(i))) Then
                FindCol = c
                Exit Function
            End If
        Next i
    Next c
    FindCol = 0
End Function


' 对齐到最近的「本周或未来的」锚定周六（含当日若即为周六）
Public Function AlignToAnchorSaturday(ByVal d As Date) As Date
    Dim wd As Integer
    wd = Weekday(d, vbSunday)
    ' 周日=1, 周六=7；若 wd=7 则就是周六
    AlignToAnchorSaturday = d + ((vbSaturday - wd + 7) Mod 7)
End Function

' 安全读取字符串
Public Function NzText(ByVal v As Variant, Optional ByVal defaultText As String = "") As String
    If IsError(v) Then
        NzText = defaultText
    ElseIf IsNull(v) Or IsEmpty(v) Then
        NzText = defaultText
    Else
        NzText = Trim$(CStr(v))
    End If
End Function

' 安全读取日期
Public Function NzDate(ByVal v As Variant, Optional ByVal def As Date = 0) As Date
    On Error Resume Next
    If IsDate(v) Then NzDate = CDate(v): Exit Function
    On Error GoTo 0

    Dim s As String, a() As String
    s = Trim(CStr(v))
    If Len(s) = 0 Then NzDate = def: Exit Function

    ' ?????
    s = Replace(Replace(s, "/", "-"), ".", "-")
    a = Split(s, "-")
    If UBound(a) = 2 Then
        Dim y&, m&, d&
        If Len(a(0)) = 4 Then               ' yyyy-mm-dd
            y = CLng(a(0)): m = CLng(a(1)): d = CLng(a(2))
            On Error Resume Next: NzDate = DateSerial(y, m, d): On Error GoTo 0
            If NzDate <> 0 Then Exit Function
        ElseIf Len(a(2)) = 4 Then           ' dd-mm-yyyy ? mm-dd-yyyy(??????????)
            y = CLng(a(2)): m = CLng(a(1)): d = CLng(a(0))
            On Error Resume Next: NzDate = DateSerial(y, m, d): On Error GoTo 0
            If NzDate <> 0 Then Exit Function
        End If
    End If

    NzDate = def
End Function


' 安全读取数值
Public Function NzDouble(ByVal v As Variant, Optional ByVal def As Double = 0#) As Double
    ' ??:??????????????????????(123) ??
    If IsNumeric(v) Then NzDouble = CDbl(v): Exit Function

    Dim s As String
    s = Trim$(CStr(v))
    If Len(s) = 0 Then NzDouble = def: Exit Function

    s = WorksheetFunction.Clean(s)
    s = Replace$(s, Chr(160), " ")  ' NBSP
    s = Replace$(s, " ", "")
    s = Replace$(s, ",", "")
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then s = "-" & Mid$(s, 2, Len(s) - 2)

    On Error Resume Next
    NzDouble = CDbl(s)
    If Err.Number <> 0 Then NzDouble = def: Err.Clear
    On Error GoTo 0
End Function


' 清空/重建输出表的表头
Public Sub PrepareOutputSheet(ByVal sheetName As String, ByVal headers As Variant)
    Dim ws As Worksheet
    Set ws = GetSheetByNameSafe(sheetName, True)
    ws.Cells.Clear  ' ??????,??????�????�
    Dim i&
    For i = LBound(headers) To UBound(headers)
        ws.Cells(1, i - LBound(headers) + 1).Value = headers(i)
    Next i
    ws.Rows(1).Font.Bold = True
End Sub

' 将二维变体数组一次性写入（从第2行开始）
Public Sub WriteArrayToSheet(ByVal sheetName As String, ByRef a As Variant)
    If IsEmpty(a) Then Exit Sub
    Dim ws As Worksheet
    Set ws = GetSheetByNameSafe(sheetName, True)
    Dim nRows&, nCols&
    nRows = UBound(a, 1) - LBound(a, 1) + 1
    nCols = UBound(a, 2) - LBound(a, 2) + 1
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + nRows, nCols)).Value = a
    ws.Columns.AutoFit
End Sub

' 四舍五入到小数位
Public Function RoundTo(ByVal x As Double, ByVal digits As Long) As Double
    RoundTo = WorksheetFunction.Round(x, digits)
End Function

' ?????????(??=1 ... ??=7)
Public Function AlignToAnchorWeekday(ByVal d As Date, ByVal anchorWd As Long) As Date
    Dim wd As Integer
    wd = Weekday(d, vbSunday)
    AlignToAnchorWeekday = d + ((anchorWd - wd + 7) Mod 7)
End Function

Public Function FactorFromFGType(ByVal s As String) As Double
    s = LCase$(Trim$(CStr(s)))
    Select Case s
        Case "10ml", "10 ml": FactorFromFGType = 10.4
        Case "5ml", "5 ml":   FactorFromFGType = 5.4
        Case "3ml", "3 ml":   FactorFromFGType = 3.4
        Case Else:            FactorFromFGType = 0#
    End Select
End Function

Sub ProbeColumns()
    Dim sh$: sh = ReadSourceSheetName()
    Dim ws As Worksheet: Set ws = Worksheets(sh)
    Dim cStart&, cEnd&, cQty&, cOrderID&, cFactor&, cUsageCol&, cFGType&
    cStart = FindCol(ws, Array("FG start date", "Start date"))
    cEnd = FindCol(ws, Array("FG end date", "End date"))
    cQty = FindCol(ws, Array("plan order qty", "plan order quantity", "Plan order qty", "Plan Qty"))
    cOrderID = FindCol(ws, Array("Order ID", "OrderID", "Order", "Document"))
    cFactor = FindCol(ws, Array("Multiply factor", "Factor", "multiply factor"))
    cUsageCol = FindCol(ws, Array("usage (t)", "usage", "Derived compounding usage (t)"))
    cFGType = FindCol(ws, Array("FG type", "FGtype", "Type"))

    MsgBox "Sheet = " & sh & vbCrLf & _
           "Start=" & cStart & "  End=" & cEnd & "  Qty=" & cQty & vbCrLf & _
           "OrderID=" & cOrderID & "  Factor=" & cFactor & "  usage(t)=" & cUsageCol & "  FGtype=" & cFGType, _
           vbInformation, "ProbeColumns"
End Sub

Sub QuickProbe()
    Dim ws As Worksheet: Set ws = Worksheets(ReadSourceSheetName())
    Dim hz&, runD As Date: hz = ReadHorizonDays(): runD = ReadRunDate()

    Dim cStart&, cEnd&, cQty&, cFactor&, cUsageCol&, cFGType&, lastRow&, r&, n&, total#
    cStart = FindCol(ws, Array("FG start date", "Start date"))
    cEnd = FindCol(ws, Array("FG end date", "End date"))
    cQty = FindCol(ws, Array("plan order qty", "plan order quantity", "Plan order qty", "Plan Qty"))
    cFactor = FindCol(ws, Array("Multiply factor", "Factor", "multiply factor"))
    cUsageCol = FindCol(ws, Array("usage (t)", "usage"))
    cFGType = FindCol(ws, Array("FG type", "Type"))

    lastRow = ws.Cells(ws.Rows.Count, IIf(cStart > 0, cStart, 1)).End(xlUp).Row
    For r = 2 To lastRow
        Dim d As Date, q#, u#, f#
        d = NzDate(ws.Cells(r, cStart).Value, 0)
        If d >= runD And d <= runD + hz Then
            q = NzDouble(ws.Cells(r, cQty).Value, 0)
            If cUsageCol > 0 And IsNumeric(ws.Cells(r, cUsageCol).Value) Then
                u = NzDouble(ws.Cells(r, cUsageCol).Value, 0)
            Else
                If cFactor > 0 Then f = NzDouble(ws.Cells(r, cFactor).Value, 0) Else f = FactorFromFGType(ws.Cells(r, cFGType).Value)
                If f = 0 Then f = FACTOR_UNIT
                u = q * f / 1000000#
            End If
            If u > 0 Then n = n + 1: total = total + u
        End If
        Debug.Print r, Format(d, "yyyy-mm-dd"), q, f, u
    Next r
    MsgBox "Orders in window = " & n & vbCrLf & _
           "Total demand (t) = " & Format(total, "0.000"), vbInformation, "QuickProbe"
End Sub

Public Function Max3(ByVal a As Long, ByVal b As Long, ByVal c As Long) As Long
    Max3 = Application.Max(a, b, c)
End Function

' === NEW: inclusive overlap of two [start, end] date ranges, returns number of days (>=0) ===
Public Function DaysOverlapInclusive(ByVal aStart As Date, ByVal aEnd As Date, _
                                     ByVal bStart As Date, ByVal bEnd As Date) As Long
    If aEnd < aStart Then
        DaysOverlapInclusive = 0
        Exit Function
    End If
    If bEnd < bStart Then
        DaysOverlapInclusive = 0
        Exit Function
    End If
    Dim s As Date, e As Date
    If aStart > bStart Then s = aStart Else s = bStart
    If aEnd < bEnd Then e = aEnd Else e = bEnd
    If e < s Then
        DaysOverlapInclusive = 0
    Else
        DaysOverlapInclusive = DateDiff("d", s, e) + 1
    End If
End Function

' ?????H??�?�,?????? ????? ??
' ?:offsetCols=+2 ???H??J
Public Function ReadNamedValueAtOffset(ByVal sheetName As String, _
                                       ByVal keyText As String, _
                                       ByVal offsetCols As Long, _
                                       Optional ByVal defaultValue As Variant) As Variant
    Dim ws As Worksheet, lastRow&, r&
    Set ws = GetSheetByNameSafe(sheetName, True)
    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    For r = 1 To lastRow
        If Trim$(CStr(ws.Cells(r, "H").Value)) = keyText Then
            ReadNamedValueAtOffset = ws.Cells(r, "H").Offset(0, offsetCols).Value
            Exit Function
        End If
    Next r
    ReadNamedValueAtOffset = defaultValue
End Function

' ??�??FG???????���??????J?(H????2?)
Public Function ReadMinTonsByFGType(ByVal fgType As String) As Double
    Dim key$
    Select Case LCase$(Trim$(fgType))
        Case "10ml": key = KEY_MIN_QTY_10ML
        Case "5ml":  key = KEY_MIN_QTY_5ML
        Case "3ml":  key = KEY_MIN_QTY_3ML
        Case Else
            ' ????????
            ReadMinTonsByFGType = 0#
            Exit Function
    End Select
    ' H?J = ?? +2
    ReadMinTonsByFGType = CDbl(NzDouble(ReadNamedValueAtOffset(SHEET_CFG, key, 2, 0#)))
End Function



'=== ? OrderID ?????????(??) ===
Public Function LookupByOrderId(ByVal wb As Workbook, ByVal sheetName As String, _
                                ByVal orderId As Long, ByVal targetHeader As String) As Variant
    Dim ws As Worksheet, lastRow As Long
    Dim colOrder As Long, colTarget As Long
    Dim rowFound As Variant

    On Error GoTo EH
    Set ws = wb.Worksheets(sheetName)

    '?????????????;???,????? Find
    colOrder = FindCol(ws, "Order ID")
    colTarget = FindCol(ws, targetHeader)
    If colOrder = 0 Or colTarget = 0 Then GoTo EH

    lastRow = ws.Cells(ws.Rows.Count, colOrder).End(xlUp).Row
    If lastRow < 2 Then GoTo EH

    rowFound = Application.Match(orderId, ws.Range(ws.Cells(2, colOrder), ws.Cells(lastRow, colOrder)), 0)
    If IsError(rowFound) Then GoTo EH

    LookupByOrderId = ws.Cells(1 + rowFound, colTarget).Value
    Exit Function
EH:
    LookupByOrderId = vbNullString   '???????
End Function

'=== ? FG type ===
Public Function GetFGTypeForOrder(ByVal wb As Workbook, ByVal runSheetName As String, _
                                  ByVal orderId As Long) As String
    GetFGTypeForOrder = CStr(LookupByOrderId(wb, runSheetName, orderId, "FG type"))
End Function

'=== ? plan order qty(pcs)===
Public Function GetPlanQtyForOrder(ByVal wb As Workbook, ByVal runSheetName As String, _
                                   ByVal orderId As Long) As Double
    Dim v As Variant
    v = LookupByOrderId(wb, runSheetName, orderId, "plan order qty")
    If IsNumeric(v) Then
        GetPlanQtyForOrder = CDbl(v)
    Else
        GetPlanQtyForOrder = 0#
    End If
End Function

