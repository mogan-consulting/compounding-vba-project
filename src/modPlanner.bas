Attribute VB_Name = "modPlanner"
Option Explicit


' ===== È°∂Â±ÇÂÖ•Âè£ =====
Public Sub RunCompoundingPlan()
    PlanFixed40Core True, True
    ' === FINAL: rebuild summary from actual allocation ===
    On Error Resume Next
    RebuildBatchSummaryFromAllocation
    On Error GoTo 0

    Call Sanity_Post_AllocWindow
End Sub

' ‰∏ªÊµÅÁ®ã
Public Sub PlanFixed40Core(ByVal writeSheets As Boolean, ByVal writeText As Boolean)
    Dim ws As Worksheet
    Set ws = GetSheetByNameSafe(ReadSourceSheetName(), False)
    If ws Is Nothing Then Err.Raise 1001, , "Missing source sheet: " & ReadSourceSheetName()
    
    ' --- ????????(???) ---
    Dim srcName As String
    Dim runD As Date
    Dim hz As Long
    
    srcName = ReadSourceSheetName()
    runD = ReadRunDate()
    hz = ReadHorizonDays()
    
    MsgBox "Run with:" & vbCrLf & _
           "  Source sheet = " & srcName & vbCrLf & _
           "  RunDate = " & Format(runD, "yyyy-mm-dd") & vbCrLf & _
           "  HorizonDays = " & CStr(hz), _
           vbInformation, "Compounding"
' --- /end ?? ---


    ' ÊâæÂÖ≥ÈîÆÂàó
    Dim cStart&, cEnd&, cQty&, cCompany&, cProduct&, cOrderID&, cFactor&, cUsageCol&, cFGType&
    cStart = FindCol(ws, Array("FG start date", "Start date"))
    cEnd = FindCol(ws, Array("FG end date", "End date"))
    cQty = FindCol(ws, Array("plan order qty", "plan order quantity", "Plan order qty", "Plan Qty"))
    
    If cStart = 0 Or cQty = 0 Then
        Err.Raise 1001, , "Missing required columns: Start date and/or plan order qty."
    End If
    
    cCompany = FindCol(ws, Array("Company", "Customer", "BU"))         ' ??
    cProduct = FindCol(ws, Array("Product", "Material", "Item"))       ' ??
    cOrderID = FindCol(ws, Array("Order ID", "OrderID", "Order", "Document"))   ' ???????
    cFactor = FindCol(ws, Array("Multiply factor", "Factor", "multiply factor"))
    cUsageCol = FindCol(ws, Array("usage (t)", "usage", "Derived compounding usage (t)"))
    cFGType = FindCol(ws, Array("FG type", "FGtype", "Type"))

    ' ËØªÂèñÂèÇÊï∞
    Dim horizonDays&, windowDays&
    Dim effCapPerBatch#
    horizonDays = ReadHorizonDays()
    effCapPerBatch = ReadEffCapPerBatch()
    windowDays = ReadWindowDays()

    ' ËØªÂèñÊ∫êÊï∞ÊçÆ
    Dim orders() As FGOrder, nOrders&
    ' ????
    ReadOrders ws, cStart, cEnd, cQty, cCompany, cProduct, cOrderID, _
               cFactor, cUsageCol, cFGType, _
               horizonDays, orders, nOrders
    
    MsgBox "ReadOrders -> orders in window = " & nOrders, vbInformation
    
    If nOrders = 0 Then
        If writeSheets Then
            PrepareOutputSheet SHEET_ALLOC, Array("Company", "Product", "Order ID", "Start date", "End date", "Batch#", "Anchor", "usage (t)", "Horizon(d)")
            PrepareOutputSheet SHEET_SUM, Array("Batch#", "Anchor", "First start", "Last start", "Allocated (t)", "EffCap (t)", "Remaining (t)", "Window(d)", "Horizon(d)")
        End If
        Exit Sub
    End If

    ' ÁîüÊàêÊâπÂàóË°®ÔºàÊåâÂë®ÂÖ≠ÈîöÁÇπÔºåË¶ÜÁõñÁ™óÂÜÖÔºâ
    Dim batches() As tBatch, nbatches As Long
    Dim alloc()   As tAlloc, nAlloc    As Long
    
    ' ??????????,????? ReDim Preserve
    ReDim batches(1 To 1): nbatches = 0
    ReDim alloc(1 To 1): nAlloc = 0

    'AllocateOrders orders, nOrders, batches, nbatches, HorizonDays, WindowDays, alloc, nAlloc
    ' ???????(???????)
    AllocateOrdersAsNeeded _
    batches, nbatches, _
    alloc, nAlloc, _
    orders, nOrders, _
    effCapPerBatch, windowDays, horizonDays

    ' Ê†°È™å‰∏ÄËá¥ÊÄß
    ValidateConsistency orders, nOrders, alloc, nAlloc, batches, nbatches

    ' ËæìÂá∫
    If writeSheets Then
        WriteAllocationSheet alloc, nAlloc
        WriteBatchSummarySheet batches, nbatches
        ShowRunSummary orders, nOrders, alloc, nAlloc, batches, nbatches, horizonDays, effCapPerBatch
    End If

    ' ÂèØÈÄâÊñáÊú¨ËæìÂá∫
    If writeText Then
        ' ËøôÈáå‰øùÁïôÊó•ÂøóËæìÂá∫ÁÇπÔºàÂ¶ÇÈúÄÂÜôÂá∫Âà∞ÊñáÊú¨ÔºåÂèØÂÆûÁé∞ SaveTextFileÔºâ
    End If
End Sub

' ------- ËØªÂèñÊ∫êÊï∞ÊçÆÂπ∂Êç¢ÁÆó usage(t) -------
Private Sub ReadOrders(ByVal ws As Worksheet, _
                       ByVal cStart&, ByVal cEnd&, ByVal cQty&, _
                       ByVal cCompany&, ByVal cProduct&, ByVal cOrderID&, _
                       ByVal cFactor&, ByVal cUsageCol&, ByVal cFGType&, _
                       ByVal horizonDays&, _
                       ByRef orders() As FGOrder, ByRef nOrders&)

    Dim lastRow&, r&, runDate As Date, dStart As Date
    runDate = ReadRunDate()
    lastRow = ws.Cells(ws.Rows.Count, cStart).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ReDim orders(1 To lastRow - 1)
    nOrders = 0

    For r = 2 To lastRow
        dStart = NzDate(ws.Cells(r, cStart).Value, 0)
        If dStart = 0 Then GoTo ContinueRow

        ' ËÆ°ÂàíÁ™óËøáÊª§ÔºöStartDate ‚àà [today, today + horizonDays]
        If dStart < runDate Or dStart > (runDate + horizonDays) Then GoTo ContinueRow

        Dim qty#, UsageT#, dEnd As Date, factor#
        qty = NzDouble(ws.Cells(r, cQty).Value, 0#)
        If qty <= 0 Then GoTo ContinueRow
        
        ' 1) ???????? usage(t)
        If cUsageCol > 0 And IsNumeric(ws.Cells(r, cUsageCol).Value) Then
            UsageT = NzDouble(ws.Cells(r, cUsageCol).Value, 0#)
        Else
            ' 2) ?????? Multiply factor;??? FG type ??;??????
            If cFactor > 0 Then
                factor = NzDouble(ws.Cells(r, cFactor).Value, 0#)
            ElseIf cFGType > 0 Then
                factor = FactorFromFGType(ws.Cells(r, cFGType).Value)
                If factor = 0# Then factor = FACTOR_UNIT
            Else
                factor = FACTOR_UNIT
            End If
            UsageT = qty * factor / 1000000#
            If APPLY_LOSS_FACTOR Then UsageT = UsageT * FACTOR_LOSS
        End If
        UsageT = RoundTo(UsageT, 4)
        If UsageT <= 0 Then GoTo ContinueRow
        
        ' 3) EndDate(??)
        If cEnd > 0 Then dEnd = NzDate(ws.Cells(r, cEnd).Value, dStart) Else dEnd = dStart
        
        ' ======== ??????(??? OrderID) ========
        nOrders = nOrders + 1
        With orders(nOrders)
            ' Company/Product ????
            If cCompany > 0 Then .Company = NzText(ws.Cells(r, cCompany).Value, "N/A") Else .Company = "N/A"
            If cProduct > 0 Then .Product = NzText(ws.Cells(r, cProduct).Value, "") Else .Product = ""
        
            ' OrderID:????,??? ROW-<??> ??
            Dim oid$
            If cOrderID > 0 Then oid = CStr(ws.Cells(r, cOrderID).Value)
            If Len(Trim$(oid)) = 0 Then oid = "ROW-" & CStr(r)
            .orderId = oid
        
            .StartDate = dStart
            .EndDate = dEnd
            .PlanQty = qty
            .UsageT = UsageT
        End With
            If cFGType > 0 Then orders(nOrders).fgType = NzText(ws.Cells(r, cFGType).Value, "") Else orders(nOrders).fgType = ""

        'End If
ContinueRow:
    Next r

    If nOrders > 0 Then ReDim Preserve orders(1 To nOrders) Else Erase orders
End Sub

Private Function ComputeUsageT(ByVal PlanQty As Double) As Double
    ' usage(t) = qty √ó 1.07 √ó 10.4 √∑ 1,000 √∑ 1,000
    Dim t#
    t = PlanQty * FACTOR_LOSS * FACTOR_UNIT / 1000# / 1000#
    ComputeUsageT = RoundTo(t, 4) ' ‰∏≠Èó¥‰øùÁïô 4 ‰Ωç
End Function

' ------- ÊâπÁîüÊàê -------
Private Sub BuildBatches(ByVal runDate As Date, ByVal horizonDays&, ByVal effCap#, ByVal windowDays&, _
                         ByRef batches() As Batch, ByRef nbatches&)
    Dim horizonEnd As Date
    horizonEnd = runDate + horizonDays

    ' Á¨¨‰∏ÄÊù°ÈîöÁÇπÔºö‰ªé runDate ÂØπÈΩêÂà∞ÊúÄËøëÁöÑ„ÄåÊú¨Âë®ÊàñÊú™Êù•Âë®ÂÖ≠„Äç
    Dim anchor As Date
    anchor = AlignToAnchorSaturday(runDate)

    nbatches = 0
    ReDim batches(1 To 1)

    Do While anchor <= horizonEnd
        nbatches = nbatches + 1
        If nbatches > UBound(batches) Then ReDim Preserve batches(1 To nbatches)
        With batches(nbatches)
            .BatchNo = nbatches
            .anchor = anchor
            .FirstStart = anchor
            .LastStart = anchor + (windowDays - 1)
            .EffCapT = effCap
            .AllocatedT = 0#
            .windowDays = windowDays
            .horizonDays = horizonDays
        End With
        anchor = anchor + 7 ' ‰∏ã‰∏ÄÂë®ÂÖ≠
    Loop
End Sub

' ------- ÂàÜÈÖçÈÄªËæëÔºà‰∏çË∂ÖÈÖçÔºåÊ∫¢Âá∫ÂêéÊé®Ôºâ -------


Private Function FindBatchIndexByAnchor(ByRef batches() As Batch, ByVal nbatches&, ByVal anchor As Date) As Long
    Dim i&
    For i = 1 To nbatches
        If batches(i).anchor = anchor Then
            FindBatchIndexByAnchor = i
            Exit Function
        End If
    Next i
    FindBatchIndexByAnchor = 0
End Function

' ------- ‰∏ÄËá¥ÊÄßÊ†°È™å -------
Public Sub ValidateConsistency( _
    ByRef orders() As FGOrder, ByVal nOrders As Long, _
    ByRef alloc() As tAlloc, ByVal nAlloc As Long, _
    ByRef batches() As tBatch, ByVal nbatches As Long)
    
    Dim sumOrder#, sumAlloc#, i&
    For i = 1 To nOrders
        sumOrder = sumOrder + orders(i).UsageT
    Next i
    For i = 1 To nAlloc
        sumAlloc = sumAlloc + alloc(i).UsageT
    Next i
    sumOrder = RoundTo(sumOrder, 3)
    sumAlloc = RoundTo(sumAlloc, 3)

    If Abs(sumOrder - sumAlloc) > 0.001 Then
        Err.Raise 3001, , "Consistency error: total allocated <> total demand within horizon. Demand=" & _
            Format(sumOrder, "0.000") & " Alloc=" & Format(sumAlloc, "0.000")
    End If

    ' ‰∏çË∂ÖÈÖçÊ£ÄÊü•
    For i = 1 To nbatches
        If batches(i).AllocatedT - batches(i).EffCapT > 0.0001 Then
            Err.Raise 3002, , "Over-allocation on batch #" & batches(i).BatchNo
        End If
    Next i
End Sub

Public Sub WriteAllocationSheet(ByRef alloc() As tAlloc, ByVal nAlloc As Long)
    Dim ws As Worksheet
    Dim i As Long, windowDays As Long
    Dim a() As Variant

    windowDays = ReadWindowDays()                          ' 21(?????)
    Set ws = GetSheetByNameSafe(SHEET_ALLOC, True)         ' "CompoundingAllocation"

    ws.Cells.Clear
    ' ??? 8 ??? 10 ?,?? I/J
    ws.Range("A1").Resize(1, 10).Value = Array( _
        "Order ID", "Start date", "End date", "Batch#", "Anchor", "Valid thru", _
        "usage (t)", "Horizon(d)", "FG type", "plan order qty")

    If nAlloc <= 0 Then Exit Sub

    ' ????:8 -> 10 ?(??????)
    ReDim a(1 To nAlloc, 1 To 10)

    'óó ?????????(??????/??)óó
    Static initDone As Boolean
    'óó ?????????(? SRC Sheet ?????????)óó
    Static wsSrc As Worksheet, rngOrder As Range
    Static cOrder As Long, cFG As Long, cQty As Long, cUsage As Long
    Static lastSrcName As String
    Dim wsCur As Worksheet, lr As Long
    
    Set wsCur = ResolveRunSheet()                    ' ?? CompoundingTab ? SRC Sheet
    
    ' ???? ? ???? ? ?????
    If (wsSrc Is Nothing) Or (wsCur Is Nothing) Or (wsCur.name <> lastSrcName) Then
        Set wsSrc = wsCur
        lastSrcName = ""
        Set rngOrder = Nothing
        cOrder = 0: cFG = 0: cQty = 0: cUsage = 0
    
        If Not wsSrc Is Nothing Then
            lastSrcName = wsSrc.name
    
            ' ????????(?????)
            cOrder = FindCol(wsSrc, Array("Order ID"))
            cFG = FindCol(wsSrc, Array("FG type"))
            cQty = FindCol(wsSrc, Array("plan order qty"))
            cUsage = FindCol(wsSrc, Array("usage (t)"))
    
            ' ????(? Compounding_* ???:D/E/G)
            If cQty = 0 Then cQty = 4          ' D
            If cFG = 0 Then cFG = 5            ' E
            If cUsage = 0 Then cUsage = 7      ' G
    
            ' ??????? A?(????,?? 9 ???)
            If cOrder > 0 Then
                On Error Resume Next
                lr = wsSrc.Cells(wsSrc.Rows.Count, cOrder).End(xlUp).Row
                On Error GoTo 0
                If lr >= 2 Then
                    Set rngOrder = wsSrc.Range(wsSrc.Cells(2, cOrder), wsSrc.Cells(lr, cOrder))
                End If
            End If
        End If
    End If

    If Not initDone Then
        Set wsSrc = ResolveRunSheet()                      ' ????:???CompoundingTab? SRC Sheet

        If Not wsSrc Is Nothing Then
            ' ? ???????(??????,????)
            cOrder = FindCol(wsSrc, Array("Order ID"))
            cFG = FindCol(wsSrc, Array("FG type"))
            cQty = FindCol(wsSrc, Array("plan order qty"))
            cUsage = FindCol(wsSrc, Array("usage (t)"))      ' ? ??:?? usage(t) ?
            ' ????(??? Compounding_Test_3:D=4, E=5)
            If cQty = 0 Then cQty = 4
            If cFG = 0 Then cFG = 5
            If cUsage = 0 Then cUsage = 7    ' G ?
            ' ?????? rngOrder(cOrder=0 ???,?? 9 ??)
            If cOrder > 0 Then
                On Error Resume Next
                lr = wsSrc.Cells(wsSrc.Rows.Count, cOrder).End(xlUp).Row
                On Error GoTo 0
                If lr >= 2 Then
                    Set rngOrder = wsSrc.Range(wsSrc.Cells(2, cOrder), wsSrc.Cells(lr, cOrder))
                Else
                    Set rngOrder = Nothing
                End If
            Else
                Set rngOrder = Nothing
            End If
        End If
        initDone = True
    End If

    'óó ????(??????)óó
    Dim posVar As Variant, key As Variant, rFound As Long
    For i = 1 To nAlloc
        a(i, 1) = alloc(i).orderId
        a(i, 2) = alloc(i).StartDate
        a(i, 3) = alloc(i).EndDate
        a(i, 4) = alloc(i).BatchNo
        a(i, 5) = alloc(i).anchor
        a(i, 6) = DateAdd("d", windowDays - 1, alloc(i).anchor)   ' Valid thru
        a(i, 7) = RoundTo(alloc(i).UsageT, 3)
        a(i, 8) = alloc(i).horizonDays

        'óó ??:I/J ??(? Order ID ???????)óó
        a(i, 9) = vbNullString
        a(i, 10) = vbNullString

        If Not wsSrc Is Nothing Then
            ' ????????? rngOrder,???????(????????)
            If rngOrder Is Nothing And cOrder > 0 Then
                On Error Resume Next
                lr = wsSrc.Cells(wsSrc.Rows.Count, cOrder).End(xlUp).Row
                On Error GoTo 0
                If lr >= 2 Then
                    Set rngOrder = wsSrc.Range(wsSrc.Cells(2, cOrder), wsSrc.Cells(lr, cOrder))
                End If
            End If

            If Not rngOrder Is Nothing Then
            ' ?????,?????,?? "1" vs 1
            posVar = Application.Match(CDbl(alloc(i).orderId), rngOrder, 0)
            If IsError(posVar) Or IsEmpty(posVar) Then
                posVar = Application.Match(CStr(alloc(i).orderId), rngOrder, 0)
            End If
        
            If Not IsError(posVar) And Len(posVar) > 0 Then
                rFound = 1 + CLng(posVar)
        
                ' I ?:???? FG type
                a(i, 9) = wsSrc.Cells(rFound, cFG).Value
        
                ' J ?:???? = ???? ◊ (??usage / ??usage) ,? RoundUp ???
                Dim planTotal As Double, usageTotal As Double, qtySlice As Double, v As Variant
        
                v = wsSrc.Cells(rFound, cQty).Value
                If IsNumeric(v) Then planTotal = CDbl(v) Else planTotal = 0
        
                v = wsSrc.Cells(rFound, cUsage).Value
                If IsNumeric(v) Then usageTotal = CDbl(v) Else usageTotal = 0
        
                If usageTotal > 0 Then
                    qtySlice = planTotal * (alloc(i).UsageT / usageTotal)
                Else
                    qtySlice = planTotal   ' ??
                End If
        
                ' ?????:?? RoundUp ???(???10/100,??????)
                a(i, 10) = Application.WorksheetFunction.RoundUp(qtySlice, -1)
            End If
End If

        End If
        'óó end I/J óó
    Next i

    ' ????? 10 ?
    ws.Range("A2").Resize(nAlloc, 10).Value = a

    ' ???? + ?? J ???
    ws.Columns("B:C").NumberFormat = "yyyy-mm-dd"    ' Start / End
    ws.Columns("E:F").NumberFormat = "yyyy-mm-dd"    ' Anchor / Valid thru
    ws.Columns("G:G").NumberFormat = "0.000"         ' usage (t)
    ws.Columns("J:J").NumberFormat = "#,##0"         ' plan order qty
    ws.Columns.AutoFit
End Sub




Public Sub WriteBatchSummarySheet(ByRef batches() As tBatch, ByVal nbatches As Long)
    Dim ws As Worksheet
    Dim i As Long
    Dim a() As Variant

    Set ws = GetSheetByNameSafe(SHEET_SUM, True)       ' "CompoundingBatchSummary"

    ws.Cells.Clear
    ws.Range("A1").Resize(1, 10).Value = Array( _
        "Batch#", "Anchor", "Valid thru", "First start", "Last start", _
        "Allocated (t)", "EffCap (t)", "Remaining (t)", "Window(d)", "Horizon(d)")

    If nbatches <= 0 Then Exit Sub

    ReDim a(1 To nbatches, 1 To 10)
    For i = 1 To nbatches
        a(i, 1) = batches(i).BatchNo
        a(i, 2) = batches(i).anchor
        a(i, 3) = DateAdd("d", batches(i).windowDays - 1, batches(i).anchor) ' –ß⁄£Àµ„£©
        a(i, 4) = batches(i).FirstStart
        a(i, 5) = batches(i).LastStart
        a(i, 6) = RoundTo(batches(i).AllocatedT, 3)
        a(i, 7) = batches(i).EffCapT
        a(i, 8) = RoundTo(batches(i).EffCapT - batches(i).AllocatedT, 3)
        a(i, 9) = batches(i).windowDays
        a(i, 10) = batches(i).horizonDays
    Next i

    ws.Range("A2").Resize(nbatches, 10).Value = a

    ws.Columns("B:E").NumberFormat = "yyyy-mm-dd"      ' Anchor/ValidThru/First/Last start
    ws.Columns("F:H").NumberFormat = "0.000"           ' Allocated/EffCap/Remaining
    ws.Columns.AutoFit
    RebuildBatchSummaryFromAllocation
End Sub

Public Sub AllocateOrdersAsNeeded( _
    ByRef batches() As tBatch, ByRef nbatches As Long, _
    ByRef alloc() As tAlloc, ByRef nAlloc As Long, _
    ByRef orders() As FGOrder, ByVal nOrders As Long, _
    ByVal effCap As Double, ByVal windowDays As Long, ByVal horizonDays As Long)

    Dim i As Long, remain As Double, capLeft As Double, take As Double
    Dim batchMinStart As Date     ' ???????????
    Dim firstAllocIdx As Long     ' ??? alloc() ?????
    Dim leadDays As Long: leadDays = ReadLeadDays()
    Dim orderStartIdx As Long
    ' Added for daily split
    Dim totalDays As Long, plannedDays As Long
    Dim dailyU As Double
    Dim cursorDate As Date

    '  ?????(Anchor ?? 0,??????????)
    If nbatches = 0 Then
        StartNewBatch batches, nbatches, 0, effCap, windowDays, horizonDays
        batchMinStart = 0
        firstAllocIdx = nAlloc + 1
    End If
    ' ????????(?? SAP ???????????)
    If nOrders > 1 Then SortOrdersByStartThenEnd orders, nOrders

    For i = 1 To nOrders
        remain = RoundTo(orders(i).UsageT, 4)
        If remain <= 0 Then GoTo ContinueOrder
        orderStartIdx = nAlloc + 1
        Do While remain > 0
            ' ???????
            capLeft = RoundTo(batches(nbatches).EffCapT - batches(nbatches).AllocatedT, 4)

            ' ????/????,????(Anchor ??? 0)
            If capLeft <= 0 Then
                StartNewBatch batches, nbatches, 0, effCap, windowDays, horizonDays
                batchMinStart = 0
                firstAllocIdx = nAlloc + 1
                capLeft = batches(nbatches).EffCapT
            End If
                           
                If batches(nbatches).anchor > 0 Then
                    Dim windowStart As Date, validThru As Date
                    windowStart = batches(nbatches).anchor
                    validThru = DateAdd("d", windowDays - 1, windowStart)
                
                    ' per-day usage & unplanned cursor day
                    ' totalDays/dailyU declared above
                    totalDays = DateDiff("d", orders(i).StartDate, orders(i).EndDate) + 1
                    If totalDays < 1 Then totalDays = 1
                    dailyU = orders(i).UsageT / totalDays
                
                    ' plannedDays/cursorDate declared above
                    If dailyU > 0# Then
                        plannedDays = CLng(Fix((orders(i).UsageT - remain) / dailyU + 0.0000001))
                    Else
                        plannedDays = 0
                    End If
                    If plannedDays < 0 Then plannedDays = 0
                    If plannedDays > totalDays Then plannedDays = totalDays
                    cursorDate = DateAdd("d", plannedDays, orders(i).StartDate)
                
                    ' next unplanned day already beyond window? ? new batch
                    If cursorDate > validThru Then
        StartNewBatch batches, nbatches, 0, effCap, windowDays, horizonDays
                        batchMinStart = 0
                        firstAllocIdx = nAlloc + 1
                        capLeft = batches(nbatches).EffCapT
        ' preset anchor based on cursorDate to avoid empty batch
        batchMinStart = cursorDate
        batches(nbatches).anchor = DateAdd("d", -leadDays, batchMinStart)
        firstAllocIdx = nAlloc + 1
        windowStart = batches(nbatches).anchor
        validThru = DateAdd("d", windowDays - 1, windowStart)
        End If
                
                    ' recompute bounds (batch may have changed)
                    windowStart = batches(nbatches).anchor
                    validThru = DateAdd("d", windowDays - 1, windowStart)
                
                    ' only allocate the overlap days within this window
                    Dim inDays As Long
                    inDays = DaysOverlapInclusive(cursorDate, orders(i).EndDate, windowStart, validThru)
                
                    Dim maxByWindow As Double
                    maxByWindow = RoundTo(inDays * dailyU, 4)
                
                    ' if nothing fits by window in this batch ? new batch
                    If maxByWindow <= 0# Then
                        StartNewBatch batches, nbatches, 0, effCap, windowDays, horizonDays
                        batchMinStart = 0
                        firstAllocIdx = nAlloc + 1
                        capLeft = batches(nbatches).EffCapT
                    End If
                
                    ' tighten cap by window portion
                    If maxByWindow < remain Then
                        If capLeft > maxByWindow Then capLeft = maxByWindow
                    End If

' ------------------------------------------------------------------------------------------

                
            End If
' -------------------------------------------------------------------------------

            ' ??????
            take = IIf(remain <= capLeft, remain, capLeft)

            '  ?????? StartDate ???????????? Anchor
            Dim effStart As Date
            If plannedDays > 0 Then
                effStart = cursorDate
            Else
                effStart = orders(i).StartDate
            End If
            If (batchMinStart = 0) Or (effStart < batchMinStart) Then
                batchMinStart = effStart
                batches(nbatches).anchor = DateAdd("d", -leadDays, batchMinStart)
                Dim k As Long
                For k = firstAllocIdx To nAlloc
                    alloc(k).anchor = batches(nbatches).anchor
                Next k
            End If
            ' === ?? alloc() ???????? ===
            If take > 0# Then
                If (Not Not alloc) = 0 Then
                    ' ?????:???????
                    ReDim alloc(1 To 100)
                ElseIf nAlloc + 1 > UBound(alloc) Then
                    ' ????:??
                    ReDim Preserve alloc(1 To UBound(alloc) + 100)
                End If
' === ??????? nAlloc ? 1 ?? ===

                ' ??? allocation(???? Anchor)
                nAlloc = nAlloc + 1
                With alloc(nAlloc)
                    .orderId = orders(i).orderId
                    .StartDate = effStart
                    .EndDate = orders(i).EndDate
                    .BatchNo = batches(nbatches).BatchNo         ' ??????
                    .anchor = batches(nbatches).anchor
                    .UsageT = RoundTo(take, 4)
                    .horizonDays = horizonDays
                End With

                With batches(nbatches)
                    If .FirstStart = 0 Or effStart < .FirstStart Then
                        .FirstStart = effStart
                    End If
                    If .LastStart = 0 Or effStart > .LastStart Then
                        .LastStart = effStart
                    End If
                End With

                ' ??????
                batches(nbatches).AllocatedT = RoundTo(batches(nbatches).AllocatedT + take, 4)

                ' ?????
                remain = RoundTo(remain - take, 4)
            End If
        Loop
        If nAlloc >= orderStartIdx Then
            EnforceMinTonsForRange alloc, nAlloc, orderStartIdx, nAlloc, orders(i).fgType
        End If
ContinueOrder:
    Next i
    Call RebalanceBatchesByCap(alloc, nAlloc, batches, nbatches, effCap, windowDays)
    Call ConsolidateAllocAdjacent(alloc, nAlloc)
    Call RebuildBatchSummaryFromAllocation
    ' ??:?? alloc ??
    If nAlloc > 0 Then ReDim Preserve alloc(1 To nAlloc)
End Sub
Public Sub StartNewBatch(ByRef batches() As tBatch, ByRef nbatches As Long, _
                         ByVal anchor As Date, ByVal effCap As Double, _
                         ByVal windowDays As Long, ByVal horizonDays As Long)
    nbatches = nbatches + 1
    If nbatches > UBound(batches) Then ReDim Preserve batches(1 To nbatches + 100)

    With batches(nbatches)
        .BatchNo = nbatches
        .anchor = anchor               ' ??? 0,??????
        .FirstStart = 0                ' ???
        .LastStart = 0                 ' ???
        .EffCapT = effCap
        .AllocatedT = 0
        .windowDays = windowDays
        .horizonDays = horizonDays
    End With
End Sub

Public Sub ShowRunSummary( _
    ByRef orders() As FGOrder, ByVal nOrders As Long, _
    ByRef alloc() As tAlloc, ByVal nAlloc As Long, _
    ByRef batches() As tBatch, ByVal nbatches As Long, _
    ByVal horizonDays As Long, ByVal effCapPerBatch As Double)

    Dim sumOrder As Double, sumAlloc As Double
    Dim i As Long
    Dim dMin As Date, dMax As Date

    For i = 1 To nOrders
        sumOrder = sumOrder + orders(i).UsageT
        If dMin = 0 Or orders(i).StartDate < dMin Then dMin = orders(i).StartDate
        If dMax = 0 Or orders(i).StartDate > dMax Then dMax = orders(i).StartDate
    Next i

    For i = 1 To nAlloc
        sumAlloc = sumAlloc + alloc(i).UsageT
    Next i

    MsgBox "Compounding run summary" & vbCrLf & _
           "HorizonDays: " & horizonDays & vbCrLf & _
           "EffCap per batch (t): " & Format(effCapPerBatch, "0.000") & vbCrLf & _
           "Orders in window: " & nOrders & _
           "   Total demand (t): " & Format(sumOrder, "0.000") & vbCrLf & _
           "Batches created: " & nbatches & "   Date span: " & _
               IIf(dMin = 0, "-", Format(dMin, "yyyy-mm-dd")) & " ~ " & _
               IIf(dMax = 0, "-", Format(dMax, "yyyy-mm-dd")) & vbCrLf & _
           "Allocated (t): " & Format(sumAlloc, "0.000") & _
           "   Unallocated (t): " & Format(sumOrder - sumAlloc, "0.000"), _
           vbInformation, "Compounding"
End Sub


'  ??????(? StartDate ??)
Private Sub SortOrdersByStartDate(ByRef a() As FGOrder, ByVal n&)
    Dim i&, j&
    Dim key As FGOrder
    For i = 2 To n
        key = a(i)
        j = i - 1
        Do While j >= 1 And a(j).StartDate > key.StartDate
            a(j + 1) = a(j)
            j = j - 1
        Loop
        a(j + 1) = key
    Next i
End Sub

' ????????,????????(????????)
Private Sub MergeTailIfPossible(ByRef batches() As Batch, ByRef nbatches&, _
                                ByRef alloc() As AllocationLine, ByRef nAlloc&)
    If nbatches < 2 Then Exit Sub

    Const TAIL_EPS As Double = 1#     ' ??:????<=1 ??????(????,?? 0.5)
    Dim last&, prev&, tailT#, headroom#
    last = nbatches: prev = nbatches - 1

    tailT = batches(last).AllocatedT
    headroom = batches(prev).EffCapT - batches(prev).AllocatedT

    ' ?????????????????;????????(?????????????)
    If tailT > 0# And tailT <= headroom And tailT <= TAIL_EPS Then
        Dim i&, moved#
        For i = 1 To nAlloc
            If alloc(i).BatchNo = last Then
                ' ???????????????
                If alloc(i).StartDate >= batches(prev).FirstStart And alloc(i).StartDate <= batches(prev).LastStart Then
                    ' ???????
                    alloc(i).BatchNo = prev
                    alloc(i).anchor = batches(prev).anchor
                    moved = moved + alloc(i).UsageT
                Else
                    ' ????????,????
                    moved = 0#: Exit For
                End If
            End If
        Next i

        If moved > 0# And Abs(moved - tailT) < 0.0001 Then
            ' ?????
            batches(prev).AllocatedT = RoundTo(batches(prev).AllocatedT + moved, 4)
            batches(last).AllocatedT = 0#
            ' ??????(???????;????????)
            nbatches = nbatches - 1
        End If
    End If
End Sub

Public Sub Sanity_Post_AllocWindow()
    Dim ws As Worksheet: Set ws = GetSheetByNameSafe(SHEET_ALLOC, True)  ' "CompoundingAllocation"
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim r As Long, bad As Long
    For r = 2 To lastRow
        Dim startD As Date, validThru As Date, anchor As Date, windowDays As Long
        startD = NzDate(ws.Cells(r, "B").Value)  ' Start date ??B
        anchor = NzDate(ws.Cells(r, "E").Value)  ' ?????? Anchor ??E(????????)
        ' ? Allocation ?????? ValidThru,?? Anchor+WindowDays-1 ??
        ' ?? WindowDays ??????,???? Window(d) ???? Allocation,???????
        windowDays = ReadWindowDays()
        validThru = DateAdd("d", windowDays - 1, anchor)

        If startD > validThru Then bad = bad + 1
    Next

    If bad > 0 Then
        Err.Raise 1021, , bad & " allocation row(s) violate batch window (Start date > Valid thru)."
    End If
End Sub

' ?????? alloc() ???? [iStart..iEnd] ?ì????î??:
' ???????,???????(??????);???????,???????
Public Sub EnforceMinTonsForRange(ByRef alloc() As tAlloc, ByRef nAlloc As Long, _
                                  ByVal iStart As Long, ByVal iEnd As Long, _
                                  ByVal fgType As String)
    Dim minT As Double, i As Long, j As Long
    minT = ReadMinTonsByFGType(fgType)
    If minT <= 0# Then Exit Sub
    If iStart < 1 Or iEnd < iStart Or iEnd > nAlloc Then Exit Sub

    i = iStart
    Do While i <= iEnd
        If alloc(i).UsageT >= minT Then
            i = i + 1
        Else
            If i < iEnd Then
                ' ?????(??)
                alloc(i + 1).UsageT = RoundTo(alloc(i + 1).UsageT + alloc(i).UsageT, 4)
                ' ???? i:????
                For j = i To nAlloc - 1
                    alloc(j) = alloc(j + 1)
                Next j
                nAlloc = nAlloc - 1
                iEnd = iEnd - 1
                ' ?? i ??,??????? i ????
            Else
                ' ??????:?????
                If i > iStart Then
                    alloc(i - 1).UsageT = RoundTo(alloc(i - 1).UsageT + alloc(i).UsageT, 4)
                    For j = i To nAlloc - 1
                        alloc(j) = alloc(j + 1)
                    Next j
                    nAlloc = nAlloc - 1
                    iEnd = iEnd - 1
                    i = i - 1            ' ????????????
                Else
                    ' ??????????:??(????)
                    Exit Do
                End If
            End If
        End If
    Loop
End Sub

Public Sub RebalanceBatchesByCap( _
    ByRef alloc() As tAlloc, ByVal nAlloc As Long, _
    ByRef batches() As tBatch, ByRef nbatches As Long, _
    ByVal effCap As Double, ByVal windowDays As Long)

    Dim b As Long, k As Long, idx As Long
    Dim sumT As Double

    If nAlloc <= 0 Or nbatches <= 0 Then Exit Sub

    For b = 1 To nbatches
        sumT = 0#
        For k = 1 To nAlloc
            If alloc(k).BatchNo = b Then
                sumT = RoundTo(sumT + alloc(k).UsageT, 4)
            End If
        Next k

        Do While sumT > effCap + 0.0001
            idx = 0
            For k = nAlloc To 1 Step -1
                If alloc(k).BatchNo = b Then
                    idx = k: Exit For
                End If
            Next k
            If idx = 0 Then Exit Do

            If b = nbatches Then
                nbatches = nbatches + 1
                batches(nbatches).BatchNo = nbatches
                batches(nbatches).anchor = DateAdd("d", windowDays, batches(nbatches - 1).anchor)
            End If

            ' ????????;ValidThru ??????
            alloc(idx).BatchNo = b + 1
            alloc(idx).anchor = batches(b + 1).anchor

            sumT = RoundTo(sumT - alloc(idx).UsageT, 4)
        Loop
    Next b
End Sub

' ? alloc() ?ì??????î????????:
' ? = OrderID + BatchNo + Anchor + StartDate + EndDate
Public Sub ConsolidateAllocAdjacent(ByRef alloc() As tAlloc, ByRef nAlloc As Long)
    Dim i As Long, j As Long
    If nAlloc <= 1 Then Exit Sub

    i = 2
    Do While i <= nAlloc
        If alloc(i).BatchNo = alloc(i - 1).BatchNo _
           And alloc(i).orderId = alloc(i - 1).orderId _
           And alloc(i).anchor = alloc(i - 1).anchor _
           And alloc(i).StartDate = alloc(i - 1).StartDate _
           And alloc(i).EndDate = alloc(i - 1).EndDate Then

            ' ????
            alloc(i - 1).UsageT = RoundTo(alloc(i - 1).UsageT + alloc(i).UsageT, 4)
            ' ??? i ?:????
            For j = i To nAlloc - 1
                alloc(j) = alloc(j + 1)
            Next j
            nAlloc = nAlloc - 1
            ' ????????(??? i ???)
        Else
            i = i + 1
        End If
    Loop
End Sub

' ===============================================================
' Rebuilds CompoundingBatchSummary *from CompoundingAllocation*.
' “¿¿µº´…Ÿ£∫÷±Ω”¥” SHEET_CFG(Hº¸/I÷µ) ∂¡»° Useable vol / Batch window / Horizon days°£
' ‘⁄ÕÍ≥…£∫◊Ó–°∂÷Œª∫œ≤¢ °˙ »›¡øªÿ ’ °˙ œ‡¡⁄∫œ≤¢ °˙ –¥ÕÍ Allocation ∫Ûµ˜”√°£
' Ωˆ÷ÿ–¥ Summary ±Ì£¨≤ª∏ƒ Allocation°£
' –Ë“™£∫GetSheetByNameSafe, FindCol, SHEET_ALLOC, SHEET_SUM, SHEET_CFG
' ===============================================================
Public Sub RebuildBatchSummaryFromAllocation()
    Dim wsA As Worksheet, wsS As Worksheet, wsCfg As Worksheet
    Dim lastRow As Long, r As Long
    Dim cStart As Long, cEnd As Long, cBatch As Long, cAnchor As Long, cUsage As Long
    Dim maxBatch As Long, b As Long
    Dim effCap As Double, windowDays As Long, horizonDays As Long
    
    On Error GoTo EH
    
    Set wsA = GetSheetByNameSafe(SHEET_ALLOC, True)
    Set wsS = GetSheetByNameSafe(SHEET_SUM, True)
    Set wsCfg = GetSheetByNameSafe(SHEET_CFG, True)
    
    ' Allocation ±ÿ“™¡–∂®Œª
    cStart = FindCol(wsA, Array("Start date", "Startdate", "Start"))
    cEnd = FindCol(wsA, Array("End date", "Enddate", "End"))
    cBatch = FindCol(wsA, Array("Batch#", "Batch No", "BatchNo"))
    cAnchor = FindCol(wsA, Array("Anchor"))
    cUsage = FindCol(wsA, Array("usage (t)", "UsageT", "usage_t"))
    
    If cBatch = 0 Or cUsage = 0 Then Err.Raise 513, , "Allocation sheet missing required columns (Batch#/usage)."
    
    lastRow = wsA.Cells(wsA.Rows.Count, cBatch).End(xlUp).Row
    
    ' –¥±ÌÕ∑
    wsS.Cells.ClearContents
    wsS.Range("A1:J1").Value = Array("Batch#", "Anchor", "Valid thru", "First start", "Last start", _
                                     "Allocated (t)", "EffCap (t)", "Remaining (t)", "Window(d)", "Horizon(d)")
    If lastRow < 2 Then Exit Sub
    
    ' ÷±Ω”¥”≈‰÷√±Ì∂¡»°≤Œ ˝£®Hº¸/I÷µ£©
    effCap = SafeReadCfgDouble(wsCfg, "Useable vol", 37.5)
    windowDays = CLng(SafeReadCfgDouble(wsCfg, "Batch window", 21))
    horizonDays = CLng(SafeReadCfgDouble(wsCfg, "Horizon days", 42))
    
    ' ◊Ó¥Û≈˙∫≈
    maxBatch = 0
    For r = 2 To lastRow
        If Len(wsA.Cells(r, cBatch).Value) > 0 Then
            b = CLng(wsA.Cells(r, cBatch).Value)
            If b > maxBatch Then maxBatch = b
        End If
    Next r
    If maxBatch = 0 Then Exit Sub
    
    ' ¿€º”∆˜
    Dim sumT() As Double, minStart() As Date, maxStart() As Date, anchorArr() As Date, hasBatch() As Boolean
    ReDim sumT(1 To maxBatch)
    ReDim minStart(1 To maxBatch)
    ReDim maxStart(1 To maxBatch)
    ReDim anchorArr(1 To maxBatch)
    ReDim hasBatch(1 To maxBatch)
    
    ' ∞¥≈˙æ€∫œ
    Dim curB As Long, u As Double, s As Date, anc As Date, sv As Variant, av As Variant
    For r = 2 To lastRow
        If Len(wsA.Cells(r, cBatch).Value) > 0 Then
            curB = CLng(wsA.Cells(r, cBatch).Value)
            If curB >= 1 And curB <= maxBatch Then
                u = ToDouble(wsA.Cells(r, cUsage).Value, 0#)
                sumT(curB) = Round4(sumT(curB) + u)
                
                sv = wsA.Cells(r, cStart).Value
                If IsDate(sv) Then
                    s = CDate(sv)
                    If minStart(curB) = 0 Or s < minStart(curB) Then minStart(curB) = s
                    If maxStart(curB) = 0 Or s > maxStart(curB) Then maxStart(curB) = s
                End If
                
                av = wsA.Cells(r, cAnchor).Value
                If IsDate(av) Then anchorArr(curB) = CDate(av)
                
                hasBatch(curB) = True
            End If
        End If
    Next r
    
    '  ‰≥ˆ Summary
    Dim outRow As Long: outRow = 2
    For b = 1 To maxBatch
        If hasBatch(b) Then
            wsS.Cells(outRow, 1).Value = b
            
            ' Anchor£∫”≈œ»”√æ€∫œµΩµƒ Anchor£ª»Ù»± ß£¨”√…œ“ª≈˙ Anchor + windowDays Õ∆µº
            If anchorArr(b) = 0 And outRow > 2 Then
                anchorArr(b) = DateAdd("d", windowDays, wsS.Cells(outRow - 1, 2).Value)
            End If
            wsS.Cells(outRow, 2).Value = anchorArr(b)
            If anchorArr(b) <> 0 Then
                wsS.Cells(outRow, 3).Value = DateAdd("d", windowDays - 1, anchorArr(b))
            End If
            
            wsS.Cells(outRow, 4).Value = IIf(minStart(b) = 0, Empty, minStart(b))
            wsS.Cells(outRow, 5).Value = IIf(maxStart(b) = 0, Empty, maxStart(b))
            
            ' πÿº¸–ﬁ∏¥£∫∞¥ Allocation ’Ê÷µº∆À„
            wsS.Cells(outRow, 6).Value = Round3(sumT(b))                             ' Allocated (t)
            wsS.Cells(outRow, 7).Value = effCap                                      ' EffCap (t)
            wsS.Cells(outRow, 8).Value = Round3(Application.Max(0, effCap - sumT(b))) ' Remaining (t)
            
            wsS.Cells(outRow, 9).Value = windowDays
            wsS.Cells(outRow, 10).Value = horizonDays
            
            outRow = outRow + 1
        End If
    Next b
    Exit Sub
EH:
    MsgBox "RebuildBatchSummaryFromAllocation failed: " & Err.Description, vbExclamation
End Sub

' ======= ±æµÿ–°π§æﬂ£®±‹√‚Õ‚≤ø“¿¿µ£© =======
Private Function SafeReadCfgDouble(wsCfg As Worksheet, keyText As String, defaultVal As Double) As Double
    Dim lastRow As Long, r As Long, v As Variant
    lastRow = wsCfg.Cells(wsCfg.Rows.Count, "H").End(xlUp).Row
    For r = 1 To lastRow
        If Trim$(CStr(wsCfg.Cells(r, "H").Value)) = keyText Then
            v = wsCfg.Cells(r, "I").Value
            SafeReadCfgDouble = ToDouble(v, defaultVal)
            Exit Function
        End If
    Next r
    SafeReadCfgDouble = defaultVal
End Function

Private Function ToDouble(v As Variant, defaultVal As Double) As Double
    If IsNumeric(v) Then
        ToDouble = CDbl(v)
    Else
        ToDouble = defaultVal
    End If
End Function

Private Function Round4(x As Double) As Double
    Round4 = WorksheetFunction.Round(x, 4)
End Function

Private Function Round3(x As Double) As Double
    Round3 = WorksheetFunction.Round(x, 3)
End Function

' ???????(???????????)????????
Public Sub SortOrdersByStartThenEnd(ByRef orders() As FGOrder, ByVal nOrders As Long)
    Dim i As Long, j As Long
    Dim tmp As FGOrder
    Dim s1 As Date, s2 As Date, e1 As Date, e2 As Date

    If nOrders <= 1 Then Exit Sub

    ' ?? O(n^2) ??;??????,????????
    For i = 1 To nOrders - 1
        For j = i + 1 To nOrders
            s1 = orders(i).StartDate: e1 = orders(i).EndDate
            s2 = orders(j).StartDate: e2 = orders(j).EndDate

            If (s1 > s2) Or (s1 = s2 And e1 > e2) Then
                tmp = orders(i)
                orders(i) = orders(j)
                orders(j) = tmp
            End If
        Next j
    Next i
End Sub

'??????(CompoundingTab)? SRC Sheet ????;???????????
Private Function ResolveRunSheet() As Worksheet
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim ws As Worksheet
    Dim srcName As String

    ' 1) ?????? SRC Sheet(H???,I???)
    srcName = ReadCfgKey_Simple("CompoundingTab", "SRC Sheet")
    If Len(Trim$(srcName)) > 0 Then
        On Error Resume Next
        Set ws = wb.Worksheets(srcName)
        On Error GoTo 0
        If Not ws Is Nothing Then
            If FindCol(ws, Array("Order ID")) > 0 Then
                Set ResolveRunSheet = ws
                Exit Function
            End If
        End If
    End If

    ' 2) ??:????????????ìOrder IDî??
    Dim candidates As Variant, i As Long
    candidates = Array("Compounding_Test_3", "Compounding_Test_2", "Compounding_Test_1", _
                       "Compounding_ECC extraction", "Extract")

    For i = LBound(candidates) To UBound(candidates)
        On Error Resume Next
        Set ws = wb.Worksheets(CStr(candidates(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            If FindCol(ws, Array("Order ID")) > 0 Then
                Set ResolveRunSheet = ws
                Exit Function
            End If
            Set ws = Nothing
        End If
    Next i

    ' 3) ????
    Set ResolveRunSheet = Nothing
End Function

'????(CompoundingTab)???????:H?=?,I?=?
Private Function ReadCfgKey_Simple(ByVal cfgSheetName As String, ByVal keyText As String) As String
    Dim ws As Worksheet, lastRow As Long, r As Range
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(cfgSheetName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    lastRow = ws.Cells(ws.Rows.Count, "H").End(xlUp).Row
    If lastRow < 1 Then Exit Function

    Set r = ws.Range("H1:H" & lastRow).Find(What:=keyText, LookAt:=xlWhole, MatchCase:=False)
    If Not r Is Nothing Then
        ReadCfgKey_Simple = CStr(ws.Cells(r.Row, "I").Value)
    End If
End Function


Public Sub TestResolve()
    Dim t As Worksheet
    Set t = ResolveRunSheet()
    If t Is Nothing Then
        MsgBox "Nothing found"
    Else
        MsgBox "ResolveRunSheet got: " & t.name
    End If
End Sub

