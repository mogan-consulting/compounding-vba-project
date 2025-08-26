Attribute VB_Name = "modPlanner"
Option Explicit

' ===== 顶层入口 =====
Public Sub RunCompoundingPlan()
    PlanFixed40Core True, True
    Call Sanity_Post_AllocWindow
End Sub

' 主流程
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


    ' 找关键列
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

    ' 读取参数
    Dim horizonDays&, windowDays&
    Dim effCapPerBatch#
    horizonDays = ReadHorizonDays()
    effCapPerBatch = ReadEffCapPerBatch()
    windowDays = ReadWindowDays()

    ' 读取源数据
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

    ' 生成批列表（按周六锚点，覆盖窗内）
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

    ' 校验一致性
    ValidateConsistency orders, nOrders, alloc, nAlloc, batches, nbatches

    ' 输出
    If writeSheets Then
        WriteAllocationSheet alloc, nAlloc
        WriteBatchSummarySheet batches, nbatches
        ShowRunSummary orders, nOrders, alloc, nAlloc, batches, nbatches, horizonDays, effCapPerBatch
    End If

    ' 可选文本输出
    If writeText Then
        ' 这里保留日志输出点（如需写出到文本，可实现 SaveTextFile）
    End If
End Sub

' ------- 读取源数据并换算 usage(t) -------
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

        ' 计划窗过滤：StartDate ∈ [today, today + horizonDays]
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
            .OrderID = oid
        
            .StartDate = dStart
            .EndDate = dEnd
            .PlanQty = qty
            .UsageT = UsageT
        End With


        'End If
ContinueRow:
    Next r

    If nOrders > 0 Then ReDim Preserve orders(1 To nOrders) Else Erase orders
End Sub

Private Function ComputeUsageT(ByVal PlanQty As Double) As Double
    ' usage(t) = qty × 1.07 × 10.4 ÷ 1,000 ÷ 1,000
    Dim t#
    t = PlanQty * FACTOR_LOSS * FACTOR_UNIT / 1000# / 1000#
    ComputeUsageT = RoundTo(t, 4) ' 中间保留 4 位
End Function

' ------- 批生成 -------
Private Sub BuildBatches(ByVal runDate As Date, ByVal horizonDays&, ByVal effCap#, ByVal windowDays&, _
                         ByRef batches() As Batch, ByRef nbatches&)
    Dim horizonEnd As Date
    horizonEnd = runDate + horizonDays

    ' 第一条锚点：从 runDate 对齐到最近的「本周或未来周六」
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
        anchor = anchor + 7 ' 下一周六
    Loop
End Sub

' ------- 分配逻辑（不超配，溢出后推） -------


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

' ------- 一致性校验 -------
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

    ' 不超配检查
    For i = 1 To nbatches
        If batches(i).AllocatedT - batches(i).EffCapT > 0.0001 Then
            Err.Raise 3002, , "Over-allocation on batch #" & batches(i).BatchNo
        End If
    Next i
End Sub

'  1ϸҳCompoundingAllocation
'  Anchor  Valid thru = Anchor + Window(d) - 1
'======================
Public Sub WriteAllocationSheet(ByRef alloc() As tAlloc, ByVal nAlloc As Long)
    Dim ws As Worksheet
    Dim i As Long, windowDays As Long
    Dim a() As Variant

    windowDays = ReadWindowDays()                      '  21
    Set ws = GetSheetByNameSafe(SHEET_ALLOC, True)     ' "CompoundingAllocation"

    ws.Cells.Clear
    ws.Range("A1").Resize(1, 8).Value = Array( _
        "Order ID", "Start date", "End date", "Batch#", "Anchor", "Valid thru", "usage (t)", "Horizon(d)")

    If nAlloc <= 0 Then Exit Sub

    ReDim a(1 To nAlloc, 1 To 8)
    For i = 1 To nAlloc
        a(i, 1) = alloc(i).OrderID
        a(i, 2) = alloc(i).StartDate
        a(i, 3) = alloc(i).EndDate
        a(i, 4) = alloc(i).BatchNo
        a(i, 5) = alloc(i).anchor
        a(i, 6) = DateAdd("d", windowDays - 1, alloc(i).anchor) ' Чڣ˵㣩
        a(i, 7) = RoundTo(alloc(i).UsageT, 3)
        a(i, 8) = alloc(i).horizonDays
    Next i

    ws.Range("A2").Resize(nAlloc, 8).Value = a

    ws.Columns("B:C").NumberFormat = "yyyy-mm-dd"      ' Start / End
    ws.Columns("E:F").NumberFormat = "yyyy-mm-dd"      ' Anchor / Valid thru
    ws.Columns("G:G").NumberFormat = "0.000"           ' usage (t)
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
        a(i, 3) = DateAdd("d", batches(i).windowDays - 1, batches(i).anchor) ' Чڣ˵㣩
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

    For i = 1 To nOrders
        remain = RoundTo(orders(i).UsageT, 4)
        If remain <= 0 Then GoTo ContinueOrder

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
                    .OrderID = orders(i).OrderID
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
ContinueOrder:
    Next i

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



