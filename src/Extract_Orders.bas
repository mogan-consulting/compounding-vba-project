Attribute VB_Name = "Extract_Orders"
Sub Extract_orders()

Dim Mainbook As String
Dim Extract As String
Dim Savedfolderlist As String
Dim Coois_variant As String
Dim Display_variant As String
Dim MyMessage

        Mainbook = ActiveWorkbook.name
           
'-----Clean old extract file from system-------------
        Savedfolderlist = Dir("C:\temp\coois.xls")
        If Savedfolderlist = "" Then
        Range("A2").Select
        Else
        Kill "C:\temp\coois.xls"
        End If
'----------------------------------------
    
        
        'Call Initiate_Values
        If Not Find_Correct_Session("ECP") Then
            MsgBox ("You have no SAP connection opened, please open one one ECC session")
            End
        End If

Session.findById("wnd[0]").maximize
Session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncoois"
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/tbar[1]/btn[17]").press
Session.findById("wnd[1]/usr/txtV-LOW").Text = "/Comp_Auto"
Session.findById("wnd[1]/usr/txtENAME-LOW").Text = ""
Session.findById("wnd[1]/usr/txtENAME-LOW").SetFocus
Session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 0
Session.findById("wnd[1]").sendVKey 8
Session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").Text = "/Comp_Auto"
Session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").SetFocus
Session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").caretPosition = 10
Session.findById("wnd[0]").sendVKey 0
Session.findById("wnd[0]/tbar[1]/btn[8]").press
Session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarButton "&NAVIGATION_PROFILE_TOOLBAR_EXPAND"
Session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").pressToolbarContextButton "&MB_EXPORT"
Session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").selectContextMenuItem "&PC"
Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
Session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "C:\temp\"
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "coois.xls"
Session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 9
Session.findById("wnd[1]/tbar[0]/btn[0]").press
Session.findById("wnd[0]/tbar[0]/btn[3]").press

Call Release_SAP_Session


Windows(Mainbook).Activate
Sheets("Extract").Select
Rows("2:2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp

Range("A2").Select


'Open Coois Extract
ChDir "C:\Temp"
    Workbooks.OpenText Filename:="C:\Temp\coois.xls", Origin:=xlWindows, _
        StartRow:=1, DataType:=xlDelimited, TextQualifier:=xlDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, Comma:=False _
        , Space:=False, Other:=False, FieldInfo:=Array(Array(1, 1), Array(2, 1), _
        Array(3, 1), Array(4, 1), Array(5, 1), Array(6, 1)), TrailingMinusNumbers:=True
    Cells.Select
    Cells.EntireColumn.AutoFit
    Columns("A:A").Select
    Range("A13").Activate
    Selection.Delete Shift:=xlToLeft
    Rows("1:3").Select
    Selection.Delete Shift:=xlUp
 
Extract = ActiveWorkbook.name



'-----iF there is no data in extract, close file and diplay error

    If Range("B3").Value = "List contains no data" Then
    Windows(Mainbook).Activate
    MyMessage = MsgBox("No data", vbOKOnly, "Attention!!")
    GoTo Jump_close
    End If

 Columns("A:A").Select
    Selection.SpecialCells(xlCellTypeBlanks).Select
    Selection.EntireRow.Delete
   

 'cOPY AND PASTE IN MAIN SHEET
    Range("A2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select
    Selection.Copy
    Windows(Mainbook).Activate
    Range("A2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select



Jump_close:
'close extract file
    Windows(Extract).Activate
    Application.DisplayAlerts = False
    ActiveWindow.Close
    Application.DisplayAlerts = True

'Categorise by size
    Range("F2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-1],CompoundingTab!C[-5]:C[-4],2,0)"
    Range("F2").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    ActiveCell.Select
    Application.CutCopyMode = False


'Input Factor & Usage
    Range("G2").Select
    ActiveCell.FormulaR1C1 = "=VLOOKUP(RC[-2],CompoundingTab!C[-6]:C[-4],3,0)"
    Range("H2").Select
    ActiveCell.FormulaR1C1 = "=RC[-4]*RC[-1]*1.07/1000/1000"
    Range("G2:H2").Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Range("A1").Select
    Selection.End(xlDown).Select
    ActiveCell.Offset(0, 1).Range("A1").Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("A1").Select


'copy data to main sheet
Sheets("Compounding_ECC Extraction").Select
    Range("A2").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.ClearContents
    Range("A2").Select
    Sheets("Extract").Select
    Range("A2:D2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("Compounding_ECC Extraction").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Extract").Select
    Range("F2:H2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Compounding_ECC Extraction").Select
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Range("A2").Select

'formatting
Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With



If MyMessage = "" Then
MyMessage = MsgBox("Data extracted successfully", vbOKOnly, "Attention!!")
End If


End Sub



