Attribute VB_Name = "Shared_Functions"
Function Initiate_Values()
    
    Set Prod = Sheets("Production Order to Create")
    'Set TempB = Sheets("Temp_BOM_Extract")
    
    Application.DisplayAlerts = False
    
    

    'record user calculation mode
    CalculationMode = Application.Calculation
    'set calculation mode to manual
    Application.Calculation = xlManual
    
End Function




Function Clean_Sheet(Feuil As Worksheet, FirstLine As Long)
    'macro to delete all exiting lines of a given sheet (except the headers line)
    On Error Resume Next
    Feuil.Activate
    HeaderLine = FirstLine - 1
    If FirstLine = 1 Then HeaderLine = 1
    Rows(HeaderLine).AutoFilter
    Rows(HeaderLine).AutoFilter
    Range(Rows(FirstLine), Rows(Feuil.UsedRange.Rows.Count + 10)).Delete
    
End Function


Function CleanEmptyColumns(S As Worksheet)
       
       'delete all columns empty (ie no values)
       S.Activate
       col = 1
       While col < S.UsedRange.Columns.Count And col < 50
        
            If Application.CountA(S.Columns(col)) = 0 Then
                S.Columns(col).Delete
            Else
                col = col + 1
            End If
       
       Wend
       
End Function

Function CleanEmptyRows(S As Worksheet)
       
       'delete all columns empty (ie no values)
       S.Activate
       lin = 1
       While lin < S.UsedRange.Rows.Count And lin < 30
        
            If Application.CountA(S.Rows(lin)) = 0 Then
                S.Rows(lin).Delete
            Else
                lin = lin + 1
            End If
       
       Wend
       
End Function


Function Release_SAP_Session()

    Set Session = Nothing
    Set Connection = Nothing
    Set SAPGUI = Nothing
    Set SAP = Nothing
    Set Connections = Nothing
    Set Sessions = Nothing
    
      
End Function


Function Check_Sbar() As Boolean
      Dim saperror As String
      Check_Sbar = True
        'if the message is just a warning then send another enter to by pass
        If Session.findById("wnd[0]/sbar").MessageType = "E" Then
            'if this is an error then record the error text to display to the user
            saperror = Session.findById("wnd[0]/sbar").Text
            Check_Sbar = False
        End If
        
        While Session.findById("wnd[0]/sbar").MessageType = "W"
            Session.findById("wnd[0]").sendVKey 0
        Wend
        

End Function


Function Find_SOP_family(ItemLookUp As String)
    'look up the SOP famliy for item
    'function used several time
    lineSOP = Application.Match(ItemLookUp, SOP.Columns(colItemF), 0)
    If IsError(lineSOP) Then
        Find_SOP_family = "Please Update SOP Family (from WW Access DB) tab, ask WW planner to do so (it is a downlaod from WW SOP database)"
        SOP_Sub_Family = ""
    Else
        Find_SOP_family = SOP.Cells(lineSOP, ColItemFamilyF)
        SOP_Sub_Family = SOP.Cells(lineSOP, ColItemSubFamilyF)
    End If

End Function


Function List_Plant_Products()
    
    'for each item in production list, check if item alrady listed in Job aid, if not then add it
    PList.Activate
    Cells.Select
    Selection.Sort Key1:=Columns(colMaterialPR), order1:=xlAscending, Header:=xlYes
    
    i = 2
    While PList.Cells(i, colMaterialPR) <> ""
        lineJobAid = Application.Match(PList.Cells(i, colMaterialPR), Job.Columns(ColMaterialJ), 0)
        If IsError(lineJobAid) Then
            Job.Cells(Application.CountA(Job.Columns(ColMaterialJ)) + 1, ColMaterialJ) = PList.Cells(i, colMaterialPR)
            
        End If
        i = i + Application.CountIf(PList.Columns(colMaterialPR), PList.Cells(i, colMaterialPR))
    Wend


End Function
Function Convert_to_Text(col As Long)
'
' Convert_to_Text Macro
'
    Columns(col).Select
    Selection.TextToColumns Destination:=Columns(col), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 2), TrailingMinusNumbers:=True

End Function

Function Convert_to_Number(col As Long, SCMDecimalSep As String, SCMThousandSep As String)
'
' Convert_to_Text Macro
'
    Columns(col).Select
    Selection.TextToColumns Destination:=Columns(col), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), DecimalSeparator:=SCMDecimalSep, ThousandsSeparator:=SCMThousandSep, TrailingMinusNumbers:=True

End Function


Function Convert_to_Date(col As Long, DateImportFormat As Long)
'
' Convert_to_Text Macro
'
    Columns(col).Select
    Selection.TextToColumns Destination:=Columns(col), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, DateImportFormat), TrailingMinusNumbers:=True
   
End Function

Sub Add_Commands_Menu(str As String)
    
    'Add the macros to manage Check and FOSY Import
    Dim cBut As CommandBarButton
    
    'On Error Resume Next
        
        
        'add several commmands
        Set cBut = Application.CommandBars("Cell").Controls.Add(temporary:=True)
        With cBut
            .Caption = "----- To " & str & " -------"
            .Style = msoButtonCaption
         End With
        Select Case str
            Case "Create Production Orders"
                Set cBut = Application.CommandBars("Cell").Controls.Add(temporary:=True)
                With cBut
                    .Caption = "Create Production Orders"
                    .Style = msoButtonCaption
                    .OnAction = "Create_Prod_Orders_CO01"
                End With
            
        End Select
        
        
        
        
        
End Sub


Function Delete_Commands_Menu(str As String)
    
    'delete all right click macros
    Dim cBut As CommandBarButton
    On Error Resume Next

           Application.CommandBars("Cell").Controls(str).Delete
           Application.CommandBars("Cell").Controls("----- To " & str & " -------").Delete
           
           
End Function

Function Name_Range()

Dim n As Integer
Job.Activate

n = Application.WorksheetFunction.CountA(Job.Columns(ColMaterialJ))


ActiveWorkbook.Names.Add name:="Items", RefersToR1C1:= _
        "='Job Aid'!R2C" & ColMaterialJ & ":R" & n & "C" & ColMaterialJ

End Function

Function Latest_Run_Date(ReportName As String)

    lineReport = Application.Match(ReportName, Job.Columns(colReportNameJ), 0)
    If Not IsError(lineReport) Then
        Job.Cells(lineReport, colUpdateDateJ) = Now
    End If
  
End Function





Function Convert_to_Qty(col As Long)

    Columns(col).Select
    Selection.TextToColumns Destination:=Columns(col), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
    Range("M12").Select
    
End Function

Function R3DateNumberFormat()
'Make sur that a R3 session is opened
'we will use the SU01D transaction
        'go to SU01D to catch format used in R3
    If Not Find_Correct_Session("R3") Then
        MsgBox ("You have no " & "R3" & " connection opened, please open one and rerun the macro")
        End
    End If
        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nsu3"
        Session.findById("wnd[0]").sendVKey 0
        'go to default tab
        Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA").Select
        
        If Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/subMAINAREA:SAPLSUU5:0105/radRADIO_DATFM1").Selected = True Then
            DateSAP = Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/subMAINAREA:SAPLSUU5:0105/radRADIO_DATFM1").Text
            R3DateFormat = "DDMMYYYY"
            R3ImportDateFormat = 4

        ElseIf Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/subMAINAREA:SAPLSUU5:0105/radRADIO_DATFM2").Selected = True Or Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/subMAINAREA:SAPLSUU5:0105/radRADIO_DATFM3").Selected = True Then
            DateSAP = Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/subMAINAREA:SAPLSUU5:0105/radRADIO_DATFM1").Text
            R3DateFormat = "MMDDYYYY"
            R3ImportDateFormat = 3
                
        ElseIf Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/subMAINAREA:SAPLSUU5:0105/radRADIO_DATFM4").Selected = True Or Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/subMAINAREA:SAPLSUU5:0105/radRADIO_DATFM5").Selected = True Or Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/subMAINAREA:SAPLSUU5:0105/radRADIO_DATFM6").Selected = True Then
            DateSAP = Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/subMAINAREA:SAPLSUU5:0105/radRADIO_DATFM1").Text
            R3DateFormat = "YYYYMMDD"
            R3ImportDateFormat = 5
        End If
        
        'comma / decimal
        If Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/subMAINAREA:SAPLSUU5:0105/radRADIO_DCPFM2").Selected = True Then
            
            R3DecimalSep = "."
            R3ThousandSep = ","
           
        Else
            R3DecimalSep = ","
            R3ThousandSep = "."
        End If
        
        'determine the language of the session to adapt to Production order text send
        If Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/subMAINAREA:SAPLSUU5:0105/radRADIO_DCPFM1").Text = "Virg." Then
            'french version
            ProductionOrderStatus = "Ouv."
            DateFinReel = "Date fin réelle"
            DateReleaseReel = "Date lancem. réelle"
        Else
            'enlgish version
            ProductionOrderStatus = "CRTE"
            DateFinReel = "Actual finish date"
            DateReleaseReel = "Release date (a*"
            
        End If
        
        Call Release_SAP_Session
            
End Function


Function SCMDateNumberFormat(SearchedSAP As String)
'Make sur that a SCM session is opened
'we will use the SU01D transaction
        'go to SU01D to catch format used in SCM
        If Not Find_Correct_Session(SearchedSAP) Then
            MsgBox ("You have no " & SearchedSAP & " connection opened, please open one a rerun macro")
            End
        End If
        
        
        ID = Session.info.user
        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSU01D"
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").Text = ID
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/tbar[1]/btn[7]").press
        Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA").Select
        DateSAP = Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DATFM").Value
        NumberSAP = Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").Value
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        
        '__________________Update Date format from SCM ___________________________________
        Select Case Left(DateSAP, 1)
            Case Is = "D"
                SCMDateFormat = "DDMMYYYY"
                SCMDateImport = 4
            Case Is = "M"
                SCMDateFormat = "MMDDYYYY"
                SCMDateImport = 3
            Case Is = "Y"
                SCMDateFormat = "YYYYMMDD"
                SCMDateImport = 5
        End Select
        '_____________________Update Number format from SCM ______________________________
        ' 3 case for number
            '1.234.567,89 -> first
            '1 234 567,89 -> second
            '1,234,456.89 -> third
        If InStr(NumberSAP, " ") <> 0 Then
            'second case
            SCMDecimalSep = ","
            SCMThousandSep = " "
        Else
            If InStr(Left(NumberSAP, 2), ".") <> 0 Then
                'first case
                SCMDecimalSep = ","
                SCMThousandSep = "."
            Else
                SCMDecimalSep = "."
                SCMThousandSep = ","
            End If
        End If
 
End Function


Function ECCDateNumberFormat(SearchedSAP As String)
'Make sur that a ECC session is opened
'we will use the SU01D transaction
        'go to SU01D to catch format used in ECC
        If Not Find_Correct_Session(SearchedSAP) Then
            MsgBox ("You have no " & SearchedSAP & " connection opened, please open one a rerun macro")
            End
        End If
        
        
        ID = Session.info.user
        Session.findById("wnd[0]/tbar[0]/okcd").Text = "/nSU01D"
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/usr/ctxtSUID_ST_BNAME-BNAME").Text = ID
        Session.findById("wnd[0]").sendVKey 0
        Session.findById("wnd[0]/tbar[1]/btn[7]").press
        Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA").Select
        DateSAP = Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DATFM").Value
        NumberSAP = Session.findById("wnd[0]/usr/tabsTABSTRIP1/tabpDEFA/ssubMAINAREA:SAPLSUID_MAINTENANCE:1105/cmbSUID_ST_NODE_DEFAULTS-DCPFM").Value
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        Session.findById("wnd[0]/tbar[0]/btn[3]").press
        
        '__________________Update Date format from ECC ___________________________________
        Select Case Left(DateSAP, 1)
            Case Is = "D"
                ECCDateFormat = "DDMMYYYY"
                ECCDateImport = 4
            Case Is = "M"
                ECCDateFormat = "MMDDYYYY"
                ECCDateImport = 3
            Case Is = "Y"
                ECCDateFormat = "YYYYMMDD"
                ECCDateImport = 5
        End Select
        '_____________________Update Number format from ECC ______________________________
        ' 3 case for number
            '1.234.567,89 -> first
            '1 234 567,89 -> second
            '1,234,456.89 -> third
        If InStr(NumberSAP, " ") <> 0 Then
            'second case
            ECCDecimalSep = ","
            ECCThousandSep = " "
        Else
            If InStr(Left(NumberSAP, 2), ".") <> 0 Then
                'first case
                ECCDecimalSep = ","
                ECCThousandSep = "."
            Else
                ECCDecimalSep = "."
                ECCThousandSep = ","
            End If
        End If
 
End Function

Function ExcelDateFormat()
'Macro that will determine what are the user set of parameters as per date format
'thiusand sperator and decimal
'these are important data to achieve to import succesfully the file exported from SAP

    Dim DateOrder As String
    
    Dim DateSeparator As String
   
    With Application
        
        Select Case .International(xlDateOrder)
            Case Is = 0
                DateOrder = "month-day-year"
                ExcelDateFormat = 3
                DateSeparator = .International(xlDateSeparator)
                ExcelDecimalSep = .International(xlDecimalSeparator)
                ExcelThousandSep = .International(xlThousandsSeparator)
                ExcelDateImport = "MMDDYYYY"
                
            Case Is = 1
                DateOrder = "day-month-year"
                ExcelDateFormat = 4
                DateSeparator = .International(xlDateSeparator)
                ExcelDecimalSep = .International(xlDecimalSeparator)
                ExcelThousandSep = .International(xlThousandsSeparator)
                ExcelDateImport = "DDMMYYYY"
            Case Is = 2
                'hungarian case,
                ExcelDateFormat = 5
                DateOrder = "year-month-day"
                DateSeparator = .International(xlDateSeparator)
                ExcelDecimalSep = .International(xlDecimalSeparator)
                ExcelThousandSep = .International(xlThousandsSeparator)
                ExcelDateImport = "YYYYMMDD"
            Case Else
                DateOrder = "Error"
        End Select
       
        
        
    End With

End Function


Function keep_ECC_Awake()
'sending one trigger to ecc to keep the session opened


    If Not Find_Correct_Session("ECC") Then
        MsgBox ("You have no " & "ECC" & " connection opened, please open one a rerun macro Extract_InterCoSo_PurchaseOrders")
        End
    End If
    
    'call for va05n
    Session.findById("wnd[0]/tbar[0]/okcd").Text = "/n"
    Session.findById("wnd[0]").sendVKey 0
    
    Call Release_SAP_Session
End Function



Function Increment_String(CurrentString As String, NewValue As String)

    If CurrentString = "" Then
        Increment_String = "'" & NewValue
    Else
        Increment_String = CurrentString & Chr(10) & NewValue
    End If
End Function


Public Sub KillProperly(Directory, Killfile)
     If Len(Dir$(Directory & Killfile)) > 0 Then
        For Each wbk In Workbooks
           If wbk.name = Killfile Then
               Workbooks(Killfile).Close savechanges:=False
           End If
        Next
        SetAttr Directory & Killfile, vbNormal
        Kill Directory & Killfile
     End If
 End Sub

Function GetDateLastUpdated(AddrFichier)

    Dim oFS As Object
    Dim strFilename As String

    'This creates an instance of the MS Scripting Runtime FileSystemObject class
    Set oFS = CreateObject("Scripting.FileSystemObject")

    GetDateLastUpdated = oFS.GetFile(AddrFichier).Datelastmodified

    Set oFS = Nothing

End Function






Sub CreateSheet(Nom As String)
     
    toAdd = True
    For Each Sheet In ActiveWorkbook.Sheets
        If Sheet.name = Nom Then
            toAdd = False
            Sheet.Activate
        End If
    Next
    If toAdd = True Then
        Sheets.Add After:=Sheets(Sheets.Count)
        ActiveSheet.name = Nom

    End If
End Sub


