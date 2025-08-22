Attribute VB_Name = "Module1"
' ???:Tools ? References ??
'    Microsoft Visual Basic for Applications Extensibility 5.3
' ?? Excel:File ? Options ? Trust Center ? Trust Center Settings ? Macro Settings
'    ?? “Trust access to the VBA project object model”

Sub ExportAllStdModules()
    Dim vbComp As VBIDE.VBComponent
    Dim outDir As String

    ' ? OneDrive ?????(????????????)
    outDir = Environ$("OneDrive") & "\Yi' Career\Mogan-Consulting B.V\BD\Yixing\compounding-vba-project\src\"

    If Right$(outDir, 1) <> "\" Then outDir = outDir & "\"
    If Dir(outDir, vbDirectory) = "" Then
        MsgBox "???????: " & outDir, vbExclamation
        Exit Sub
    End If

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        If vbComp.Type = vbext_ct_StdModule Then
            vbComp.Export outDir & vbComp.name & ".bas"
        End If
    Next vbComp

    MsgBox "Modules ???? ? " & outDir, vbInformation
End Sub

