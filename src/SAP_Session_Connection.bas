Attribute VB_Name = "SAP_Session_Connection"
'Function to test if an SAP session is opened that correspond to the system (SCM / PR1 / ECC) the user has to act on
Public SAP, SAPGUI, Connections, Connection, Sessions, Session As Object
Public cntConnection, cntSession As Long

Function Find_Correct_Session(keyWord As String) As Boolean
        'Keyword  "SCP = SCM, ECP = ECC, PR1 = R3"
 On Error GoTo NoConnection
      
        Select Case keyWord
            Case "SCM"
                keyWord = "SCP"
            Case "ECC"
                keyWord = "ECP"
            Case "PR1"
                keyWord = "PR1"
            Case "SCQ"
                keyWord = "SCQ"
            Case "R3"
                keyWord = "PR1"
            Case "SCI"
                keyWord = "SCI"
        End Select
        
    Find_Correct_Session = False
      
      Set SAP = GetObject("SAPGUI")
      Set SAPGUI = SAP.GetScriptingEngine()
      Set Connections = SAPGUI.Connections()
      'count the number of SAPGUI connections
      'if several instance of SAP are Opened (SCM/PR1/ECC) then the number of connection > 1
      'Set Connection = SAPGUI.Connections(0)
      cntConnection = Connections.Count()

        '-Loop on connection to check each session of the connection-----------------------------

        For i = 0 To cntConnection - 1
          Set Connection = SAPGUI.Connections(CLng(i))
          If IsObject(Connection) Then

            Set Sessions = Connection.Sessions()
            If IsObject(Sessions) Then

              cntSession = Sessions.Count()

              '-Here the loop on the sessions------------------------
                'the first session of SCM/ECC is always  the netweaver which is a small screen
                'it is better to use /osmen session so we will start by the last session
                'an if the number of session is less than 5 then create a new one
                'if it is more than 1 then we are by defult in the gui
                For j = cntSession - 1 To 0 Step -1
                  Set Session = Connection.Sessions(CLng(j))
                  If IsObject(Session) Then

                    'for ECC and SCM we find the keyword in the PassportSystemId
                    If InStr(Session.passportsystemid, keyWord) > 0 Then
                        'Session is a global constant, thus it will be identified for the script
                        Find_Correct_Session = True
                        'send an /osmen to make sure to avoid netweaver format (as for laptop this can block the navigation
                        If cntSession = 1 Then
                            Session.findById("wnd[0]/tbar[0]/okcd").Text = "/osmen"
                            Session.findById("wnd[0]").sendVKey 0
                            Application.Wait (Now + TimeValue("00:00:02"))
                            Set Session = Connection.Sessions(CLng(j) + 1)
                        End If
                        Exit Function
                    End If '
                    'for PR it is in the Session Description
                    If InStr(Session.Parent.Description, keyWord) > 0 Then
                        Find_Correct_Session = True
                        Exit Function
                    End If '
                  End If
                Next
              End If
          End If
        Next
        
        Exit Function

'if there is an error then it means either the is no connection opened
NoConnection:
        MsgBox ("You have no SAP session Opened, the Netweaver is not enought you need at least one transaction screen")
        
End Function

