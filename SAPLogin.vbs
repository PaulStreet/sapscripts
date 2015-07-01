REM The following script was written to log into the SAP server automatically.
REM To view historical information and credit for this script please see
REM the following thread on the SAP Community Network:
REM http://scn.sap.com/thread/3763970

REM This script was last updated by Paul Street on 7/1/15

REM Directives
    Option Explicit

  REM Variables!  Must declare before using because of Option Explicit
    Dim WSHShell, SAPGUIPath, SID, InstanceNo, WinTitle, SapGuiAuto, application, connection, session

  REM Main
    Set WSHShell = WScript.CreateObject("WScript.Shell")
    If IsObject(WSHShell) Then

      REM Set the path to the SAP GUI directory
        SAPGUIPath = "C:\Program Files\SAP\FrontEnd\SAPgui\"

      REM Set the SAP system ID
        SID = "NBP"

      REM Set the instance number of the SAP system
        InstanceNo = "00"

      REM Starts the SAP GUI
        WSHShell.Exec SAPGUIPath & "SAPgui.exe " & SID & " " & _
          InstanceNo
 
      REM Set the title of the SAP GUI window here
        WinTitle = "SAP"
 
      While Not WSHShell.AppActivate(WinTitle)
        WScript.Sleep 250
      Wend
 
      Set WSHShell = Nothing
    End If
 
	REM Remove this if you need to test the above script and want a message box at the end launching the login screen.
    REM MsgBox "Here now your script..."
 
If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/txtRSYST-BNAME").text = "USERNAME"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").text = "PASSWORD"
session.findById("wnd[0]/usr/pwdRSYST-BCODE").setFocus
session.findById("wnd[0]/usr/pwdRSYST-BCODE").caretPosition = 10
session.findById("wnd[0]").sendVKey 0
