Attribute VB_Name = "mdlMain"
'******************************************************************************
'*                                                                            *
'*                                  WinSubst                                  *
'*                                                                            *
'*                        GUI Version of Subst for DOS                        *
'*                                                                            *
'*                              By John Priestley                             *
'*                                                                            *
'******************************************************************************
Option Explicit
'
' If you want the drives to be mapped when you start your computer
' put a shortcut of the program into your Startup folder, open the properties and
' add /map at the end of the command.
'


'
'
'
Public Sub Main()
    Dim strCmd As String
    Dim cSubst As clsSubst

    strCmd = Trim$(LCase$(Command$()))
    If strCmd = "/map" Then
        Set cSubst = New clsSubst
        cSubst.ScanDrives
        cSubst.LoadDefinedPaths
        Set cSubst = Nothing
    Else
        Load frmMain
        frmMain.Show
    End If
End Sub
