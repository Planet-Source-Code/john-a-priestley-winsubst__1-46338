VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "WinSubst"
   ClientHeight    =   5085
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   5325
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5085
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvDrives 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Drive"
         Object.Width           =   1235
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Type"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Path"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Status"
         Object.Width           =   1587
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFile_Exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuSubst 
      Caption         =   "&Subst"
      Begin VB.Menu mnuSubst_Add 
         Caption         =   "&Add Subst..."
      End
      Begin VB.Menu mnuSubst_Delete 
         Caption         =   "&Delete Subst"
      End
      Begin VB.Menu mnuSubst_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubst_Start 
         Caption         =   "&Start Subst"
      End
      Begin VB.Menu mnuSubst_Stop 
         Caption         =   "S&top Subst"
      End
      Begin VB.Menu mnuSubst_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSubst_Refresh 
         Caption         =   "&Refresh"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp_About 
         Caption         =   "&About..."
      End
   End
   Begin VB.Menu mnuPopupList 
      Caption         =   "PopupList"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupList_Add 
         Caption         =   "Add Subst..."
      End
      Begin VB.Menu mnuPopupList_Delete 
         Caption         =   "Delete Subst"
      End
      Begin VB.Menu mnuPopupList_Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupList_Start 
         Caption         =   "Start Subst"
      End
      Begin VB.Menu mnuPopupList_Stop 
         Caption         =   "Stop Subst"
      End
      Begin VB.Menu mnuPopupList_Sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupList_Refresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private mvarSubst As clsSubst

'******************************************************************************
'* Form Loading                                                               *
'******************************************************************************
Private Sub Form_Load()
    Set mvarSubst = New clsSubst
    
    mvarSubst.ScanDrives
    mvarSubst.LoadDefinedPaths

    RefreshDriveList
    lvDrives_Click
End Sub

'******************************************************************************
'* Form Closing                                                               *
'******************************************************************************
Private Sub Form_Unload(Cancel As Integer)
    mvarSubst.SaveDefinedPaths
    Set mvarSubst = Nothing
End Sub

'******************************************************************************
' Refresh Drive List                                                          *
'******************************************************************************
Private Sub RefreshDriveList()
    Dim node As ListItem
    Dim iDrive As Integer
    Dim iType As DriveTypeEnum
    Dim strPath As String
    Dim bSubsted As Boolean

    lvDrives.ListItems.Clear
    mvarSubst.ScanDrives
    
    For iDrive = 0 To 25
        iType = mvarSubst.DriveType(Chr$(Asc("A") + iDrive) & ":")
        strPath = mvarSubst.DrivePath(Chr$(Asc("A") + iDrive) & ":")
        bSubsted = mvarSubst.DriveSubsted(Chr$(Asc("A") + iDrive) & ":")
        
        If iType <> DT_NOTDEFINED Then
            Set node = lvDrives.ListItems.Add(, "DRIVE" & Format(iDrive, "00"), Chr$(Asc("A") + iDrive))
            Select Case iType
                Case DriveTypeEnum.DT_FLOPPY, DriveTypeEnum.DT_HARDDRIVE, DriveTypeEnum.DT_CDROM
                    node.ListSubItems.Add , , "Device"
                Case DriveTypeEnum.DT_LAN
                    node.ListSubItems.Add , , "LAN"
                Case DriveTypeEnum.DT_SUBST
                    node.ListSubItems.Add , , "SUBST"
                Case Else
                    node.ListSubItems.Add , , "UNKNOWN"
            End Select
            
            node.ListSubItems.Add , , strPath

            If bSubsted = True Then
                node.ListSubItems.Add , , "Mapped"
            Else
                node.ListSubItems.Add , , ""
            End If
            
            If iType <> DT_SUBST Then
                node.ForeColor = RGB(0, 140, 0)
                node.ListSubItems(1).ForeColor = node.ForeColor
                node.ListSubItems(2).ForeColor = node.ForeColor
                node.ListSubItems(3).ForeColor = node.ForeColor
            End If
        End If
    Next
End Sub

'******************************************************************************
'* Clicked Drive List                                                         *
'******************************************************************************
Private Sub lvDrives_Click()
    Dim iDrive As Integer
    Dim strDrive As String

    If lvDrives.ListItems.Count = 0 Then Exit Sub
    iDrive = CInt(Mid$(lvDrives.SelectedItem.Key, 6))
    strDrive = Chr$(Asc("A") + iDrive) & ":"
    
    If mvarSubst.CanSubstDrive(strDrive) = True Then
        mnuSubst_Delete.Enabled = True
        mnuSubst_Start.Enabled = True
        mnuSubst_Stop.Enabled = True
        
        mnuPopupList_Delete.Enabled = True
        mnuPopupList_Start.Enabled = True
        mnuPopupList_Stop = True
    Else
        mnuSubst_Delete.Enabled = False
        mnuSubst_Start.Enabled = False
        mnuSubst_Stop.Enabled = False
        
        mnuPopupList_Delete.Enabled = False
        mnuPopupList_Start.Enabled = False
        mnuPopupList_Stop = False
    End If
End Sub

'******************************************************************************
'* Mouse Click                                                                *
'******************************************************************************
Private Sub lvDrives_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton Then
        lvDrives_Click
        PopupMenu mnuPopupList
    End If
End Sub

'******************************************************************************
'* File Menu - Exit                                                           *
'******************************************************************************
Private Sub mnuFile_Exit_Click()
    Unload Me
End Sub

'******************************************************************************
'* Help Menu - About                                                          *
'******************************************************************************
Private Sub mnuHelp_About_Click()

End Sub

'******************************************************************************
'* Popup List - Add                                                           *
'******************************************************************************
Private Sub mnuPopupList_Add_Click()
    mnuSubst_Add_Click
End Sub

'******************************************************************************
'* Popup List - Delete                                                        *
'******************************************************************************
Private Sub mnuPopupList_Delete_Click()
    mnuSubst_Delete_Click
End Sub

'******************************************************************************
'* Popup List - Refresh                                                       *
'******************************************************************************
Private Sub mnuPopupList_Refresh_Click()
    mnuSubst_Refresh_Click
End Sub

'******************************************************************************
'* Popup List - Start                                                         *
'******************************************************************************
Private Sub mnuPopupList_Start_Click()
    mnuSubst_Start_Click
End Sub

'******************************************************************************
'* Popup List - Stop                                                          *
'******************************************************************************
Private Sub mnuPopupList_Stop_Click()
    mnuSubst_Stop_Click
End Sub

'******************************************************************************
'* Subst Menu - Add                                                           *
'******************************************************************************
Private Sub mnuSubst_Add_Click()
    Dim fAdd As frmAdd
    Dim strDrive As String
    Dim strPath As String

    Set fAdd = New frmAdd
    If fAdd.comboDrive.ListCount > 0 Then fAdd.Show vbModal
    
    strDrive = fAdd.Drive
    strPath = fAdd.Path

    Set fAdd = Nothing

    If Len(strDrive) = 0 Then Exit Sub
    
    Call mvarSubst.AddSubst(strDrive, strPath)
    RefreshDriveList
End Sub

'******************************************************************************
'* Subst Menu - Delete                                                        *
'******************************************************************************
Private Sub mnuSubst_Delete_Click()
    Dim iDrive As Integer

    If lvDrives.ListItems.Count = 0 Then Exit Sub
    iDrive = CInt(Mid$(lvDrives.SelectedItem.Key, 6))
    If mvarSubst.CanSubstDrive(Chr$(Asc("A") + iDrive) & ":") = False Then
        MsgBox "You cannot delete Drive " & Chr$(Asc("A") + iDrive) & ":", vbExclamation + vbOKOnly, "Delete Drive"
        Exit Sub
    End If
    
    mvarSubst.DeleteSubst Chr$(Asc("A") + iDrive) & ":"
    RefreshDriveList
End Sub

'******************************************************************************
'* Refresh Subst List                                                         *
'******************************************************************************
Private Sub mnuSubst_Refresh_Click()
    RefreshDriveList
End Sub

'******************************************************************************
'* Subst Menu - Start                                                         *
'******************************************************************************
Private Sub mnuSubst_Start_Click()
    Dim iDrive As Integer

    If lvDrives.ListItems.Count = 0 Then Exit Sub
    iDrive = CInt(Mid$(lvDrives.SelectedItem.Key, 6))
    If mvarSubst.CanSubstDrive(Chr$(Asc("A") + iDrive) & ":") = False Then
        MsgBox "You cannot start Drive " & Chr$(Asc("A") + iDrive) & ":" & vbCrLf & "It is not a substed drive.", vbExclamation + vbOKOnly, "Start Drive"
        Exit Sub
    End If
    
    mvarSubst.StartSubst Chr$(Asc("A") + iDrive) & ":"
    RefreshDriveList
End Sub

'******************************************************************************
'* Subst Menu - Stop                                                          *
'******************************************************************************
Private Sub mnuSubst_Stop_Click()
    Dim iDrive As Integer

    If lvDrives.ListItems.Count = 0 Then Exit Sub
    iDrive = CInt(Mid$(lvDrives.SelectedItem.Key, 6))
    If mvarSubst.CanSubstDrive(Chr$(Asc("A") + iDrive) & ":") = False Then
        MsgBox "You cannot stop Drive " & Chr$(Asc("A") + iDrive) & ":" & vbCrLf & "It is not a substed drive.", vbExclamation + vbOKOnly, "Stop Drive"
        Exit Sub
    End If

    mvarSubst.StopSubst Chr$(Asc("A") + iDrive) & ":"
    RefreshDriveList
End Sub
