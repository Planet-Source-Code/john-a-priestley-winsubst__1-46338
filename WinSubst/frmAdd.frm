VERSION 5.00
Begin VB.Form frmAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Subst"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4260
   Icon            =   "frmAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdPath 
      Caption         =   "..."
      Height          =   285
      Left            =   3840
      TabIndex        =   2
      ToolTipText     =   "Select Path"
      Top             =   120
      Width           =   285
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      ToolTipText     =   "Path"
      Top             =   120
      Width           =   3015
   End
   Begin VB.ComboBox comboDrive 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Drive Letter"
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*                                                                            *
'*                                  WinSubst                                  *
'*                                                                            *
'*                               Add Subst Form                               *
'*                                                                            *
'******************************************************************************
Option Explicit

Private mvarDrive As String
Private mvarPath As String

'******************************************************************************
'* Cancel Selection                                                           *
'******************************************************************************
Private Sub cmdCancel_Click()
    mvarPath = ""
    mvarDrive = ""
    Unload Me
End Sub

'******************************************************************************
'* Okay Selection                                                             *
'******************************************************************************
Private Sub cmdOkay_Click()
    If Len(txtPath.Text) = 0 Then Exit Sub
    
    mvarPath = txtPath.Text
    mvarDrive = comboDrive

    Unload Me
End Sub

'******************************************************************************
'* Select Path                                                                *
'******************************************************************************
Private Sub cmdPath_Click()
    Dim fDir As frmDirectory
    Dim strPath As String

    Set fDir = New frmDirectory
    fDir.Show vbModal
    strPath = fDir.Path
    Set fDir = Nothing
    
    If Len(strPath) > 0 Then txtPath.Text = strPath
End Sub

'******************************************************************************
'* Form Loading                                                               *
'******************************************************************************
Private Sub Form_Load()
    mvarPath = ""
    mvarDrive = ""
    RefreshDriveCombo
    
    If comboDrive.ListCount = 0 Then
        MsgBox "There are no free drive letters to add a new subst." & vbCrLf & "You must delete a substed drive to add a new one.", vbExclamation + vbOKOnly, "Add Subst"
    End If
End Sub

'******************************************************************************
'* Refresh combo drive box                                                    *
'******************************************************************************
Private Sub RefreshDriveCombo()
    Dim cSubst As clsSubst
    Dim iDrive As Integer
    Dim strDrive As String

    Set cSubst = New clsSubst
    cSubst.ScanDrives
    
    comboDrive.Clear
    For iDrive = 0 To 25
        strDrive = Chr$(Asc("A") + iDrive) & ":"
        If cSubst.DriveType(strDrive) = DT_NOTDEFINED Then
            comboDrive.AddItem strDrive
        End If
    Next

    If comboDrive.ListCount > 0 Then
        comboDrive.ListIndex = 0
    End If

    Set cSubst = Nothing
End Sub

'******************************************************************************
'* Path Property - Read Only                                                  *
'******************************************************************************
Public Property Get Path() As String
    Path = mvarPath
End Property

'******************************************************************************
'* Drive Property - Read Only                                                 *
'******************************************************************************
Public Property Get Drive() As String
    Drive = mvarDrive
End Property
