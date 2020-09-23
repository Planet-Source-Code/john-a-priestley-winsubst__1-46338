VERSION 5.00
Begin VB.Form frmDirectory 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Directory"
   ClientHeight    =   2955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4230
   Icon            =   "frmDirectory.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2955
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdOkay 
      Caption         =   "Okay"
      Height          =   375
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.DriveListBox driveList 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   2520
      Width           =   2655
   End
   Begin VB.DirListBox dirList 
      Height          =   2340
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
   End
End
Attribute VB_Name = "frmDirectory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************************
'*                                                                            *
'*                                  WinSubst                                  *
'*                                                                            *
'*                               Directory Form                               *
'*                                                                            *
'******************************************************************************
Option Explicit

Private mvarPath As String

'******************************************************************************
'* Cancel Selection                                                           *
'******************************************************************************
Private Sub cmdCancel_Click()
    mvarPath = ""
    Unload Me
End Sub

'******************************************************************************
'* Okay Selection                                                             *
'******************************************************************************
Private Sub cmdOkay_Click()
    mvarPath = dirList.Path
    Unload Me
End Sub

'******************************************************************************
'* Drive list changed                                                         *
'******************************************************************************
Private Sub driveList_Change()
    dirList.Path = UCase$(driveList.Drive)
End Sub

'******************************************************************************
'* Form Loading                                                               *
'******************************************************************************
Private Sub Form_Load()
    mvarPath = ""
End Sub

'******************************************************************************
'* Path Property - Read Only                                                  *
'******************************************************************************
Public Property Get Path() As String
    Path = mvarPath
End Property
