VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'******************************************************************************
'*                                                                            *
'*                                Subst Class                                 *
'*                                                                            *
'******************************************************************************
Option Explicit

Private Declare Function DefineDosDevice Lib "kernel32" Alias "DefineDosDeviceA" (ByVal dwFlags As Long, ByVal lpDeviceName As String, Optional ByVal lpTargetPath As String = vbNullString) As Long
Private Declare Function QueryDosDevice Lib "kernel32" Alias "QueryDosDeviceA" (ByVal lpDeviceName As String, ByVal lpTargetPath As String, ByVal ucchMax As Long) As Long

Private Const DDD_RAW_TARGET_PATH = &H1
Private Const DDD_REMOVE_DEFINITION = &H2
Private Const DDD_EXACT_MATCH_ON_REMOVE = &H4

Public Enum DriveTypeEnum
    DT_NOTDEFINED = 0
    DT_FLOPPY = 1
    DT_HARDDRIVE = 2
    DT_CDROM = 3
    DT_LAN = 4
    DT_SUBST = 5
    DT_UNKNOWN = 99
End Enum

Private Type DriveTypeType
    dtt_DriveType As DriveTypeEnum      ' Drive Type (See DriveTypeEnum)
    dtt_Path As String                  ' Path
    dtt_Substed As Boolean              ' Is Subst for this drive active?
End Type

Private mvarDrive(0 To 25) As DriveTypeType
Private mvarDefinedPath(0 To 25) As String

'******************************************************************************
'* Add Subst drive                                                            *
'******************************************************************************
Public Function AddSubst(ByVal strDrive As String, ByVal strPath As String) As Boolean
On Error GoTo ERR_ADD
    Dim lResult As Long
    Dim iDrive As Integer
    Dim iPos As Integer

    If Len(strDrive) <> 2 Then
        AddSubst = False
        Exit Function
    End If
    
    If Right$(strDrive, 1) <> ":" Then
        AddSubst = False
        Exit Function
    End If
    
    strDrive = UCase$(strDrive)
    If Left$(strDrive, 1) < "A" Or Left$(strDrive, 1) > "Z" Then
        AddSubst = False
        Exit Function
    End If
    
    If Len(strPath) = 0 Then
        AddSubst = False
        Exit Function
    End If
    
    If Len(Dir$(strPath, vbDirectory)) = 0 Then
        AddSubst = False
        Exit Function
    End If

    iDrive = Asc(Left$(strDrive, 1)) - Asc("A")
    
    If mvarDrive(iDrive).dtt_DriveType = DT_SUBST Then DeleteSubst strDrive

    lResult = DefineDosDevice(0&, strDrive, strPath)
    If lResult = 0 Then
        AddSubst = False
        mvarDefinedPath(iDrive) = ""
    Else
        AddSubst = True
        mvarDefinedPath(iDrive) = strPath
    End If

    ScanDrives

    Exit Function
ERR_ADD:
    If Err.Number = 52 Then
        iPos = InStr(strPath, "?")
        If iPos > 0 Then
            strPath = Left$(strPath, iPos - 1)
            Resume
        End If
    End If
    AddSubst = False
End Function

'******************************************************************************
'* Can Subst Drive Property - Read Only                                       *
'******************************************************************************
Public Property Get CanSubstDrive(ByVal strDrive As String) As Boolean
    Dim iDrive As Integer

    If Len(strDrive) <> 2 Then
        CanSubstDrive = False
        Exit Property
    End If
    
    If Right$(strDrive, 1) <> ":" Then
        CanSubstDrive = False
        Exit Property
    End If
    
    strDrive = UCase$(strDrive)
    If Left$(strDrive, 1) < "A" Or Left$(strDrive, 1) > "Z" Then
        CanSubstDrive = False
        Exit Property
    End If

    iDrive = Asc(Left$(strDrive, 1)) - Asc("A")
    If mvarDrive(iDrive).dtt_DriveType = DT_SUBST Or mvarDrive(iDrive).dtt_DriveType = DT_NOTDEFINED Then
        CanSubstDrive = True
    Else
        CanSubstDrive = False
    End If
End Property

'******************************************************************************
'* Initialise Class                                                           *
'******************************************************************************
Private Sub Class_Initialize()
    Dim iDrive As Integer
    
    For iDrive = 0 To 25
        mvarDrive(iDrive).dtt_DriveType = DT_NOTDEFINED
        mvarDrive(iDrive).dtt_Path = ""
        mvarDrive(iDrive).dtt_Substed = False
    Next
End Sub

'******************************************************************************
' Delete Subst Drive                                                          *
'******************************************************************************
Public Sub DeleteSubst(ByVal strDrive As String)
    Dim lResult As Long
    Dim iDrive As Integer

    If Len(strDrive) <> 2 Then Exit Sub
    If Right$(strDrive, 1) <> ":" Then Exit Sub
    
    strDrive = UCase$(strDrive)
    If Left$(strDrive, 1) < "A" Or Left$(strDrive, 1) > "Z" Then Exit Sub
    
    iDrive = Asc(Left$(strDrive, 1)) - Asc("A")
    
    If mvarDrive(iDrive).dtt_DriveType = DT_SUBST Then
        lResult = DefineDosDevice(DDD_REMOVE_DEFINITION, strDrive)
        mvarDefinedPath(iDrive) = ""
    End If
    ScanDrives
End Sub

'******************************************************************************
' Drive Type Property - Read Only                                             *
'******************************************************************************
Public Property Get DriveType(ByVal strDrive As String) As DriveTypeEnum
    Dim iDrive As Integer
    
    If Len(strDrive) <> 2 Then
        DriveType = DT_UNKNOWN
        Exit Property
    End If
    
    If Right$(strDrive, 1) <> ":" Then
        DriveType = DT_UNKNOWN
        Exit Property
    End If
    
    strDrive = UCase$(strDrive)
    If Left$(strDrive, 1) < "A" Or Left$(strDrive, 1) > "Z" Then
        DriveType = DT_UNKNOWN
        Exit Property
    End If
    
    iDrive = Asc(Left$(strDrive, 1)) - Asc("A")
    DriveType = mvarDrive(iDrive).dtt_DriveType
End Property

'******************************************************************************
'* Drive Path Property - Read Only                                            *
'******************************************************************************
Public Property Get DrivePath(ByVal strDrive As String) As String
    Dim iDrive As Integer
    
    If Len(strDrive) <> 2 Then
        DrivePath = DT_UNKNOWN
        Exit Property
    End If
    
    If Right$(strDrive, 1) <> ":" Then
        DrivePath = DT_UNKNOWN
        Exit Property
    End If
    
    strDrive = UCase$(strDrive)
    If Left$(strDrive, 1) < "A" Or Left$(strDrive, 1) > "Z" Then
        DrivePath = DT_UNKNOWN
        Exit Property
    End If
    
    iDrive = Asc(Left$(strDrive, 1)) - Asc("A")
    DrivePath = mvarDrive(iDrive).dtt_Path
End Property

'******************************************************************************
'* Drive Substed Property - Read Only                                         *
'******************************************************************************
Public Property Get DriveSubsted(ByVal strDrive As String) As Boolean
    Dim iDrive As Integer
    
    If Len(strDrive) <> 2 Then
        DriveSubsted = DT_UNKNOWN
        Exit Property
    End If
    
    If Right$(strDrive, 1) <> ":" Then
        DriveSubsted = DT_UNKNOWN
        Exit Property
    End If
    
    strDrive = UCase$(strDrive)
    If Left$(strDrive, 1) < "A" Or Left$(strDrive, 1) > "Z" Then
        DriveSubsted = DT_UNKNOWN
        Exit Property
    End If
    
    iDrive = Asc(Left$(strDrive, 1)) - Asc("A")
    DriveSubsted = mvarDrive(iDrive).dtt_Substed
End Property

'******************************************************************************
'* Load Defined Paths                                                         *
'******************************************************************************
Public Sub LoadDefinedPaths()
    Dim iFile As Integer
    Dim strData As String
    Dim strCmd As String
    Dim strOption As String
    Dim iPos As Integer

    If Len(Dir$(App.Path & "\winsubst.dat")) = 0 Then Exit Sub
    
    iFile = FreeFile
    Open App.Path & "\winsubst.dat" For Input Access Read As #iFile
    Do While EOF(iFile) = False
        Line Input #iFile, strData
        strData = Trim$(strData)
        '
        ' Check to see if the line is empty or starts with a remark (#)
        '
        If Len(strData) <> 0 And Left$(strData, 1) <> "#" Then
            iPos = InStr(strData, "=")
            If iPos > 0 Then
                strCmd = UCase$(Trim$(Left$(strData, iPos - 1)))
                strOption = Trim$(Mid$(strData, iPos + 1))
                
                If Len(strCmd) = 1 And strCmd >= "A" And strCmd <= "Z" Then
                    Call AddSubst(strCmd & ":", strOption)
                End If
            End If
        End If
    Loop
    Close #iFile
End Sub

'******************************************************************************
'* Save Defined Paths                                                         *
'******************************************************************************
Public Sub SaveDefinedPaths()
    Dim iFile As Integer
    Dim iDrive As Integer
    Dim strData As String

    iFile = FreeFile
    Open App.Path & "\winsubst.dat" For Output Access Write As #iFile
    
    Print #iFile, "#" & vbCrLf & "# WinSubst Data File" & vbCrLf & "#"
    For iDrive = 0 To 25
        If mvarDrive(iDrive).dtt_DriveType = DT_SUBST Then
            Print #iFile, Chr$(Asc("A") + iDrive) & "=" & mvarDrive(iDrive).dtt_Path
        End If
    Next
    Close #iFile
End Sub

'******************************************************************************
'* Scan Drives                                                                *
'******************************************************************************
Public Sub ScanDrives()
    Dim iDrive As Integer
    Dim strInfo As String
    Dim lResult As Long
    Dim strDrive As String
    Dim strParam() As String
    Dim iPos As Integer

    '
    ' Loop though drives
    '
    For iDrive = 0 To 25
        mvarDrive(iDrive).dtt_DriveType = DT_NOTDEFINED
        mvarDrive(iDrive).dtt_Path = ""
        mvarDrive(iDrive).dtt_Substed = False

        '
        ' Prepare drive letter
        '
        strDrive = Chr$(Asc("A") + iDrive) & ":"
        strInfo = String$(255, Chr$(0))
        '
        ' Query drive
        '
        lResult = QueryDosDevice(strDrive, strInfo, Len(strInfo))
        If lResult <> 0 Then
            strInfo = Replace(strInfo, Chr$(0), "")
            If Left$(strInfo, 1) = "\" Then strInfo = Mid$(strInfo, 2)
            '
            ' Break the information down
            '
            Erase strParam
            strParam = Split(strInfo, "\")
            '
            '
            '
            strParam(0) = LCase$(strParam(0))
            Select Case strParam(0)
                Case "device"   'Floppy, HD, CD, LAN
                    strParam(1) = LCase$(strParam(1))
                    If InStr(strParam(1), "harddisk") = 1 Then
                        '
                        ' Harddrive
                        '
                        mvarDrive(iDrive).dtt_DriveType = DT_HARDDRIVE
                        mvarDrive(iDrive).dtt_Path = "Harddrive " & Mid$(strParam(1), 15)
                        mvarDrive(iDrive).dtt_Substed = False
                    ElseIf InStr(strParam(1), "floppy") = 1 Then
                        '
                        ' Floppy
                        '
                        mvarDrive(iDrive).dtt_DriveType = DT_FLOPPY
                        mvarDrive(iDrive).dtt_Path = "Floppy " & Mid$(strParam(1), 7)
                        mvarDrive(iDrive).dtt_Substed = False
                    ElseIf InStr(strParam(1), "cdrom") = 1 Then
                        '
                        ' CDRom
                        '
                        mvarDrive(iDrive).dtt_DriveType = DT_CDROM
                        mvarDrive(iDrive).dtt_Path = "CDRom " & Mid$(strParam(1), 6)
                        mvarDrive(iDrive).dtt_Substed = False
                    ElseIf InStr(strParam(1), "lanman") = 1 Then
                        '
                        ' LAN Mapped Drive
                        '
                        mvarDrive(iDrive).dtt_DriveType = DT_LAN
                        mvarDrive(iDrive).dtt_Substed = False
                        mvarDrive(iDrive).dtt_Path = "\"
                        For iPos = 3 To UBound(strParam)
                            mvarDrive(iDrive).dtt_Path = mvarDrive(iDrive).dtt_Path & "\" & strParam(iPos)
                        Next
                    Else
                        '
                        ' Unknown
                        '
                        mvarDrive(iDrive).dtt_DriveType = DT_UNKNOWN
                        mvarDrive(iDrive).dtt_Substed = False
                        For iPos = 2 To UBound(strParam)
                            If Len(mvarDrive(iDrive).dtt_Path) > 0 Then mvarDrive(iDrive).dtt_Path = mvarDrive(iDrive).dtt_Path & "\"
                            mvarDrive(iDrive).dtt_Path = mvarDrive(iDrive).dtt_Path & strParam(iPos)
                        Next
                    End If
                    
                Case "??"       'Subst
                    mvarDrive(iDrive).dtt_DriveType = DT_SUBST
                    mvarDrive(iDrive).dtt_Substed = True
                    For iPos = 1 To UBound(strParam)
                        If Len(mvarDrive(iDrive).dtt_Path) > 0 Then mvarDrive(iDrive).dtt_Path = mvarDrive(iDrive).dtt_Path & "\"
                        mvarDrive(iDrive).dtt_Path = mvarDrive(iDrive).dtt_Path & strParam(iPos)
                        mvarDefinedPath(iDrive) = mvarDrive(iDrive).dtt_Path
                    Next

                Case Else       'Dont know
                    mvarDrive(iDrive).dtt_DriveType = DT_UNKNOWN
                    mvarDrive(iDrive).dtt_Substed = False
                    For iPos = 1 To UBound(strParam)
                        If Len(mvarDrive(iDrive).dtt_Path) > 0 Then mvarDrive(iDrive).dtt_Path = mvarDrive(iDrive).dtt_Path & "\"
                        mvarDrive(iDrive).dtt_Path = mvarDrive(iDrive).dtt_Path & strParam(iPos)
                    Next
            End Select
        Else
            '
            ' Subst not running but defined
            '
            If Len(mvarDefinedPath(iDrive)) > 0 Then
                If Len(Dir$(mvarDefinedPath(iDrive), vbDirectory)) > 0 Then
                    mvarDrive(iDrive).dtt_DriveType = DT_SUBST
                    mvarDrive(iDrive).dtt_Path = mvarDefinedPath(iDrive)
                    mvarDrive(iDrive).dtt_Substed = False
                End If
            End If
        End If
    Next
End Sub

'******************************************************************************
'* Start Drive                                                                *
'******************************************************************************
Public Sub StartSubst(ByVal strDrive As String)
    Dim iDrive As Integer
    Dim lResult As Long

    If Len(strDrive) <> 2 Then Exit Sub
    If Right$(strDrive, 1) <> ":" Then Exit Sub
    strDrive = Left$(UCase$(strDrive), 1)
    If strDrive < "A" Or strDrive > "Z" Then Exit Sub
    
    iDrive = Asc(strDrive) - Asc("A")
    
    ScanDrives

    If mvarDrive(iDrive).dtt_DriveType <> DT_SUBST Then Exit Sub
    If mvarDrive(iDrive).dtt_Substed = True Then Exit Sub
    
    lResult = DefineDosDevice(0&, strDrive & ":", mvarDrive(iDrive).dtt_Path)
    If lResult <> 0 Then mvarDrive(iDrive).dtt_Substed = True
End Sub

'******************************************************************************
'* Stop Drive                                                                 *
'******************************************************************************
Public Sub StopSubst(ByVal strDrive As String)
    Dim iDrive As Integer
    Dim lResult As Long

    If Len(strDrive) <> 2 Then Exit Sub
    If Right$(strDrive, 1) <> ":" Then Exit Sub
    strDrive = Left$(UCase$(strDrive), 1)
    If strDrive < "A" Or strDrive > "Z" Then Exit Sub
    
    iDrive = Asc(strDrive) - Asc("A")
    
    ScanDrives

    If mvarDrive(iDrive).dtt_DriveType <> DT_SUBST Then Exit Sub
    If mvarDrive(iDrive).dtt_Substed = False Then Exit Sub
    
    lResult = DefineDosDevice(DDD_REMOVE_DEFINITION, strDrive & ":")
    If lResult <> 0 Then mvarDrive(iDrive).dtt_Substed = False
End Sub
