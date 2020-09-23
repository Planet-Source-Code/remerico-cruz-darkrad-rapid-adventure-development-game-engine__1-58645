Attribute VB_Name = "modFunc"
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function GetFileName(ByRef Fn As String, ByVal Filter As String, ByVal Title As String, Optional Save As Boolean = False) As Boolean
On Error Resume Next
Dim m_dlgDialog As cCommonDialog
    If Trim(Fn) <> "" Then GetFileName = True: Exit Function
    Set m_dlgDialog = New cCommonDialog
    If Save Then
        GetFileName = m_dlgDialog.VBGetSaveFileName(Fn, , True, Filter, , , Title)
    Else
        GetFileName = m_dlgDialog.VBGetOpenFileName(Fn, , True, False, False, True, Filter, , , Title)
    End If
    If Trim(Fn) = "" Then
        GetFileName = False
    End If
    Set m_dlgDialog = Nothing
End Function

Sub MakeTmpDir()
On Error Resume Next
MkDir WrkDir
End Sub

Public Function GetInitEntry(ByVal sSection As String, ByVal sKeyName As String, Optional ByVal sDefault As String = "", Optional ByVal sInitFileName As String = "") As String

'This Function Reads In a String From The Init File.
'Returns Value From Init File or sDefault If No Value Exists.
'sDefault Defaults to an Empty String ("").
'Creates and Uses sDefInitFileName (AppPath\AppEXEName.Ini)
'if sInitFileName Parameter Is Not Passed In.

Dim sBuffer As String
Dim sInitFile As String

    'If Init Filename NOT Passed In
    If Len(sInitFileName) = 0 Then
        'If Static Init FileName NOT Already Created
        If Len(sDefInitFileName) = 0 Then
            'Create Static Init FileName
            sDefInitFileName = App.Path
            If Right$(sDefInitFileName, 1) <> "\" Then
                sDefInitFileName = sDefInitFileName & "\"
            End If
            sDefInitFileName = sDefInitFileName & App.EXEName & ".ini"
        End If
        sInitFile = sDefInitFileName
    Else    'If Init Filename Passed In
        sInitFile = sInitFileName
    End If
    
    sBuffer = String$(2048, " ")
    GetInitEntry = Left$(sBuffer, GetPrivateProfileString(sSection, ByVal sKeyName, sDefault, sBuffer, Len(sBuffer), sInitFile))

End Function
Public Function SetInitEntry(ByVal sSection As String, Optional ByVal sKeyName As String, Optional ByVal sValue As String, Optional ByVal sInitFileName As String = "") As Long

'This Function Writes a String To The Init File.
'Returns WritePrivateProfileString Success or Error.
'Creates and Uses sDefInitFileName (AppPath\AppEXEName.Ini)
'if sInitFileName Parameter Is Not Passed In.

'***** CAUTION *****
'If sValue is Null then sKeyName is deleted from the Init File.
'If sKeyName is Null then sSection is deleted from the Init File.

Dim sInitFile As String

    'If Init Filename NOT Passed In
    If Len(sInitFileName) = 0 Then
        'If Static Init FileName NOT Already Created
        If Len(sDefInitFileName) = 0 Then
            'Create Static Init FileName
            sDefInitFileName = App.Path
            If Right$(sDefInitFileName, 1) <> "\" Then
                sDefInitFileName = sDefInitFileName & "\"
            End If
            sDefInitFileName = sDefInitFileName & App.EXEName & ".ini"
        End If
        sInitFile = sDefInitFileName
    Else    'If Init Filename Passed In
        sInitFile = sInitFileName
    End If
    
    If Len(sKeyName) > 0 And Len(sValue) > 0 Then
        SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, ByVal sValue, sInitFile)
    ElseIf Len(sKeyName) > 0 Then
        SetInitEntry = WritePrivateProfileString(sSection, ByVal sKeyName, vbNullString, sInitFile)
    Else
        SetInitEntry = WritePrivateProfileString(sSection, vbNullString, vbNullString, sInitFile)
    End If

End Function

Function GetFName(Path As String) As String

    For Findsep = 1 To Len(Path)
        If Mid(Path, Len(Path) - (Findsep - 1), 1) = "\" Or Mid(Path, Len(Path) - (Findsep - 1), 1) = "/" Then
            GetFName = Right(Path, Findsep - 1)
            Exit Function
        End If
    Next Findsep

End Function

Function GetPath(FullPath As String) As String
    
    Dim C As Integer
    Dim s As Integer
    Dim J As Integer

    C = 0: s = 0: J = 0
    
    For m = 1 To Len(FullPath)
        GetChr0 = Right(FullPath, m): GetChr1 = Left(GetChr0, 1)
        If GetChr1 = "\" Or GetChr1 = "/" Then C = C + 1
    Next m
    For m = 1 To Len(FullPath)
        GetChr0 = Left(FullPath, m): GetChr1 = Right(GetChr0, 1)
        J = J + 1
        If GetChr1 = "\" Or GetChr1 = "/" Then
            J = 0: s = s + 1
            If s = C Then GetPath = Right(GetChr0, m - J): Exit Function
        End If
    Next m

End Function

Function SetCaption(gTitle As String)
SetCaption = "Dark Adventure Toolkit (" + gTitle + ")"
End Function

Function GetPercent(wSoFar, wTotal)

If wTotal < wSoFar Then Exit Function

GetPercent = (wSoFar / wTotal) * 100
End Function

Sub Progress(pb As Control, ByVal Percent)
Dim num$
    If Not pb.AutoRedraw Then
      pb.AutoRedraw = -1
    End If
    
    If wCCVU = 1 Then
      If Percent >= 0 And Percent <= 33 Then
        pb.ForeColor = RGB(0, 255, 0)
      ElseIf Percent > 33 And Percent <= 75 Then
        pb.ForeColor = RGB(255, 255, 0)
      Else
        pb.ForeColor = RGB(255, 0, 0)
      End If
    End If
    
    pb.Cls
    pb.ScaleWidth = 100
    pb.DrawMode = 10
    num$ = Format$(Percent, "###") + "%"
    pb.CurrentX = 50 - pb.TextWidth(num$) / 2
    pb.CurrentY = (pb.ScaleHeight - pb.TextHeight(num$)) / 2
    'pb.Print num$
    pb.Line (0, 0)-(Percent, pb.ScaleHeight), , BF
    pb.Refresh
End Sub
