Attribute VB_Name = "modEngine"
' Main Engine

Public Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Const HWND_TOPMOST = -1
Const SWP_NOSIZE = &H1
Const SWP_NOMOVE = &H2

' Engine variables
Public WrkDir As String
Public DVar As New Collection
Public Inventory As New Collection
Public PolyArea() As POINTAPI
Public AllAreas As New Areas
Public SaveMode As String   ' Is user loading or saving game?

' Runtime variables
Public GName As String      ' Name of the Adventure game
Public GScreen As String    ' What screen is the player in?
Public GMusic As String     ' The BG Music that is playing
Public GRoom As String      ' The room the player currently is

' Pass-on variables (whatever)
Public uOK As Boolean
Public uChoiceMode As Boolean
Public uChoiceLbl As Integer
Public uSlot As Integer

' Game variables
Public SaveExt As String   ' Save game extension
Public DBug As Integer     ' Debug mode or not??

' Current Room
Type QRoom
  RName As String
  RSaved As Boolean
  RBG As String
  RLocName As String
  Desc As String
  
  dNorthWest As String
  dNorth As String
  dNorthEast As String
  dWest As String
  dEast As String
  dSouthWest As String
  dSouth As String
  dSouthEast As String
  
  Trans As Integer
  
  RMusic As String
  RConst(10) As String
  
  ItmCount As Integer
  Itm(50, 5) As String

  '0 = Name
  '1 = Desc
  '2 = Prog
  '3 = X
  '4 = Y
  '5 = Gfx
  
  AreaCount As Integer
  Area(50, 4) As String
  
  DontSave As Integer
  DontShowNav As Integer
  DisableMenu As Integer
  DisableTooltip As Integer
  
  '0 = Area Name
  '1 = Desc
  '2 = Prog
  '3 = X
  '4 = Y
  
End Type

Public DaRoom As QRoom

Sub Main()
Dim FName As String

If Command$ = "" Then
  If GetFileName(FName, "Quest Adventure Files|*.adv", "Open Adventure Game") = False Then End
Else
  FName = Trim$(Command$)
End If

WrkDir = App.Path + "\Game.Tmp\"

MakeTmpDir

ExtractGame FName
LoadGameSettings

frmMain.Show

End Sub

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
' Create the temporary working directory, if it isn't in there.
On Error Resume Next
MkDir WrkDir
End Sub

Sub DeleteTmp()
' Deletes Temporary Working Files for safety
Dim PName As String

  PName = Dir(WrkDir + "*.*")
  Do Until PName = ""
    Kill WrkDir + PName
    PName = Dir()
  Loop

End Sub

Sub ExtractGame(dFName As String)
' Open it, let's get down and dirty....:P
Dim PName As String

If JpkList(dFName, frmLstFiles.LstFiles) = True Then
  For a = 0 To frmLstFiles.LstFiles.ListCount - 1
    JpkExtract dFName, frmLstFiles.LstFiles.List(a), WrkDir + frmLstFiles.LstFiles.List(a)
  Next a
Else
  MsgBox "Failed to open " & dFName, vbCritical
  End
End If
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

Sub LoadGameSettings()
SaveExt = GetInitEntry("General", "SaveExt", "sav", WrkDir + "Config.cfg")
DBug = GetInitEntry("General", "Debug", vbUnchecked, WrkDir + "Config.cfg")
GName = GetInitEntry("General", "Title", , WrkDir + "Config.cfg")
End Sub

Sub Echo(wText As String, Optional wErr As Boolean = False)
' This is where we output all debug text messages (no, not cellphones)

' If the debug option is off and/or the message is not a critical error,
' just forget all about it.
If DBug = vbUnchecked And wErr = False Then Exit Sub

frmConsole.Show
StayOnTop frmConsole

frmConsole.txtConsole.Text = frmConsole.txtConsole.Text + wText + vbCrLf
frmConsole.txtConsole.SelStart = Len(frmConsole.txtConsole.Text)
End Sub

Public Sub StayOnTop(Frm As Form)
' Make the form stay on top (duh)
  TopProp = True
  SetWindowPos Frm.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub

Public Sub NotOnTop(Frm As Form)
' Make the form not on top
  TopProp = False
  SetWindowPos Frm.hwnd, -2, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
End Sub
