Attribute VB_Name = "modPack"
Dim ItemLength As Long
Dim ItemString As String
Dim ItemNumber(0 To 1) As Integer

Dim BytesExtract As String
Dim BytesAdd As String

Dim ItemBinary As String
Dim Position As Long
Dim LastPosition As Long

Dim FileListStart As Long
Dim FilePosition As Long
Dim ExitDo As Boolean

Dim PutLength As String
Dim PutPosition As Long

Function JpkAdd(JpkFile As String, FileName As String, AddName As String) As Boolean

    On Error GoTo FinaliseError

    AddName = AddName & Chr(0)
    
    ItemNumber(0) = FreeFile
    Open JpkFile For Binary As #ItemNumber(0)
        ItemNumber(1) = FreeFile
        Open FileName For Binary As #ItemNumber(1)
            PutLength = LOF(ItemNumber(1)) & Chr(0)
            Put ItemNumber(0), LOF(ItemNumber(0)) + 1, AddName
            Put ItemNumber(0), LOF(ItemNumber(0)) + 1, PutLength
            PutPosition = LOF(ItemNumber(0))
            If LOF(ItemNumber(1)) > 1000000 Then
                Position = -999999
                Do
                    Position = Position + 1000000
                    If Position + 999999 > LOF(ItemNumber(1)) Then BytesAdd = String(LOF(ItemNumber(1)) - Position + 1, Chr$(0)) Else BytesAdd = String(1000000, Chr$(0))
                    Get ItemNumber(1), Position, BytesAdd
                    Put ItemNumber(0), PutPosition + 1, BytesAdd
                    PutPosition = LOF(ItemNumber(0))
                Loop Until Position + 999999 > LOF(ItemNumber(1))
            Else
                BytesAdd = String(LOF(ItemNumber(1)), Chr$(0))
                Get ItemNumber(1), , BytesAdd
                Put ItemNumber(0), PutPosition + 1, BytesAdd
            End If
        Close ItemNumber(1)
    Close #ItemNumber(0)
    JpkAdd = True
    Exit Function
    
FinaliseError:
    JpkAdd = False

End Function

Function JpkList(JpkFile As String, ListItem As ListBox) As Boolean
ListItem.Clear
    On Error GoTo FinaliseError

    ItemNumber(0) = FreeFile
    Open JpkFile For Binary As #ItemNumber(0)
        Position = 1
        Do
            ItemString = Space(256)
            Get #ItemNumber(0), Position, ItemString
            ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
            Position = Position + Len(ItemString) + 1
            ListItem.AddItem ItemString
            
            ItemString = Space(256)
            Get #ItemNumber(0), Position, ItemString
            ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
            ItemLength = CLng(ItemString)
            Position = Position + Len(ItemString) + ItemLength + 1
        Loop Until Position > LOF(ItemNumber(0))
    Close #ItemNumber(0)
    JpkList = True
    Exit Function
    
FinaliseError:
    JpkList = False

End Function

Function JpkExtract(JpkFile As String, FileName As String, Destination As String) As Boolean

    On Error GoTo FinaliseError

    ItemNumber(0) = FreeFile
    Open JpkFile For Binary As ItemNumber(0)
        ItemNumber(1) = FreeFile
        Open Destination For Binary As ItemNumber(1)
            Position = 1
            ExitDo = False
            Do
                ItemString = Space(256)
                Get #ItemNumber(0), Position, ItemString
                ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
                Position = Position + Len(ItemString) + 1
                If LCase(ItemString) = LCase(FileName) Then ExitDo = True
                
                ItemString = Space(256)
                Get #ItemNumber(0), Position, ItemString
                ItemString = Mid(ItemString, 1, InStr(1, ItemString, Chr(0)) - 1)
                ItemLength = CLng(ItemString)
                Position = Position + Len(ItemString) + ItemLength + 1
                If ExitDo = True Then Exit Do
            Loop Until Position > LOF(ItemNumber(0))
            
            FileListStart = Position - ItemLength
            If ItemLength > 1000000 Then
                FilePosition = -999999
                Do
                    FilePosition = FilePosition + 1000000
                    If FilePosition + 999999 > ItemLength Then BytesExtract = Space(ItemLength - FilePosition + 1) Else BytesExtract = Space(1000000)
                    Get ItemNumber(0), FileListStart, BytesExtract
                    Put ItemNumber(1), FilePosition, BytesExtract
                    FileListStart = FileListStart + Len(BytesExtract)
                Loop Until FilePosition + 999999 > LOF(ItemNumber(1))
            Else
                BytesExtract = Space(ItemLength)
                Get ItemNumber(0), Position - ItemLength, BytesExtract
                Put ItemNumber(1), 1, BytesExtract
            End If
        Close ItemNumber(1)
    Close ItemNumber(0)
    JpkExtract = True
    Exit Function
    
FinaliseError:
    JpkExtract = False

End Function
