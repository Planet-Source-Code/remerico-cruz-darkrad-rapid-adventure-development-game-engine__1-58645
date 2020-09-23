VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Areas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"Area"
Attribute VB_Ext_KEY = "Member0" ,"Area"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'local variable to hold collection
Private mCol As Collection
Public ParentHDC As Long





Public Function Add(AreaNumber As Long, AreaBrush As Long, AreaAlertBrush As Long, AreaPen As Long, AreaName As String, AreaNormalColor As Long, AreaMouseColor As Long, AreaSelected As Boolean, AreaState As Long, AreaComments As String, AreaStyle As Long, AreaType As String, Optional sKey As String) As Area
    'create a new object
    Dim objNewMember As Area
    Set objNewMember = New Area


    'set the properties passed into the method
    objNewMember.AreaNumber = AreaNumber
    objNewMember.AreaBrush = AreaBrush
    objNewMember.AreaAlertBrush = AreaAlertBrush
    objNewMember.AreaPen = AreaPen
    objNewMember.AreaName = AreaName
    objNewMember.AreaNormalColor = AreaNormalColor
    objNewMember.AreaMouseColor = AreaMouseColor
    objNewMember.AreaSelected = AreaSelected
    objNewMember.AreaState = AreaState
    objNewMember.AreaComments = AreaComments
    objNewMember.AreaStyle = AreaStyle
    objNewMember.AreaType = AreaType

        mCol.Add objNewMember
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As Area
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

Public Sub SelectNone()
Dim I As Integer
If mCol.Count = 0 Then Exit Sub
For I = 1 To mCol.Count
    mCol.Item(I).AreaSelected = False
Next I
End Sub


Public Sub SetSelected(Inum As Long)
Dim I As Integer
If mCol.Count = 0 Then Exit Sub
For I = 1 To mCol.Count
    If mCol.Item(I).AreaNumber = Inum Then
         mCol.Item(I).AreaSelected = False
    Else
        mCol.Item(I).AreaSelected = False
    End If
Next I
End Sub

Public Sub ClearAll()
Dim I As Integer
If mCol.Count = 0 Then Exit Sub
For I = mCol.Count To 1 Step -1
    DestroyRGN mCol.Item(I).AreaNumber, mCol.Item(I).AreaPen, mCol.Item(I).AreaBrush
    mCol.Remove I
Next I
End Sub

Public Sub RemoveRegion(LRGN As Long)
Dim I As Integer
If mCol.Count = 0 Then Exit Sub

For I = 1 To mCol.Count
If mCol.Item(I).AreaNumber = LRGN Then
     DestroyRGN mCol.Item(I).AreaNumber, mCol.Item(I).AreaPen, mCol.Item(I).AreaBrush
    mCol.Remove I
    Exit For
End If
Next I
End Sub

Public Sub DestroyRGN(LRGN As Long, LPen As Long, LBrush As Long)
'Destroy all items we asked windows to do for us
'or bad things happen in memory
        DeleteObject LRGN
        DeleteObject LPen
        DeleteObject LBrush
End Sub
