VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Area"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"

'local variable(s) to hold property value(s)
Private mvarAreaNumber As Long 'local copy
Private mvarAreaBrush As Long 'local copy
Private mvarAreaAlertBrush As Long 'local copy
Private mvarAreaPen As Long 'local copy
Private mvarAreaName As String 'local copy
Private mvarAreaNormalColor As Long 'local copy
Private mvarAreaMouseColor As Long 'local copy
Private mvarAreaSelected As Boolean 'local copy
Private mvarAreaState As Long 'local copy
Private mvarAreaComments As String 'local copy
Private mvarAreaStyle As Long 'local copy
Private mvarAreaType As String 'local copy


Public Property Let AreaType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaType = 5
    mvarAreaType = vData
End Property


Public Property Get AreaType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaType
    AreaType = mvarAreaType
End Property



Public Property Let AreaStyle(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaStyle = 5
    mvarAreaStyle = vData
End Property


Public Property Get AreaStyle() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaStyle
    AreaStyle = mvarAreaStyle
End Property



Public Property Let AreaComments(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaComments = 5
    mvarAreaComments = vData
End Property


Public Property Get AreaComments() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaComments
    AreaComments = mvarAreaComments
End Property



Public Property Let AreaState(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaState = 5
    mvarAreaState = vData
End Property


Public Property Get AreaState() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaState
    AreaState = mvarAreaState
End Property



Public Property Let AreaSelected(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaSelected = 5
    mvarAreaSelected = vData
End Property


Public Property Get AreaSelected() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaSelected
    AreaSelected = mvarAreaSelected
End Property



Public Property Let AreaMouseColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaMouseColor = 5
    mvarAreaMouseColor = vData
End Property


Public Property Get AreaMouseColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaMouseColor
    AreaMouseColor = mvarAreaMouseColor
End Property



Public Property Let AreaNormalColor(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaNormalColor = 5
    mvarAreaNormalColor = vData
End Property


Public Property Get AreaNormalColor() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaNormalColor
    AreaNormalColor = mvarAreaNormalColor
End Property



Public Property Let AreaName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaName = 5
    mvarAreaName = vData
End Property


Public Property Get AreaName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaName
    AreaName = mvarAreaName
End Property



Public Property Let AreaPen(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaPen = 5
    mvarAreaPen = vData
End Property


Public Property Get AreaPen() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaPen
    AreaPen = mvarAreaPen
End Property

Public Property Let AreaAlertBrush(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaBrush = 5
    mvarAreaAlertBrush = vData
End Property


Public Property Get AreaAlertBrush() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaBrush
    AreaAlertBrush = mvarAreaAlertBrush
End Property

Public Property Let AreaBrush(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaBrush = 5
    mvarAreaBrush = vData
End Property


Public Property Get AreaBrush() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaBrush
    AreaBrush = mvarAreaBrush
End Property



Public Property Let AreaNumber(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AreaNumber = 5
    mvarAreaNumber = vData
End Property


Public Property Get AreaNumber() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AreaNumber
    AreaNumber = mvarAreaNumber
End Property



