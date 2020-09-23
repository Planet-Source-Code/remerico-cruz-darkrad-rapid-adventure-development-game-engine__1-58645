Attribute VB_Name = "modRegion"
' Regional Graphics Engine
' This is a part of the graphics engine...
' When you point at an object on the screen, this
' is where it all activates!


Public Enum RGNAreaType
    Rectangle = 0
    Elipse = 1
    Polygon = 2
End Enum


Public Enum RGNHatchType
    RGN_HS_HORIZONTAL = 0 '-----
    RGN_HS_VERTICAL = 1 '|||||
    RGN_HS_FDIAGONAL = 2 '\\\\\
    RGN_HS_BDIAGONAL = 3 '/////
    RGN_HS_CROSS = 4 '+++++
    RGN_HS_DIAGCROSS = 5 'xxxxx
    RGN_HS_FDIAGONAL1 = 6
    RGN_HS_BDIAGONAL1 = 7
    RGN_HS_SOLID = 8
    RGN_HS_DENSE1 = 9
    RGN_HS_DENSE2 = 10
    RGN_HS_DENSE3 = 11
    RGN_HS_DENSE4 = 12
    RGN_HS_DENSE5 = 13
    RGN_HS_DENSE6 = 14
    RGN_HS_DENSE7 = 15
    RGN_HS_DENSE8 = 16
    RGN_HS_NOSHADE = 17
    RGN_HS_HALFTONE = 18
    RGN_HS_SOLIDCLR = 19
    RGN_HS_DITHEREDCLR = 20
    RGN_HS_SOLIDTEXTCLR = 21
    RGN_HS_DITHEREDTEXTCLR = 22
    RGN_HS_SOLIDBKCLR = 23
    RGN_HS_DITHEREDBKCLR = 24
    RGN_HS_API_MAX = 25
End Enum

Public Enum RGNBrushType
    RGN_BS_Solid = 0
    RGN_BS_NULL = 1
    RGN_BS_HOLLOW = 1
    RGN_BS_HATCHED = 2
    RGN_BS_PATTERN = 3
    RGN_BS_INDEXED = 4
    RGN_BS_DIBPATTERN = 5
    RGN_BS_DIBPATTERNPT = 6
    RGN_BS_PATTERN8X8 = 7
    RGN_BS_DIBPATTERN8X8 = 8
End Enum

''''''''''''''
Public Declare Function CreatePolyPolygonRgn Lib "gdi32" (lpPoint As POINTAPI, lpPolyCounts As Long, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Public Declare Function CreateBrushIndirect Lib "gdi32" (lpLogBrush As LOGBRUSH) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateRectRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function CreateEllipticRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreateEllipticRgnIndirect Lib "gdi32" (lpRect As RECT) As Long
Public Declare Function FillRgn Lib "gdi32" (ByVal Hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Public Declare Function FrameRgn Lib "gdi32" (ByVal Hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal Hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function EqualRgn Lib "gdi32" (ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long) As Long
Public Declare Function GetPolyFillMode Lib "gdi32" (ByVal Hdc As Long) As Long
Public Declare Function SetPolyFillMode Lib "gdi32" (ByVal Hdc As Long, ByVal nPolyFillMode As Long) As Long
Public Declare Function GetRgnBox Lib "gdi32" (ByVal hRgn As Long, lpRect As RECT) As Long
Public Declare Function OffsetRgn Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function PaintRgn Lib "gdi32" (ByVal Hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function InvertRgn Lib "gdi32" (ByVal Hdc As Long, ByVal hRgn As Long) As Long
Public Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function SetRectRgn Lib "gdi32" (ByVal hRgn As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function CreateHatchBrush Lib "gdi32" (ByVal nIndex As Long, ByVal crColor As Long) As Long

Public Const RGN_AND = 1
Public Const RGN_COPY = 5
Public Const RGN_DIFF = 4
Public Const RGN_MAX = RGN_COPY
Public Const RGN_MIN = RGN_AND
Public Const RGN_OR = 2
Public Const RGN_XOR = 3

' Brush Styles
Private Const BS_SOLID = 0
Private Const BS_NULL = 1
Private Const BS_HOLLOW = BS_NULL
Private Const BS_HATCHED = 2
Private Const BS_PATTERN = 3
Private Const BS_INDEXED = 4
Private Const BS_DIBPATTERN = 5
Private Const BS_DIBPATTERNPT = 6
Private Const BS_PATTERN8X8 = 7
Private Const BS_DIBPATTERN8X8 = 8

'  Hatch Styles
Private Const HS_HORIZONTAL = 0              '  -----
Private Const HS_VERTICAL = 1                '  |||||
Private Const HS_FDIAGONAL = 2               '  \\\\\
Private Const HS_BDIAGONAL = 3               '  /////
Private Const HS_CROSS = 4                   '  +++++
Private Const HS_DIAGCROSS = 5               '  xxxxx
Private Const HS_FDIAGONAL1 = 6
Private Const HS_BDIAGONAL1 = 7
Private Const HS_SOLID = 8
Private Const HS_DENSE1 = 9
Private Const HS_DENSE2 = 10
Private Const HS_DENSE3 = 11
Private Const HS_DENSE4 = 12
Private Const HS_DENSE5 = 13
Private Const HS_DENSE6 = 14
Private Const HS_DENSE7 = 15
Private Const HS_DENSE8 = 16
Private Const HS_NOSHADE = 17
Private Const HS_HALFTONE = 18
Private Const HS_SOLIDCLR = 19
Private Const HS_DITHEREDCLR = 20
Private Const HS_SOLIDTEXTCLR = 21
Private Const HS_DITHEREDTEXTCLR = 22
Private Const HS_SOLIDBKCLR = 23
Private Const HS_DITHEREDBKCLR = 24
Private Const HS_API_MAX = 25

'  Pen Styles
Private Const PS_SOLID = 0
Private Const PS_DASH = 1                    '  -------
Private Const PS_DOT = 2                     '  .......
Private Const PS_DASHDOT = 3                 '  _._._._
Private Const PS_DASHDOTDOT = 4              '  _.._.._
Private Const PS_NULL = 5
Private Const PS_INSIDEFRAME = 6
Private Const PS_USERSTYLE = 7
Private Const PS_ALTERNATE = 8
Private Const PS_STYLE_MASK = &HF

Private Const PS_ENDCAP_ROUND = &H0
Private Const PS_ENDCAP_SQUARE = &H100
Private Const PS_ENDCAP_FLAT = &H200
Private Const PS_ENDCAP_MASK = &HF00

Private Const PS_JOIN_ROUND = &H0
Private Const PS_JOIN_BEVEL = &H1000
Private Const PS_JOIN_MITER = &H2000
Private Const PS_JOIN_MASK = &HF000

Private Const PS_COSMETIC = &H0
Private Const PS_GEOMETRIC = &H10000
Private Const PS_TYPE_MASK = &HF0000

Private Const AD_COUNTERCLOCKWISE = 1
Private Const AD_CLOCKWISE = 2

' PolyFill() Modes
Private Const ALTERNATE = 1
Private Const WINDING = 2
Private Const POLYFILL_LAST = 2

'  Object Definitions for EnumObjects()
Private Const OBJ_PEN = 1
Private Const OBJ_BRUSH = 2
Private Const OBJ_DC = 3
Private Const OBJ_METADC = 4
Private Const OBJ_PAL = 5
Private Const OBJ_FONT = 6
Private Const OBJ_BITMAP = 7
Private Const OBJ_REGION = 8
Private Const OBJ_METAFILE = 9
Private Const OBJ_MEMDC = 10
Private Const OBJ_EXTPEN = 11
Private Const OBJ_ENHMETADC = 12
Private Const OBJ_ENHMETAFILE = 13

Public Type LOGBRUSH
        lbStyle As Long
        lbColor As Long
        lbHatch As Long
End Type
Public Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public Type POINTAPI
        x As Long
        y As Long
End Type




Public Function AddRGNRectangle(AllAreas As Areas, RC As RECT, sName As String, NormalColor As Long, AlertColor As Long, RGNhatch As RGNHatchType, RGNBrush As RGNBrushType) As Boolean
On Error GoTo EH
Dim ANumber As Long
Dim Apen As Long
Dim ABrush As Long 'normal brush
Dim AABrush As Long 'alert brush
Dim LB As LOGBRUSH

ANumber = CreateRectRgnIndirect(RC)
Apen = CreatePen(PS_SOLID, 2, NormalColor)
LB.lbColor = NormalColor
LB.lbStyle = CLng(RGNBrush)
LB.lbHatch = CLng(RGNhatch)
ABrush = CreateBrushIndirect(LB)

LB.lbColor = AlertColor
LB.lbStyle = CLng(RGNBrush)
LB.lbHatch = CLng(RGNhatch)
AABrush = CreateBrushIndirect(LB)

If ANumber = 0 Then
AddRGNRectangle = False
Exit Function
End If

AllAreas.Add ANumber, ABrush, AABrush, Apen, sName, NormalColor, AlertColor, False, 0, "", CLng(RGNhatch), "Rectangle"
AddRGNRectangle = True
Exit Function
EH:
AddRGNRectangle = False
MsgBox Err.Description, vbCritical, "Add Rectangle Region"
Exit Function
End Function

Public Function AddRGNElliptic(AllAreas As Areas, RC As RECT, sName As String, NormalColor As Long, AlertColor As Long, RGNhatch As RGNHatchType, RGNBrush As RGNBrushType) As Boolean
On Error GoTo EH
Dim ANumber As Long
Dim Apen As Long
Dim ABrush As Long 'normal brush
Dim AABrush As Long 'alert brush
Dim LB As LOGBRUSH

ANumber = CreateEllipticRgnIndirect(RC)
Apen = CreatePen(PS_SOLID, 2, NormalColor)
LB.lbColor = NormalColor
LB.lbStyle = CLng(RGNBrush)
LB.lbHatch = CLng(RGNhatch)
ABrush = CreateBrushIndirect(LB)

LB.lbColor = AlertColor
LB.lbStyle = CLng(RGNBrush)
LB.lbHatch = CLng(RGNhatch)
AABrush = CreateBrushIndirect(LB)

If ANumber = 0 Then
    AddRGNElliptic = False
    Exit Function
End If

AllAreas.Add ANumber, ABrush, AABrush, Apen, sName, NormalColor, AlertColor, False, 0, "", CLng(RGNhatch), "Elliptic"
AddRGNElliptic = True
Exit Function
EH:
AddRGNElliptic = False
MsgBox Err.Description, vbCritical, "Add Elliptic Region"
Exit Function
End Function

Public Function AddRGNPoly(AllAreas As Areas, P() As POINTAPI, nCount As Long, sName As String, sComment As String, NormalColor As Long, AlertColor As Long, RGNhatch As RGNHatchType, RGNBrush As RGNBrushType) As Boolean
On Error GoTo EH
Dim ANumber As Long
Dim Apen As Long
Dim ABrush As Long 'normal brush
Dim AABrush As Long 'alert brush
Dim LB As LOGBRUSH

ANumber = CreatePolygonRgn(P(1), nCount, 1)
Apen = CreatePen(PS_SOLID, 2, NormalColor)
LB.lbColor = NormalColor
LB.lbStyle = CLng(RGNBrush)
LB.lbHatch = CLng(RGNhatch)
ABrush = CreateBrushIndirect(LB)

LB.lbColor = AlertColor
LB.lbStyle = CLng(RGNBrush)
LB.lbHatch = CLng(RGNhatch)
AABrush = CreateBrushIndirect(LB)

If ANumber = 0 Then
    AddRGNPoly = False
    Exit Function
End If

AllAreas.Add ANumber, ABrush, AABrush, Apen, sName, NormalColor, AlertColor, False, 0, sComment, CLng(RGNhatch), "Polygon"
AddRGNPoly = True
Exit Function
EH:
AddRGNPoly = False
MsgBox Err.Description, vbCritical, "Add Polygon Region"
Exit Function
End Function

Public Function IsInRegion(AllAreas As Areas, x As Single, y As Single) As Long
'returns 0 if the mouse is not in a region boundries
'or returns the region number as assigned by windows
Dim IRGN As Long
Dim bFound As Boolean
IRGN = 0
For I = 1 To AllAreas.Count
    If PtInRegion(AllAreas(I).AreaNumber, x \ Screen.TwipsPerPixelX, y \ Screen.TwipsPerPixelY) Then
        IRGN = AllAreas(I).AreaNumber
        bFound = True
        Exit For
    End If
Next I
IsInRegion = IRGN
End Function

Public Sub PaintARGN(Hdc As Long, ANumber As Long, Apen As Long, ABrush As Long)
'Called to paint the region
        SelectObject Hdc, Apen
        SelectObject Hdc, ABrush
        PaintRgn Hdc, ANumber
End Sub

