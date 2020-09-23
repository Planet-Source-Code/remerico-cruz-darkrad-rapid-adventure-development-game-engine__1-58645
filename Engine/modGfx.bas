Attribute VB_Name = "modGfx"
' Graphics Engine
' (Almost) all things you see on the screen came here...
' Cool, huh?

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x _
      As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As _
      Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As _
      Long, ByVal dwRop As Long) As Long

Const AC_SRC_OVER = &H0
' This structure holds the arguments required by Alphablend function to work
Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
' This is the main API that is blending the pictures
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal Hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal Hdc As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal lInt As Long, ByVal BLENDFUNCT As Long) As Long
' This is a commenly used API function(maybe by me only) which is very helpful to Tranfer ALL the values of a 'Structure'(Type) to a Long variable
Private Declare Sub RtlMoveMemory Lib "kernel32.dll" (Destination As Any, Source As Any, ByVal Length As Long)
' Being used by the Timer
Dim Counter As Long
' The BlendFunction 'Structure' is used by the 'AlphaBlend' API function
Dim BF As BLENDFUNCTION
' Actually the AlphaBlend API Function requires a refrence to a "LONG" value containing the values of BlendFunction structure!. This Variale holds the values done in the BlendFunction Structure.
' A Structure (Type) can be converted into a 'Long' value by using the 'RtlMoveMemory' API Function.. See below for its example ;)
Dim lBF As Long



Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean

Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
    Const CCDEVICENAME = 32
    Const CCFORMNAME = 32
    Const DM_PELSWIDTH = &H80000
    Const DM_PELSHEIGHT = &H100000

Private Type DevMode
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
    End Type
    Dim DevM As DevMode
    
Sub ChangeRes(iWidth As Single, iHeight As Single)

    Dim a As Boolean
    Dim I&
    I = 0

    Do
        a = EnumDisplaySettings(0&, I&, DevM)
        I = I + 1
    Loop Until (a = False)

    Dim b&
    DevM.dmFields = DM_PELSWIDTH Or DM_PELSHEIGHT
    DevM.dmPelsWidth = iWidth
    DevM.dmPelsHeight = iHeight
    b = ChangeDisplaySettings(DevM, 0)
End Sub

Sub Blend(dVal As Long, wPic1 As PictureBox, wPic2 As PictureBox)
    
    'wPic2.Picture = wPic1.Picture

    'set the parameters
    With BF
        .BlendOp = AC_SRC_OVER
        .BlendFlags = 0
        .SourceConstantAlpha = dVal
        .AlphaFormat = 0
    End With
    
    'copy the BLENDFUNCTION-structure to a Long
    RtlMoveMemory lBF, BF, 4
    
    'AlphaBlend the picture from wPic1 over the picture of Picture2
    AlphaBlend wPic2.Hdc, 0, 0, wPic2.ScaleWidth, wPic2.ScaleHeight, wPic1.Hdc, 0, 0, wPic1.ScaleWidth, wPic1.ScaleHeight, lBF
    wPic2.Refresh
    

End Sub

Function GetImgSize(ByVal Filename As String, ByRef SizeW As Long, ByRef SizeH As Long, Optional Ext As String) As Boolean

'Inputs:
'
'fileName is a string containing the path name of the image file.
'
'ImgDim is passed as an empty type var and contains the height
'and width that's passed back.
'
'Ext is passed as an empty string and contains the image type
'as a 3 letter description that's passed back.
'
'
'Returns:
'
'True if the function was successful.


  'declare vars
  Dim handle As Integer, isValidImage As Boolean
  Dim byteArr(255) As Byte, I As Integer

  'init vars
  isValidImage = False
  SizeH = 0
  SizeW = 0
  
  'open file and get 256 byte chunk
  handle = FreeFile
  On Error GoTo endFunction
  Open Filename For Binary Access Read As #handle
  Get handle, , byteArr
  Close #handle
  

  'check for jpg header (SOI): &HFF and &HD8
  ' contained in first 2 bytes
  If byteArr(0) = &HFF And byteArr(1) = &HD8 Then
    isValidImage = True
  Else
    GoTo checkGIF
  End If
  
  'check for SOF marker: &HFF and &HC0 TO &HCF
  For I = 0 To 255
    If byteArr(I) = &HFF And byteArr(I + 1) >= &HC0 _
                         And byteArr(I + 1) <= &HCF Then
      SizeH = byteArr(I + 5) * 256 + byteArr(I + 6)
      SizeW = byteArr(I + 7) * 256 + byteArr(I + 8)
      Exit For
    End If
  Next I
  
  'get image type and exit
  Ext = "jpg"
  GoTo endFunction


checkGIF:
  
  'check for GIF header
  If byteArr(0) = &H47 And byteArr(1) = &H49 And byteArr(2) = &H46 _
  And byteArr(3) = &H38 Then
    SizeW = byteArr(7) * 256 + byteArr(6)
    SizeH = byteArr(9) * 256 + byteArr(8)
    isValidImage = True
  Else
    GoTo checkBMP
  End If
  
  'get image type and exit
  Ext = "gif"
  GoTo endFunction

  
checkBMP:
  
  'check for BMP header
  If byteArr(0) = 66 And byteArr(1) = 77 Then
    isValidImage = True
  Else
    GoTo checkPNG
  End If
  
  'get record type info
  If byteArr(14) = 40 Then
    
    'get width and height of BMP
    SizeW = byteArr(21) * 256 ^ 3 + byteArr(20) * 256 ^ 2 _
                 + byteArr(19) * 256 + byteArr(18)
    
    SizeH = byteArr(25) * 256 ^ 3 + byteArr(24) * 256 ^ 2 _
                  + byteArr(23) * 256 + byteArr(22)
  
  'another kind of BMP
  ElseIf byteArr(17) = 12 Then
  
    'get width and height of BMP
    SizeW = byteArr(19) * 256 + byteArr(18)
    SizeH = byteArr(21) * 256 + byteArr(20)
    
  End If
  
  'get image type and exit
  Ext = "bmp"
  GoTo endFunction

  
checkPNG:

  'check for PNG header
  If byteArr(0) = &H89 And byteArr(1) = &H50 And byteArr(2) = &H4E _
  And byteArr(3) = &H47 Then
    SizeW = byteArr(18) * 256 + byteArr(19)
    SizeH = byteArr(22) * 256 + byteArr(23)
    isValidImage = True
  Else
    GoTo endFunction
  End If
  
  Ext = "png"


endFunction:

  'return function's success status
  getImgDim = isValidImage


End Function

Sub Wipe(wPic1 As PictureBox, wPic2 As PictureBox)
Dim ImgX, ImgY As Integer
Dim NumLoop As Integer
Dim HalfHeight As Integer
Dim I As Integer
Dim suc&
Dim dwRop&
Dim blocco&
Dim blocco1&
Dim blocco2&
Dim tempor&
Dim Xtemp&

hDestDC& = wPic2.Hdc
TRWait = 5
hSrcDC& = wPic1.Hdc
ImgX = wPic1.ScaleWidth
ImgY = wPic1.ScaleHeight
HalfHeight = ImgY / 2
dwRop& = &HCC0020
blocco = 20
blocco1 = blocco
blocco2 = blocco
For I = 0 To HalfHeight Step blocco1
    y = I
    For x = I To ImgX - I Step blocco1
        If x + blocco1 > ImgX Then
            blocco1 = ImgX - x
        End If
        suc& = BitBlt(hDestDC&, x, y, blocco1, blocco2, hSrcDC&, _
      x, y, dwRop&)
      tempor = x
    Next
    Wait TRWait
    x = tempor
    For y = I + blocco1 To ImgY - I Step blocco2
        If y + blocco2 > ImgY Then
            blocco2 = ImgY - y
        End If
        suc& = BitBlt(hDestDC&, x, y, blocco1, blocco2, hSrcDC&, _
      x, y, dwRop&)
      tempor = y
    Next
    Wait TRWait
    y = tempor
    tempor = x
    For x = tempor - blocco To I Step -blocco
        suc& = BitBlt(hDestDC&, x, y, blocco, blocco2, hSrcDC&, _
      x, y, dwRop&)
        Xtemp = x
    Next
    Wait TRWait
    x = Xtemp
    tempor = y
    For y = tempor - blocco To I - blocco Step -blocco
        suc& = BitBlt(hDestDC&, x, y, blocco, blocco, hSrcDC&, _
      x, y, dwRop&)
    Next
    Wait TRWait
    DoEvents
    blocco1 = blocco
    blocco2 = blocco
Next
End Sub

Sub HourDouble(wPic1 As PictureBox, wPic2 As PictureBox)
Const PI = 3.1415
Dim ray, angle As Double
dwRop& = &HCC0020
hDestDC& = wPic2.Hdc
hSrcDC& = wPic1.Hdc
DoEvents


    For angle = 0 To 2 * PI Step 0.01
        a = Tan(angle)
        b = Cos(angle)
        c = Sin(angle)
        If Abs(a * (wPic1.ScaleWidth / 2)) < (wPic1.ScaleHeight / 2) Then
            For x = -0.5 * (1 + Sgn(b)) * (wPic1.ScaleWidth / 2) To 0.5 * (1 + Sgn(b)) * (wPic1.ScaleWidth / 2) Step Sgn(b)
            suc& = BitBlt(hDestDC&, (wPic1.ScaleWidth / 2) + x, (wPic1.ScaleHeight / 2) + a * x, 5, 5, hSrcDC&, _
      (wPic1.ScaleWidth / 2) + x, (wPic1.ScaleHeight / 2) + a * x, dwRop&)
        Next
        Else
            For y = -0.5 * (1 + Sgn(c)) * (wPic1.ScaleWidth / 2) To 0.5 * (1 + Sgn(c)) * (wPic1.ScaleWidth / 2) Step Sgn(c)
            suc& = BitBlt(hDestDC&, (wPic1.ScaleWidth / 2) + y / a, (wPic1.ScaleHeight / 2) + y, 5, 5, hSrcDC&, _
      (wPic1.ScaleWidth / 2) + y / a, (wPic1.ScaleHeight / 2) + y, dwRop&)
            Next
        End If
    Next
End Sub

Sub HourInverse(wPic1 As PictureBox, wPic2 As PictureBox)
Const PI = 3.1415
Dim ray, angle As Double
dwRop& = &HCC0020
hDestDC& = wPic2.Hdc
hSrcDC& = wPic1.Hdc
DoEvents


    For angle = 2 * PI To 0 Step -0.01
        a = Tan(angle)
        b = Cos(angle)
        c = Sin(angle)
        If Abs(a * (wPic1.ScaleWidth / 2)) < (wPic1.ScaleHeight / 2) Then
            For x = 0.5 * (Sgn(b) - 1) * (wPic1.ScaleWidth / 2) To 0.5 * (1 + Sgn(b)) * (wPic1.ScaleWidth / 2)
            suc& = BitBlt(hDestDC&, (wPic1.ScaleWidth / 2) + x, (wPic1.ScaleHeight / 2) + a * x, 5, 5, hSrcDC&, _
      (wPic1.ScaleWidth / 2) + x, (wPic1.ScaleHeight / 2) + a * x, dwRop&)
        Next
        Else
            For y = 0.5 * (Sgn(c) - 1) * (wPic1.ScaleWidth / 2) To 0.5 * (1 + Sgn(c)) * (wPic1.ScaleWidth / 2)
            suc& = BitBlt(hDestDC&, (wPic1.ScaleWidth / 2) + y / a, (wPic1.ScaleHeight / 2) + y, 5, 5, hSrcDC&, _
      (wPic1.ScaleWidth / 2) + y / a, (wPic1.ScaleHeight / 2) + y, dwRop&)
            Next
        End If
    Next
End Sub

Sub CircleTrans(wPic1 As PictureBox, wPic2 As PictureBox)
    Const PI = 3.1415
    Dim ray, angle As Double

    dwRop& = &HCC0020
    hDestDC& = wPic2.Hdc
    hSrcDC& = wPic1.Hdc
    DoEvents
    ray = Sqr(wPic1.ScaleHeight ^ 2 + wPic1.ScaleWidth ^ 2) / 2
    For I = ray To 0 Step -1.9
        For angle = 0 To 2 * PI Step 0.01
            x = I * Cos(angle) + (wPic1.ScaleWidth / 2)
            y = I * Sin(angle) + (wPic1.ScaleHeight / 2)
            suc& = BitBlt(hDestDC&, x, y, 3, 3, hSrcDC&, _
          x, y, dwRop&)
        Next
    Next
End Sub

Sub Implode(wPic1 As PictureBox, wPic2 As PictureBox)
Const PI = 3.1415
Dim ray, angle As Double
dwRop& = &HCC0020
hDestDC& = wPic2.Hdc
hSrcDC& = wPic1.Hdc
DoEvents
ray = Sqr(wPic1.ScaleHeight ^ 2 + wPic1.ScaleWidth ^ 2) / 2
For I = ray To 0 Step -1.1
    For angle = 0 To 5 * PI Step 0.01
        x = I * Tan(angle) + (wPic1.ScaleWidth / 2)
        y = I * Cos(angle) + (wPic1.ScaleHeight / 2)
        suc& = BitBlt(hDestDC&, x, y, 16, 16, hSrcDC&, _
      x, y, dwRop&)
    Next
Next
End Sub

Sub Tenda(wPic1 As PictureBox, wPic2 As PictureBox)
Dim ImgX, ImgY As Integer
Dim NumLoop As Integer
Dim HalfHeight As Integer
Dim I As Integer
Dim suc&
Dim dwRop&

hDestDC& = wPic2.Hdc
hSrcDC& = wPic1.Hdc

dwRop& = &HCC0020
For I = 0 To wPic1.ScaleHeight / 2
For a = 1 To CLng(100000): Next
        suc& = BitBlt(hDestDC&, 0, I, wPic1.ScaleWidth, 1, hSrcDC&, 0, I, dwRop&)
        suc& = BitBlt(hDestDC&, 0, wPic1.ScaleHeight - I, wPic1.ScaleWidth, 1, _
        hSrcDC&, 0, wPic1.ScaleHeight - I, dwRop&)
Next
End Sub
