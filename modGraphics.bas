Attribute VB_Name = "modGraphics"
Option Explicit
Public Type POINTAPI
    X As Long
    Y As Long

End Type
Public Enum TileSizeEnum
    [32 x 32] = 32
    [48 x 48] = 48
    [64 x 64] = 64
End Enum
Public Enum DirectionEnum
    HorizontalN
    HorizontalS
    VerticalE
    Verticalw
End Enum
Public Declare Function GetTickCount Lib "kernel32" () As Long
Public Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Public Declare Function StrokePath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function StrokeAndFillPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function BeginPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function EndPath Lib "gdi32" (ByVal hdc As Long) As Long
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCERASE = &H440328        ' (DWORD) dest = source AND (NOT dest )
Public Const SRCINVERT = &H660046       ' (DWORD) dest = source XOR dest
Public Const SRCPAINT = &HEE0086        ' (DWORD) dest = source OR dest
Public Const PI = 3.141592654
Public Tilesize As TileSizeEnum
Private Function AirBrush(lngHdc As Long, X, Y, Size As Integer, Count As Integer, color As Long)

  Dim i As Integer
  Dim xadd As Integer
  Dim yadd As Integer

    For i = 0 To Count
        Select Case Int(Rnd * 2)
          Case 0
            xadd = X + Int(Rnd * Size)
          Case 1
            xadd = X - Int(Rnd * Size)
        End Select
        Select Case Int(Rnd * 2)
          Case 0
            yadd = Y + Int(Rnd * Size)
          Case 1
            yadd = Y - Int(Rnd * Size)
        End Select
        SetPixel lngHdc, xadd, yadd, color
    Next i

End Function

Public Sub DrawFlatGraph(ByVal lngHdc As Long, ByVal CenterX As Long, ByVal CenterY As Long, Orientation As DirectionEnum, ByVal LineMaxLength As Integer, ByVal Spokes As Integer, UseSpray As Boolean, Optional SprayLength As Integer = 10, Optional SprayDensity As Integer = 100, Optional SprayColor As Long = vbBlue)

  Dim X As Integer
  Dim Y As Integer
  Dim Q As Integer
  Dim Degrees As Single
  Dim LineWidth As Integer
  Dim PTS() As POINTAPI

    ReDim PTS(Spokes + 3)
    LineWidth = LineMaxLength

    Select Case Orientation
      Case 0 'HZS

        If UseSpray Then
            LineMaxLength = LineMaxLength  '+ (SprayLength)
        End If
        BeginPath lngHdc
        For Q = 0 To (Spokes)
            Randomize Timer
            Degrees = Q * ((LineWidth * 2) \ (Spokes))
            Select Case Q
              Case 0
                PTS(Q).X = CenterX
                PTS(Q).Y = LineMaxLength + CenterY
              Case (Spokes)
                PTS(Q).X = (LineWidth * 2) + CenterX
                PTS(Q).Y = LineMaxLength + CenterY
              Case Else
                PTS(Q).X = Degrees + CenterX
                PTS(Q).Y = CInt(Rnd * (LineMaxLength / 4)) + ((LineMaxLength / 4) * 3) + CenterY
            End Select

        Next Q
        PTS(Q).X = (LineWidth * 2) + CenterX
        PTS(Q).Y = (LineMaxLength * 2) + CenterY
        PTS(Q + 1).X = CenterX
        PTS(Q + 1).Y = (LineMaxLength * 2) + CenterY
        Polygon lngHdc, PTS(0), (Spokes + 3)
        EndPath lngHdc

      Case 1
        If UseSpray Then
            LineMaxLength = LineMaxLength  '- (SprayLength)
        End If
        BeginPath lngHdc
        For Q = 0 To (Spokes)
            Randomize Timer
            Degrees = Q * ((LineWidth * 2) \ (Spokes))
            Select Case Q
              Case 0
                PTS(Q).X = CenterX
                PTS(Q).Y = LineMaxLength + CenterY
              Case (Spokes)
                PTS(Q).X = (LineWidth * 2) + CenterX
                PTS(Q).Y = LineMaxLength + CenterY
              Case Else
                PTS(Q).X = Degrees + CenterX
                PTS(Q).Y = CInt(Rnd * (LineMaxLength / 4)) + ((LineMaxLength / 4) * 3) + CenterY
            End Select

        Next Q
        PTS(Q).X = (LineWidth * 2) + CenterX
        PTS(Q).Y = CenterY
        PTS(Q + 1).X = CenterX
        PTS(Q + 1).Y = CenterY
        Polygon lngHdc, PTS(0), (Spokes + 3)
        EndPath lngHdc
      Case 2
        If UseSpray Then
            LineMaxLength = LineMaxLength  ' + (SprayLength)
        End If
        BeginPath lngHdc
        For Q = 0 To (Spokes)
            Randomize Timer
            Degrees = Q * ((LineWidth * 2) \ (Spokes))
            Select Case Q
              Case 0
                PTS(Q).Y = CenterX
                PTS(Q).X = LineMaxLength + CenterY
              Case (Spokes)
                PTS(Q).Y = (LineWidth * 2) + CenterX
                PTS(Q).X = LineMaxLength + CenterY
              Case Else
                PTS(Q).Y = Degrees + CenterX
                PTS(Q).X = CInt(Rnd * (LineMaxLength / 4)) + ((LineMaxLength / 4) * 3) + CenterY
            End Select

        Next Q
        PTS(Q).Y = (LineWidth * 2) + CenterX
        PTS(Q).X = (LineMaxLength * 2) + CenterY
        PTS(Q + 1).Y = CenterX
        PTS(Q + 1).X = (LineMaxLength * 2) + CenterY
        Polygon lngHdc, PTS(0), (Spokes + 3)
        EndPath lngHdc
      Case 3
        If UseSpray Then
            LineMaxLength = LineMaxLength  ' - (SprayLength)
        End If
        BeginPath lngHdc
        For Q = 0 To (Spokes)
            Randomize Timer
            Degrees = Q * ((LineWidth * 2) \ (Spokes))
            Select Case Q
              Case 0
                PTS(Q).Y = CenterY
                PTS(Q).X = LineMaxLength + CenterX
              Case (Spokes)
                PTS(Q).Y = (LineWidth * 2) + CenterY
                PTS(Q).X = LineMaxLength + CenterX
              Case Else
                PTS(Q).Y = Degrees + CenterY
                PTS(Q).X = CInt(Rnd * (LineMaxLength / 4)) + ((LineMaxLength / 4) * 3) + CenterX
            End Select
            
        Next Q
        PTS(Q).Y = (LineWidth * 2) + CenterY
        PTS(Q).X = CenterX
        PTS(Q + 1).Y = CenterY
        PTS(Q + 1).X = CenterX
        Polygon lngHdc, PTS(0), (Spokes + 3)
        EndPath lngHdc
    End Select
    StrokeAndFillPath lngHdc
    If UseSpray Then
        For Q = 0 To (Spokes)
            AirBrush lngHdc, PTS(Q).X, PTS(Q).Y, SprayLength, SprayDensity, SprayColor
        Next Q
    End If

End Sub

Public Sub DrawRadarGraph(ByVal lngHdc As Long, ByVal CenterX As Long, ByVal CenterY As Long, ByVal LineMaxLength As Integer, ByVal Spokes As Integer, UseSpray As Boolean, Optional SprayLength As Integer = 10, Optional SprayDensity As Integer = 100, Optional SprayColor As Long = vbBlue)

  Dim X As Integer
  Dim Y As Integer
  Dim Q As Integer
  Dim Z As Integer
  Dim Radians As Single
  Dim Degrees As Single
  Dim LineLen As Integer
  Dim PTS() As POINTAPI

    ReDim PTS(Spokes)
    Z = 360 \ Spokes
    If UseSpray Then
        LineMaxLength = LineMaxLength  '- (SprayLength)
    End If
    BeginPath lngHdc
    For Q = 0 To (Spokes)
        Randomize Timer
        Degrees = Q * Z
        Select Case Degrees
          Case 0, 90, 180, 270, 360
            LineLen = LineMaxLength
          Case Else
            LineLen = CInt(Rnd * (LineMaxLength / 4)) + ((LineMaxLength / 4) * 3)
        End Select
        Radians = Degrees * (PI / 180)
        PTS(Q).X = CenterX + (Sin(Radians) * LineLen)
        PTS(Q).Y = CenterY - (Cos(Radians) * LineLen)
    Next Q
    Polygon lngHdc, PTS(0), Spokes
    EndPath lngHdc
    StrokeAndFillPath lngHdc
    If UseSpray Then
        For Q = 0 To (Spokes)
            AirBrush lngHdc, PTS(Q).X, PTS(Q).Y, SprayLength, SprayDensity, SprayColor
        Next Q
    End If

End Sub

Function TransparentBlt(hDestDC As Long, nDestX, nDestY, nWidth, nHeight, hSourceDC As Long, nSourceX, nSourceY, lTransColor As Long)

  Dim lOldColor As Long
  Dim hMaskDC As Long
  Dim hMaskBmp As Long
  Dim hOldMaskBmp As Long
  Dim hTempBmp As Long
  Dim hTempDC As Long
  Dim hOldTempBmp As Long
  Dim hDummy As Long

    lOldColor = SetBkColor&(hSourceDC, lTransColor)
    lOldColor = SetBkColor&(hDestDC, lTransColor)
    hMaskDC = CreateCompatibleDC(hDestDC)
    hMaskBmp = CreateCompatibleBitmap(hDestDC, nWidth, nHeight)
    hOldMaskBmp = SelectObject(hMaskDC, hMaskBmp)
    hTempBmp = CreateBitmap(nWidth, nHeight, 1, 1, 0&)
    hTempDC = CreateCompatibleDC(hDestDC)
    hOldTempBmp = SelectObject(hTempDC, hTempBmp)
    If BitBlt(hTempDC, 0, 0, nWidth, nHeight, hSourceDC, nSourceX, nSourceY, SRCCOPY) Then
        hDummy = BitBlt(hMaskDC, 0, 0, nWidth, nHeight, hTempDC, 0, 0, SRCCOPY)
    End If
    hTempBmp = SelectObject(hTempDC, hOldTempBmp)
    hDummy = DeleteObject(hTempBmp)
    hDummy = DeleteDC(hTempDC)
    If BitBlt(hDestDC, nDestX, nDestY, nWidth, nHeight, hSourceDC, nSourceX, nSourceY, SRCINVERT) Then
        If BitBlt(hDestDC, nDestX, nDestY, nWidth, nHeight, hMaskDC, 0, 0, SRCAND) Then
            If BitBlt(hDestDC, nDestX, nDestY, nWidth, nHeight, hSourceDC, nSourceX, nSourceY, SRCINVERT) Then
                TransparentBlt = True
            End If
        End If
    End If
    hMaskBmp = SelectObject(hMaskDC, hOldMaskBmp)
    hDummy = DeleteObject(hMaskBmp)
    hDummy = DeleteDC(hMaskDC)
End Function

':) Ulli's VB Code Formatter V2.5.12 (1/24/2002 12:41:18 AM) 40 + 237 = 277 Lines
