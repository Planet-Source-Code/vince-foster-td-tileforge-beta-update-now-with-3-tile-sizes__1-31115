VERSION 5.00
Begin VB.Form frmMap 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Map"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8010
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   442
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   534
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   495
      Left            =   4020
      TabIndex        =   4
      Top             =   6060
      Width           =   1935
   End
   Begin VB.CommandButton cmdRandomMap 
      Caption         =   "Random Map"
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   6060
      Width           =   1935
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Map"
      Height          =   495
      Left            =   60
      TabIndex        =   2
      Top             =   6060
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   6000
      TabIndex        =   1
      Top             =   6060
      Width           =   1935
   End
   Begin VB.PictureBox picMap 
      AutoRedraw      =   -1  'True
      Height          =   5835
      Left            =   120
      ScaleHeight     =   385
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   513
      TabIndex        =   0
      Top             =   120
      Width           =   7755
      Begin VB.Shape Shape1 
         BorderColor     =   &H0000FF00&
         Height          =   975
         Left            =   0
         Top             =   0
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type TileType
    CellID As Integer
    Flags As Integer
    VarrID As Integer
    Varrnum As Integer
    Animated As Integer
    EnemyID As Integer
    EnemyFlags As Integer
    ItemId As Integer
    ItemFlags As Integer
End Type
Dim MapX As Long
Dim MapY As Long
Dim Map(24, 24) As TileType
Public Sub BltMap()

  Dim X As Integer
  Dim Y As Integer
  Dim TileNum As Integer
  Dim PointString As String
  Dim TType As Integer

    picMap.Cls
    For Y = 0 To 24
        For X = 0 To 24
            If X + 1 > 24 And Y + 1 > 24 Then
                PointString = CStr(Map(X, Y).CellID & 0 & 0 & 0)
              ElseIf X + 1 > 24 Then
                PointString = CStr(Map(X, Y).CellID & 0 & 0 & Map(X, Y + 1).CellID)
              ElseIf Y + 1 > 24 Then
                PointString = CStr(Map(X, Y).CellID & Map(X + 1, Y).CellID & 0 & 0)
              Else
                PointString = CStr(Map(X, Y).CellID & Map(X + 1, Y).CellID & Map(X + 1, Y + 1).CellID & Map(X, Y + 1).CellID)
            End If

            Select Case PointString

              Case "0000"
                TileNum = 0
              Case "0001"
                TileNum = 1
              Case "0010"
                TileNum = 2
              Case "0011"
                TileNum = 3
              Case "0100"
                TileNum = 4
              Case "0101"
                TileNum = 5
              Case "0110"
                TileNum = 6
              Case "0111"
                TileNum = 7
              Case "1000"
                TileNum = 8
              Case "1001"
                TileNum = 9
              Case "1010"
                TileNum = 10
              Case "1011"
                TileNum = 11
              Case "1100"
                TileNum = 12
              Case "1101"
                TileNum = 13
              Case "1110"
                TileNum = 14
              Case "1111"
                TileNum = 15
            End Select

            BitBlt picMap.hdc, (X * Tilesize), (Y * Tilesize), Tilesize, Tilesize, frmMain.picTile(TileNum).hdc, 0, 0, SRCCOPY
        Next X
    Next Y
                
End Sub

Private Sub cmdClear_Click()

    NewMap
    BltMap

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdGenerate_Click()

    frmMain.MakeNewTiles
    BltMap

End Sub

Private Sub cmdRandomMap_Click()

    On Error Resume Next
    Dim MX As Integer
    Dim MY As Integer
    Dim AX As Integer
    Dim AY As Integer
    Dim R As Integer
      For MX = 0 To 24
          For MY = 0 To 24

              For AY = (MY - 1) To (MY + 1)
                  For AX = (MX - 1) To (MX + 1)
                      Randomize Timer
                      R = CInt(Rnd * 10)
                      If R >= 6 Then
                          Map(AX, AY).CellID = 1
                          Map(AX, AY).VarrID = CInt(Rnd * 2)
                        Else
                          Map(AX, AY).CellID = 0
                          Map(AX, AY).VarrID = CInt(Rnd * 2)
                      End If
                  Next AX
              Next AY
          Next MY
      Next MX
      BltMap

End Sub

Public Sub DrawOnMap(Button As Integer, Shift As Integer, X As Single, Y As Single)

    On Error Resume Next
    Dim R As Integer
    Dim AX As Integer
    Dim AY As Integer
      MapX = (X \ Tilesize) * Tilesize
      MapY = (Y \ Tilesize) * Tilesize
      If MapX > 512 Then MapX = 512
      If MapY > 352 Then MapY = 352
      If MapX < 0 Then MapX = 0
      If MapY < 0 Then MapY = 0
      Shape1.Move MapX, MapY, Tilesize + 1, Tilesize + 1
      If Button = 1 Then

          For AY = ((MapY \ Tilesize) - 1) To ((MapY \ Tilesize) + 1)
              For AX = ((MapX \ Tilesize) - 1) To ((MapX \ Tilesize) + 1)
                  R = CInt((Rnd * 2))
                  Map(AX, AY).VarrID = R
              Next AX
          Next AY

          For AY = ((MapY \ Tilesize)) To ((MapY \ Tilesize) + 1)
              For AX = ((MapX \ Tilesize)) To ((MapX \ Tilesize) + 1)
                  If Map(AX, AY).CellID = 0 Then Map(AX, AY).CellID = 1
                  Map(AX, AY).Flags = 0
              Next AX
          Next AY

          Map((MapX \ Tilesize), (MapY \ Tilesize)).CellID = 1
          Map(((MapX \ Tilesize) + 1), (MapY \ Tilesize)).CellID = 1
          Map((MapX \ Tilesize), ((MapY \ Tilesize) + 1)).CellID = 1
          Map(((MapX \ Tilesize) + 1), ((MapY \ Tilesize) + 1)).CellID = 1
            
          BltMap
      End If

End Sub

Private Sub Form_Load()

    cmdRandomMap_Click
    BltMap

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmMap = Nothing

End Sub

Public Sub LoadMap()

    On Error Resume Next
    Dim sFile As String
    Dim X As Integer
    Dim Y As Integer
    Dim Z As Integer
    Dim Tempmap As String
    Dim FFile As Integer
      FFile = FreeFile
      sFile = OpenDialog(Me.hwnd, "(*.map)|*.map", 1, 255, App.Path, "", "Open Map")

      Open sFile For Input As #FFile
      For Y = 0 To 24
          Input #FFile, Tempmap
          For X = 0 To 24
              Map(X, Y).CellID = Mid$(Tempmap, X * 3 + 1, 1)
              Map(X, Y).Flags = Mid$(Tempmap, X * 3 + 2, 1)
              Map(X, Y).VarrID = Mid$(Tempmap, X * 3 + 3, 1)
          Next X
      Next Y
      Close #FFile
      BltMap
  
Exit Sub

ErrHandler:
      Select Case Err.Number
        Case 75
        Case Else
          MsgBox "Error Number " & Err.Number & vbCrLf & " " & Err.Description, vbExclamation, "Error!"
      End Select

End Sub

Public Sub NewMap()

  Dim X As Long
  Dim Y As Long

    picMap.Cls
    For X = 0 To 24
        For Y = 0 To 24
            Randomize Timer
            Map(X, Y).CellID = 0
            Map(X, Y).Flags = 0
            Map(X, Y).VarrID = CInt(Rnd * 2)
        Next Y
    Next X

End Sub

Private Sub picMap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    DrawOnMap Button, Shift, X, Y

End Sub

Private Sub picMap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    DrawOnMap Button, Shift, X, Y

End Sub

Public Sub SaveMap()

    On Error GoTo ErrHandler
  Dim X As Integer
  Dim Y As Integer
  Dim FFile As Integer
  Dim Tempmap As String
  Dim sFile As String
    FFile = FreeFile
    sFile = SaveDialog(Me.hwnd, "(*.map)|*.map", 1, 255, App.Path, "", "Save Map")
    Open sFile For Output As #FFile
    For Y = 0 To 24
        For X = 0 To 24
            Tempmap = Tempmap & Map(X, Y).CellID & Map(X, Y).Flags & Map(X, Y).VarrID
        Next X
        Print #FFile, Tempmap
        Tempmap = ""
    Next Y
    Close #FFile
    MsgBox "Map Saved.", vbInformation

Exit Sub

ErrHandler:
    Select Case Err.Number
      Case 0
      Case 75
      Case Else
        MsgBox "Error Number " & Err.Number & vbCrLf & " " & Err.Description, vbExclamation, "Error!"
    End Select
End Sub

':) Ulli's VB Code Formatter V2.5.12 (1/24/2002 12:40:46 AM) 15 + 262 = 277 Lines
