VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TD-TileForge"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   8310
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   436
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   554
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picTex 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   2
      Left            =   4200
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   64
      Top             =   720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picTex 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   1
      Left            =   2160
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   63
      Top             =   720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picStrip 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   15360
      Left            =   11460
      ScaleHeight     =   1024
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   62
      Top             =   360
      Width           =   960
   End
   Begin VB.PictureBox picAngleB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   4
      Left            =   13320
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   61
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picAngleB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   3
      Left            =   12300
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   60
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picAngleB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   2
      Left            =   11280
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   59
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picAngleB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   1
      Left            =   10260
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   58
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picAngleB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   0
      Left            =   9240
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   57
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picMaskInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      ForeColor       =   &H00FFFFFF&
      Height          =   2880
      Index           =   0
      Left            =   11340
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   27
      Top             =   3300
      Width           =   2880
   End
   Begin VB.PictureBox picAngleA 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   1
      Left            =   9300
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   56
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picAngleA 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   3
      Left            =   11340
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   55
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picAngleA 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   2
      Left            =   10320
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   54
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   15
      Left            =   7200
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   53
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   14
      Left            =   6180
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   52
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   13
      Left            =   5160
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   51
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   12
      Left            =   4140
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   50
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   11
      Left            =   3120
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   49
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   10
      Left            =   2100
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   48
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   9
      Left            =   1080
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   47
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   8
      Left            =   60
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   46
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   7
      Left            =   7200
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   45
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   6
      Left            =   6180
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   44
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   5
      Left            =   5160
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   43
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   4
      Left            =   4140
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   42
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   3
      Left            =   3120
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   41
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   2
      Left            =   2100
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   40
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   1
      Left            =   1080
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   39
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picTile 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   0
      Left            =   60
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   38
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picAngleA 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   4
      Left            =   8220
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   37
      Top             =   6780
      Width           =   960
   End
   Begin VB.PictureBox picAngleA 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   0
      Left            =   8280
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   36
      Top             =   7800
      Width           =   960
   End
   Begin VB.PictureBox picMaskInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2880
      Index           =   4
      Left            =   11340
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   35
      Top             =   3300
      Width           =   2880
   End
   Begin VB.PictureBox picMaskInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2880
      Index           =   3
      Left            =   11340
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   34
      Top             =   3300
      Width           =   2880
   End
   Begin VB.PictureBox picMaskInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2880
      Index           =   2
      Left            =   8340
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   33
      Top             =   3600
      Width           =   2880
   End
   Begin VB.PictureBox picMaskInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2880
      Index           =   1
      Left            =   11340
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   32
      Top             =   3300
      Width           =   2880
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2880
      Index           =   0
      Left            =   11340
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   26
      Top             =   360
      Width           =   2880
   End
   Begin VB.CommandButton cmdDirectory 
      Caption         =   "Select Texture Directory"
      Height          =   375
      Left            =   6180
      TabIndex        =   23
      Top             =   120
      Width           =   1995
   End
   Begin VB.CommandButton cmdTestMap 
      Caption         =   "Show Test Map"
      Height          =   495
      Left            =   6240
      TabIndex        =   21
      Top             =   3360
      Width           =   1935
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   495
      Left            =   6240
      TabIndex        =   20
      Top             =   3900
      Width           =   1935
   End
   Begin VB.HScrollBar SprayAmount 
      Height          =   255
      Left            =   6240
      Max             =   100
      Min             =   20
      TabIndex        =   13
      Top             =   2160
      Value           =   40
      Width           =   1935
   End
   Begin VB.HScrollBar SprayLength 
      Height          =   255
      Left            =   6240
      Max             =   16
      Min             =   1
      TabIndex        =   12
      Top             =   1560
      Value           =   10
      Width           =   1935
   End
   Begin VB.HScrollBar EdgePixels 
      Height          =   255
      Left            =   6240
      Max             =   8
      Min             =   1
      TabIndex        =   11
      Top             =   960
      Value           =   2
      Width           =   1935
   End
   Begin VB.OptionButton optBoth 
      Caption         =   "Both"
      Height          =   255
      Left            =   7260
      TabIndex        =   10
      Top             =   3060
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.OptionButton optNone 
      Caption         =   "None"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   2760
      Width           =   855
   End
   Begin VB.OptionButton opt3DEdge 
      Caption         =   "3D Edge"
      Height          =   255
      Left            =   7260
      TabIndex        =   8
      Top             =   2760
      Width           =   975
   End
   Begin VB.OptionButton optSpray 
      Caption         =   "Spray"
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   3060
      Width           =   915
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Index           =   2
      Left            =   4200
      Pattern         =   "*.bmp;*.jpg"
      TabIndex        =   6
      Top             =   2760
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Index           =   1
      Left            =   2160
      Pattern         =   "*.bmp;*.jpg"
      TabIndex        =   5
      Top             =   2760
      Width           =   1935
   End
   Begin VB.FileListBox File1 
      Height          =   1650
      Index           =   0
      Left            =   120
      Pattern         =   "*.bmp;*.jpg"
      TabIndex        =   4
      Top             =   2760
      Width           =   1935
   End
   Begin VB.PictureBox picTex 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      DrawWidth       =   2
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   960
      Index           =   0
      Left            =   120
      ScaleHeight     =   64
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   64
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2880
      Index           =   2
      Left            =   8340
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   28
      Top             =   660
      Width           =   2880
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2880
      Index           =   4
      Left            =   11340
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   31
      Top             =   360
      Width           =   2880
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2880
      Index           =   1
      Left            =   11340
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   30
      Top             =   360
      Width           =   2880
   End
   Begin VB.PictureBox picMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H00000000&
      Height          =   2880
      Index           =   3
      Left            =   11340
      ScaleHeight     =   192
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   192
      TabIndex        =   29
      Top             =   360
      Width           =   2880
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   15
      Left            =   7260
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   14
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   13
      Left            =   5220
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   12
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   11
      Left            =   3180
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   10
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   9
      Left            =   1140
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   8
      Left            =   120
      Stretch         =   -1  'True
      Top             =   5520
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   7
      Left            =   7260
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   6
      Left            =   6240
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   5
      Left            =   5220
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   4
      Left            =   4200
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   3
      Left            =   3180
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   2
      Left            =   2160
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   1
      Left            =   1140
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   960
   End
   Begin VB.Image imgTile 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Index           =   0
      Left            =   120
      Stretch         =   -1  'True
      Top             =   4500
      Width           =   960
   End
   Begin VB.Image imgTex 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1920
      Index           =   2
      Left            =   4200
      Picture         =   "frmMain.frx":030A
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1920
   End
   Begin VB.Image imgTex 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1920
      Index           =   1
      Left            =   2160
      Picture         =   "frmMain.frx":C34C
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1920
   End
   Begin VB.Image imgTex 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1920
      Index           =   0
      Left            =   120
      Picture         =   "frmMain.frx":1838E
      Stretch         =   -1  'True
      Top             =   720
      Width           =   1920
   End
   Begin VB.Label Label1 
      Caption         =   "Texture Directory: = "
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Edge Style:"
      Height          =   255
      Index           =   3
      Left            =   6240
      TabIndex        =   24
      Top             =   2460
      Width           =   1215
   End
   Begin VB.Label lblDirectory 
      BackStyle       =   0  'Transparent
      Height          =   315
      Left            =   1620
      TabIndex        =   22
      Top             =   120
      Width           =   4515
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Spray Amount: = "
      Height          =   195
      Index           =   0
      Left            =   6240
      TabIndex        =   19
      Top             =   1860
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Spray Length: = "
      Height          =   195
      Index           =   1
      Left            =   6240
      TabIndex        =   18
      Top             =   1260
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Edge Pixels: = "
      Height          =   195
      Index           =   2
      Left            =   6240
      TabIndex        =   17
      Top             =   660
      Width           =   1035
   End
   Begin VB.Label lbl3Dedge 
      BackStyle       =   0  'Transparent
      Caption         =   "002"
      Height          =   195
      Left            =   7320
      TabIndex        =   16
      Top             =   660
      Width           =   375
   End
   Begin VB.Label lblsprayLength 
      BackStyle       =   0  'Transparent
      Caption         =   "010"
      Height          =   195
      Left            =   7440
      TabIndex        =   15
      Top             =   1260
      Width           =   495
   End
   Begin VB.Label lblSprayAmount 
      BackStyle       =   0  'Transparent
      Caption         =   "040"
      Height          =   195
      Left            =   7440
      TabIndex        =   14
      Top             =   1860
      Width           =   495
   End
   Begin VB.Label lblTexture 
      Caption         =   "Foreground Texture:"
      Height          =   195
      Index           =   2
      Left            =   4200
      TabIndex        =   3
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label lblTexture 
      Caption         =   "Edge Texture:"
      Height          =   195
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label lblTexture 
      Caption         =   "Background Texture:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1500
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save Tile Strip"
      End
      Begin VB.Menu mnuFileSeperator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuTileSizes 
      Caption         =   "&Tilesize"
      Begin VB.Menu mnuTileSize 
         Caption         =   "32 x 32"
         Index           =   0
      End
      Begin VB.Menu mnuTileSize 
         Caption         =   "48 x 48"
         Index           =   1
      End
      Begin VB.Menu mnuTileSize 
         Caption         =   "64 x 64"
         Checked         =   -1  'True
         Index           =   2
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'******************************************************************
'******************************************************************
'© Copyright 2002 Vincent Foster VBVince Software Co.
'******************************************************************
'******************************************************************
Option Explicit

Public Sub ClearTiles()

  Dim D As Integer

    For D = 0 To 4
        picMask(D).Cls
        picMaskInv(D).Cls
        picAngleA(D).Cls
        picAngleB(D).Cls
    Next D
    For D = 0 To 15
        picTile(D).Cls
    Next D

End Sub

Private Sub cmdDirectory_Click()

    On Error GoTo ErrHandler
  Dim D As Integer
  Dim strDir As String
    strDir = DirBox(Me.hwnd, "Select Texture Directory", App.Path)
    If strDir = "" Then Exit Sub
    lblDirectory.Caption = strDir
    For D = 0 To 2
        File1(D).Path = strDir
    Next D

Exit Sub

ErrHandler:
    MsgBox "err"

End Sub

Private Sub cmdGenerate_Click()

    MakeNewTiles

End Sub

Private Sub cmdTestMap_Click()

  Dim X As Integer

    For X = 0 To 15
        Set picTile(X).Picture = picTile(X).Image
    Next X
    frmMap.Show 1, Me

End Sub

Public Sub DrawTiles()
Dim S As Integer
    BitBlt picTile(15).hdc, 0, 0, Tilesize, Tilesize, picMask(2).hdc, Tilesize, Tilesize, SRCCOPY
    BitBlt picTile(1).hdc, 0, 0, Tilesize, Tilesize, picMask(2).hdc, Tilesize * 2, 0, SRCCOPY
    BitBlt picTile(2).hdc, 0, 0, Tilesize, Tilesize, picMask(2).hdc, 0, 0, SRCCOPY
    BitBlt picTile(3).hdc, 0, 0, Tilesize, Tilesize, picMask(2).hdc, Tilesize, 0, SRCCOPY
    BitBlt picTile(4).hdc, 0, 0, Tilesize, Tilesize, picMask(2).hdc, 0, Tilesize * 2, SRCCOPY
    BitBlt picTile(5).hdc, 0, 0, Tilesize, Tilesize, picAngleB(2).hdc, 0, 0, SRCCOPY
    BitBlt picTile(6).hdc, 0, 0, Tilesize, Tilesize, picMask(2).hdc, 0, Tilesize, SRCCOPY
    BitBlt picTile(7).hdc, 0, 0, Tilesize, Tilesize, picMaskInv(2).hdc, Tilesize * 2, Tilesize * 2, SRCCOPY
    BitBlt picTile(8).hdc, 0, 0, Tilesize, Tilesize, picMask(2).hdc, Tilesize * 2, Tilesize * 2, SRCCOPY
    BitBlt picTile(9).hdc, 0, 0, Tilesize, Tilesize, picMask(2).hdc, Tilesize * 2, Tilesize, SRCCOPY
    BitBlt picTile(10).hdc, 0, 0, Tilesize, Tilesize, picAngleA(2).hdc, 0, 0, SRCCOPY
    BitBlt picTile(11).hdc, 0, 0, Tilesize, Tilesize, picMaskInv(2).hdc, 0, Tilesize * 2, SRCCOPY
    BitBlt picTile(12).hdc, 0, 0, Tilesize, Tilesize, picMask(2).hdc, Tilesize, Tilesize * 2, SRCCOPY
    BitBlt picTile(13).hdc, 0, 0, Tilesize, Tilesize, picMaskInv(2).hdc, 0, 0, SRCCOPY
    BitBlt picTile(14).hdc, 0, 0, Tilesize, Tilesize, picMaskInv(2).hdc, Tilesize * 2, 0, SRCCOPY
    BitBlt picTile(0).hdc, 0, 0, Tilesize, Tilesize, picMaskInv(2).hdc, Tilesize, Tilesize, SRCCOPY
For S = 0 To 15
Set imgTile(S).Picture = picTile(S).Image
Next
End Sub

Private Sub EdgePixels_Change()

    lbl3Dedge.Caption = Format$(EdgePixels.Value, "000")
    cmdGenerate_Click

End Sub

Private Sub EdgePixels_Scroll()

    lbl3Dedge.Caption = Format$(EdgePixels.Value, "000")

End Sub

Private Sub File1_Click(Index As Integer)
On Error GoTo ErrHandler
  Dim D As Integer
Set picTex(Index).Picture = Nothing
'picTex(Index).Cls

    If Len(File1(Index).Path) = 3 Then
        imgTex(Index).Picture = LoadPicture(File1(Index).Path & File1(Index).FileName)
      Else
        imgTex(Index).Picture = LoadPicture(File1(Index).Path & "\" & File1(Index).FileName)
    End If
    picTex(Index).PaintPicture imgTex(Index).Picture, 0, 0, Tilesize, Tilesize
    Set picTex(Index).Picture = picTex(Index).Image


    cmdGenerate_Click
Exit Sub
ErrHandler:
File1(0).Refresh
File1(1).Refresh
File1(2).Refresh
End Sub

Private Sub Form_Load()
On Error GoTo ErrHandler
Dim T As Long
Dim D As Integer
    T = GetTickCount
    frmSplash.Show 0, Me
    frmSplash.Refresh
    Do Until (T + 2000) < (GetTickCount)
    Loop
    Unload frmSplash
    Me.Refresh
    Set frmSplash = Nothing

    lblDirectory.Caption = App.Path & "\Textures"
    For D = 0 To 2
        File1(D).Path = App.Path & "\Textures"
    Next D
    Tilesize = [64 x 64]
    For T = 0 To 2
    picTex(T).PaintPicture imgTex(T).Picture, 0, 0, Tilesize, Tilesize
    Set picTex(T).Picture = picTex(T).Image
    Next
    cmdGenerate_Click

Exit Sub

ErrHandler:
    lblDirectory.Caption = App.Path
    For D = 0 To 2
        File1(D).Path = App.Path
    Next D
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmMain = Nothing

End Sub

Public Sub MakeNewTiles()

  Dim J As Integer
  Dim K As Integer
  Dim L As Integer
  Dim R As Integer
  Dim M As Integer
  Dim S1 As Integer
  Dim S2 As Integer
  Select Case Tilesize
  Case 32
  S1 = 18
  S2 = 8
  Case 48, 64
  S1 = 36
  S2 = 16
  End Select
    ClearTiles

    If optSpray.Value = True Or optBoth.Value = True Then

        DrawRadarGraph picMask(0).hdc, Tilesize * 2, Tilesize, Tilesize \ 2, S1, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawRadarGraph picMask(0).hdc, Tilesize * 2, Tilesize * 2, Tilesize \ 2, S1, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawRadarGraph picMask(0).hdc, Tilesize, Tilesize, Tilesize \ 2, S1, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawRadarGraph picMask(0).hdc, Tilesize, Tilesize * 2, Tilesize \ 2, S1, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawFlatGraph picMask(0).hdc, Tilesize, 0, HorizontalN, Tilesize \ 2, S2, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawFlatGraph picMask(0).hdc, Tilesize, Tilesize * 2, HorizontalS, Tilesize \ 2, S2, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawFlatGraph picMask(0).hdc, Tilesize, 0, VerticalE, Tilesize \ 2, S2, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawFlatGraph picMask(0).hdc, Tilesize * 2, Tilesize, Verticalw, Tilesize \ 2, S2, True, SprayLength.Value, SprayAmount.Value, vbBlack
        picMask(0).Line (Tilesize, Tilesize)-(Tilesize * 2, Tilesize * 2), vbBlack, BF
        
        DrawRadarGraph picMaskInv(0).hdc, Tilesize * 2, Tilesize, Tilesize \ 2, S1, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawRadarGraph picMaskInv(0).hdc, Tilesize * 2, Tilesize * 2, Tilesize \ 2, S1, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawRadarGraph picMaskInv(0).hdc, Tilesize, Tilesize, Tilesize \ 2, S1, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawRadarGraph picMaskInv(0).hdc, Tilesize, Tilesize * 2, Tilesize \ 2, S1, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawFlatGraph picMaskInv(0).hdc, Tilesize, 0, HorizontalN, Tilesize \ 2, S2, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawFlatGraph picMaskInv(0).hdc, Tilesize, Tilesize * 2, HorizontalS, Tilesize \ 2, S2, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawFlatGraph picMaskInv(0).hdc, Tilesize, 0, VerticalE, Tilesize \ 2, S2, True, SprayLength.Value, SprayAmount.Value, vbBlack
        DrawFlatGraph picMaskInv(0).hdc, Tilesize * 2, Tilesize, Verticalw, Tilesize \ 2, S2, True, SprayLength.Value, SprayAmount.Value, vbBlack
        picMaskInv(0).Line (Tilesize, Tilesize)-(Tilesize * 2, Tilesize * 2), vbWhite, BF
      Else
        DrawRadarGraph picMask(0).hdc, Tilesize * 2, Tilesize, Tilesize \ 2, S1, False
        DrawRadarGraph picMask(0).hdc, Tilesize * 2, Tilesize * 2, Tilesize \ 2, S1, False
        DrawRadarGraph picMask(0).hdc, Tilesize, Tilesize, Tilesize \ 2, S1, False
        DrawRadarGraph picMask(0).hdc, Tilesize, Tilesize * 2, Tilesize \ 2, S1, False
        DrawFlatGraph picMask(0).hdc, Tilesize, 0, HorizontalN, Tilesize \ 2, S2, False
        DrawFlatGraph picMask(0).hdc, Tilesize, Tilesize * 2, HorizontalS, Tilesize \ 2, S2, False
        DrawFlatGraph picMask(0).hdc, Tilesize, 0, VerticalE, Tilesize \ 2, S2, False
        DrawFlatGraph picMask(0).hdc, Tilesize * 2, Tilesize, Verticalw, Tilesize \ 2, S2, False
        picMask(0).Line (Tilesize, Tilesize)-(Tilesize * 2, Tilesize * 2), vbBlack, BF
        DrawRadarGraph picMaskInv(0).hdc, Tilesize * 2, Tilesize, Tilesize \ 2, S1, False
        DrawRadarGraph picMaskInv(0).hdc, Tilesize * 2, Tilesize * 2, Tilesize \ 2, S1, False
        DrawRadarGraph picMaskInv(0).hdc, Tilesize, Tilesize, Tilesize \ 2, S1, False
        DrawRadarGraph picMaskInv(0).hdc, Tilesize, Tilesize * 2, Tilesize \ 2, S1, False
        DrawFlatGraph picMaskInv(0).hdc, Tilesize, 0, HorizontalN, Tilesize \ 2, S2, False
        DrawFlatGraph picMaskInv(0).hdc, Tilesize, Tilesize * 2, HorizontalS, Tilesize \ 2, S2, False
        DrawFlatGraph picMaskInv(0).hdc, Tilesize, 0, VerticalE, Tilesize \ 2, S2, False
        DrawFlatGraph picMaskInv(0).hdc, Tilesize * 2, Tilesize, Verticalw, Tilesize \ 2, S2, False
        picMaskInv(0).Line (Tilesize, Tilesize)-(Tilesize * 2, Tilesize * 2), vbWhite, BF
    End If
    'make angle tiles

    'Build 3d Edges
    For K = 0 To 2
        For J = 0 To 2
            For L = 0 To EdgePixels.Value
                TransparentBlt picMask(1).hdc, K * Tilesize + L, J * Tilesize + L, Tilesize, Tilesize, picMask(0).hdc, K * Tilesize, J * Tilesize, vbWhite
                TransparentBlt picMaskInv(1).hdc, K * Tilesize + L, J * Tilesize + L, Tilesize, Tilesize, picMaskInv(0).hdc, K * Tilesize, J * Tilesize, vbWhite
            Next L
        Next J
    Next K
    For K = 0 To 2
        For J = 0 To 2
            BitBlt picMask(2).hdc, K * Tilesize, J * Tilesize, 192, 192, picTex(0).hdc, 0, 0, SRCCOPY
            BitBlt picMask(3).hdc, K * Tilesize, J * Tilesize, 192, 192, picTex(1).hdc, 0, 0, SRCCOPY
            BitBlt picMask(4).hdc, K * Tilesize, J * Tilesize, 192, 192, picTex(2).hdc, 0, 0, SRCCOPY
            BitBlt picMaskInv(2).hdc, K * Tilesize, J * Tilesize, 192, 192, picTex(0).hdc, 0, 0, SRCCOPY
            BitBlt picMaskInv(3).hdc, K * Tilesize, J * Tilesize, 192, 192, picTex(1).hdc, 0, 0, SRCCOPY
            BitBlt picMaskInv(4).hdc, K * Tilesize, J * Tilesize, 192, 192, picTex(2).hdc, 0, 0, SRCCOPY
        Next J
    Next K
    'Draw AngleA
    BitBlt picAngleA(2).hdc, 0, 0, Tilesize, Tilesize, picTex(0).hdc, 0, 0, SRCCOPY
    BitBlt picAngleA(3).hdc, 0, 0, Tilesize, Tilesize, picTex(1).hdc, 0, 0, SRCCOPY
    BitBlt picAngleA(4).hdc, 0, 0, Tilesize, Tilesize, picTex(2).hdc, 0, 0, SRCCOPY

    BitBlt picAngleA(0).hdc, 0, 0, Tilesize, Tilesize, picMask(0).hdc, 0, 0, SRCCOPY
    TransparentBlt picAngleA(0).hdc, 0, 0, Tilesize, Tilesize, picMask(0).hdc, Tilesize * 2, Tilesize * 2, vbWhite
    BitBlt picAngleA(1).hdc, 0, 0, Tilesize, Tilesize, picMask(1).hdc, 0, 0, SRCCOPY
    TransparentBlt picAngleA(1).hdc, 0, 0, Tilesize, Tilesize, picMask(1).hdc, Tilesize * 2, Tilesize * 2, vbWhite

    TransparentBlt picAngleA(4).hdc, 0, 0, Tilesize, Tilesize, picAngleA(0).hdc, 0, 0, vbBlack
    TransparentBlt picAngleA(3).hdc, 0, 0, Tilesize, Tilesize, picAngleA(1).hdc, 0, 0, vbBlack
    If opt3DEdge.Value = True Or optBoth.Value = True Then
        TransparentBlt picAngleA(2).hdc, 0, 0, Tilesize, Tilesize, picAngleA(3).hdc, 0, 0, vbWhite
    End If
    TransparentBlt picAngleA(2).hdc, 0, 0, Tilesize, Tilesize, picAngleA(4).hdc, 0, 0, vbWhite
    'DrawAngleB
    BitBlt picAngleB(2).hdc, 0, 0, Tilesize, Tilesize, picTex(0).hdc, 0, 0, SRCCOPY
    BitBlt picAngleB(3).hdc, 0, 0, Tilesize, Tilesize, picTex(1).hdc, 0, 0, SRCCOPY
    BitBlt picAngleB(4).hdc, 0, 0, Tilesize, Tilesize, picTex(2).hdc, 0, 0, SRCCOPY

    BitBlt picAngleB(0).hdc, 0, 0, Tilesize, Tilesize, picMask(0).hdc, Tilesize * 2, 0, SRCCOPY
    TransparentBlt picAngleB(0).hdc, 0, 0, Tilesize, Tilesize, picMask(0).hdc, 0, Tilesize * 2, vbWhite
    BitBlt picAngleB(1).hdc, 0, 0, Tilesize, Tilesize, picMask(1).hdc, Tilesize * 2, 0, SRCCOPY
    TransparentBlt picAngleB(1).hdc, 0, 0, Tilesize, Tilesize, picMask(1).hdc, 0, Tilesize * 2, vbWhite

    TransparentBlt picAngleB(4).hdc, 0, 0, Tilesize, Tilesize, picAngleB(0).hdc, 0, 0, vbBlack
    TransparentBlt picAngleB(3).hdc, 0, 0, Tilesize, Tilesize, picAngleB(1).hdc, 0, 0, vbBlack
    If opt3DEdge.Value = True Or optBoth.Value = True Then
        TransparentBlt picAngleB(2).hdc, 0, 0, Tilesize, Tilesize, picAngleB(3).hdc, 0, 0, vbWhite
    End If
    TransparentBlt picAngleB(2).hdc, 0, 0, Tilesize, Tilesize, picAngleB(4).hdc, 0, 0, vbWhite

    'DrawPics
    For K = 0 To 2
        For J = 0 To 2
            TransparentBlt picMask(4).hdc, K * Tilesize, J * Tilesize, Tilesize, Tilesize, picMask(0).hdc, K * Tilesize, J * Tilesize, vbBlack
            TransparentBlt picMask(3).hdc, K * Tilesize, J * Tilesize, Tilesize, Tilesize, picMask(1).hdc, K * Tilesize, J * Tilesize, vbBlack
            TransparentBlt picMaskInv(4).hdc, K * Tilesize, J * Tilesize, Tilesize, Tilesize, picMaskInv(0).hdc, K * Tilesize, J * Tilesize, vbBlack
            TransparentBlt picMaskInv(3).hdc, K * Tilesize, J * Tilesize, Tilesize, Tilesize, picMaskInv(1).hdc, K * Tilesize, J * Tilesize, vbBlack
            If opt3DEdge.Value = True Or optBoth.Value = True Then
                TransparentBlt picMask(2).hdc, K * Tilesize, J * Tilesize, Tilesize, Tilesize, picMask(3).hdc, K * Tilesize, J * Tilesize, vbWhite
                TransparentBlt picMaskInv(2).hdc, K * Tilesize, J * Tilesize, Tilesize, Tilesize, picMaskInv(3).hdc, K * Tilesize, J * Tilesize, vbWhite
            End If
            TransparentBlt picMask(2).hdc, K * Tilesize, J * Tilesize, Tilesize, Tilesize, picMask(4).hdc, K * Tilesize, J * Tilesize, vbWhite
            TransparentBlt picMaskInv(2).hdc, K * Tilesize, J * Tilesize, Tilesize, Tilesize, picMaskInv(4).hdc, K * Tilesize, J * Tilesize, vbWhite
        Next J
    Next K

    DrawTiles

End Sub

Private Sub mnuFileExit_Click()

    SaveAllSettings
    Unload Me

End Sub

Private Sub mnuFileSave_Click()
  On Error GoTo ErrHandler
  Dim sFile As String
  Dim Y As Integer
   picStrip.Cls
   picStrip.Width = Tilesize
   picStrip.Height = Tilesize * 16
   
    For Y = 0 To 15
        BitBlt picStrip.hdc, 0, Y * Tilesize, Tilesize, Tilesize, picTile(Y).hdc, 0, 0, SRCCOPY
    Next Y
    sFile = fncGetFileNametoSave("(*.bmp)|*.bmp", "*.bmp", App.Path, "Save Tile Strip")
    If Len(sFile) > 3 Then
        SavePicture picStrip.Image, sFile
    End If

Exit Sub

ErrHandler:

End Sub

Private Sub mnuHelpAbout_Click()

    frmAbout.Show 1, Me

End Sub

Private Sub mnuTileSize_Click(Index As Integer)
Dim S As Integer
    For S = 0 To 2
        mnuTileSize(S).Checked = False
        picTex(S).Width = Tilesize
    Next
    mnuTileSize(Index).Checked = True
    Select Case Index
        Case 0
            Tilesize = [32 x 32]
        Case 1
            Tilesize = [48 x 48]
        Case 2
            Tilesize = [64 x 64]
    End Select
        For S = 0 To 2
        picTex(S).Width = Tilesize
        picTex(S).Height = Tilesize
        Next
        For S = 0 To 15
        picTile(S).Width = Tilesize
        picTile(S).Height = Tilesize
        
        Next
    For S = 0 To 2
    picTex(S).PaintPicture imgTex(S).Picture, 0, 0, Tilesize, Tilesize
    Set picTex(S).Picture = picTex(S).Image
    Next
        cmdGenerate_Click

End Sub

Private Sub opt3DEdge_Click()

    cmdGenerate_Click

End Sub

Private Sub optBoth_Click()

    cmdGenerate_Click

End Sub

Private Sub optNone_Click()

    cmdGenerate_Click

End Sub

Private Sub optSpray_Click()

    cmdGenerate_Click

End Sub

Private Sub picMask_Click(Index As Integer)

    SavePicture picMask(Index).Image, App.Path & "\test.bmp"

End Sub

Private Sub SprayAmount_Change()

    lblSprayAmount.Caption = Format$(SprayAmount.Value, "000")
    cmdGenerate_Click

End Sub

Private Sub SprayAmount_Scroll()

    lblSprayAmount.Caption = Format$(SprayAmount.Value, "000")

End Sub

Private Sub SprayLength_Change()

    lblsprayLength.Caption = Format$(SprayLength.Value, "000")
    cmdGenerate_Click

End Sub

Private Sub SprayLength_Scroll()

    lblsprayLength.Caption = Format$(SprayLength.Value, "000")
End Sub

':) Ulli's VB Code Formatter V2.5.12 (1/24/2002 12:41:25 AM) 1 + 349 = 350 Lines
