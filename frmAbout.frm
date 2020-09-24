VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7155
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   477
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   900
      Left            =   60
      Picture         =   "frmAbout.frx":030A
      ScaleHeight     =   840
      ScaleWidth      =   6960
      TabIndex        =   3
      Top             =   60
      Width           =   7020
   End
   Begin VB.Label lblCopyright 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail: alienheretic@attbi.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   6945
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Beta - Version 1.0.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   5340
      TabIndex        =   1
      Top             =   1080
      Width           =   1725
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "frmAbout.frx":133CE
      Top             =   960
      Width           =   480
   End
   Begin VB.Label lblEmail 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-Mail: alienheretic@attbi.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   600
      TabIndex        =   0
      Top             =   1080
      Width           =   2610
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()

    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    lblVersion.Caption = "Beta - Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblCopyright.Caption = App.LegalCopyright

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With lblEmail
        .ForeColor = vbBlack
        .FontUnderline = False
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set frmAbout = Nothing

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Call lblEmail_MouseDown(Button, Shift, X, Y)

End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With lblEmail
        .ForeColor = vbBlue
        .FontUnderline = True
    End With

End Sub

Private Sub lblEmail_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    sendemail

End Sub

Private Sub lblEmail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With lblEmail
        .ForeColor = vbBlue
        .FontUnderline = True
    End With

End Sub

Private Sub Picture1_Click()

    Unload Me
End Sub

':) Ulli's VB Code Formatter V2.5.12 (1/24/2002 12:40:46 AM) 1 + 73 = 74 Lines
