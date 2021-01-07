VERSION 5.00
Begin VB.Form frmAlveograma_Captura 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captura de alveograma"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14445
   Icon            =   "frmAlveograma_Captura.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   7050
      Left            =   12780
      TabIndex        =   5
      Top             =   315
      Width           =   1590
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Limpiar"
      Height          =   1050
      Left            =   12825
      Picture         =   "frmAlveograma_Captura.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7470
      Width           =   1545
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   1050
      Left            =   12825
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8550
      Width           =   1545
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      DrawWidth       =   3
      ForeColor       =   &H80000008&
      Height          =   9285
      Left            =   45
      ScaleHeight     =   802.602
      ScaleMode       =   0  'User
      ScaleWidth      =   1026.438
      TabIndex        =   0
      Top             =   315
      Width           =   12660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "Captura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   12735
      TabIndex        =   6
      Top             =   90
      Width           =   1635
   End
   Begin VB.Label ly 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ly"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1575
      TabIndex        =   4
      Top             =   45
      Width           =   1635
   End
   Begin VB.Label lx 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   1455
   End
End
Attribute VB_Name = "frmAlveograma_Captura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ALTO = 800
Dim X1 As Long      'This Is The X Position Of The Last Line Drawn
Dim y1 As Long      'This Is The Y Position Of The Last Line Drawn
Dim X2 As Long      'This Is The Start Mark Of The Box or Circle
Dim y2 As Long      'This Is The Satrt Mark Of The Box or Circle

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    Picture1.Cls
    List1.Clear
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
End Sub

Private Sub Picture1_Click()
'    If Button = vbLeftButton Then
        'Draw Continuous Line...
'        Picture1.Line (X1, y1)-(x + 1, y + 1), vbBlack
        Picture1.Line (X1, y1)-(X1, y1), vbBlack
        
        If X2 <> X1 And y2 <> y1 Then
            List1.AddItem X1 & "-" & ALTO - y1
            X2 = X1
            y2 = y1
        End If
'        Picture1.Point X2, y2
'        X1 = x: y1 = y
'    End If

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    X1 = x
    y1 = y
'    X2 = x
'    y2 = y
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'    If Button = vbLeftButton Then
'        'Draw Continuous Line...
'        Picture1.Line (X1, y1)-(x, y), vbBlack
'        X1 = x: y1 = y
'    End If
    lx = "X: " & x
    ly = "Y: " & y
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'    X1 = 0
'    y1 = 0
'    lx = x
'    ly = y
End Sub
