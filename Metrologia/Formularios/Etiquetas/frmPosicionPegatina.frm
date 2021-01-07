VERSION 5.00
Begin VB.Form frmPosicionPegatina 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Posición Pegatina"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4605
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1260
      Picture         =   "frmPosicionPegatina.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3735
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2385
      Picture         =   "frmPosicionPegatina.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3735
      Width           =   1050
   End
   Begin VB.CommandButton peg 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Index           =   1
      Left            =   45
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1740
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Posición Pegatina"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   3450
   End
End
Attribute VB_Name = "frmPosicionPegatina"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    pegatina = 0
    Unload Me
End Sub

Private Sub cmdok_Click()
    For i = 1 To peg.Count
        If peg(i).BackColor = vbYellow Then
            pegatina = i
            Unload Me
        End If
    Next
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim filas As Integer
    Dim col As Integer
'    filas = ReadINI(App.Path + "\config.ini", "Otros", "Pegatinas_Filas")
'    col = ReadINI(App.Path + "\config.ini", "Otros", "Pegatinas_Columnas")
    filas = 4
    col = 2
    For i = 2 To filas * col
        Load peg(i)
        If i Mod col = 0 Then
            peg(i).Left = peg(i - 1).Left + peg(i - 1).Width + 5
            peg(i).Top = peg(i - 1).Top
        Else
            peg(i).Left = peg(1).Left
            peg(i).Top = peg(i - 1).Top + peg(i - 1).Height + 1
        End If
        peg(i).Caption = i
        peg(i).Visible = True
    Next
    Me.Height = (peg(1).Height * filas) + lbltitulo.Height + cmdok.Height + 500
    cmdok.Top = (peg(1).Height * filas) + lbltitulo.Height + 100
    cmdcancel.Top = cmdok.Top
    peg(1).BackColor = vbYellow
End Sub

Private Sub peg_Click(Index As Integer)
    For i = 1 To peg.Count
        peg(i).BackColor = vbWhite
    Next
    peg(Index).BackColor = vbYellow
End Sub
