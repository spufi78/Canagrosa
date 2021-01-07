VERSION 5.00
Begin VB.Form frmPosicionPegatina 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Posición Pegatina"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3465
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   3465
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3735
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   2385
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
      Height          =   735
      Index           =   1
      Left            =   45
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   360
      Width           =   1680
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
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
      Width           =   3390
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
    log Me.Name
    cargar_botones Me
    Dim i As Integer
    Dim filas As Integer
    Dim Col As Integer
    filas = ReadINI(App.Path + "\config.ini", "Otros", "Pegatinas_Filas")
    Col = ReadINI(App.Path + "\config.ini", "Otros", "Pegatinas_Columnas")
    For i = 2 To filas * Col
        Load peg(i)
        If i Mod Col = 0 Then
            peg(i).Left = peg(i - 1).Left + peg(i - 1).Width + 5
            peg(i).top = peg(i - 1).top
        Else
            peg(i).Left = peg(1).Left
            peg(i).top = peg(i - 1).top + peg(i - 1).Height + 1
        End If
        peg(i).Caption = i
        peg(i).visible = True
    Next
    Me.Height = (peg(1).Height * filas) + lbltitulo.Height + cmdok.Height + 500
    cmdok.top = (peg(1).Height * filas) + lbltitulo.Height + 100
    cmdcancel.top = cmdok.top
    peg(1).BackColor = vbYellow
End Sub

Private Sub peg_Click(Index As Integer)
    For i = 1 To peg.Count
        peg(i).BackColor = vbWhite
    Next
    peg(Index).BackColor = vbYellow
End Sub
