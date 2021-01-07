VERSION 5.00
Begin VB.Form frmMotivo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Introduzca motivo"
   ClientHeight    =   4215
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4215
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   5895
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3285
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   4770
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3285
      Width           =   1050
   End
   Begin VB.TextBox txttexto 
      Appearance      =   0  'Flat
      Height          =   2580
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   675
      Width           =   6900
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   6435
      Picture         =   "frmMotivo.frx":0000
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique el motivo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   5115
      WordWrap        =   -1  'True
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   7005
   End
End
Attribute VB_Name = "frmMotivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    MOTIVO = ""
    Unload Me
End Sub

Private Sub cmdok_Click()
    MOTIVO = txttexto
    Unload Me
End Sub

Private Sub Form_Load()
    MOTIVO = ""
    log (Me.Name)
    cargar_botones Me
End Sub
