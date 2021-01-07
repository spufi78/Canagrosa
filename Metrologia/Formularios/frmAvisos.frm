VERSION 5.00
Begin VB.Form frmAvisos 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Avisos del Cliente"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5100
   Icon            =   "frmAvisos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   5100
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "ESC-Salir"
      Height          =   855
      Left            =   3810
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3150
      Width           =   1245
   End
   Begin VB.TextBox txtaviso 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2715
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   5055
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Avisos del cliente"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   5070
   End
End
Attribute VB_Name = "frmAvisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    cargar_botones Me
    If PK <> 0 Then
        Dim oObra As New clsObras
        oObra.Carga (PK)
        txtaviso.Text = oObra.getAVISOS
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 27
        cmdcancel_Click
    End Select
End Sub

