VERSION 5.00
Begin VB.Form frmDocumentos_Observaciones 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Observaciones de la linea del pedido"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   Icon            =   "frmDocumentos_Observaciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6660
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   4320
      Picture         =   "frmDocumentos_Observaciones.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3150
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   5520
      Picture         =   "frmDocumentos_Observaciones.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3150
      Width           =   1155
   End
   Begin VB.TextBox txtaviso 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
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
      Width           =   6585
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Observaciones de la linea del pedido"
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
      Width           =   6600
   End
End
Attribute VB_Name = "frmDocumentos_Observaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public OB As String
Private Sub cmdAceptar_Click()
    frmDocumento.OB = txtaviso.Text
    Unload Me
End Sub

Private Sub cmdcancel_Click()
    frmDocumento.OB = "---"
    Unload Me
End Sub

Private Sub Form_Load()
    txtaviso.Text = OB
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 27
        cmdcancel_Click
    End Select
End Sub

