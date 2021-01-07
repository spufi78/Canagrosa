VERSION 5.00
Object = "{91ED0830-8EBD-4FB3-BBE6-C5253C9895EF}#1.0#0"; "DVRClient.ocx"
Begin VB.Form frmCamaras 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Camaras"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11850
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCamaras.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8505
   ScaleWidth      =   11850
   WindowState     =   1  'Minimized
   Begin DVRClient.ClientMain ClientMain1 
      Height          =   8520
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   15028
   End
End
Attribute VB_Name = "frmCamaras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End Sub

Private Sub Form_Load()
    ClientMain1.Connect "camaras", "canagr0sa", "canagrosa"

End Sub
