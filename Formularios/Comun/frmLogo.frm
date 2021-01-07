VERSION 5.00
Begin VB.Form frmLogo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7860
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7860
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   6525
      Left            =   90
      Picture         =   "frmLogo.frx":0000
      Stretch         =   -1  'True
      Top             =   45
      Width           =   6465
   End
End
Attribute VB_Name = "frmLogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CentrarForma Me
End Sub

