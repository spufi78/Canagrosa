VERSION 5.00
Begin VB.Form frmLogo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5565
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12435
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5565
   ScaleWidth      =   12435
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image2 
      Height          =   2370
      Left            =   4050
      Picture         =   "frmLogo.frx":0000
      Top             =   1470
      Width           =   2970
   End
   Begin VB.Image Image1 
      Height          =   1185
      Left            =   90
      Picture         =   "frmLogo.frx":88F3
      Top             =   90
      Width           =   11910
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
