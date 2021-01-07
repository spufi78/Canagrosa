VERSION 5.00
Begin VB.Form frmLogoPrueba 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7575
   ControlBox      =   0   'False
   Enabled         =   0   'False
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   Begin VB.Image Image1 
      Height          =   2370
      Left            =   45
      Picture         =   "frmLogoPrueba.frx":0000
      Top             =   -180
      Width           =   7500
   End
End
Attribute VB_Name = "frmLogoPrueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    CentrarForma Me
End Sub
