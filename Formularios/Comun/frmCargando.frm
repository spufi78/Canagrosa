VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCargando 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Cargando contenido"
   ClientHeight    =   1065
   ClientLeft      =   5415
   ClientTop       =   5010
   ClientWidth     =   8025
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1065
   ScaleWidth      =   8025
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar pgbar 
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Min             =   1
   End
   Begin VB.Timer tmr 
      Interval        =   250
      Left            =   7320
      Top             =   660
   End
   Begin VB.Label lblMensaje 
      Alignment       =   2  'Center
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "frmCargando"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Event pasarela(ByRef Cancel As Boolean)
Private cancelar As Boolean
Private Sub Form_Load()
cancelar = False
pgbar.value = 1
End Sub


Private Sub tmr_Timer()

RaiseEvent pasarela(cancelar)

If pgbar.value < 100 Then
    pgbar.value = pgbar.value + 1
End If

If cancelar Then
    tmr.Enabled = False
    Unload Me
End If

End Sub


