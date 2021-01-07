VERSION 5.00
Begin VB.Form frmCambioUsuario 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Cambio Usuario"
   ClientHeight    =   2985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1755
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   1755
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   90
      Top             =   2070
   End
   Begin VB.CommandButton cmdCambiar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cambiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   90
      Picture         =   "frmCambioUsuario.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2160
      Width           =   1545
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   1605
      Begin VB.Image img 
         Appearance      =   0  'Flat
         Height          =   1755
         Left            =   90
         Stretch         =   -1  'True
         Top             =   225
         Width           =   1425
      End
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H0080FFFF&
      Caption         =   "Mensajes sin leer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   2655
      Visible         =   0   'False
      Width           =   1500
   End
End
Attribute VB_Name = "frmCambioUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCambiar_Click()
    glogin = 1
    frmMenu.cambiar_usuario
    cargar_usuario
End Sub
Private Sub Form_Load()
    log (Me.Name)
    mensajes
    Me.Left = Screen.Width - Me.Width - frmMenu.ButtonBar.Width - 500
    Me.Top = Screen.Height - Me.Height - frmMenu.SmartMenuXP1.Height - frmMenu.StatusBar1.Height - 1000
    cargar_usuario
End Sub

Public Sub cargar_usuario()
'    On Error Resume Next
   On Error GoTo cargar_usuario_Error

    Frame3.Caption = "Usuario: " & USUARIO.getUSUARIO
    If USUARIO.getIMAGEN <> "" Then
        If Dir(USUARIO.getIMAGEN) <> "" Then
            Set img.Picture = LoadPicture(USUARIO.getIMAGEN)
        End If
    Else
        Set img.Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "no"))
    End If

   On Error GoTo 0
   Exit Sub

cargar_usuario_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_usuario of Formulario frmCambioUsuario"
End Sub
Private Sub mensajes()
    Dim oMensaje As New clsMensajes
    If oMensaje.Mensajes_Sin_Leer Then
        lblMsg.Visible = True
    Else
        lblMsg.Visible = False
    End If
End Sub
Private Sub Timer1_Timer()
    mensajes
End Sub
