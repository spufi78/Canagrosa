VERSION 5.00
Begin VB.Form frmMEN_Nuevo 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Consultas"
   ClientHeight    =   6690
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3675
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmMEN_Nuevo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   3675
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox texto 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   2715
      Index           =   0
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   315
      Width           =   3480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   600
      Left            =   990
      TabIndex        =   2
      Top             =   5940
      Width           =   1635
   End
   Begin VB.PictureBox mensaje1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080FFFF&
      ForeColor       =   &H80000008&
      Height          =   3255
      Index           =   0
      Left            =   45
      ScaleHeight     =   3225
      ScaleWidth      =   3540
      TabIndex        =   0
      Top             =   315
      Width           =   3570
   End
   Begin VB.Label lbltitulo 
      BackColor       =   &H0080FFFF&
      Caption         =   "Label2"
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
      Height          =   420
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   3435
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de mensajes del usuario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   450
      TabIndex        =   1
      Top             =   45
      Width           =   2805
   End
End
Attribute VB_Name = "frmMEN_Nuevo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
'    Dim i As Integer
'    For i = 0 To mensaje1.Count - 1
'        Set mensaje1(i) = Nothing
'    Next
    Unload Me
End Sub

Public Sub carga()

    Dim oMensaje As New clsMensajes
    Dim rs As ADODB.Recordset
    Set rs = oMensaje.Listado
    If rs.RecordCount > 0 Then
        Do
            Index = mensaje1.Count
            lbltitulo(Index - 1).Caption = rs("asunto")
            texto(Index - 1).Text = rs("texto")
            Set lbltitulo(Index - 1).Container = mensaje1(Index - 1)
            Set texto(Index - 1).Container = mensaje1(Index - 1)
            mensaje1(Index - 1).ZOrder Index - 1
            rs.MoveNext
            If rs.EOF = False Then
                Load mensaje1(Index)
                Load lbltitulo(Index)
                Load texto(Index)
                Set lbltitulo(Index).Container = mensaje1(Index)
                Set texto(Index).Container = mensaje1(Index)
                mensaje1(Index).Top = mensaje1(Index - 1).Top + lbltitulo(Index).Height + 500
                mensaje1(Index).Visible = True
                texto(Index).Visible = True
                lbltitulo(Index).Visible = True
            End If
        Loop Until rs.EOF
    End If
    Set oMensaje = Nothing
    Set rs = Nothing

End Sub

Private Sub Form_Load()
    Me.Left = 11000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMEN_Nuevo = Nothing
End Sub

Private Sub lbltitulo_Click(Index As Integer)
    If mensaje1(Index).Height = 3255 Then
        mensaje1(Index).Height = 510
    Else
        mensaje1(Index).Height = 3255
    End If
End Sub
