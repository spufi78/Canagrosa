VERSION 5.00
Begin VB.Form frmEscaner 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Esperando Escáner..."
   ClientHeight    =   2295
   ClientLeft      =   4485
   ClientTop       =   3195
   ClientWidth     =   6885
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdjuntar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntar"
      Enabled         =   0   'False
      Height          =   870
      Left            =   1575
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1350
      Width           =   1500
   End
   Begin VB.CommandButton CMDMOSTRAR 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Documento"
      Enabled         =   0   'False
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1350
      Width           =   1500
   End
   Begin VB.DirListBox Dir1 
      Height          =   540
      Left            =   3465
      TabIndex        =   5
      Top             =   1170
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   3465
      TabIndex        =   4
      Top             =   1755
      Visible         =   0   'False
      Width           =   2220
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   5805
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1350
      Width           =   1050
   End
   Begin VB.TextBox txtlog 
      Alignment       =   2  'Center
      BackColor       =   &H00404040&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   -135
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   585
      Width           =   7035
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5985
      Top             =   900
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Introduzca el documento en el Escáner"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   45
      TabIndex        =   7
      Top             =   45
      Width           =   5685
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   6390
      Picture         =   "frmEscaner.frx":0000
      Top             =   45
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Height          =   570
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   6900
   End
End
Attribute VB_Name = "frmEscaner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdjuntar_Click()
    documento_escaner = Dir1.Path & "\" & File1.List(0)
    documento_escaner_nombre = File1.List(0)
    Unload Me
End Sub

Private Sub cmdcancel_Click()
    documento_escaner = ""
    documento_escaner_nombre = ""
    Unload Me
End Sub

Private Sub cmdMostrar_Click()
'    If MsgBox("¿Desea visualizar el documento escaneado?", vbQuestion + vbYesNo, App.Title) = vbYes Then
   On Error GoTo CMDMOSTRAR_Click_Error

         r = Shell("rundll32.exe url.dll,FileProtocolHandler " & Dir1.Path & "\" & File1.List(0), vbNormalFocus)
'    End If

   On Error GoTo 0
   Exit Sub

CMDMOSTRAR_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CMDMOSTRAR_Click of Formulario frmEscaner"
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Form_Load()
Dim strRuta As String

   On Error GoTo Form_Load_Error

    log Me.Name
    cargar_botones Me
    documento_escaner = ""
    documento_escaner_nombre = ""
    strRuta = localizar_directorio_escaneo_equipo()
    borrar_documentos strRuta
    
    'Dir1.Path = ReadINI(App.Path & "\config.ini", "Documentos", "Escaner")
    Dir1.Path = strRuta

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmEscaner"
End Sub

Private Sub Timer1_Timer()
    txtlog = "Esperando documento...."
    File1.Refresh
    If File1.ListCount > 0 Then
        txtlog.ForeColor = &HFF00&
        txtlog = "Documento localizado. Recuperando documento."
        cmdMostrar.Enabled = True
        cmdAdjuntar.Enabled = True
'        borrar_documentos
'        File1.Refresh
    End If
End Sub

Private Sub borrar_documentos(Optional ByVal ruta As String = "")
    On Error Resume Next
    If Trim(ruta) = "" Then
        Kill ReadINI(App.Path & "\config.ini", "Documentos", "Escaner") & "\*.*"
    Else
        If Right(ruta, 1) = "\" Then
            Kill ruta & "*.*"
        Else
            Kill ruta & "\*.*"
        End If
    End If
End Sub
