VERSION 5.00
Begin VB.Form frmInicio 
   BackColor       =   &H00808080&
   Caption         =   "Servidor de PDF"
   ClientHeight    =   2430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4455
   Icon            =   "frmInicio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2430
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   2025
      Top             =   1350
   End
   Begin VB.CommandButton cmdPrueba 
      Caption         =   "B.D. Prueba"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   2340
      Picture         =   "frmInicio.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   720
      Width           =   2040
   End
   Begin VB.CommandButton cmdReal 
      Caption         =   "B.D. Real"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      Left            =   135
      Picture         =   "frmInicio.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   2040
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "No seleccione ninguna opcion. El sistema la identificara en 3 segundos."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   90
      TabIndex        =   2
      Top             =   90
      Width           =   4290
   End
End
Attribute VB_Name = "frmInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdPrueba_Click()
    database = ReadINI(App.Path + "\config.ini", "server", "bd_prueba")
    Unload Me
    frmPDF.Show 1
End Sub

Private Sub cmdReal_Click()
    database = ReadINI(App.Path + "\config.ini", "server", "bd")
    Unload Me
    frmPDF.Show 1
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then
        Timer1.Enabled = False
        MsgBox "El servidor de impresión ya se encuentra en ejecución. Verifique la ejecución anterior.", vbInformation, App.Title
        End
    End If
End Sub

Private Sub Timer1_Timer()
    cmdReal_Click
End Sub
