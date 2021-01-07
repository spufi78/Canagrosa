VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfiguracion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuración del programa"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11370
   Icon            =   "frmConfiguracion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   11370
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkimprimir 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Imprimir directamente a la impresora"
      Height          =   225
      Left            =   5850
      TabIndex        =   38
      Top             =   2520
      Width           =   5295
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   14
      Left            =   6855
      TabIndex        =   37
      Top             =   2010
      Width           =   4335
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   13
      Left            =   9240
      TabIndex        =   35
      Top             =   1200
      Width           =   1965
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   12
      Left            =   9240
      TabIndex        =   34
      Top             =   1590
      Width           =   1965
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   11
      Left            =   9240
      TabIndex        =   33
      Top             =   420
      Width           =   1965
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   10
      Left            =   9240
      TabIndex        =   32
      Top             =   810
      Width           =   1965
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   9
      Left            =   975
      TabIndex        =   26
      Top             =   2400
      Width           =   4455
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   8
      Left            =   975
      TabIndex        =   25
      Top             =   2790
      Width           =   4455
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   5
      Left            =   975
      TabIndex        =   24
      Top             =   1620
      Width           =   4455
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   6
      Left            =   975
      TabIndex        =   23
      Top             =   2010
      Width           =   4455
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   7
      Left            =   690
      TabIndex        =   18
      Top             =   4050
      Width           =   3945
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Explorar"
      Height          =   345
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4050
      Width           =   795
   End
   Begin VB.CommandButton cmdEXplorar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Explorar"
      Height          =   345
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   3660
      Width           =   795
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   1770
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   4800
      Width           =   3705
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   690
      TabIndex        =   2
      Top             =   3660
      Width           =   3945
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   0
      Left            =   1770
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   3
      Left            =   1770
      TabIndex        =   4
      Top             =   4440
      Width           =   3705
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   1
      Left            =   1770
      TabIndex        =   1
      Top             =   840
      Width           =   3735
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   9780
      Picture         =   "frmConfiguracion.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4215
      Width           =   1365
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   8310
      Picture         =   "frmConfiguracion.frx":4344
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4215
      Width           =   1365
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6900
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Msg. Cert."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   14
      Left            =   5850
      TabIndex        =   36
      Top             =   2040
      Width           =   885
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Retención de Garantía"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   13
      Left            =   5820
      TabIndex        =   31
      Top             =   1230
      Width           =   2010
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Copias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   12
      Left            =   5835
      TabIndex        =   30
      Top             =   1620
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Parámetros generales"
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
      Index           =   3
      Left            =   5835
      TabIndex        =   29
      Top             =   60
      Width           =   5460
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "I.v.a."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   11
      Left            =   5835
      TabIndex        =   28
      Top             =   450
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Retención de Calidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   10
      Left            =   5820
      TabIndex        =   27
      Top             =   840
      Width           =   1950
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   5700
      X2              =   5700
      Y1              =   -30
      Y2              =   5220
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Docum."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   9
      Left            =   45
      TabIndex        =   22
      Top             =   2430
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Log"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   8
      Left            =   60
      TabIndex        =   21
      Top             =   2820
      Width           =   345
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Documentos"
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
      Index           =   2
      Left            =   60
      TabIndex        =   20
      Top             =   1260
      Width           =   5460
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fondo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   7
      Left            =   60
      TabIndex        =   19
      Top             =   4110
      Width           =   585
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Plantillas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   5
      Left            =   45
      TabIndex        =   16
      Top             =   2040
      Width           =   810
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   4
      Left            =   60
      TabIndex        =   15
      Top             =   4860
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Logo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   6
      Left            =   60
      TabIndex        =   14
      Top             =   3720
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Otras configuraciones"
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
      Index           =   1
      Left            =   60
      TabIndex        =   13
      Top             =   3270
      Width           =   5460
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   3
      Left            =   60
      TabIndex        =   12
      Top             =   1650
      Width           =   765
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usuario defecto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   4500
      Width           =   1425
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ruta Programa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   10
      Top             =   900
      Width           =   1365
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Configuración"
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
      Height          =   345
      Index           =   0
      Left            =   60
      TabIndex        =   9
      Top             =   60
      Width           =   5460
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "IP Base datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   8
      Top             =   540
      Width           =   1260
   End
End
Attribute VB_Name = "frmConfiguracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    ' Configuración
    WriteINI App.path & "\config.ini", "server", "ip", txtDatos(0)
    WriteINI App.path & "\config.ini", "server", "ruta", txtDatos(1)
    ' Documentos
    WriteINI App.path & "\config.ini", "documentos", "Informes", txtDatos(5)
    WriteINI App.path & "\config.ini", "documentos", "plantillas", txtDatos(6)
    WriteINI App.path & "\config.ini", "documentos", "documentos", txtDatos(9)
    WriteINI App.path & "\config.ini", "documentos", "log", txtDatos(8)
    ' Otros
    WriteINI App.path & "\config.ini", "logo", "logo", txtDatos(2)
    WriteINI App.path & "\config.ini", "logo", "fondo", txtDatos(7)
    WriteINI App.path & "\config.ini", "usuario", "usuario", txtDatos(3)
    WriteINI App.path & "\config.ini", "usuario", "pass", txtDatos(4)
    'Parametros
    WriteINI App.path & "\config.ini", "parametros", "iva", txtDatos(11)
    WriteINI App.path & "\config.ini", "parametros", "Retencion_calidad", txtDatos(10)
    WriteINI App.path & "\config.ini", "parametros", "Retencion_garantia", txtDatos(13)
    WriteINI App.path & "\config.ini", "parametros", "Copias", txtDatos(12)
    WriteINI App.path & "\config.ini", "parametros", "Certificacion", txtDatos(14)
    If chkimprimir.Value = Checked Then
        WriteINI App.path & "\config.ini", "parametros", "imprimir", "1"
    Else
        WriteINI App.path & "\config.ini", "parametros", "imprimir", "0"
    End If
    
    WriteINI App.path & "\config.ini", "parametros", "Horas_trabajadores", txtDatos(5)
    WriteINI App.path & "\config.ini", "parametros", "Precio_kilometro", txtDatos(6)
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    cd.DialogTitle = "Abrir fichero de fondo"
    cd.InitDir = App.path & "\recursos\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        txtDatos(7).Text = cd.FileName  ' cd.FileTitle
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 27 ' ESC
        cmdSalir_Click
    Case 121 ' F10
        cmdAceptar_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    ' Configuración
    txtDatos(0) = ReadINI(App.path & "\config.ini", "server", "ip")
    txtDatos(1) = ReadINI(App.path & "\config.ini", "server", "ruta")
    ' Documentos
    txtDatos(5) = ReadINI(App.path & "\config.ini", "documentos", "Informes")
    txtDatos(6) = ReadINI(App.path & "\config.ini", "documentos", "plantillas")
    txtDatos(9) = ReadINI(App.path & "\config.ini", "documentos", "documentos")
    txtDatos(8) = ReadINI(App.path & "\config.ini", "documentos", "log")
    ' Otros
    txtDatos(2) = ReadINI(App.path & "\config.ini", "logo", "logo")
    txtDatos(7) = ReadINI(App.path & "\config.ini", "logo", "fondo")
    txtDatos(3) = ReadINI(App.path & "\config.ini", "usuario", "usuario")
    txtDatos(4) = ReadINI(App.path & "\config.ini", "usuario", "pass")
    'Parametros
    txtDatos(11) = ReadINI(App.path & "\config.ini", "parametros", "iva")
    txtDatos(10) = ReadINI(App.path & "\config.ini", "parametros", "Retencion_calidad")
    txtDatos(13) = ReadINI(App.path & "\config.ini", "parametros", "Retencion_garantia")
    txtDatos(12) = ReadINI(App.path & "\config.ini", "parametros", "Copias")
    txtDatos(14) = ReadINI(App.path & "\config.ini", "parametros", "Certificacion")
    If ReadINI(App.path & "\config.ini", "parametros", "imprimir") = 0 Then
        chkimprimir.Value = Unchecked
    Else
        chkimprimir.Value = Checked
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cmdEXplorar_Click()
    cd.DialogTitle = "Abrir fichero de logo"
    cd.InitDir = App.path & "\recursos\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        txtDatos(2).Text = cd.FileName  ' cd.FileTitle
    End If
End Sub
