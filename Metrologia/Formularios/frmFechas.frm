VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFechas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Seleccione rango de fechas para el listado"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1725
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   780
      Width           =   1155
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   780
      Width           =   1155
   End
   Begin MSComCtl2.DTPicker fdesde 
      Height          =   330
      Left            =   810
      TabIndex        =   0
      Top             =   120
      Width           =   1350
      _ExtentX        =   2381
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   12632256
      Format          =   16842753
      CurrentDate     =   38002
   End
   Begin MSComCtl2.DTPicker fhasta 
      Height          =   330
      Left            =   3150
      TabIndex        =   1
      Top             =   120
      Width           =   1305
      _ExtentX        =   2302
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   12632256
      Format          =   16842753
      CurrentDate     =   38002
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Desde"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   210
      Width           =   465
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hasta"
      Height          =   195
      Index           =   1
      Left            =   2460
      TabIndex        =   2
      Top             =   210
      Width           =   420
   End
End
Attribute VB_Name = "frmFechas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAceptar_Click()
    gFECHA_DESDE = fdesde
    gFECHA_HASTA = fhasta
    Unload Me
End Sub

Private Sub cmdcancel_Click()
    gFECHA_DESDE = ""
    gFECHA_HASTA = ""
    Unload Me
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    fdesde = Date - 90
    fhasta = Date
End Sub
