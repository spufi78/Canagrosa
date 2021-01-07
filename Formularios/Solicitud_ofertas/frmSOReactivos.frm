VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSOReactivos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud Oferta Reactivos"
   ClientHeight    =   4005
   ClientLeft      =   1665
   ClientTop       =   2820
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   12555
   Begin VB.CommandButton cmdCrearPediido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear Pedido"
      Height          =   870
      Index           =   0
      Left            =   1320
      Picture         =   "frmSOReactivos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3090
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   3
      Left            =   2490
      MaxLength       =   250
      TabIndex        =   14
      Top             =   1800
      Width           =   4440
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1905
      Index           =   33
      Left            =   7560
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   1110
      Width           =   4920
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   2490
      MaxLength       =   100
      TabIndex        =   10
      Top             =   2520
      Width           =   1530
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Solicitar Oferta"
      Height          =   870
      Index           =   4
      Left            =   60
      Picture         =   "frmSOReactivos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3090
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3090
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11460
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3090
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   2490
      MaxLength       =   250
      TabIndex        =   1
      Top             =   1110
      Width           =   4440
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
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
      Index           =   0
      Left            =   2490
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   0
      Top             =   690
      Width           =   1110
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   2490
      TabIndex        =   17
      Top             =   1440
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbResponsable 
      Height          =   315
      Left            =   2490
      TabIndex        =   19
      Top             =   2160
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSComCtl2.DTPicker frecepcion 
      Height          =   315
      Left            =   5520
      TabIndex        =   20
      Top             =   720
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
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
      CalendarTitleBackColor=   14737632
      Format          =   75563009
      CurrentDate     =   2
      MinDate         =   2
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Solicitud"
      Height          =   195
      Index           =   12
      Left            =   4320
      TabIndex        =   21
      Top             =   810
      Width           =   1095
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nº de Solicitud Oferta"
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
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   810
      Width           =   1875
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código Producto"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   15
      Top             =   1875
      Width           =   1185
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   55
      Left            =   7590
      TabIndex        =   13
      Top             =   870
      Width           =   1065
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   11
      Top             =   2580
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   6990
      Picture         =   "frmSOReactivos.frx":1194
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   285
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Proveedor"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   6
      Top             =   1530
      Width           =   735
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Responsable Técnico"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   2250
      Width           =   1560
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de la Solicitud de Oferta de Reactivos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   315
      Width           =   3225
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11970
      Picture         =   "frmSOReactivos.frx":1A5E
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Solicitud Oferta Reactivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   2715
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre Comercial"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1155
      Width           =   1290
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   12555
   End
End
Attribute VB_Name = "frmSOReactivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarobjTipoEdicion As enumTipoEdicion
Private mvarlngidSolicitud As Long
Private mvarobjSolicitud As New clsSolicitud_ofertas

Public Property Get idSolicitud() As Long
    idSolicitud = mvarlngidSolicitud
End Property
Public Property Let idSolicitud(ByVal lngidSolicitud As Long)
    mvarlngidSolicitud = lngidSolicitud
End Property

Public Property Get TipoEdicion() As enumTipoEdicion
    TipoEdicion = mvarobjTipoEdicion
End Property
Public Property Let TipoEdicion(objTipoEdicion As enumTipoEdicion)
    mvarobjTipoEdicion = objTipoEdicion
End Property

Private Sub Label4_Click()

End Sub

Private Sub cmdcancel_Click()
Unload Me
End Sub


