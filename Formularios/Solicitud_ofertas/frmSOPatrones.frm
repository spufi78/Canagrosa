VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmSOPatrones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud Oferta Material de Referencia / Patrones"
   ClientHeight    =   5610
   ClientLeft      =   1260
   ClientTop       =   2775
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   12555
   Begin VB.CommandButton cmdCrearPediido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear Pedido"
      Height          =   870
      Index           =   0
      Left            =   1320
      Picture         =   "frmSOPatrones.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   2490
      MaxLength       =   250
      TabIndex        =   27
      Top             =   4230
      Width           =   4440
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   2490
      MaxLength       =   100
      TabIndex        =   20
      Top             =   3540
      Width           =   1530
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   2490
      MaxLength       =   100
      TabIndex        =   18
      Top             =   3210
      Width           =   1530
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   2490
      MaxLength       =   100
      TabIndex        =   15
      Top             =   2880
      Width           =   1530
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   2490
      MaxLength       =   100
      TabIndex        =   13
      Top             =   2550
      Width           =   1530
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   2490
      MaxLength       =   250
      TabIndex        =   11
      Top             =   1830
      Width           =   4440
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   3165
      Index           =   33
      Left            =   7590
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1440
      Width           =   4950
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Solicitar Oferta"
      Height          =   870
      Index           =   4
      Left            =   60
      Picture         =   "frmSOPatrones.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10410
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4680
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11490
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4680
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
      TabIndex        =   24
      Top             =   1440
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo2 
      Height          =   315
      Left            =   2490
      TabIndex        =   26
      Top             =   2160
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo3 
      Height          =   315
      Left            =   2490
      TabIndex        =   28
      Top             =   3870
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
      TabIndex        =   30
      Top             =   750
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
      Format          =   70320129
      CurrentDate     =   2
      MinDate         =   2
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Solicitud"
      Height          =   195
      Index           =   12
      Left            =   4350
      TabIndex        =   31
      Top             =   840
      Width           =   1095
   End
   Begin VB.Image Image2 
      Height          =   285
      Left            =   6990
      Picture         =   "frmSOPatrones.frx":1194
      Stretch         =   -1  'True
      Top             =   4230
      Width           =   285
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
      Index           =   2
      Left            =   120
      TabIndex        =   25
      Top             =   780
      Width           =   1875
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   55
      Left            =   7590
      TabIndex        =   23
      Top             =   1170
      Width           =   1065
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Uso al que está destinado"
      Height          =   195
      Index           =   10
      Left            =   150
      TabIndex        =   22
      Top             =   2250
      Width           =   1845
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Cantidad"
      Height          =   195
      Index           =   9
      Left            =   120
      TabIndex        =   21
      Top             =   3570
      Width           =   630
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "% Vida Útil"
      Height          =   195
      Index           =   8
      Left            =   120
      TabIndex        =   19
      Top             =   3240
      Width           =   750
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Equipo Destinatario"
      Height          =   195
      Index           =   7
      Left            =   180
      TabIndex        =   17
      Top             =   4320
      Width           =   1380
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tolerancia"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   2910
      Width           =   750
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incertidumbre"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   14
      Top             =   2580
      Width           =   960
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Código Producto"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   12
      Top             =   1875
      Width           =   1185
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   6990
      Picture         =   "frmSOPatrones.frx":1A5E
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
      Left            =   150
      TabIndex        =   5
      Top             =   3960
      Width           =   1560
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de la Solicitud de Oferta de Material de Referencia / Patrones"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   315
      Width           =   4905
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11970
      Picture         =   "frmSOPatrones.frx":2328
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Solicitud Oferta Material de Referencia / Patrones"
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
      Width           =   5160
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
Attribute VB_Name = "frmSOPatrones"
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

Private Sub cmdcancel_Click()
Unload Me
End Sub


