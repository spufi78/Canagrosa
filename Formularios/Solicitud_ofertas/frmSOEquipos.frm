VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#35.0#0"; "miCombo.ocx"
Object = "{EDA716CF-5FC0-409B-8637-01EBB93C0182}#3.0#0"; "AjaxText.ocx"
Begin VB.Form frmSOEquipos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud Oferta Equipo"
   ClientHeight    =   6930
   ClientLeft      =   3210
   ClientTop       =   1845
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   12555
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   2310
      MaxLength       =   250
      TabIndex        =   42
      Top             =   2640
      Width           =   4440
   End
   Begin VB.CommandButton cmdCrearPediido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear Pedido"
      Height          =   870
      Index           =   0
      Left            =   1320
      Picture         =   "frmSOEquipos.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminarAccesorio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   570
      Left            =   11610
      Picture         =   "frmSOEquipos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Eliminar accesorio"
      Top             =   990
      Width           =   870
   End
   Begin VB.CommandButton cmdAnadirAccesorio 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir "
      Height          =   570
      Left            =   10680
      Picture         =   "frmSOEquipos.frx":0A5E
      Style           =   1  'Graphical
      TabIndex        =   34
      ToolTipText     =   "Añadir accesorio"
      Top             =   990
      Width           =   870
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Solicitar Oferta"
      Height          =   870
      Index           =   4
      Left            =   60
      Picture         =   "frmSOEquipos.frx":0C83
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   6000
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11460
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6000
      Width           =   1050
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3855
      Left            =   6960
      TabIndex        =   28
      Top             =   1950
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   6800
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   2310
      MaxLength       =   250
      TabIndex        =   25
      Top             =   1440
      Width           =   4440
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   39
      Left            =   2310
      MaxLength       =   100
      TabIndex        =   15
      Top             =   3780
      Width           =   690
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   40
      Left            =   3120
      MaxLength       =   100
      TabIndex        =   14
      Top             =   3780
      Width           =   690
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   41
      Left            =   2310
      MaxLength       =   100
      TabIndex        =   13
      Top             =   4110
      Width           =   1860
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   42
      Left            =   2310
      MaxLength       =   100
      TabIndex        =   12
      Top             =   4440
      Width           =   1860
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   43
      Left            =   2310
      MaxLength       =   100
      TabIndex        =   11
      Top             =   4770
      Width           =   1860
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   2310
      MaxLength       =   250
      TabIndex        =   3
      Top             =   1110
      Width           =   4440
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   465
      Index           =   33
      Left            =   2310
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   2130
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
      Left            =   2310
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   0
      Top             =   690
      Width           =   1110
   End
   Begin MSComCtl2.DTPicker frecepcion 
      Height          =   315
      Left            =   5340
      TabIndex        =   9
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
      Format          =   75431937
      CurrentDate     =   2
      MinDate         =   2
   End
   Begin pryCombo.miCombo miCombo3 
      Height          =   330
      Left            =   4620
      TabIndex        =   16
      Top             =   3750
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   582
   End
   Begin prjAjaxText.AjaxText txtAccesorio 
      Height          =   285
      Left            =   6960
      TabIndex        =   33
      Top             =   1620
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   503
      Backcolor       =   0
      FixedListSize   =   0
      Object.Height          =   0
      Object.Width           =   0
      MaxRows         =   5
      Field4Show      =   "NOMBRE"
      FieldId         =   "ID_ACCESORIO"
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbResponsable 
      Height          =   315
      Left            =   2310
      TabIndex        =   36
      Top             =   5100
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbNorma 
      Height          =   315
      Left            =   2310
      TabIndex        =   37
      Top             =   5460
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo cmbProveedor 
      Height          =   315
      Left            =   2310
      TabIndex        =   39
      Top             =   2970
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
      Left            =   2310
      TabIndex        =   41
      Top             =   1770
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fabricante"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   43
      Top             =   2685
      Width           =   750
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Familia"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   40
      Top             =   1860
      Width           =   480
   End
   Begin VB.Image imgBuscarAccesorios 
      Height          =   285
      Left            =   12180
      Picture         =   "frmSOEquipos.frx":154D
      Stretch         =   -1  'True
      Top             =   1590
      Width           =   285
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   6870
      Picture         =   "frmSOEquipos.frx":1E17
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   285
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Accesorios"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8970
      TabIndex        =   29
      Top             =   1260
      Width           =   1335
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Conforme la Norma"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   27
      Top             =   5580
      Width           =   1350
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modelo"
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   26
      Top             =   1485
      Width           =   525
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Responsable Técnico"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   24
      Top             =   5190
      Width           =   1560
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "  Minimo   -   Máximo"
      Height          =   195
      Index           =   63
      Left            =   2310
      TabIndex        =   23
      Top             =   3570
      Width           =   1500
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      Height          =   195
      Index           =   64
      Left            =   3030
      TabIndex        =   22
      Top             =   3840
      Width           =   45
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "R. Medida"
      Height          =   195
      Index           =   65
      Left            =   120
      TabIndex        =   21
      Top             =   3825
      Width           =   735
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Unidades"
      Height          =   195
      Index           =   66
      Left            =   3840
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incertidumbre Máx. Admisible"
      Height          =   195
      Index           =   67
      Left            =   120
      TabIndex        =   19
      Top             =   4155
      Width           =   2055
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tolerancia Máx."
      Height          =   195
      Index           =   68
      Left            =   120
      TabIndex        =   18
      Top             =   4485
      Width           =   1140
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Precisión"
      Height          =   195
      Index           =   69
      Left            =   120
      TabIndex        =   17
      Top             =   4815
      Width           =   645
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Solicitud"
      Height          =   195
      Index           =   12
      Left            =   4080
      TabIndex        =   10
      Top             =   810
      Width           =   1095
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de la Solicitud de Oferta de Equipo"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   8
      Top             =   315
      Width           =   3000
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11970
      Picture         =   "frmSOEquipos.frx":26E1
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Solicitud Oferta Equipo"
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
      TabIndex        =   7
      Top             =   45
      Width           =   2400
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   1155
      Width           =   585
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción"
      Height          =   195
      Index           =   55
      Left            =   120
      TabIndex        =   5
      Top             =   2190
      Width           =   840
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Proveedor"
      Height          =   195
      Index           =   9
      Left            =   150
      TabIndex        =   4
      Top             =   3060
      Width           =   735
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
      TabIndex        =   1
      Top             =   780
      Width           =   1875
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
Attribute VB_Name = "frmSOEquipos"
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


