VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#35.0#0"; "miCombo.ocx"
Begin VB.Form frmSOCalibracion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Solicitud Oferta Calibraci�n"
   ClientHeight    =   4365
   ClientLeft      =   1845
   ClientTop       =   7095
   ClientWidth     =   12555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4365
   ScaleWidth      =   12555
   Begin VB.CommandButton cmdCrearPediido 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear Pedido"
      Height          =   870
      Index           =   0
      Left            =   1320
      Picture         =   "frmSOCalibracion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3450
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   2265
      Index           =   33
      Left            =   7560
      MaxLength       =   250
      MultiLine       =   -1  'True
      TabIndex        =   19
      Top             =   1110
      Width           =   4920
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   3300
      MaxLength       =   100
      TabIndex        =   16
      Top             =   2040
      Width           =   690
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   2490
      MaxLength       =   100
      TabIndex        =   15
      Top             =   2040
      Width           =   690
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Solicitar Oferta"
      Height          =   870
      Index           =   4
      Left            =   60
      Picture         =   "frmSOCalibracion.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3450
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10380
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3450
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11460
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3450
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   39
      Left            =   2490
      MaxLength       =   100
      TabIndex        =   7
      Top             =   2370
      Width           =   690
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   40
      Left            =   3300
      MaxLength       =   100
      TabIndex        =   6
      Top             =   2370
      Width           =   690
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
   Begin MSComCtl2.DTPicker frecepcion 
      Height          =   315
      Left            =   5520
      TabIndex        =   4
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
      Format          =   69795841
      CurrentDate     =   2
      MinDate         =   2
   End
   Begin pryCombo.miCombo miCombo3 
      Height          =   330
      Left            =   5040
      TabIndex        =   23
      Top             =   2040
      Width           =   2655
      _ExtentX        =   4683
      _ExtentY        =   582
   End
   Begin MSDataListLib.DataCombo cmbResponsable 
      Height          =   315
      Left            =   2490
      TabIndex        =   24
      Top             =   2700
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
      Left            =   2490
      TabIndex        =   25
      Top             =   3060
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin MSDataListLib.DataCombo DataCombo1 
      Height          =   315
      Left            =   2490
      TabIndex        =   29
      Top             =   1470
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
      TabIndex        =   30
      Top             =   1110
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Proveedor"
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   28
      Top             =   1470
      Width           =   735
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Unidades"
      Height          =   195
      Index           =   66
      Left            =   4140
      TabIndex        =   27
      Top             =   2130
      Width           =   735
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Conforme la Norma"
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   26
      Top             =   3180
      Width           =   1350
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "N� de Solicitud Oferta"
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
      Index           =   4
      Left            =   120
      TabIndex        =   21
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
      TabIndex        =   20
      Top             =   870
      Width           =   1065
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "R. Medida"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   18
      Top             =   2085
      Width           =   735
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      Height          =   195
      Index           =   1
      Left            =   3210
      TabIndex        =   17
      Top             =   2070
      Width           =   45
   End
   Begin VB.Image Image1 
      Height          =   285
      Left            =   6990
      Picture         =   "frmSOCalibracion.frx":1194
      Stretch         =   -1  'True
      Top             =   1110
      Width           =   285
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Responsable T�cnico"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   11
      Top             =   2790
      Width           =   1560
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "  Minimo   -   M�ximo"
      Height          =   195
      Index           =   63
      Left            =   2520
      TabIndex        =   10
      Top             =   1830
      Width           =   1470
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "-"
      Height          =   195
      Index           =   64
      Left            =   3210
      TabIndex        =   9
      Top             =   2400
      Width           =   45
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tolerancia"
      Height          =   195
      Index           =   65
      Left            =   120
      TabIndex        =   8
      Top             =   2415
      Width           =   750
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Solicitud"
      Height          =   195
      Index           =   12
      Left            =   4320
      TabIndex        =   5
      Top             =   810
      Width           =   1095
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de la Solicitud de Oferta de Calibraci�n"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   315
      Width           =   3285
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11970
      Picture         =   "frmSOCalibracion.frx":1A5E
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Solicitud Oferta Calibraci�n"
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
      TabIndex        =   2
      Top             =   45
      Width           =   2850
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre Equipo"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   1140
      Width           =   1095
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
Attribute VB_Name = "frmSOCalibracion"
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


