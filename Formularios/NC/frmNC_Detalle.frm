VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form frmNC_Detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "DETALLE DE INCIDENCIA"
   ClientHeight    =   9645
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10485
   Icon            =   "frmNC_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   10485
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtasignanumero 
      Height          =   330
      Left            =   4995
      TabIndex        =   43
      Top             =   8775
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.CommandButton cmddocumentacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Asociar Documentación"
      Height          =   870
      Left            =   2070
      Picture         =   "frmNC_Detalle.frx":2AFA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8730
      Width           =   1905
   End
   Begin VB.CommandButton cmdAccion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Acciones Correctoras"
      Height          =   870
      Left            =   90
      Picture         =   "frmNC_Detalle.frx":33C4
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8730
      Width           =   1905
   End
   Begin VB.Frame frmevaluacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Evaluación e impacto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   45
      TabIndex        =   25
      Top             =   6750
      Width           =   10350
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1005
         Index           =   1
         Left            =   5580
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   810
         Width           =   4680
      End
      Begin VB.CheckBox chkimpacto 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Impacto"
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
         Height          =   285
         Left            =   5580
         TabIndex        =   7
         Top             =   225
         Width           =   1725
      End
      Begin VB.CheckBox cnknoprocede 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "No Procedente"
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
         Height          =   330
         Left            =   1710
         TabIndex        =   6
         Top             =   225
         Width           =   2040
      End
      Begin VB.CheckBox chkevaluada 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Evaluada"
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
         Height          =   330
         Left            =   135
         TabIndex        =   5
         Top             =   225
         Width           =   1410
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   1005
         Index           =   5
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   810
         Width           =   5160
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   5445
         X2              =   5445
         Y1              =   90
         Y2              =   1890
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción del Impacto"
         Height          =   195
         Index           =   9
         Left            =   5580
         TabIndex        =   30
         Top             =   585
         Width           =   1710
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción de la Evaluación"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   29
         Top             =   585
         Width           =   2070
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Número y fechas de la no conformidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1095
      Index           =   0
      Left            =   0
      TabIndex        =   20
      Top             =   720
      Width           =   10410
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   4
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   32
         Top             =   630
         Width           =   1515
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   2
         Left            =   4320
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   12
         Top             =   225
         Width           =   1515
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1305
         TabIndex        =   10
         Top             =   270
         Width           =   1335
         _ExtentX        =   2355
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
         CalendarTitleBackColor=   14737632
         Format          =   52887553
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha_cierre 
         Height          =   330
         Left            =   1305
         TabIndex        =   11
         Top             =   675
         Width           =   1335
         _ExtentX        =   2355
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
         CalendarTitleBackColor=   14737632
         Format          =   52887553
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   480
         Left            =   7065
         TabIndex        =   13
         Top             =   360
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   847
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número Particular"
         Height          =   195
         Index           =   11
         Left            =   2790
         TabIndex        =   33
         Top             =   720
         Width           =   1260
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   5
         Left            =   6435
         TabIndex        =   24
         Top             =   540
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número General"
         Height          =   195
         Index           =   7
         Left            =   2790
         TabIndex        =   23
         Top             =   315
         Width           =   1155
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Alta"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   22
         Top             =   360
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Cierre"
         Height          =   195
         Index           =   10
         Left            =   225
         TabIndex        =   21
         Top             =   720
         Width           =   900
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8730
      Width           =   1050
   End
   Begin VB.Frame frmanalisis 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Análisis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   45
      TabIndex        =   31
      Top             =   4725
      Width           =   10350
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   870
         Index           =   3
         Left            =   1125
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   945
         Width           =   9000
      End
      Begin MSDataListLib.DataCombo cmbtipo 
         Height          =   315
         Left            =   1125
         TabIndex        =   35
         Top             =   225
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbDepartamento 
         Height          =   315
         Left            =   6390
         TabIndex        =   36
         Top             =   225
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbAfectado 
         Height          =   315
         Left            =   1125
         TabIndex        =   39
         Top             =   585
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "An. Causas"
         Height          =   195
         Left            =   180
         TabIndex        =   41
         Top             =   1215
         Width           =   1860
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Afectado"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   40
         Top             =   675
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   38
         Top             =   315
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Departamento"
         Height          =   195
         Index           =   1
         Left            =   5265
         TabIndex        =   37
         Top             =   270
         Width           =   1005
      End
   End
   Begin VB.Frame Frame 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción de la Incidencia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2805
      Index           =   1
      Left            =   45
      TabIndex        =   17
      Top             =   1845
      Width           =   10350
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   870
         Index           =   10
         Left            =   1215
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   1485
         Width           =   8910
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   825
         Index           =   0
         Left            =   1215
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   630
         Width           =   8910
      End
      Begin MSDataListLib.DataCombo cmbUsuario 
         Height          =   315
         Left            =   1215
         TabIndex        =   0
         Top             =   270
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmborigen 
         Height          =   315
         Left            =   1215
         TabIndex        =   3
         Top             =   2385
         Width           =   8910
         _ExtentX        =   15716
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Origen"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   42
         Top             =   2430
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Acc.inmediata"
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   34
         Top             =   1710
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   28
         Top             =   315
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   19
         Top             =   900
         Width           =   840
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Describa los datos de la incidencia, rellenando los siguientes campos."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   27
      Top             =   360
      Width           =   4920
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9855
      Picture         =   "frmNC_Detalle.frx":3C8E
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Incidencia"
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
      Index           =   0
      Left            =   90
      TabIndex        =   26
      Top             =   90
      Width           =   2220
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   -45
      Width           =   10485
   End
End
Attribute VB_Name = "frmNC_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmbTipo_Change()
    If cmbTipo.Text <> "" And Trim(txtasignanumero) = "" Then
        Dim oNC As New clsNc
        txtDatos(4) = oNC.Calcular_Numero(CLng(cmbTipo.BoundText))
    End If
End Sub

Private Sub cmdAccion_Click()
    frmNC_Acciones.PK = PK
    frmNC_Acciones.Show 1
End Sub

Private Sub chkevaluada_Click()
    On Error Resume Next
    If chkevaluada.value = Checked Then
        txtDatos(5).Enabled = True
        txtDatos(5).BackColor = vbWhite
        txtDatos(5).SetFocus
    Else
        txtDatos(5).Enabled = False
        txtDatos(5) = ""
        txtDatos(5).BackColor = &HE0E0E0
    End If
End Sub

Private Sub chkimpacto_Click()
'    txtDatos(1).Enabled = chkimpacto.value
    On Error Resume Next
    If chkimpacto.value = Checked Then
        txtDatos(1).Enabled = True
        txtDatos(1).BackColor = vbWhite
        txtDatos(1).SetFocus
    Else
        txtDatos(1).Enabled = False
        txtDatos(1) = ""
        txtDatos(1).BackColor = &HE0E0E0
    End If
End Sub

Private Sub cmbestados_Change()
    If cmbestados.BoundText = C_NC_ESTADOS.cerrada Then
        fecha_cierre.Enabled = True
        fecha_cierre = Date
    Else
        fecha_cierre.Enabled = False
    End If
End Sub
Private Sub cmdDocumentacion_Click()
    'M0499-I
'    frmNC_Adjuntos.PK = PK
'    frmNC_Adjuntos.Show 1
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_NC
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
    'M0499-F
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error
    If validar = True Then
      Dim documento As Long
      Dim oNC As New clsNc
      With oNC
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setFECHA_CIERRE = Format(fecha_cierre, "yyyy-mm-dd")
            .setID_NC = txtDatos(2)
            If txtDatos(4) = "" Then
                .setNUMERO = 0
            Else
                .setNUMERO = txtDatos(4)
            End If
            .setESTADO_ID = cmbestados.BoundText
            If cmbTipo.Text = "" Then
                .setTIPO_HECHO_ID = 0
            Else
                .setTIPO_HECHO_ID = cmbTipo.BoundText
            End If
            .setORIGEN_ID = cmborigen.BoundText
            If cmbafectado.Text = "" Then
                .setAFECTADO_ID = 0
            Else
                .setAFECTADO_ID = cmbafectado.BoundText
            End If
            .setUSUARIO_ID = cmbUsuario.BoundText
            If cmbDepartamento.Text = "" Then
                .setDEPARTAMENTO_ID = 0
            Else
                .setDEPARTAMENTO_ID = cmbDepartamento.BoundText
            End If
            .setDESCRIPCION = txtDatos(0)
            .setACCION_INMEDIATA = txtDatos(10)
            .setANALISIS_CAUSAS = txtDatos(3)
            .setEVALUADA = chkevaluada.value
            .setNO_PROCEDENTE = cnknoprocede.value
            .setIMPACTO = chkimpacto.value
            .setEVALUACION = txtDatos(5)
            .setIMPACTO_TEXTO = txtDatos(1)
      End With
      If PK = 0 Then
        If MsgBox("Va a introducir una incidencia. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            documento = oNC.Insertar
            enviar_mensaje (1)
        Else
            Exit Sub
        End If
      Else
        If cmbestados.BoundText = C_NC_ESTADOS.cerrada Then
            If MsgBox("Va a cerrar la Incidencia. Una vez cerrada no será posible modificarla. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            Else
                oNC.Modificar (PK)
                enviar_mensaje (2)
            End If
        Else
            If MsgBox("Va a modificar la incidencia. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                oNC.Modificar (PK)
                enviar_mensaje (2)
            Else
                Exit Sub
            End If
        End If
      End If
      If PK = 0 Then
          If MsgBox("La incidencia se ha introducido correctamente. ¿Desea asignar documentación?", vbInformation + vbYesNo, App.Title) = vbYes Then
                PK = documento
                cmdDocumentacion_Click
          End If
      Else
          MsgBox "La incidencia se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmNC_Detalle"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    fecha = Date
    fecha_cierre = Date
    fecha_cierre.Enabled = False
    If PK <> 0 Then
        lbltitulo(0) = "Modificación de Incidencia"
        cargar_documento
    Else
        Dim oNC As New clsNc
        oNC.CrearID
        txtDatos(2) = oNC.getID_NC
'        txtDatos(2) = oNC.Calcular_Numero(fecha)
        cmbUsuario.BoundText = usuario.getID_EMPLEADO
        ' Al crear incidencia, debe ser PDTE. REVISION
        cmbestados.BoundText = C_NC_ESTADOS.PDTE_REVISION
        cmdDocumentacion.Enabled = False
    End If
    ' Deshabiliar controles
    If PK = 0 Or usuario.getPER_NC = False Then
        cmbestados.Locked = True
        frmanalisis.Enabled = False
        frmevaluacion.Enabled = False
        cmdAccion.Enabled = False
    End If
    ' Si el tipo esta informado y el numero, no deja modificar
    
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargar_documento()
    Dim oNC As New clsNc
    With oNC
        If .Carga(PK) = True Then
            fecha = .getFECHA
            txtDatos(2) = .getID_NC
            If .getNUMERO <> 0 Then
                txtDatos(4) = .getNUMERO
                txtasignanumero = .getNUMERO
            End If
            cmbestados.BoundText = .getESTADO_ID
            fecha_cierre = .getFECHA_CIERRE
            cmbTipo.BoundText = .getTIPO_HECHO_ID
            cmborigen.BoundText = .getORIGEN_ID
            
            cmbDepartamento.BoundText = .getDEPARTAMENTO_ID
            cmbUsuario.BoundText = .getUSUARIO_ID
            cmbafectado.BoundText = .getAFECTADO_ID
            
            txtDatos(0) = .getDESCRIPCION
            txtDatos(10) = .getACCION_INMEDIATA
            txtDatos(3) = .getANALISIS_CAUSAS
            chkevaluada.value = .getEVALUADA
            cnknoprocede.value = .getNO_PROCEDENTE
            txtDatos(5) = .getEVALUACION
            chkimpacto.value = .getIMPACTO
            txtDatos(1) = .getIMPACTO_TEXTO
            If .getTIPO_HECHO_ID <> 0 Then
                cmbTipo.Enabled = False
            End If
        End If
    End With
    If cmbestados.BoundText = C_NC_ESTADOS.cerrada Then
        Frame(0).Enabled = False
'        Frame(1).Enabled = False
        cmbUsuario.Locked = True
        txtDatos(0).Locked = True
        txtDatos(10).Locked = True
        cmborigen.Locked = True
'        frmanalisis.Enabled = False
        cmbTipo.Locked = True
        cmbDepartamento.Locked = True
        cmbafectado.Locked = True
        txtDatos(3).Locked = True
'        frmevaluacion.Enabled = False
        chkevaluada.Enabled = False
        cnknoprocede.Enabled = False
        chkimpacto.Enabled = False
        txtDatos(5).Locked = True
        txtDatos(1).Locked = True
        cmdok.Enabled = False
    End If
    
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle una descripción.", vbExclamation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(10)) = "" Then
        MsgBox "Debe rellenar las acciones inmediatas.", vbExclamation, App.Title
        txtDatos(10).SetFocus
        validar = False
        Exit Function
    End If
'    If cmbtipo.BoundText = "" Then
'        MsgBox "Debe asignar un tipo.", vbExclamation, App.Title
'        cmbtipo.SetFocus
'        validar = False
'        Exit Function
'    End If
    If cmborigen.BoundText = "" Then
        MsgBox "Debe asignar un origen.", vbExclamation, App.Title
        cmborigen.SetFocus
        validar = False
        Exit Function
    End If
    If cmbestados.BoundText = "" Then
        MsgBox "Debe asignar un estado.", vbExclamation, App.Title
        cmbEstado.SetFocus
        validar = False
        Exit Function
    End If
'    If cmbDepartamento.BoundText = "" Then
'        MsgBox "Debe asignar un departamento.", vbExclamation, App.Title
'        cmbDepartamento.SetFocus
'        validar = False
'        Exit Function
'    End If
End Function

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, decodificadora.NC_TIPOS_HECHOS
    oDeco.cargar_combo cmborigen, decodificadora.NC_ORIGENES
    oDeco.cargar_combo cmbestados, decodificadora.NC_ESTADOS
    oDeco.cargar_combo cmbDepartamento, decodificadora.NC_DEPARTAMENTOS
    oDeco.cargar_combo cmbafectado, decodificadora.NC_AFECTADO
    cargar_combo cmbUsuario, New clsUsuarios
End Sub
Private Sub enviar_mensaje(tipo As Integer)
    ' Enviar aviso
    Dim oMensaje As New clsMensajes
    Dim asunto As String
    Dim texto As String
    Dim mens As Integer
    With oMensaje
        If tipo = 1 Then
            asunto = "Alta incidencia, nº: " & txtDatos(2)
            texto = texto & "El usuario " & cmbUsuario.Text & " ha dado de alta una incidencia. " & vbNewLine & vbNewLine
        Else
            asunto = "Modificación incidencia, nº: " & txtDatos(2)
            texto = texto & "El usuario " & cmbUsuario.Text & " ha modificado una incidencia. " & vbNewLine & vbNewLine
        End If
        texto = texto & "Fecha de Alta : " & Format(fecha, "dd-mm-yyyy") & vbNewLine & vbNewLine
        texto = texto & "Descripción : " & txtDatos(0) & vbNewLine & vbNewLine
        texto = texto & "Acc.Inmediata : " & txtDatos(10) & vbNewLine & vbNewLine
        texto = texto & "Origen : " & cmborigen.Text & vbNewLine
        
        .setASUNTO = asunto
        .setTEXTO = texto
        .setEMPLEADO_ID = usuario.getID_EMPLEADO
        .setFECHA_INICIO = Format(fecha.value, "yyyy-mm-dd")
        .setFECHA_FIN = Format(fecha.value + 7, "yyyy-mm-dd")
        .setACCION = "frmNC_Detalle;" & Trim(txtDatos(2))
        
        .setHORA_INICIO = Format(Time, "hh:mm:ss")
        .setHORA_FIN = Format(Time, "hh:mm:ss")
        .setCATEGORIA = MENSAJES_CATEGORIAS.MENSAJES_CATEGORIAS_PROCN
        .setDURACION = 0
        
        mens = .Insertar
        If mens > 0 Then
            Dim omu As New clsMensajes_usuarios
            Dim i As Integer
            Dim usuarios As New clsUsuarios
            Dim rs As ADODB.RecordSet
            Set rs = usuarios.Listado
            If rs.RecordCount > 0 Then
                Do
                    If rs("PER_NC") = 1 And rs("ID_EMPLEADO") <> usuario.getID_EMPLEADO Then
                        omu.setEMPLEADO_ID = rs("ID_EMPLEADO")
                        omu.setMENSAJE_ID = mens
                        omu.Insertar
                    End If
                    rs.MoveNext
                Loop Until rs.EOF
                frmCalendario.cargar_eventos
            End If
        End If
    End With
End Sub
