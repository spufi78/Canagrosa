VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmTA_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Análisis"
   ClientHeight    =   10620
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11520
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTA_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10620
   ScaleWidth      =   11520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmInforme 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción en el informe de ensayo"
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
      Left            =   45
      TabIndex        =   42
      Top             =   4950
      Width           =   11445
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   6
         Left            =   810
         TabIndex        =   10
         Top             =   270
         Width           =   10560
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   810
         TabIndex        =   11
         Top             =   675
         Width           =   10560
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Español"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   44
         Top             =   315
         Width           =   570
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ingles"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   43
         Top             =   705
         Width           =   420
      End
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9585
      Width           =   1365
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Archivos Adjuntos"
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
      Height          =   1200
      Left            =   45
      TabIndex        =   33
      Top             =   6075
      Width           =   11400
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   840
         Left            =   10440
         Picture         =   "frmTA_Detalle.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   270
         Width           =   825
      End
      Begin VB.CommandButton cmdDocumentacion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   840
         Left            =   9540
         Picture         =   "frmTA_Detalle.frx":08D6
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   270
         Width           =   825
      End
      Begin MSComctlLib.ListView listaAdjuntos 
         Height          =   840
         Left            =   120
         TabIndex        =   34
         Top             =   270
         Width           =   9315
         _ExtentX        =   16431
         _ExtentY        =   1482
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14609914
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Multitarifa"
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
      Height          =   3270
      Left            =   45
      TabIndex        =   29
      Top             =   7275
      Width           =   5865
      Begin VB.CheckBox chkrevisarfactura 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Revisar factura"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   3735
         TabIndex        =   13
         Top             =   270
         Width           =   1905
      End
      Begin VB.CheckBox chkFD 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Factura por Determinaciones"
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
         Height          =   240
         Left            =   135
         TabIndex        =   12
         Top             =   270
         Width           =   2940
      End
      Begin VB.TextBox txttarifa 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   4365
         TabIndex        =   14
         Top             =   1845
         Width           =   1425
      End
      Begin MSComctlLib.ListView tarifas 
         Height          =   2550
         Left            =   135
         TabIndex        =   30
         Top             =   585
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   4498
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14609914
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   12
         Left            =   4365
         TabIndex        =   31
         Top             =   1545
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdDatosEspecificos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos Específicos"
      Height          =   870
      Left            =   9705
      Picture         =   "frmTA_Detalle.frx":11A0
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7335
      Width           =   1620
   End
   Begin VB.CommandButton cmdDeterminaciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Determinaciones"
      Height          =   870
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7335
      Width           =   1620
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10395
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9585
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   9315
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9585
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   45
      TabIndex        =   19
      Top             =   585
      Width           =   11430
      Begin VB.Frame frmTrigo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Trigo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4275
         TabIndex        =   39
         Top             =   3690
         Visible         =   0   'False
         Width           =   5145
         Begin VB.OptionButton opTipoTrigo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Blando"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   3105
            TabIndex        =   41
            Top             =   225
            Width           =   1125
         End
         Begin VB.OptionButton opTipoTrigo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Duro"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   0
            Left            =   1485
            TabIndex        =   40
            Top             =   225
            Width           =   1440
         End
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   5
         Left            =   2835
         TabIndex        =   9
         Top             =   3735
         Width           =   1245
      End
      Begin VB.TextBox txtnormativa 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   585
         Left            =   1650
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1440
         Width           =   9555
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   630
         Index           =   4
         Left            =   1650
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2700
         Width           =   9555
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   585
         Index           =   3
         Left            =   1650
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   2070
         Width           =   9555
      End
      Begin VB.OptionButton opDuplicado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   5160
         TabIndex        =   2
         Top             =   720
         Width           =   540
      End
      Begin VB.OptionButton opDuplicado 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   5820
         TabIndex        =   3
         Top             =   720
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1650
         TabIndex        =   0
         Top             =   270
         Width           =   9570
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   1
         Left            =   1650
         TabIndex        =   1
         Top             =   660
         Width           =   1785
      End
      Begin MSDataListLib.DataCombo cmbTM 
         Height          =   315
         Left            =   1650
         TabIndex        =   4
         Top             =   1065
         Width           =   9555
         _ExtentX        =   16854
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin pryCombo.miCombo cmbtarifa 
         Height          =   375
         Left            =   1650
         TabIndex        =   8
         Top             =   3375
         Width           =   7770
         _ExtentX        =   13705
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Duración en días del ensayo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   135
         TabIndex        =   38
         Top             =   3780
         Width           =   2595
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cod. Tarifa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   135
         TabIndex        =   32
         Top             =   3420
         Width           =   990
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Parámetro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   135
         TabIndex        =   26
         Top             =   2925
         Width           =   975
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Método"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   135
         TabIndex        =   25
         Top             =   2385
         Width           =   705
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Normativa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   24
         Top             =   1665
         Width           =   945
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Normalizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3750
         TabIndex        =   23
         Top             =   690
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Muestra"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   22
         Top             =   1095
         Width           =   1215
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   21
         Top             =   300
         Width           =   765
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   20
         Top             =   690
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8055
      Top             =   8550
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Tipos de Análisis"
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
      TabIndex        =   28
      Top             =   45
      Width           =   1830
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10935
      Picture         =   "frmTA_Detalle.frx":1A6A
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del tipo de análisis"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   27
      Top             =   315
      Width           =   1830
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   11475
   End
End
Attribute VB_Name = "frmTA_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Private tarifa_modificada As Boolean

Private Sub cmbTM_Change()
    frmTrigo.visible = False
    If cmbTM.Text <> "" Then
        If cmbTM.BoundText = TIPOS_MUESTRAS.TRIGO Or _
           cmbTM.BoundText = TIPOS_MUESTRAS.HARINA_BLANDA Or _
           cmbTM.BoundText = TIPOS_MUESTRAS.HARINA_DURO Then
            frmTrigo.visible = True
        End If
    End If
End Sub

Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_TIPO_ANALISIS
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Tipo Análisis " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub

Private Sub cmdDocumentacion_Click()
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
    cd.ShowOpen
    If cd.FileName <> "" Then
       listaAdjuntos.ListItems.Add , , Replace(cd.FileName, "\", "/")
    End If
End Sub

Private Sub Command2_Click()
    If listaAdjuntos.ListItems.Count > 0 Then
        listaAdjuntos.ListItems.Remove listaAdjuntos.selectedItem.Index
    End If
End Sub

Private Sub listaAdjuntos_DblClick()
   On Error GoTo listaAdjuntos_DblClick_Error
    
    If listaAdjuntos.ListItems.Count > 0 Then
        Dim r As Long
        If Dir(listaAdjuntos.ListItems(listaAdjuntos.selectedItem.Index)) <> "" Then
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & listaAdjuntos.ListItems(listaAdjuntos.selectedItem.Index), vbMaximizedFocus)
        End If
    End If

   On Error GoTo 0
   Exit Sub

listaAdjuntos_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure listaAdjuntos_DblClick of Formulario frmTA_Detalle"
End Sub

Private Sub tarifas_Click()
    If tarifas.ListItems.Count > 0 Then
         txttarifa = Trim(tarifas.ListItems(tarifas.selectedItem.Index).SubItems(2))
         txttarifa.SetFocus
    End If
End Sub

Private Sub txttarifa_GotFocus()
    txttarifa.SelStart = 0
    txttarifa.SelLength = Len(txttarifa.Text)
End Sub

Private Sub txttarifa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
       KeyAscii = 44
    End If
    If KeyAscii = 13 Then
        anadir_precio
'        KeyAscii = 0
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDatosEspecificos_Click()
    With frmTDE_Analisis
        .PK_ANALISIS = PK
        .PK_BANO = 0
        .Show 1
    End With
End Sub

Private Sub cmdDeterminaciones_Click()
    With frmDeterminaciones_analisis
        .PK_ANALISIS = PK
        .PK_BANO = 0
        .Show 1
    End With
End Sub

Private Sub cmdok_Click()
    Dim i As Integer
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim oTA As New clsTipos_analisis
      Dim ANALISIS As Long
      With oTA
           .setNOMBRE = txtDatos(0)
           .setNOMBRE_INGLES = txtDatos(2)
           .setNOMBRE_SPA = txtDatos(6)
           .setTIPO_MUESTRA_ID = cmbTM.BoundText
           If opDuplicado(0).Value = True Then
              .setNORMALIZADO = 1
           Else
              .setNORMALIZADO = 0
           End If
           If txtnormativa <> "" Then
              .setNORMATIVA = txtnormativa.Text
           End If
           If txtDatos(3) <> "" Then
              .setMETODO = txtDatos(3)
           End If
           If txtDatos(1) <> "" Then
              .setPRECIO = moneda_bd(txtDatos(1))
           Else
              .setPRECIO = moneda_bd("0")
           End If
           If txtDatos(4) <> "" Then
              .setPARAMETRO = txtDatos(4)
           End If
           If txtDatos(5) <> "" Then
              .setDIAS_TRABAJO = txtDatos(5)
           Else
              .setDIAS_TRABAJO = 0
           End If
           If cmbtarifa.getTEXTO = "" Then
            .setTARIFA_CODIGO_ID = 0
           Else
            .setTARIFA_CODIGO_ID = cmbtarifa.getPK_SALIDA
           End If
           .setTIPO_TRIGO = 0
           If opTipoTrigo(1).Value = True Then
            .setTIPO_TRIGO = 1
           End If
           .setFACTURA_DETERMINACIONES = chkFD.Value
           .setREVISAR_FACTURA = chkrevisarfactura.Value
           ' Adjuntos
           .setADJUNTOS = ""
           For i = 1 To listaAdjuntos.ListItems.Count
             .setADJUNTOS = .getADJUNTOS & Replace(listaAdjuntos.ListItems(i).Text, "\", "/") & ";"
           Next
      End With
      Dim ohc As New clsHistorial_cambios
      If PK = 0 Then
        If MsgBox("Va a introducir un nuevo tipo de análisis. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            ANALISIS = oTA.Insertar
            If ANALISIS > 0 Then
                With ohc
                    .setTIPO = HC_TIPOS.HC_TIPO_ANALISIS
                    .setIDENTIFICADOR = ANALISIS
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = HC_CREACION
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el análisis. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación del tipo de análisis."
            frmMotivo.Show 1
            If Trim(MOTIVO) = "" Then
                MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                Exit Sub
            End If
            oTA.Modificar (PK)
            ANALISIS = PK
            With ohc
                .setTIPO = HC_TIPOS.HC_TIPO_ANALISIS
                .setIDENTIFICADOR = PK
                .setIDENTIFICADOR_TEXTO = txtDatos(0)
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .setMOTIVO = Trim(MOTIVO)
                .Insertar
            End With
        Else
            Exit Sub
        End If
      End If
      Set ohc = Nothing
      Me.MousePointer = 11
      ' Tarifas
      ' Enviar correo si se modifica la tarifa
'      If tarifa_modificada = True Then
'            Dim oParametro As New clsParametros
'            oParametro.Carga PARAM_USUARIO_VIGILADO, ""
'            If USUARIO.getID_EMPLEADO = oParametro.getVALOR Then
'                Dim asunto As String
'                Dim DETALLE As String
'                asunto = "El usuario " & USUARIO.getUSUARIO & " ha modificado la tarifa de un tipo de análisis."
'
'                DETALLE = "" & vbNewLine
'                DETALLE = DETALLE & " Fecha : " & Format(Date, "dd-mm-yyyy") & vbNewLine
'                DETALLE = DETALLE & " Hora  : " & Time & vbNewLine & vbNewLine
'                DETALLE = DETALLE & " Tipo Análisis : " & txtDatos(0) & vbNewLine & vbNewLine
'
'                DETALLE = DETALLE & " Cambios en la tarifa " & vbNewLine
'                DETALLE = DETALLE & " -------------------- " & vbNewLine
'
'                Dim CO As String
'                Dim rs2 As ADODB.RecordSet
'                CO = "SELECT A.ID_TARIFA, A.NOMBRE, B.PRECIO " & _
'                     "  FROM TARIFAS A LEFT JOIN TARIFAS_PRECIOS B ON A.ID_TARIFA = B.TARIFA_ID  AND B.TIPO_ANALISIS_ID = " & ANALISIS & _
'                     " where A.EN_VIGOR = 1 "
'                Set rs2 = datos_bd(CO)
'                Dim PRECIO As String
'                Dim precio_ant As String
'                If rs2.RecordCount > 0 Then
'                    Do
'                            For i = 1 To tarifas.ListItems.Count
'                              If tarifas.ListItems(i).Text = rs2(0) Then
'                                If IsNull(rs2(2)) Then
'                                    precio_ant = moneda("0")
'                                Else
'                                    precio_ant = moneda(rs2(2))
'                                End If
'                                If Trim(tarifas.ListItems(i).SubItems(2)) = "" Then
'                                    PRECIO = moneda("0")
'                                Else
'                                    PRECIO = moneda(tarifas.ListItems(i).SubItems(2))
'                                End If
'                                If PRECIO <> precio_ant Then
'                                    DETALLE = DETALLE & tarifas.ListItems(i).SubItems(1) & " : " & precio_ant & " -> " & PRECIO & vbNewLine
'                                End If
'                              End If
'                            Next
'                        rs2.MoveNext
'                    Loop Until rs2.EOF
'                End If
'
'                oParametro.Carga PARAM_USUARIO_VIGILADO_CORREO, ""
'                ret = Enviar_Mail_CDO(oParametro.getVALOR, asunto, DETALLE, vbNullString)
'            End If
'      End If
      
      
      If USUARIO.getPER_FACTURACION = True Then
        Dim oTP As New clsTarifas_precios
        If PK <> 0 Then
            oTP.Eliminar_por_analisis (PK)
        End If
        If tarifas.ListItems.Count > 0 Then
            For i = 1 To tarifas.ListItems.Count
              If Trim(tarifas.ListItems(i).SubItems(2)) <> "" Then
                  With oTP
                      .setTIPO_ANALISIS_ID = ANALISIS
                      .setTARIFA_ID = tarifas.ListItems(i).Text
                      .setPRECIO = moneda_bd(tarifas.ListItems(i).SubItems(2))
                      .Insertar
                  End With
              End If
            Next
        End If
      End If
      Me.MousePointer = 0
      If PK = 0 Then
          MsgBox "El tipo de análisis se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
          PK = ANALISIS
          cargar_analisis
      Else
          MsgBox "El tipo de análisis se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
          Unload Me
      End If
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
      Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmTA_Detalle"
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combo cmbTM, New clsTipos_muestra
    llenar_combo cmbtarifa, New clsTarifas_codigos, 0, Me, ""
    cabecera
    cargar_tarifas
    perfil
    Dim titulo As String
    If PK <> 0 Then
        cargar_analisis
    Else
        lbltitulo = "Alta de Tipo de Análisis"
        cmdDeterminaciones.Enabled = False
        cmdDatosEspecificos.Enabled = False
    End If
    tarifa_modificada = False
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
    If KeyAscii = 13 And Index <> 2 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 1 Then
        If txtDatos(1) <> "" Then
            txtDatos(1) = moneda(txtDatos(1))
        End If
    End If
End Sub
Public Sub cargar_analisis()
    Dim i As Integer
    lbltitulo = "Modificación de Tipo de Análisis"
    cmdDeterminaciones.Enabled = True
    cmdDatosEspecificos.Enabled = True
    Dim oTA As New clsTipos_analisis
    If oTA.CARGAR(PK) = True Then
        With oTA
            txtDatos(0) = .getNOMBRE
            txtDatos(2) = .getNOMBRE_INGLES
            txtDatos(6) = .getNOMBRE_SPA
            txtDatos(1) = moneda(.getPRECIO)
'            txtnormativa.TextRTF = .getNORMATIVA
            txtnormativa = .getNORMATIVA
            txtDatos(3) = .getMETODO
            txtDatos(4) = .getPARAMETRO
            txtDatos(5) = .getDIAS_TRABAJO
            cmbTM.BoundText = .getTIPO_MUESTRA_ID
            cmbtarifa.MostrarElemento .getTARIFA_CODIGO_ID
            chkFD.Value = .getFACTURA_DETERMINACIONES
            chkrevisarfactura.Value = .getREVISAR_FACTURA
            
            opTipoTrigo(.getTIPO_TRIGO).Value = True
            ' Adjuntos
            If .getADJUNTOS <> "" Then
                Dim ad() As String
                ad = Split(.getADJUNTOS, ";")
                For i = LBound(ad) To UBound(ad) - 1
                   listaAdjuntos.ListItems.Add , , Replace(ad(i), "/", "\")
                Next
            End If
        End With
        ' Multitarifa
        Dim oMT As New clsTarifas_precios
        Dim rs As ADODB.Recordset
        Set rs = oMT.Listado_por_analisis(PK)
        If rs.RecordCount <> 0 Then
                Do
                        For i = 1 To tarifas.ListItems.Count
                            If CInt(tarifas.ListItems(i).Text) = CInt(rs("TARIFA_ID")) Then
                                tarifas.ListItems(i).SubItems(2) = moneda(CStr(rs("PRECIO")))
                            End If
                        Next
                    rs.MoveNext
                Loop Until rs.EOF
        End If
    End If
    Set oTA = Nothing
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un nombre al análisis.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(5)) = "" Then
        MsgBox "Debe indicar la duración del análisis.", vbInformation, App.Title
        validar = False
        Exit Function
    Else
        If Not IsNumeric(txtDatos(5)) Then
            MsgBox "La duración del análisis debe ser numérica.", vbInformation, App.Title
            validar = False
            Exit Function
        End If
    End If
    If cmbTM.BoundText = "" Then
        MsgBox "Debe asignar un tipo de muestra.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
End Function
Private Sub anadir_precio()
    If tarifas.ListItems.Count > 0 Then
        If txttarifa.Text = "" Then
            MsgBox "Introduzca el precio correctamente.", vbInformation, App.Title
            txttarifa.SetFocus
        Else
            If moneda(tarifas.ListItems(tarifas.selectedItem.Index).SubItems(2)) <> moneda(txttarifa) Then
                tarifa_modificada = True
            End If
            tarifas.ListItems(tarifas.selectedItem.Index).SubItems(2) = moneda(txttarifa)
            txttarifa = ""
            If tarifas.ListItems.Count > tarifas.selectedItem.Index Then
                Set tarifas.selectedItem = tarifas.ListItems(tarifas.selectedItem.Index + 1)
                tarifas.SetFocus
                tarifas_Click
            End If
        End If
    End If
End Sub
Public Sub cargar_tarifas()
    Dim oTarifa As New clsTarifas
    Dim rs As ADODB.Recordset
    Set rs = oTarifa.Listado_por_nombre
    If rs.RecordCount <> 0 Then
        Do
            With tarifas.ListItems.Add(, , rs(3))
                .SubItems(1) = rs(0)
                .SubItems(2) = " "
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub
Public Sub cabecera()
    With tarifas.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Tarifa", 2300, lvwColumnLeft
        .Add , , "Precio", 1600, lvwColumnRight
    End With
    With listaAdjuntos.ColumnHeaders
        .Add , , "Fichero", listaAdjuntos.Width, lvwColumnLeft
    End With
End Sub

Public Sub perfil()
    If USUARIO.getPER_FACTURACION = False Then
        Frame2.visible = False
    End If
End Sub
