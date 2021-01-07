VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmAirbus_ListadoMuestras 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Muestras en Factura"
   ClientHeight    =   10140
   ClientLeft      =   45
   ClientTop       =   120
   ClientWidth     =   15900
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAirbus_ListadoMuestras.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   10140
   ScaleWidth      =   15900
   Begin VB.Frame frmDatosEspeciales 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informar Datos"
      ForeColor       =   &H80000008&
      Height          =   3630
      Left            =   3510
      TabIndex        =   8
      Top             =   3060
      Visible         =   0   'False
      Width           =   8550
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   4
         Left            =   270
         TabIndex        =   39
         Top             =   2025
         Value           =   1  'Checked
         Width           =   285
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   3
         Left            =   270
         TabIndex        =   38
         Top             =   1620
         Value           =   1  'Checked
         Width           =   285
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   2
         Left            =   270
         TabIndex        =   37
         Top             =   1215
         Value           =   1  'Checked
         Width           =   285
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   36
         Top             =   810
         Value           =   1  'Checked
         Width           =   285
      End
      Begin VB.CheckBox chkOpt 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Index           =   0
         Left            =   270
         TabIndex        =   35
         Top             =   405
         Value           =   1  'Checked
         Width           =   285
      End
      Begin pryCombo.miCombo cmbPrograma 
         Height          =   330
         Left            =   1665
         TabIndex        =   9
         Top             =   765
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   795
         Left            =   2565
         TabIndex        =   11
         Top             =   2700
         Width           =   2760
         _Version        =   851970
         _ExtentX        =   4868
         _ExtentY        =   1402
         _StockProps     =   79
         Caption         =   "Informar Datos"
         Appearance      =   5
         Picture         =   "frmAirbus_ListadoMuestras.frx":08CA
      End
      Begin pryCombo.miCombo cmbEnsayo 
         Height          =   330
         Left            =   1665
         TabIndex        =   12
         Top             =   360
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbSection 
         Height          =   330
         Left            =   1665
         TabIndex        =   16
         Top             =   1170
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbFluid 
         Height          =   330
         Left            =   1665
         TabIndex        =   18
         Top             =   1575
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.PushButton cmdCerrarDatos 
         Height          =   795
         Left            =   6165
         TabIndex        =   32
         Top             =   2655
         Width           =   1500
         _Version        =   851970
         _ExtentX        =   2646
         _ExtentY        =   1402
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   5
         Picture         =   "frmAirbus_ListadoMuestras.frx":712C
      End
      Begin pryCombo.miCombo cmbFacility 
         Height          =   330
         Left            =   1665
         TabIndex        =   20
         Top             =   1980
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   582
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facility"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   675
         TabIndex        =   21
         Top             =   2025
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fluid"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   675
         TabIndex        =   19
         Top             =   1620
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Section"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   675
         TabIndex        =   17
         Top             =   1215
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   675
         TabIndex        =   13
         Top             =   405
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Programa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   675
         TabIndex        =   10
         Top             =   810
         Width           =   870
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7755
      Left            =   45
      TabIndex        =   4
      Top             =   1485
      Width           =   15825
      _ExtentX        =   27914
      _ExtentY        =   13679
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de búsqueda"
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   45
      TabIndex        =   1
      Top             =   315
      Width           =   15840
      Begin VB.OptionButton opT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Facility"
         Height          =   195
         Index           =   5
         Left            =   11970
         TabIndex        =   30
         Top             =   810
         Width           =   1455
      End
      Begin VB.OptionButton opT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Fluid"
         Height          =   195
         Index           =   4
         Left            =   11970
         TabIndex        =   29
         Top             =   495
         Width           =   1455
      End
      Begin VB.OptionButton opT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Section"
         Height          =   195
         Index           =   3
         Left            =   11970
         TabIndex        =   28
         Top             =   180
         Width           =   1455
      End
      Begin VB.OptionButton opT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Programa"
         Height          =   195
         Index           =   2
         Left            =   10350
         TabIndex        =   27
         Top             =   810
         Width           =   2310
      End
      Begin VB.OptionButton opT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Ensayo"
         Height          =   195
         Index           =   1
         Left            =   10350
         TabIndex        =   26
         Top             =   495
         Width           =   2310
      End
      Begin VB.OptionButton opT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin Planta"
         Height          =   195
         Index           =   0
         Left            =   10350
         TabIndex        =   25
         Top             =   180
         Width           =   1500
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   375
         Left            =   990
         TabIndex        =   6
         Top             =   270
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   661
      End
      Begin VB.CommandButton cmdiniciar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   14670
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   1035
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   13545
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   180
         Width           =   1035
      End
      Begin pryCombo.miCombo cmbAnalisis 
         Height          =   375
         Left            =   990
         TabIndex        =   14
         Top             =   675
         Width           =   9030
         _ExtentX        =   15928
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T. Análisis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   15
         Top             =   765
         Width           =   720
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   2
         Top             =   360
         Width           =   480
      End
   End
   Begin XtremeSuiteControls.PushButton cmdcancel 
      Height          =   840
      Left            =   14130
      TabIndex        =   7
      Top             =   9270
      Width           =   1725
      _Version        =   851970
      _ExtentX        =   3043
      _ExtentY        =   1482
      _StockProps     =   79
      Caption         =   "Salir"
      Appearance      =   5
      Picture         =   "frmAirbus_ListadoMuestras.frx":D98E
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2565
      Top             =   9540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAirbus_ListadoMuestras.frx":141F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAirbus_ListadoMuestras.frx":14ACA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAirbus_ListadoMuestras.frx":153A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAirbus_ListadoMuestras.frx":15C7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAirbus_ListadoMuestras.frx":16558
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAirbus_ListadoMuestras.frx":16E32
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAirbus_ListadoMuestras.frx":1D694
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeSuiteControls.PushButton PushButton4 
      Height          =   840
      Left            =   12375
      TabIndex        =   22
      Top             =   9270
      Width           =   1725
      _Version        =   851970
      _ExtentX        =   3043
      _ExtentY        =   1482
      _StockProps     =   79
      Caption         =   "Informar Datos"
      Appearance      =   5
      Picture         =   "frmAirbus_ListadoMuestras.frx":23EF6
   End
   Begin XtremeSuiteControls.PushButton cmbMarca 
      Height          =   345
      Index           =   0
      Left            =   45
      TabIndex        =   23
      Top             =   9270
      Width           =   1905
      _Version        =   851970
      _ExtentX        =   3360
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Marcar Todas"
      Appearance      =   5
      Picture         =   "frmAirbus_ListadoMuestras.frx":2A758
   End
   Begin XtremeSuiteControls.PushButton cmbMarca 
      Height          =   345
      Index           =   1
      Left            =   1980
      TabIndex        =   24
      Top             =   9270
      Width           =   1950
      _Version        =   851970
      _ExtentX        =   3440
      _ExtentY        =   609
      _StockProps     =   79
      Caption         =   "Desmarcar Todas"
      Appearance      =   5
      Picture         =   "frmAirbus_ListadoMuestras.frx":30FBA
   End
   Begin XtremeSuiteControls.PushButton PushButton2 
      Height          =   840
      Left            =   5715
      TabIndex        =   31
      Top             =   9270
      Width           =   1725
      _Version        =   851970
      _ExtentX        =   3043
      _ExtentY        =   1482
      _StockProps     =   79
      Caption         =   "Datos Cliente"
      Appearance      =   5
      Picture         =   "frmAirbus_ListadoMuestras.frx":3781C
   End
   Begin XtremeSuiteControls.PushButton PushButton3 
      Height          =   840
      Left            =   3960
      TabIndex        =   33
      Top             =   9270
      Width           =   1725
      _Version        =   851970
      _ExtentX        =   3043
      _ExtentY        =   1482
      _StockProps     =   79
      Caption         =   "Airbus Plantas"
      Appearance      =   5
      Picture         =   "frmAirbus_ListadoMuestras.frx":3E07E
   End
   Begin XtremeSuiteControls.PushButton cmdExcel 
      Height          =   840
      Left            =   10620
      TabIndex        =   40
      Top             =   9270
      Width           =   1725
      _Version        =   851970
      _ExtentX        =   3043
      _ExtentY        =   1482
      _StockProps     =   79
      Caption         =   "Generar Excel ADS"
      Appearance      =   5
      Picture         =   "frmAirbus_ListadoMuestras.frx":448E0
   End
   Begin VB.Label lblRegistros 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   285
      Left            =   13635
      TabIndex        =   34
      Top             =   0
      Width           =   2265
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Muestras en Factura"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   3
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "frmAirbus_ListadoMuestras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ID_FACTURA As Long
Public ID_MUESTRAS As String
Option Explicit

Private Enum COLS
    ID_MUESTRA = 0
    CODIGO = 1
    cliente = 2
    TIPO_ANALISIS = 3
    REFERENCIA_CLIENTE = 4
    ID_GENERAL = 5
    PLANTA_ID = 6
    ENSAYO_ID = 7
    PROGRAMA_ID = 8
    SECTION_ID = 9
    FLUID_ID = 10
    FACILITY_ID = 11
End Enum


Private Sub cmdCerrarDatos_Click()
    frmDatosEspeciales.visible = False
End Sub

Private Sub cmdExcel_Click()
    Dim fichero As String
    Dim PLANTILLA As String
   On Error GoTo cmdExcel_Click_Error

    ' Cargar parametros
    Dim op As New clsParametros
    op.Carga parametros.ADS_DATOS_EXCEL, ""
    Dim l() As String
    l = Split(op.getVALOR, ";")
    'PLANTILLA;LINEA
    PLANTILLA = ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\" & l(0)
    fichero = DIRECTORIO_TEMPORAL & "\" & l(0)
    
    Dim fso As New FileSystemObject
    fso.CopyFile PLANTILLA, fichero, True
    Set fso = Nothing
    
    ' Leer contenido
    Dim rs As ADODB.Recordset
    Dim c As String
'    c = " SELECT MONTHNAME(a.fecha_factura),concat(lpad(a.numero,4,'0000'),'/',YEAR(a.FECHA_FACTURA)) AS DELIVERY " & _
'        "      ,d0.descripcion AS PLANTA,d1.descripcion AS ENSAYO,d2.descripcion AS PROGRAMA,d3.descripcion AS SECCION,d4.descripcion AS FLUIDO " & _
'        "        ,IF(isnull(d5.DESCRIPCION),c.REFERENCIA_CLIENTE,d5.DESCRIPCION) AS FACILITY " & _
'        "      ,REPLACE(IF(b.CODIGO <> '',b.TIPO_ANALISIS,tdet.NOMBRE),'*','') AS ANALYSIS " & _
'        "        ,IF(b.CODIGO <> '',b.CODIGO,IF(isnull(det.ID_DETERMINACION),b.codigo,dpm2.CODIGO)) AS TARIFA " & _
'        "      ,IF(b.CODIGO <> '',b.PRECIO,IF(isnull(det.ID_DETERMINACION),b.PRECIO,dpm2.PRECIO)) AS COST " & _
'        "      ,COUNT(DISTINCT b.DOC_ID,b.MUESTRA_ID,b.ORDEN) AS SAMPLES " & _
'        "      ,IF(b.CODIGO <> '',b.PRECIO,IF(isnull(det.ID_DETERMINACION),b.PRECIO,dpm2.PRECIO)) * COUNT(DISTINCT b.DOC_ID,b.MUESTRA_ID,b.ORDEN) AS IMPORTE " & _
'        "        ,IF (c.OP_REPETICION = 1,'REPETICION',IF(c.OP_NORUTINARIA = 0,'SI','NO')) AS PLANED "
'        c = c & " from docs_pago a " & _
'        " inner join docs_pago_muestras b on a.id_doc = b.doc_id and b.DETERMINACION_ID = 0 " & _
'        " inner join muestras c on c.id_muestra = b.muestra_id " & _
'        " left join muestras_airbus d ON c.id_muestra = d.muestra_id " & _
'        " inner join clientes e on c.cliente_id = e.id_cliente " & _
'        " inner join tipos_muestra f on c.tipo_muestra_id = f.id_tipo_muestra " & _
'        " inner join tipos_analisis g on c.tipo_analisis_id = g.id_tipo_analisis " & _
'        " left join docs_pago_muestras dpm2 on a.id_doc = dpm2.doc_id AND dpm2.MUESTRA_ID = c.ID_MUESTRA and dpm2.DETERMINACION_ID > 0 " & _
'        " left join determinaciones det ON det.ID_DETERMINACION = dpm2.DETERMINACION_ID " & _
'        " left join tipos_determinacion tdet ON det.tipo_determinacion_id = tdet.ID_TIPO_DETERMINACION " & _
'        " left join decodificadora d0 on d0.codigo = 600 and d0.valor = e.plant_id " & _
'        " left join decodificadora d1 on d1.codigo = 601 and d1.valor = d.ensayo_id " & _
'        " left join decodificadora d2 on d2.codigo = 602 and d2.valor = d.programa_id " & _
'        " left join decodificadora d3 on d3.codigo = 603 and d3.valor = d.section_id " & _
'        " left join decodificadora d4 on d4.codigo = 604 and d4.valor = d.fluid_id " & _
'        " left join decodificadora d5 on d5.codigo = 605 and d5.valor = d.facility_id " & _
'        " where a.ID_DOC in (" & ID_FACTURA & ")" & _
'        " group by 1,2,3,4,5,6,7,8,9,10,11,14 "
    c = " SELECT a.ID_DOC,MONTHNAME(a.fecha_factura),concat(lpad(a.numero,4,'0000'),'/',YEAR(a.FECHA_FACTURA)) AS DELIVERY " & _
        "      ,d0.descripcion AS PLANTA,d1.descripcion AS ENSAYO,d2.descripcion AS PROGRAMA,d3.descripcion AS SECCION,d4.descripcion AS FLUIDO " & _
        "      ,IF(isnull(d5.DESCRIPCION),c.REFERENCIA_CLIENTE,d5.DESCRIPCION) AS FACILITY " & _
        "      ,REPLACE(IF(b.CODIGO <> '' OR isnull(dpm2.TIPO_ANALISIS),b.TIPO_ANALISIS,dpm2.TIPO_ANALISIS),'*','') AS ANALYSIS " & _
        "      ,IF(b.CODIGO <> '' OR isnull(dpm2.CODIGO),b.CODIGO, dpm2.CODIGO) AS TARIFA " & _
        "      ,IF(b.CODIGO <> '' OR isnull(dpm2.PRECIO),b.PRECIO, dpm2.PRECIO) AS COST " & _
        "      ,COUNT(DISTINCT b.DOC_ID,b.MUESTRA_ID,b.ORDEN) AS SAMPLES " & _
        "      ,IF(b.CODIGO <> '' OR isnull(dpm2.PRECIO),b.PRECIO, dpm2.PRECIO) * COUNT(DISTINCT b.DOC_ID,b.MUESTRA_ID,b.ORDEN) AS IMPORTE " & _
        "      ,IF (c.OP_REPETICION = 1,'REPETICION',IF(c.OP_NORUTINARIA = 0,'SI','NO')) AS PLANED "
        c = c & " from docs_pago a " & _
        " inner join docs_pago_muestras b on a.id_doc = b.doc_id and b.DETERMINACION_ID = 0 " & _
        " inner join muestras c on c.id_muestra = b.muestra_id " & _
        " left join muestras_airbus d ON c.id_muestra = d.muestra_id " & _
        " inner join clientes e on c.cliente_id = e.id_cliente " & _
        " left join docs_pago_muestras dpm2 on a.id_doc = dpm2.doc_id AND dpm2.MUESTRA_ID = c.ID_MUESTRA and dpm2.DETERMINACION_ID <> 0 " & _
        " left join decodificadora d0 on d0.codigo = 600 and d0.valor = e.plant_id " & _
        " left join decodificadora d1 on d1.codigo = 601 and d1.valor = d.ensayo_id " & _
        " left join decodificadora d2 on d2.codigo = 602 and d2.valor = d.programa_id " & _
        " left join decodificadora d3 on d3.codigo = 603 and d3.valor = d.section_id " & _
        " left join decodificadora d4 on d4.codigo = 604 and d4.valor = d.fluid_id " & _
        " left join decodificadora d5 on d5.codigo = 605 and d5.valor = d.facility_id " & _
        " where a.ID_DOC in (" & ID_FACTURA & ")" & _
        " group by 1,2,3,4,5,6,7,8,9,10,11,12,15 "
    
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        ' Cargar Excel
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Open(fichero)
        Set XLS = XLW.Worksheets(1)
        Dim linea As Integer
        Dim total As Currency
        linea = l(1)
        Do
            XLS.Cells(linea, 5) = rs(1)
            XLS.Cells(linea, 6) = rs(2)
            XLS.Cells(linea, 7) = rs(3)
            XLS.Cells(linea, 8) = rs(4)
            XLS.Cells(linea, 9) = rs(5)
            XLS.Cells(linea, 10) = rs(6)
            XLS.Cells(linea, 11) = rs(7)
            XLS.Cells(linea, 12) = rs(8)
            XLS.Cells(linea, 13) = rs(9)
            XLS.Cells(linea, 14) = rs(10)
            XLS.Cells(linea, 15) = rs(11)
            XLS.Cells(linea, 16) = rs(12)
            XLS.Cells(linea, 17) = rs(13)
            XLS.Cells(linea, 18) = rs(14)
            total = total + rs(13)
            rs.MoveNext
            linea = linea + 1
        Loop Until rs.EOF
        ' Incluir lineas de conceptos (solicitado por Carmen 01/08/2019
        c = "select MONTHNAME(a.fecha_factura),concat(lpad(a.numero,4,'0000'),'/',YEAR(a.FECHA_FACTURA)) AS DELIVERY " & _
            " , '' as PLANTA, '' as ENSAYO, '' as PROGRAMA, '' as SECCION, '' as FLUIDO, '' as FACILITY " & _
            " ,b.DESCRIPCION as ANALYSIS,'' as REFERENCE, b.PRECIO as COST, b.CANTIDAD as UNIT, b.TOTAL,'SI' " & _
            " from docs_pago a " & _
            " inner join docs_pago_conceptos b on a.ID_DOC = b.doc_id " & _
            " where a.ID_DOC in (" & ID_FACTURA & ")" & _
            " order by a.FECHA_FACTURA,b.ID_CONCEPTO "
        Dim rs2 As ADODB.Recordset
        Set rs2 = datos_bd(c)
        If rs2.RecordCount > 0 Then
            Do
                XLS.Cells(linea, 5) = rs2(0)
                XLS.Cells(linea, 6) = rs2(1)
                XLS.Cells(linea, 7) = rs2(2)
                XLS.Cells(linea, 8) = rs2(3)
                XLS.Cells(linea, 9) = rs2(4)
                XLS.Cells(linea, 10) = rs2(5)
                XLS.Cells(linea, 11) = rs2(6)
                XLS.Cells(linea, 12) = rs2(7)
                ' Analizar el concepto, si empieza WP es el código
                If Left(rs2(8), 2) = "WP" Then
                    XLS.Cells(linea, 13) = Mid(rs2(8), 13, Len(rs2(8)) - 13)
                    XLS.Cells(linea, 14) = Left(rs2(8), 12)
                Else
                    XLS.Cells(linea, 13) = rs2(8)
                    XLS.Cells(linea, 14) = rs2(9)
                End If
                XLS.Cells(linea, 15) = rs2(10)
                XLS.Cells(linea, 16) = rs2(11)
                XLS.Cells(linea, 17) = rs2(12)
                XLS.Cells(linea, 18) = rs2(13)
                rs2.MoveNext
                linea = linea + 1
            Loop Until rs2.EOF
        End If
        XLS.Cells(linea, 13) = "SUMA IMPORTES MUESTRAS"
        XLS.Cells(linea, 15) = moneda_bd(CStr(total))
        XLS.Cells(linea, 16) = 1
        XLS.Cells(linea, 17) = moneda_bd(CStr(total))
        ' INCLUIR LINEA DE CONCEPTOS
        linea = linea + 1
        Dim oDPC As New clsDocs_pago_conceptos
        Dim totalConceptos As Currency
        totalConceptos = oDPC.ImporteDocumento(ID_FACTURA)
        XLS.Cells(linea, 13) = "SUMA IMPORTES CONCEPTOS"
        XLS.Cells(linea, 15) = moneda_bd(CStr(totalConceptos))
        XLS.Cells(linea, 16) = 1
        XLS.Cells(linea, 17) = moneda_bd(CStr(totalConceptos))
        ' INCLUIR LINEA DE CONCEPTOS
        linea = linea + 1
        XLS.Cells(linea, 13) = "MUESTRAS + CONCEPTOS"
        XLS.Cells(linea, 15) = moneda_bd(CStr(total + totalConceptos))
        XLS.Cells(linea, 16) = 1
        XLS.Cells(linea, 17) = moneda_bd(CStr(total + totalConceptos))
        ' INCLUIR LINEA DE TOTAL FACTURA
        linea = linea + 1
        Dim oDP As New clsDocs_pago
        oDP.CargarDocumento ID_FACTURA
        XLS.Cells(linea, 13) = "TOTAL IMPORTE REGISTRADO EN LA FACTURA GESLAB"
        XLS.Cells(linea, 15) = moneda_bd(CStr(oDP.getTOTAL))
        XLS.Cells(linea, 16) = 1
        XLS.Cells(linea, 17) = moneda_bd(CStr(oDP.getTOTAL))
        XLW.Save
        XLA.visible = True
        Set XLS = Nothing
        Set XLW = Nothing
        Set XLA = Nothing
        MsgBox "Exportación finalizada correctamente.", vbOKOnly + vbInformation, App.Title
    Else
        MsgBox "No existen datos para generar el documento. En el caso de TCT asegurese de que sean los albaranes.", vbExclamation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cmdExcel_Click_Error:
    Set XLS = Nothing
    Set XLW = Nothing
    Set XLA = Nothing

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExcel_Click of Formulario frmAirbus_ListadoMuestras"

End Sub

Private Sub PushButton1_Click()
    Dim i As Integer
   On Error GoTo PushButton1_Click_Error

    Me.MousePointer = 11
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Selected = True Then
            Dim oM As New clsMuestras_airbus
            With oM
                .setMUESTRA_ID = lista.ListItems(i).Text
                .setENSAYO_ID = IIf(cmbEnsayo.getTEXTO = "", 0, cmbEnsayo.getPK_SALIDA)
                .setPROGRAMA_ID = IIf(cmbPrograma.getTEXTO = "", 0, cmbPrograma.getPK_SALIDA)
                .setSECTION_ID = IIf(cmbSection.getTEXTO = "", 0, cmbSection.getPK_SALIDA)
                .setFLUID_ID = IIf(cmbFluid.getTEXTO = "", 0, cmbFluid.getPK_SALIDA)
                .setFACILITY_ID = IIf(cmbFacility.getTEXTO = "", 0, cmbFacility.getPK_SALIDA)
                .Insertar chkOpt(0), chkOpt(1), chkOpt(2), chkOpt(3), chkOpt(4)
            End With
        End If
    Next
    frmDatosEspeciales.visible = False
    actualizarLista
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

PushButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PushButton1_Click of Formulario frmAirbus_ListadoMuestras"
End Sub

Private Sub PushButton2_Click()
    If lista.ListItems.Count > 0 Then
        Dim oM As New clsMuestra
        oM.CargaMuestra lista.ListItems(lista.selectedItem.Index).Text
        frmClientes.PK = oM.getCLIENTE_ID
        frmClientes.Show 1
        actualizarLista
    End If
End Sub
Private Sub cmbMarca_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If Index = 0 Then
            lista.ListItems(i).Selected = True
        Else
            lista.ListItems(i).Selected = False
        End If
    Next
End Sub

Private Sub cmdBuscar_Click()
    buscar
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdiniciar_Click()
    cmbClientes.limpiar
    lista.ListItems.Clear
    Dim i As Integer
    For i = 0 To 5
        opT(i).Value = False
    Next
    buscar
End Sub
Private Sub Form_Activate()
'    Me.SetFocus
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = (frmMenu.ScaleWidth - Me.Width) / 2
    Me.top = (frmMenu.ScaleHeight - Me.Height) / 2
    cargar_botones Me
    cargar_titulo
    cargar_combo
    cabecera_grid
    permisos
    buscar
    If ID_FACTURA = 0 Then
        cmdExcel.visible = False
    End If
End Sub
Private Sub permisos()
    If USUARIO.getPER_MOD_CLIENTE = False Then
        PushButton2.visible = False
    End If
End Sub

Private Sub cargar_titulo()
    If ID_MUESTRAS <> "" Then
        lbltitulo = "Listado de muestras recepcionadas para informar datos ADS"
    Else
        Dim oDP As New clsDocs_pago
        Dim oCliente As New clsCliente
        oDP.CargarDocumento ID_FACTURA
        oCliente.CargaCliente oDP.getCLIENTE_ID
        lbltitulo = "Listado de muestras del documento : " & oDP.getNUMERO_FORMATEADO & "/" & Year(oDP.getFECHA_FACTURA) & " Cliente : " & oCliente.getNOMBRE
    End If
    Me.Caption = lbltitulo
End Sub
Private Sub buscar()
    Dim i As Integer
    Dim consulta As String
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    Dim oMuestra As New clsMuestra
    Me.MousePointer = 11
    If ID_MUESTRAS <> "" Then
        Set rs = oMuestra.MuestrasAirbus(ID_MUESTRAS, IIf(cmbClientes.getTEXTO = "", 0, cmbClientes.getPK_SALIDA), IIf(cmbAnalisis.getTEXTO = "", 0, cmbAnalisis.getPK_SALIDA), opT(0).Value, opT(1).Value, opT(2).Value, opT(3).Value, opT(4).Value, opT(5).Value)
    Else
        Set rs = oMuestra.MuestrasEnFactura(CStr(ID_FACTURA), 0, IIf(cmbClientes.getTEXTO = "", 0, cmbClientes.getPK_SALIDA), IIf(cmbAnalisis.getTEXTO = "", 0, cmbAnalisis.getPK_SALIDA), opT(0).Value, opT(1).Value, opT(2).Value, opT(3).Value, opT(4).Value, opT(5).Value)
    End If
    lblRegistros = rs.RecordCount & " muestras"
    If rs.RecordCount >= 1 Then
        i = 1
        Dim objLitem As ListItem, objSI As ListSubItem
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(0))
                .SubItems(COLS.CODIGO) = rs(1)
                .SubItems(COLS.cliente) = rs(2)
                .SubItems(COLS.TIPO_ANALISIS) = rs(3)
                .SubItems(COLS.REFERENCIA_CLIENTE) = rs(4)
                .SubItems(COLS.ID_GENERAL) = rs(5)
                ' AIRBUS
                If Not IsNull(rs(6)) Then
                    .SubItems(COLS.PLANTA_ID) = rs(6)
                Else
                    .SubItems(COLS.PLANTA_ID) = ""
                End If
                If Not IsNull(rs(7)) Then
                    .SubItems(COLS.ENSAYO_ID) = rs(7)
                Else
                    .SubItems(COLS.ENSAYO_ID) = ""
                End If
                If Not IsNull(rs(8)) Then
                    .SubItems(COLS.PROGRAMA_ID) = rs(8)
                Else
                    .SubItems(COLS.PROGRAMA_ID) = ""
                End If
                If Not IsNull(rs(9)) Then
                    .SubItems(COLS.SECTION_ID) = rs(9)
                Else
                    .SubItems(COLS.SECTION_ID) = ""
                End If
                If Not IsNull(rs(10)) Then
                    .SubItems(COLS.FLUID_ID) = rs(10)
                Else
                    .SubItems(COLS.FLUID_ID) = ""
                End If
                If Not IsNull(rs(11)) Then
                    .SubItems(COLS.FACILITY_ID) = rs(11)
                Else
                    .SubItems(COLS.FACILITY_ID) = ""
                End If
            End With
            i = lista.ListItems.Count
            rs.MoveNext
        Wend
'        lblMsg.Caption = "Muestras entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy") & ". Total : " & rs.RecordCount
    Else
'        lblMsg.Caption = "No existe ninguna muestra con esos criterios."
    End If
'    Set oAnalisis = Nothing
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras.", vbCritical, Err.Description
End Sub
Private Sub actualizarLista()
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim oMuestra As New clsMuestra
    Me.MousePointer = 11
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Selected = True Then
            If ID_MUESTRAS <> "" Then
                Set rs = oMuestra.MuestrasAirbus(lista.ListItems(i).Text, IIf(cmbClientes.getTEXTO = "", 0, cmbClientes.getPK_SALIDA), IIf(cmbAnalisis.getTEXTO = "", 0, cmbAnalisis.getPK_SALIDA), opT(0).Value, opT(1).Value, opT(2).Value, opT(3).Value, opT(4).Value, opT(5).Value)
            Else
                Set rs = oMuestra.MuestrasEnFactura(CStr(ID_FACTURA), lista.ListItems(i).Text, 0, 0, opT(0).Value, opT(1).Value, opT(2).Value, opT(3).Value, opT(4).Value, opT(5).Value)
            End If
            If rs.RecordCount >= 1 Then
                While Not rs.EOF
                    With lista.ListItems(i)
                        .SubItems(COLS.CODIGO) = rs(1)
                        .SubItems(COLS.cliente) = rs(2)
                        .SubItems(COLS.TIPO_ANALISIS) = rs(3)
                        .SubItems(COLS.REFERENCIA_CLIENTE) = rs(4)
                        .SubItems(COLS.ID_GENERAL) = rs(5)
                        ' AIRBUS
                        If Not IsNull(rs(6)) Then
                            .SubItems(COLS.PLANTA_ID) = rs(6)
                        Else
                            .SubItems(COLS.PLANTA_ID) = ""
                        End If
                        If Not IsNull(rs(7)) Then
                            .SubItems(COLS.ENSAYO_ID) = rs(7)
                        Else
                            .SubItems(COLS.ENSAYO_ID) = ""
                        End If
                        If Not IsNull(rs(8)) Then
                            .SubItems(COLS.PROGRAMA_ID) = rs(8)
                        Else
                            .SubItems(COLS.PROGRAMA_ID) = ""
                        End If
                        If Not IsNull(rs(9)) Then
                            .SubItems(COLS.SECTION_ID) = rs(9)
                        Else
                            .SubItems(COLS.SECTION_ID) = ""
                        End If
                        If Not IsNull(rs(10)) Then
                            .SubItems(COLS.FLUID_ID) = rs(10)
                        Else
                            .SubItems(COLS.FLUID_ID) = ""
                        End If
                        If Not IsNull(rs(11)) Then
                            .SubItems(COLS.FACILITY_ID) = rs(11)
                        Else
                            .SubItems(COLS.FACILITY_ID) = ""
                        End If
                    End With
                    rs.MoveNext
                Wend
            End If
        End If
    Next
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras.", vbCritical, Err.Description
End Sub

Private Function contar_marcados() As Integer
    Dim i As Integer
    contar_marcados = 0
    For i = 1 To lista.ListItems.Count
       If lista.ListItems(i).Checked = True Then
        contar_marcados = contar_marcados + 1
      End If
    Next
End Function

Private Sub cabecera_grid()
    With lista.ColumnHeaders
        .Add , , "ID_MUESTRA", 0, lvwColumnLeft
        .Add , , "Código", 1200, lvwColumnCenter
        .Add , , "Cliente", 2300, lvwColumnLeft
        .Add , , "Tipo de Analisis/Solución", 2300, lvwColumnLeft
        .Add , , "Ref.Cliente", 2300, lvwColumnLeft
        .Add , , "General", 0, lvwColumnLeft
        ' muestras_airbus
        .Add , , "Planta", 2000, lvwColumnLeft
        .Add , , "Ensayo", 2300, lvwColumnLeft
        .Add , , "Programa", 2300, lvwColumnLeft
        .Add , , "Section", 2300, lvwColumnLeft
        .Add , , "Fluid", 2300, lvwColumnLeft
        .Add , , "Facility", 2300, lvwColumnLeft
    End With
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
'        cmbClienteFactura.MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(COLS.CLIENTE_ID)
'        If lista.ListItems(lista.selectedItem.Index).SubItems(COLS.PEDIDO_ID) <> 0 Then
'            cmbPedidos.MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(COLS.PEDIDO_ID)
'        Else
'            pedidos cmbClienteFactura.getPK_SALIDA
'        End If
    End If
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub

Public Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).Text
        frmVerMuestra.Show 1
        actualizarLista
        gmuestra = 0
    End If
End Sub
Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        PushButton4_Click
    End If
End Sub
Private Sub colorear(fila As Integer, color As Long)
    Dim i As Integer
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
End Sub
Private Sub cargar_combo_clientes()
    Dim consulta As String
    Dim rs As ADODB.Recordset

   On Error GoTo cargar_combo_clientes_Error

    consulta = "SELECT GROUP_CONCAT(DISTINCT ALBARAN_ID) FROM DOCS_PAGO_CONCEPTOS " & _
               " WHERE DOC_ID = " & ID_FACTURA & _
               "   AND ALBARAN_ID <> 0;"
    Set rs = datos_bd(consulta)
    Dim docs As String
    If IsNull(rs(0)) Then
        docs = ID_FACTURA
    Else
        docs = rs(0)
    End If
    
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        consulta = "SELECT DISTINCT C.ID_CLIENTE,C.NOMBRE " & _
                   "  FROM CLIENTES AS C, DOCS_PAGO AS DP " & _
                   " WHERE C.ID_CLIENTE = DP.CLIENTE_ID " & _
                   "   AND DP.ID_DOC IN (" & docs & ")"
        With cmbClientes
            .setCONN = conn
                .setFK_CAMPO = ""
                .setFK_VALOR = 0
                .setTABLA = "CLIENTES"
                .setDESCRIPCION = "Clientes"
                .setPK = "ID_CLIENTE"
                .setCAMPO = "NOMBRE"
                .setQUERY = consulta
                .setMUESTRA_DETALLE = True
                Set .FORMULARIO = frmClientes
        End With
        'MDET
        consulta = "SELECT DISTINCT TA.ID_TIPO_ANALISIS,TA.NOMBRE " & _
                   "  FROM MUESTRAS M, TIPOS_ANALISIS AS TA, DOCS_PAGO_MUESTRAS AS DPM " & _
                   " WHERE M.TIPO_ANALISIS_ID = TA.ID_TIPO_ANALISIS AND M.ID_MUESTRA = DPM.MUESTRA_ID " & _
                   "   AND DPM.DOC_ID IN (" & docs & ")" & _
                   "   AND DPM.MUESTRA_ID <> 0 AND DPM.DETERMINACION_ID = 0"
        With cmbAnalisis
            .setCONN = conn
                .setFK_CAMPO = ""
                .setFK_VALOR = 0
                .setTABLA = "TIPOS_ANALISIS"
                .setDESCRIPCION = "Tipos de Analisis"
                .setPK = "ID_TIPO_ANALISIS"
                .setCAMPO = "NOMBRE"
                .setQUERY = consulta
                .setMUESTRA_DETALLE = True
                Set .FORMULARIO = frmTA_Detalle
        End With
    End If

   On Error GoTo 0
   Exit Sub

cargar_combo_clientes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_combo_clientes of Formulario frmAirbus_ListadoMuestras"
End Sub
Private Sub cargar_combo()
    cargar_combo_clientes
End Sub
Private Sub cargar_planta(ID_PLANTA As String)
    Dim oDeco As New clsDecodificadora
    cmbEnsayo.limpiar
    cmbPrograma.limpiar
    cmbSection.limpiar
    cmbFluid.limpiar
    cmbFacility.limpiar
    oDeco.cargar_mi_combo_parametro cmbEnsayo, DECODIFICADORA.AIRBUS_TIPOS_ENSAYOS, ID_PLANTA
    oDeco.cargar_mi_combo_parametro cmbPrograma, DECODIFICADORA.AIRBUS_PROGRAMAS, ID_PLANTA
    oDeco.cargar_mi_combo_parametro cmbSection, DECODIFICADORA.AIRBUS_SECTION, ID_PLANTA
    oDeco.cargar_mi_combo_parametro cmbFluid, DECODIFICADORA.AIRBUS_FLUID, ID_PLANTA
    oDeco.cargar_mi_combo_parametro cmbFacility, DECODIFICADORA.AIRBUS_FACILITY, ID_PLANTA
End Sub

Private Sub PushButton3_Click()
    frmAirbus_Decodificadora.Show 1
    
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oM As New clsMuestra
    Dim oC As New clsCliente
    oM.CargaMuestra lista.ListItems(lista.selectedItem.Index).Text
    oC.CargaCliente oM.getCLIENTE_ID
    If oC.getPLANT_ID <> 0 Then
        cargar_planta oC.getPLANT_ID
    End If
End Sub

Private Sub PushButton4_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oM As New clsMuestra
    Dim oC As New clsCliente
    oM.CargaMuestra lista.ListItems(lista.selectedItem.Index).Text
    oC.CargaCliente oM.getCLIENTE_ID
    If oC.getPLANT_ID = 0 Then
        MsgBox "El cliente no tiene informado el campo PLANTA.", vbCritical, App.Title
    Else
        cargar_planta oC.getPLANT_ID
        frmDatosEspeciales.visible = Not frmDatosEspeciales.visible
    End If
'    If lista.ListItems.Count = 0 Then
'       frmDatosEspeciales.top = Me.Height / 2 - frmDatosEspeciales.Height
'    Else
'        frmDatosEspeciales.top = lista.ListItems(lista.selectedItem.Index).top + 600
'    End If
End Sub
