VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{67129E04-3D95-4F4C-B7F3-EE4C17D586DF}#1.1#0"; "ButtonBar.ocx"
Begin VB.MDIForm frmMenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Gestión Comercial v1.1"
   ClientHeight    =   10695
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   11400
   Icon            =   "frmMenu.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin DevPowerButtonBar.ButtonBar ButtonBar1 
      Align           =   3  'Align Left
      Height          =   10320
      Left            =   0
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   18203
      _Type           =   "00007615303F262704031E0E3E646A71017B627773147F717365656F721763003E6462737466647274"
      Style           =   1
      BackColor       =   -2147483633
      CheckColor      =   14542313
      ItemForeColor   =   -2147483630
      ItemOverBackColor=   14073525
      ItemDownBackColor=   14542313
      ItemDownForeColor=   0
      SmoothScroll    =   0   'False
      TrackNavigation =   -1  'True
      AutoMouseWheel  =   -1  'True
      ItemsStayDownWithClick=   -1  'True
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      NumHeaders      =   4
      Caption1        =   "Gestión"
      Enabled1        =   -1  'True
      ToolTip1        =   "Datos de gestión"
      View1           =   1
      Header1Items    =   9
      Header1Item1Caption=   "Clientes"
      Header1Item1ToolTip=   "Listado de Clientes"
      Header1Item1Icon=   "29"
      Header1Item2Caption=   "Obras"
      Header1Item2Icon=   "18"
      Header1Item3Caption=   "Proveedores"
      Header1Item3Icon=   "16"
      Header1Item4Caption=   "Agenda"
      Header1Item4Icon=   "24"
      Header1Item5Caption=   "Formas Pago"
      Header1Item5Icon=   "31"
      Header1Item6Caption=   "Provincias"
      Header1Item6Icon=   "47"
      Header1Item7Caption=   "Agentes"
      Header1Item7Icon=   "10"
      Header1Item8Caption=   "Vehículos"
      Header1Item8Icon=   "39"
      Header1Item9Caption=   "Portes"
      Header1Item9Icon=   "15"
      Caption2        =   "Albaranes/Facturas"
      Enabled2        =   -1  'True
      View2           =   0
      Header2Items    =   3
      Header2Item1Caption=   "Nuevo"
      Header2Item1Icon=   "33"
      Header2Item2Caption=   "Buscar"
      Header2Item2Icon=   "34"
      Header2Item3Caption=   "Facturar Albaranes"
      Header2Item3Icon=   "46"
      Caption3        =   "Almacen"
      Enabled3        =   -1  'True
      View3           =   0
      Header3Items    =   2
      Header3Item1Caption=   "Artículos"
      Header3Item1Icon=   "18"
      Header3Item2Caption=   "Tipos Artículos"
      Header3Item2Icon=   "36"
      Caption4        =   "Sistema"
      Enabled4        =   -1  'True
      View4           =   0
      Header4Items    =   2
      Header4Item1Caption=   "Usuarios"
      Header4Item1Icon=   "38"
      Header4Item2Caption=   "Salir"
      Header4Item2Icon=   "35"
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3780
      Top             =   3840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   51
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":628A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":7F94
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":BA9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":D0F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":DFD2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":F024
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":11D2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":14A38
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":16092
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1696C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":17246
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":17560
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":17E3A
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":18D14
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":195EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":19EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1A7A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1B07C
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1B956
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1C7A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1CAC2
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1CDDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1D6B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1DB08
            Key             =   ""
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1E3E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1ECBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1F596
            Key             =   ""
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":1F8B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":233BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":23C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2468E
            Key             =   ""
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":24F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":25962
            Key             =   ""
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2623C
            Key             =   ""
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":26B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":273F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":27CCA
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":285A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":28E7E
            Key             =   ""
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":29758
            Key             =   ""
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2A032
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2A90C
            Key             =   ""
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2C09E
            Key             =   ""
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2C978
            Key             =   ""
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2D252
            Key             =   ""
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2DB2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2DE46
            Key             =   ""
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2E160
            Key             =   ""
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2E6EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2EEA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2F77F
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   10320
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12347
            MinWidth        =   12347
            Text            =   "estado"
            TextSave        =   "estado"
            Object.ToolTipText     =   "Estado"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "Usuario"
            TextSave        =   "Usuario"
            Object.ToolTipText     =   "Empleado:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   4304
            MinWidth        =   4304
            Text            =   "Servidor"
            TextSave        =   "Servidor"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Text            =   "Modo"
            TextSave        =   "Modo"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList botones 
      Left            =   4995
      Top             =   3420
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":2FA99
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":30373
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":30C4D
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":31527
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":31E01
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":326DB
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":32FB5
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":3388F
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":34169
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":34A43
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":3531D
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":35BF7
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":35F11
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":367EB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu menuGestion 
      Caption         =   "Gestión"
      Begin VB.Menu opGestion 
         Caption         =   "Clientes"
         Index           =   0
      End
      Begin VB.Menu opGestion 
         Caption         =   "Obras"
         Index           =   1
      End
      Begin VB.Menu opGestion 
         Caption         =   "Proveedores"
         Index           =   2
      End
      Begin VB.Menu opGestion 
         Caption         =   "Agenda"
         Index           =   3
      End
      Begin VB.Menu opGestion 
         Caption         =   "Formas de Pago"
         Index           =   4
      End
      Begin VB.Menu opGestion 
         Caption         =   "Provincias y Municipios"
         Index           =   5
      End
      Begin VB.Menu opGestion 
         Caption         =   "Agentes"
         Index           =   6
      End
      Begin VB.Menu opGestion 
         Caption         =   "Vehículos"
         Index           =   7
      End
      Begin VB.Menu opGestion 
         Caption         =   "Tarifas de Portes"
         Index           =   8
      End
      Begin VB.Menu opGestion 
         Caption         =   "-"
         Index           =   9
      End
      Begin VB.Menu opGestion 
         Caption         =   "Conversión"
         Index           =   10
      End
   End
   Begin VB.Menu menuDocumentos 
      Caption         =   "Albaranes"
      Begin VB.Menu opdocumentos 
         Caption         =   "Nuevo Albarán"
         Index           =   0
      End
      Begin VB.Menu opdocumentos 
         Caption         =   "Buscar Albaran"
         Index           =   1
      End
      Begin VB.Menu opdocumentos 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu opdocumentos 
         Caption         =   "Listado de Albaranes"
         Index           =   3
      End
      Begin VB.Menu opdocumentos 
         Caption         =   "Albaranes no Valorados"
         Index           =   4
      End
      Begin VB.Menu opdocumentos 
         Caption         =   "Facturar Albaranes"
         Index           =   5
      End
   End
   Begin VB.Menu menuFacturas 
      Caption         =   "Facturas"
      Begin VB.Menu opFacturas 
         Caption         =   "Buscar Factura"
         Index           =   0
      End
      Begin VB.Menu opFacturas 
         Caption         =   "Listado de Facturas"
         Index           =   1
      End
      Begin VB.Menu opFacturas 
         Caption         =   "Cobro de Facturas"
         Index           =   2
      End
      Begin VB.Menu opFacturas 
         Caption         =   "Impresión de Facturas"
         Index           =   3
      End
   End
   Begin VB.Menu menuOfertas 
      Caption         =   "Ofertas"
   End
   Begin VB.Menu menuAlmacen 
      Caption         =   "Almacen"
      Begin VB.Menu opalmacen 
         Caption         =   "Listado Artículos"
         Index           =   0
      End
      Begin VB.Menu opalmacen 
         Caption         =   "Tipos de Artículos"
         Index           =   1
      End
      Begin VB.Menu opalmacen 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu opalmacen 
         Caption         =   "Estadísticas de Artículos"
         Index           =   3
      End
      Begin VB.Menu opalmacen 
         Caption         =   "Estadísticas INE"
         Index           =   4
      End
   End
   Begin VB.Menu menuvehiculos 
      Caption         =   "Vehículos"
      Begin VB.Menu opVehiculos 
         Caption         =   "Tratamiento de Vehículos"
         Index           =   0
      End
      Begin VB.Menu opVehiculos 
         Caption         =   "Estadísticas"
         Index           =   1
      End
   End
   Begin VB.Menu menuCartera 
      Caption         =   "Cartera"
      Begin VB.Menu opCartera 
         Caption         =   "Efectos"
         Index           =   0
      End
      Begin VB.Menu opCartera 
         Caption         =   "Remesas"
         Index           =   1
      End
      Begin VB.Menu opCartera 
         Caption         =   "Descuentos"
         Index           =   2
      End
      Begin VB.Menu opCartera 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu opCartera 
         Caption         =   "Crear Efecto Manual"
         Index           =   4
      End
   End
   Begin VB.Menu menucontabilidad 
      Caption         =   "Contabilidad"
      Begin VB.Menu opcontabilidad 
         Caption         =   "Contabilizar Facturas"
         Index           =   0
      End
      Begin VB.Menu opcontabilidad 
         Caption         =   "Contabilizar Descuentos"
         Index           =   1
      End
      Begin VB.Menu opcontabilidad 
         Caption         =   "Contabilizar Descuentos (Ult. Apunte)"
         Index           =   2
      End
      Begin VB.Menu opcontabilidad 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu opcontabilidad 
         Caption         =   "Abrir Base de Datos de Contabilidad"
         Index           =   4
      End
      Begin VB.Menu opcontabilidad 
         Caption         =   "Movimientos Contables"
         Index           =   5
      End
      Begin VB.Menu opcontabilidad 
         Caption         =   "-"
         Index           =   6
      End
      Begin VB.Menu opcontabilidad 
         Caption         =   "Abrir BM Conta"
         Index           =   7
      End
   End
   Begin VB.Menu menuVentanas 
      Caption         =   "Ventanas"
      WindowList      =   -1  'True
   End
   Begin VB.Menu menuSistema 
      Caption         =   "Sistema"
      Begin VB.Menu opSistema 
         Caption         =   "Usuarios"
         Index           =   0
      End
      Begin VB.Menu opSistema 
         Caption         =   "Salir a windows"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub ButtonBar1_ItemClick(ByVal Item As DevPowerButtonBar.Item)
    On Error GoTo ButtonBar1_ItemClick_Error
    Select Case Item.Parent.Index
        Case 1 ' Gestion
            opGestion_Click (Item.Index - 1)
        Case 2 ' Documentos
            opdocumentos_Click (Item.Index - 1)
        Case 3  ' Almacen
            opalmacen_Click (Item.Index - 1)
        Case 4 ' Sistema
            opSistema_Click (Item.Index - 1)
    End Select
    On Error GoTo 0
    Exit Sub
ButtonBar1_ItemClick_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure ButtonBar1_ItemClick of Formulario frmMenu"
End Sub

Private Sub MDIForm_Load()
    log (Me.Name)
    Set ButtonBar1.Icons = ImageList1
    Set ButtonBar1.SmallIcons = ImageList1
    StatusBar1.Panels(1) = "Gestión Comercial. (USUARIO : " & USUARIO.getNOMBRE & ")"
    StatusBar1.Panels(2) = "Fecha : " & Format(Date, "dd/mm/yyyy")
    StatusBar1.Panels(3) = "Server: " & ip
    StatusBar1.Panels(4) = gbd
    On Error Resume Next
    Set Me.Picture = LoadPicture(ReadINI(App.Path & "\config.ini", "logo", "fondo"))
    Dim op1 As New frmErrores
    op1.Show
    op1.WindowState = 1
    Set op1 = Nothing
    permisos
    If StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
        Me.BackColor = &H808080
    Else
        If ReadINI(App.Path & "\config.ini", "server", "bd") = StatusBar1.Panels(4) Then
            If ReadINI(App.Path & "\config.ini", "parametros", "Logo") = 1 Then
                frmLogo.Show
            End If
        Else
            frmLogoPrueba.Show
        End If
    End If
    
    If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
        menuOfertas.Enabled = False
        menuCartera.Enabled = False
'        opcontabilidad(0).Enabled = False
        opcontabilidad(1).Enabled = False
        opcontabilidad(2).Enabled = False
'        menucontabilidad.Enabled = False
    End If
    
    Dim ruta As String
    Dim oP As New clsParametros
    oP.Carga ENUM_PARAMETROS.RUTA_BMCONTA, ""
    ruta = oP.getVALOR
    Set oP = Nothing
    
    If Dir(ruta & "\bmcontaw.exe") <> "" Then
        opcontabilidad.Item(7).Enabled = True
        BMContaInstalado = True
    Else
        opcontabilidad.Item(7).Enabled = False
        BMContaInstalado = False
    End If
    
'    frmReport.iniciar
'    frmReport.Show
'    Unload frmReport
'    Dim oC As New clsContabilidad
'    oC.Actualiza_Descuento 784
'    oC.Actualiza_Factura 1578
'    oC.Actualiza_Descuento_Final 1488
'    registrar_componentes_resto Me.hWnd
End Sub
Public Sub permisos()
'    If usuario.getPER_6 = 0 Then ' Recalculo
'        opGestion(8).Enabled = False
'    Else
'        opGestion(8).Enabled = True
'    End If
End Sub

Private Sub menuOfertas_Click()
    frmOfertas_Listado.Show

End Sub

'Private Sub menuOfertas_Click()
''    frmOfertas_Detalle.pk = 100
'End Sub

Private Sub opalmacen_Click(Index As Integer)
    Select Case Index
        Case 0
            frmListadoArticulos.Show
        Case 1
            frmArticulos_Tipos.Show
        Case 3
            frmArticulos_Estadistica.Show
        Case 4
            frmArticulos_Ine.Show
    End Select
        
End Sub

Private Sub opCartera_Click(Index As Integer)
    Select Case Index
        Case 0 'Efectos
            frmEfectos_Listado.Show
        Case 1 ' Remesas
            frmRemesas_Listado.Show
        Case 2 ' Descuentos
            frmDescuentos_Listado.Show
        Case 4 ' Crear efecto manual
            frmEfectos_Creacion.Show
    End Select
End Sub

Private Sub opcontabilidad_Click(Index As Integer)
     Select Case Index
        Case 0 ' Contabilizar facturas
            frmContabilidad_Facturas.Show
        Case 1 ' Contabilizar Descuentos
            frmContabilidad_Descuentos.Show
        Case 2 ' Contabilizar descuentos final
            frmContabilidad_DescuentosFinal.Show
        Case 4  ' Abrir BD
            Dim Ret As Boolean
            If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                Ret = Mapear_Unidad_De_Red("W:", "\\" & IP_RESPALDO & "\lorenzo", "BCA", "BCA")
                  
                If Ret Then
                   MsgBox " Unidad de red mapeada correctamente. ", vbExclamation, " Mapear unidad de red"
                End If
            End If
            
            Dim oP As New clsParametros
            If oP.Carga(ENUM_PARAMETROS.BD_CONTABILIDAD, "") Then
                Dim iret As Long
                Dim s As String
                s = Dir(oP.getVALOR)
                If Trim(s) = "" Then
                    MsgBox "No existe la base de datos de contabilidad, configure correctamente el parámetro.", vbInformation, App.Title
                    Exit Sub
                Else
                    iret = ShellExecute(Me.Hwnd, vbNullString, oP.getVALOR, vbNullString, vbNullString, 1)
                End If
            End If
            Set oP = Nothing
            
            
            If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                Ret = Remover_Unidad_De_Red("W:")
            End If
'            If ret Then MsgBox " Unidad removida", vbInformation
        Case 5 ' Movimientos
            frmContabilidad_Movimientos.Show
        Case 7
                Dim ruta As String
                Set oP = New clsParametros
                oP.Carga ENUM_PARAMETROS.RUTA_BMCONTA, ""
                ruta = oP.getVALOR
                Set oP = Nothing

                iret = ShellExecute(Me.Hwnd, vbNullString, ruta & "\bmcontaw.exe", vbNullString, vbNullString, 1)
    End Select
End Sub

Private Sub opdocumentos_Click(Index As Integer)
    gDocumento = 0
    Select Case Index
       Case 0 ' Nuevo
            frmDocumento.PK_CLIENTE = 0
            frmDocumento.PK_DOCUMENTO = 0
            frmDocumento.Show 1
       Case 1 ' Buscar
            Dim oFB As New frmBuscarDocumento
            oFB.TIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.ALBARAN
            oFB.Show
            Set oFB = Nothing
       Case 3 ' Listado de albaranes
            frmAlbaranes_Listado.Show
       Case 4 ' Albaranes no valorados
            frmAlbaranesNoValorados.Show
       Case 5 ' Facturacion de albaranes
            frmFacturarAlbaranes.Show
'       Case 7
'            frmFacturas_Listado_Cobrar.Show
'       Case 9
'            frmFacturas_Impresion.Show
    End Select
End Sub

Private Sub opFacturas_Click(Index As Integer)
    Select Case Index
        Case 0
            Dim oFB As New frmBuscarDocumento
            oFB.TIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.factura
            oFB.Show
            Set oFB = Nothing
        Case 1
            frmFacturas_Listado.Show
        Case 2
            frmFacturas_Listado_Cobrar.Show
        Case 3
            frmFacturas_Impresion.Show
    End Select
End Sub

Private Sub opGestion_Click(Index As Integer)
    Select Case Index
        Case 0 ' Clientes
            frmClientes_Listado.Show
        Case 1 ' Obras
            frmObras_Listado.PK_CLIENTE = 0
            frmObras_Listado.Show
        Case 2 ' Proveedores
            frmProveedores_Listado.Show
        Case 3 ' Agenda
            frmListadoAgenda.Show
        Case 4 ' Formas de Pago
            frmFormas_Pago.Show 1
        Case 5 ' Provincias
            frmProvincias.Show 1
        Case 6 ' Agentes
            frmComerciales_Listado.Show
        Case 7 ' Vehiculos
            frmVehiculos_Listado.Show
        Case 8 ' Vehiculos
            frmTarifasPortes_Listado.Show
        Case 10 ' Conversion
            frmConversion.Show 1
    End Select
End Sub

Private Sub opSistema_Click(Index As Integer)
    Select Case Index
        Case 0 ' Usuarios
            frmListadoUsuarios.Show
        Case 1
            If MsgBox("¿Desea Cerrar " & App.Title & "?", vbCritical + vbOKCancel, App.Title) = vbOK Then
                End
            End If
    End Select
End Sub

Private Sub opVehiculos_Click(Index As Integer)
    Select Case Index
        Case 0 ' Listado
            frmVehiculos_Listado.Show
        Case 1 ' Estadísticas
            frmVehiculos_Estadistica.Show
    End Select
End Sub
