VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmSC_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Subcontratación de ensayos - Listado"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14730
   Icon            =   "frmSC_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   14730
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exportar"
      Height          =   960
      Left            =   10170
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Exportar datos a impresora o excel"
      Top             =   8100
      Width           =   1140
   End
   Begin VB.CommandButton cmdRefrescar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recargar lista"
      Height          =   960
      Left            =   11340
      Style           =   1  'Graphical
      TabIndex        =   35
      ToolTipText     =   "Salir"
      Top             =   8100
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSubtipos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mnto. Subtipos"
      Height          =   960
      Left            =   8865
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8100
      Width           =   1275
   End
   Begin VB.CommandButton cmdMail 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Correo al proveedor"
      Height          =   960
      Left            =   7605
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "crear etiqueta para envío de paquete"
      Top             =   8100
      Width           =   1215
   End
   Begin VB.CommandButton cmdTramitar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tramitar"
      Height          =   960
      Left            =   6345
      Picture         =   "frmSC_Listado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "crear etiqueta para envío de paquete"
      Top             =   8100
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
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
      Height          =   1725
      Left            =   0
      TabIndex        =   14
      Top             =   315
      Width           =   14685
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   1050
         Left            =   13320
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   360
         Width           =   1275
      End
      Begin VB.ComboBox cmbTipo 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   900
         Width           =   1905
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   510
         Left            =   8010
         TabIndex        =   21
         Top             =   855
         Width           =   5190
         Begin VB.Line Line6 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   3510
            X2              =   3600
            Y1              =   405
            Y2              =   315
         End
         Begin VB.Line Line5 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   3510
            X2              =   3600
            Y1              =   225
            Y2              =   315
         End
         Begin VB.Line Line4 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   3375
            X2              =   3600
            Y1              =   315
            Y2              =   315
         End
         Begin VB.Line Line3 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   1710
            X2              =   1800
            Y1              =   405
            Y2              =   315
         End
         Begin VB.Line Line2 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   1710
            X2              =   1800
            Y1              =   225
            Y2              =   315
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000C&
            BorderWidth     =   5
            X1              =   1575
            X2              =   1800
            Y1              =   315
            Y2              =   315
         End
         Begin VB.Label Label12 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tramitado"
            Height          =   195
            Left            =   2475
            TabIndex        =   27
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label11 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Recibido"
            Height          =   195
            Left            =   4275
            TabIndex        =   26
            Top             =   180
            Width           =   870
         End
         Begin VB.Label Label10 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pendiente"
            Height          =   195
            Left            =   630
            TabIndex        =   25
            Top             =   180
            Width           =   915
         End
         Begin VB.Label Label9 
            BackColor       =   &H0000C000&
            Height          =   240
            Left            =   3825
            TabIndex        =   24
            Top             =   180
            Width           =   285
         End
         Begin VB.Label Label8 
            BackColor       =   &H80000001&
            Height          =   240
            Left            =   2025
            TabIndex        =   23
            Top             =   180
            Width           =   285
         End
         Begin VB.Label Label7 
            BackColor       =   &H80000007&
            Height          =   240
            Left            =   225
            TabIndex        =   22
            Top             =   180
            Width           =   285
         End
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   990
         TabIndex        =   2
         Top             =   540
         Width           =   1905
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   990
         TabIndex        =   0
         Top             =   180
         Width           =   1905
      End
      Begin MSComCtl2.DTPicker datFechaDesde 
         Height          =   315
         Left            =   4185
         TabIndex        =   3
         Top             =   540
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   60293121
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker datFechaHasta 
         Height          =   315
         Left            =   6390
         TabIndex        =   4
         Top             =   540
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   60293121
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbSubcontratas 
         Height          =   330
         Left            =   4185
         TabIndex        =   1
         Top             =   180
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbSubtipo 
         Height          =   375
         Left            =   4185
         TabIndex        =   29
         Top             =   945
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbEstado 
         Height          =   375
         Left            =   9405
         TabIndex        =   34
         Top             =   540
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   661
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmSC_Listado.frx":1194
         Height          =   315
         Left            =   990
         TabIndex        =   36
         Top             =   1305
         Width           =   1875
         _ExtentX        =   3307
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
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   90
         TabIndex        =   37
         Top             =   1350
         Width           =   465
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subtipo"
         Height          =   240
         Left            =   3105
         TabIndex        =   31
         Top             =   945
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   30
         Top             =   945
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado:"
         Height          =   240
         Index           =   0
         Left            =   8640
         TabIndex        =   20
         Top             =   585
         Width           =   600
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   240
         Left            =   5670
         TabIndex        =   19
         Top             =   585
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   240
         Left            =   3105
         TabIndex        =   18
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   17
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo SC"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   16
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcontrata"
         Height          =   240
         Left            =   3105
         TabIndex        =   15
         Top             =   225
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdCrearDocumentos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Documento Solicitud"
      Height          =   960
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Crear documento de solicitud de análisis"
      Top             =   8100
      Width           =   1215
   End
   Begin VB.CommandButton CMDETIQUETA 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   960
      Left            =   5085
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "crear etiqueta para envío de paquete"
      Top             =   8100
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   960
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Modificar paquete seleccionado"
      Top             =   8100
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC - Salir"
      Height          =   960
      Left            =   13455
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   8100
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo"
      Height          =   960
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Crear nuevo paquete"
      Top             =   8100
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   960
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar paquete seleccionado"
      Top             =   8100
      Width           =   1215
   End
   Begin MSComctlLib.ListView lstPaquetes 
      Height          =   5985
      Left            =   0
      TabIndex        =   5
      Top             =   2070
      Width           =   14685
      _ExtentX        =   25903
      _ExtentY        =   10557
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   12640511
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12780
      Top             =   8325
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSC_Listado.frx":11DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSC_Listado.frx":1AB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSC_Listado.frx":238E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSC_Listado.frx":2C68
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSC_Listado.frx":3542
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSC_Listado.frx":3E1C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblsubtitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento subcontrataciones"
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
      Height          =   285
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   14685
   End
End
Attribute VB_Name = "frmSC_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------------------------------------------------------------------------------------------------
' M0957: Adición del ESTADO y la gestión de la lista por colores. Planteado sin parametrización (o al menos parametrización parcial e incompleta)
' M0959: Supresión del campo NFACTURA en el filtro de búsqueda
' M1171: Envío de correo a proveedores
' M1257: Botón para agregar facturas
' M1274: Generador de versiones (campo EDICIÓN). Cambios en las cargas de combos y código fuente para eliminar HARDCODE y sustituirlo por parametrización.
'        Nuevos IDs para los estados, reservando el 0 como comodín. DECODIFICADORA = 199.

Option Explicit
Private ANCHO_CAMPO As Integer
Public COL_EDICION  As Integer
Private Enum columnas
    SUBCONTRATA = 1
    PRESUPUESTO = 2
    FACTURADO = 3
    F_PETICION = 4
    USR_PETICION = 5
    TRAMITE = 6
    F_TRAMITE = 7
    USR_TRAMITE = 8
    ID = 9
    ID_CONTRATA = 10
    TIPO_ = 11
    TIPO_MUESTRA = 12
    subtipo = 13
    EDICION = 14
    ESTADO = 15
    Centro = 16
End Enum


Private Sub cmdImprimir_Click()
    If lstPaquetes.ListItems.Count > 0 Then
        generar_excel_listado
    Else
        MsgBox "Para exportar debe existir algún registro en la lista", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Public Sub cabecera()
    With lstPaquetes.ColumnHeaders
        .Add , , "Código SC", 1000, lvwColumnLeft
        .Add , , "Subcontrata", 4400 - (ANCHO_CAMPO / 2) - 300, lvwColumnLeft
        .Add , , "Presupuesto", (ANCHO_CAMPO / 2) - 299, lvwColumnCenter
        .Add , , "Facturado", (ANCHO_CAMPO / 2) - 299, lvwColumnCenter
        .Add , , "F. Petición", 1000, lvwColumnCenter
        .Add , , "Usr. Petición", 860, lvwColumnCenter
        .Add , , "Trámite", 1200, lvwColumnCenter
        .Add , , "F. Trámite", 1000, lvwColumnCenter
        .Add , , "Usr. Trámite", 860, lvwColumnCenter
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "ID_CONTRATA", 1, lvwColumnLeft
        .Add , , "TIPO", 1, lvwColumnLeft
        .Add , , "Tipo Muestra", 1650 - (ANCHO_CAMPO / 4), lvwColumnCenter
        .Add , , "Subtipo", 900, lvwColumnCenter
        .Add , , "Ed.", 500, lvwColumnCenter
        .Add , , "Estado", 0, lvwColumnCenter
        .Add , , "Centro", 900, lvwColumnCenter
    End With
End Sub
Private Sub generar_excel_listado()
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    
   On Error GoTo generar_excel_listado_Error

    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Add
    Set XLS = XLW.Worksheets(1)
    Me.MousePointer = 11
    XLW.Worksheets(3).Delete
    XLW.Worksheets(2).Delete
    XLW.Worksheets(1).Name = "Listado de SC"
    XLS.Range("1:1").HorizontalAlignment = xlCenter
    XLS.Range("1:1").VerticalAlignment = xlCenter
    XLS.Range("1:1").RowHeight = 30
    XLS.Range("1:1").WrapText = True
    'Cabecera
    XLS.Cells(1, 1) = "Código SC"
    XLS.Cells(1, 2) = "Subcontrata"
    XLS.Cells(1, 3) = "Presupuesto"
    XLS.Cells(1, 4) = "Facturado"
    XLS.Cells(1, 5) = "F.Petición"
    XLS.Cells(1, 6) = "Usr.Petición"
    XLS.Cells(1, 7) = "Trámite"
    XLS.Cells(1, 8) = "F.Trámite"
    XLS.Cells(1, 9) = "Usr.Trámite"
    XLS.Cells(1, 10) = "Tipo Muestra"
    XLS.Cells(1, 11) = "Subtipo"
    XLS.Cells(1, 12) = "Ed."
    XLS.Cells(1, 13) = "Centro"
    Dim i As Integer
    i = 2
    ' Datos
    For i = 1 To lstPaquetes.ListItems.Count
        XLS.Cells(i + 1, 1) = lstPaquetes.ListItems(i).Text
        XLS.Cells(i + 1, 2) = lstPaquetes.ListItems(i).SubItems(1)
        XLS.Cells(i + 1, 3) = lstPaquetes.ListItems(i).SubItems(2)
        XLS.Cells(i + 1, 4) = lstPaquetes.ListItems(i).SubItems(3) ' facturado
        XLS.Cells(i + 1, 5) = Format(lstPaquetes.ListItems(i).SubItems(4), "yyyy-mm-dd")
        XLS.Cells(i + 1, 6) = lstPaquetes.ListItems(i).SubItems(5)
        XLS.Cells(i + 1, 7) = lstPaquetes.ListItems(i).SubItems(6)
        XLS.Cells(i + 1, 8) = Format(lstPaquetes.ListItems(i).SubItems(7), "yyyy-mm-dd")
        XLS.Cells(i + 1, 9) = lstPaquetes.ListItems(i).SubItems(8)
        XLS.Cells(i + 1, 10) = lstPaquetes.ListItems(i).SubItems(12)
        XLS.Cells(i + 1, 11) = lstPaquetes.ListItems(i).SubItems(13)
        XLS.Cells(i + 1, 12) = lstPaquetes.ListItems(i).SubItems(14)
        XLS.Cells(i + 1, 13) = lstPaquetes.ListItems(i).SubItems(16) ' Centro
    Next
    For i = 1 To 13
        XLS.Columns(i).AutoFit
    Next
    XLS.Range("2:" & lstPaquetes.ListItems.Count + 1).HorizontalAlignment = xlLeft
    
    Me.MousePointer = 0
    XLA.visible = True
'    Set XLS = Nothing
'    Set XLW = Nothing
'    Set XLA = Nothing
   On Error GoTo 0
   Exit Sub

generar_excel_listado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_excel_listado of Formulario frmEquipoListado"
    
End Sub


Private Sub cmbCentro_Change()
    cargar_lista
End Sub
Private Sub cmbEstado_Change()
    Call cargar_lista
End Sub

Private Sub cmbSubtipo_Change()
    Call cargar_lista
End Sub

Private Sub cmbSubtipo_LostFocus()
    Call cargar_lista
End Sub

'Private Sub cmbTipo_Change()
'    Call cargar_lista
'End Sub

Private Sub cmbTipo_Click()
    cargar_lista
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

'Private Sub cmbTipo_LostFocus()
'    Call cargar_lista
'End Sub

Private Sub cmdetiqueta_Click()
    On Error GoTo fallo
    
'JGM    Dim generar As Boolean
    Dim strContratas As String
    
'JGM    generar = False

    log ("Comienzo impresion de etiquetas para envíos de paquetes")
    If lstPaquetes.ListItems.Count = 0 Then
        MsgBox "Seleccione alguna fila para generar la etiqueta.", vbOKOnly + vbInformation, App.Title
    Else
        strContratas = "{PROVEEDORES.ID_PROVEEDOR} = " & CLng(lstPaquetes.selectedItem.SubItems(columnas.ID_CONTRATA))
        frmReport.iniciar
        frmReport.informe = "\SC\rptSCEtiquetaCaja"
        frmReport.criterio = strContratas
        frmReport.imprimir = False
        frmReport.generar
        frmReport.visible = True
    End If
    frmReport.pdf = ""
    log ("Final impresion de etiquetas para envíos de paquetes")
    
    Exit Sub
fallo:
    MsgBox "Error al generar la etiquetas de los paquetes. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdmail_Click()
    envioCorreoProveedor
End Sub



Private Sub Form_Load()
    log (Me.Name)
    permisos
    COL_EDICION = columnas.EDICION
    Call cabecera
    Me.Left = 50
    Me.top = 50
    datFechaDesde = DateAdd("m", -1, Date)
    datFechaHasta = Date
    Call cargar_combos
    Call cargar_botones(Me)
    cargaEstados
    Label7.BackColor = SC_COLOR_PENDIENTE
    Label8.BackColor = SC_COLOR_TRAMITADO
    Label9.BackColor = SC_COLOR_RECIBIDO
    Call cargar_lista
    botonTramitar
End Sub
'M1257-I
Private Sub cmbEstado_Click()
    Call cargar_lista
End Sub

Private Sub cmdRefrescar_Click()
    Call cargar_lista
End Sub

Private Sub cmdSubtipos_Click()
    Dim oform As New frmDecodificadora
    oform.CODIGO = DECODIFICADORA.SC_SUBTIPOS
    oform.Show 0
    Set oform = Nothing
End Sub

Private Sub datFechaHasta_LostFocus()
     Call cargar_lista
End Sub
Private Sub cmdTramitar_Click()
    If lstPaquetes.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oPaquete As New clsSC_Paquetes
    Select Case CLng(lstPaquetes.selectedItem.SubItems(columnas.ESTADO))
    Case SC_ESTADO_PENDIENTE
         oPaquete.Tramitar CLng(lstPaquetes.selectedItem.SubItems(columnas.ID)), CLng(lstPaquetes.selectedItem.SubItems(columnas.EDICION))
         MsgBox "El pedido se ha tramitado.", vbOKOnly + vbInformation, App.Title
         Call cargar_lista
    Case SC_ESTADO_TRAMITADO
        oPaquete.finalizar CLng(lstPaquetes.selectedItem.SubItems(columnas.ID)), CLng(lstPaquetes.selectedItem.SubItems(columnas.EDICION))
        MsgBox "El paquete se ha marcado como recibido.", vbOKOnly + vbInformation, App.Title
        Call cargar_lista
    End Select
End Sub
'M1171-I
Private Sub envioCorreoProveedor()
'------------------- Correo tras tramitación -----------------------------'
   
    Dim oPaquete As New clsSC_Paquetes
    Dim oProveedor As New clsProveedor
    Dim mail As String
    Dim ASUNTO As String
    Dim texto As String
    Dim pdf As String
    Dim code As String
    
    On Error Resume Next
    MkDir App.Path & "\tmp"
    
    code = Replace(lstPaquetes.selectedItem.Text, "/", "_")
    pdf = App.Path & "\tmp\Pedido SC Nº" & code & ".pdf"
   
    On Error Resume Next
    Kill pdf
    
    On Error GoTo errorCorreo
    oPaquete.imprimir CLng(lstPaquetes.selectedItem.SubItems(columnas.ID)), CLng(lstPaquetes.selectedItem.SubItems(columnas.EDICION)), pdf
    If oProveedor.Carga(CLng(lstPaquetes.selectedItem.SubItems(columnas.ID_CONTRATA))) Then
         mail = oProveedor.getEMAIL
    End If

    ASUNTO = "Pedido con código: SC " & lstPaquetes.selectedItem.Text
    texto = "Solicitamos el pedido del documento pdf adjunto."
    
    Call enviar_correo(mail, "", "", True, texto, ASUNTO, pdf)
    Exit Sub
errorCorreo:
    Me.MousePointer = vbNormal
    MsgBox "Error al generar / enviar el correo de pedido " & Err.Description, vbCritical, App.Title
End Sub
'M1171-F

Private Sub datFechaDesde_Change()
 '   Call cargar_lista
End Sub


Private Sub datFechaHasta_Change()
'    Call cargar_lista
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    Call cargar_lista
End Sub

Private Sub cmbSubcontratas_Change()
    Call cargar_lista
End Sub

Private Sub txtfiltro_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"): ' no se permite introducir comillas simples
            KeyAscii = 0
    End Select
End Sub
' -------------------

' lista
Private Sub lstPaquetes_Click()
    Dim str As String
    If lstPaquetes.ListItems.Count = 0 Then
        Exit Sub
    End If
    botonTramitar
    'M1274-I
    restoBotones
    'M1274-F
End Sub

Private Sub lstPaquetes_DblClick()
    If cmdModificar.Enabled = True Then
       cmdModificar_Click
       botonTramitar
    End If
End Sub

Private Sub lstPaquetes_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lstPaquetes.ListItems.Count > 0 Then
     lstPaquetes.SortKey = ColumnHeader.Index - 1
     If lstPaquetes.SortOrder = 0 Then
        lstPaquetes.SortOrder = 1
     Else
        lstPaquetes.SortOrder = 0
     End If
     lstPaquetes.Sorted = True
   End If
End Sub
' -------------------

' botones
Private Sub cmdAnadir_Click()
'    frmSC_Muestras_NoEnviadas_listado.Show 1
     frmSC_Menu.Show 1
     cargar_lista
End Sub

Private Sub cmdModificar_Click()
    If lstPaquetes.ListItems.Count > 0 Then
        Me.MousePointer = vbHourglass
        Select Case lstPaquetes.ListItems(lstPaquetes.selectedItem.Index).SubItems(columnas.TIPO_)
        Case TOBJETO_SC_DETERMINACIONES
            frmSC_Paquete_Detalle.PK = lstPaquetes.selectedItem.SubItems(columnas.ID)
            frmSC_Paquete_Detalle.EDICION = lstPaquetes.selectedItem.SubItems(columnas.EDICION)
            frmSC_Paquete_Detalle.Show 1
        Case TOBJETO_SC_EFICACIA
            frmSC_Paquete_Detalle_CE.PK = lstPaquetes.selectedItem.SubItems(columnas.ID)
            frmSC_Paquete_Detalle_CE.EDICION = lstPaquetes.selectedItem.SubItems(columnas.EDICION)
            frmSC_Paquete_Detalle_CE.Show 1
        Case TOBJETO_SC_GENERICA, TOBJETO_SC_PEACH
            frmSC_Paquete_Detalle_Generico.PK = lstPaquetes.selectedItem.SubItems(columnas.ID)
            frmSC_Paquete_Detalle_Generico.EDICION = lstPaquetes.selectedItem.SubItems(columnas.EDICION)
            frmSC_Paquete_Detalle_Generico.Show 1
        End Select
        actualizar_lista lstPaquetes.selectedItem.SubItems(columnas.ID), lstPaquetes.selectedItem.SubItems(columnas.EDICION)
        Me.MousePointer = vbNormal
    Else
        MsgBox "Debe seleccionar el paquete que desea modificar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdEliminar_Click()
    If Not (lstPaquetes.selectedItem Is Nothing) Then
        If MsgBox("Se va a eliminar el paquete con código SC: " & lstPaquetes.selectedItem & vbCrLf & _
                  "¿Está seguro?", vbYesNo + vbInformation, App.Title) = vbYes Then
            Dim oSCPaquete As New clsSC_Paquetes
            Me.MousePointer = vbHourglass
            
'M1147-I
'            If lstPaquetes.selectedItem.SubItems(10) < 2 Then
'M1147-F
            If oSCPaquete.Eliminar(lstPaquetes.selectedItem.SubItems(columnas.ID), lstPaquetes.selectedItem.SubItems(columnas.EDICION)) Then
                MsgBox "El paquete se ha eliminado correctamente.", vbOKOnly + vbInformation, App.Title
            End If
'M1147-I
'            Else
'                If oSCPaquete.EliminarGenerico(lstPaquetes.selectedItem.SubItems(8)) Then
'                    MsgBox "El paquete se ha eliminado correctamente.", vbOKOnly + vbInformation, App.Title
'                End If
'            End If
'M1147-F
            Call cargar_lista
            Me.MousePointer = vbNormal
            Set oSCPaquete = Nothing
        End If
    Else
        MsgBox "Debe seleccionar el paquete que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdCrearDocumentos_Click()
    On Error GoTo fallo
    
    If lstPaquetes.selectedItem.SubItems(columnas.TRAMITE) = SC_TRAMITADO Or lstPaquetes.selectedItem.SubItems(columnas.TRAMITE) = SC_RECIBIDO Then
    Dim ID_PAQUETE As Long
    Dim ID_EDICION As Long
    
    Me.MousePointer = vbHourglass
    log ("Comienzo impresion de documentos sc para envíos de paquetes")
    If lstPaquetes.ListItems.Count = 0 Then
        MsgBox "Seleccione algún paquete para generar el documento de solicitud de análisis.", vbOKOnly + vbInformation, App.Title
    Else
    'M1147-I
        Me.MousePointer = vbHourglass
        ID_PAQUETE = CLng((lstPaquetes.selectedItem.SubItems(columnas.ID)))
        ID_EDICION = CLng((lstPaquetes.selectedItem.SubItems(columnas.EDICION)))
        ' JGM-I : Cambio codigo duplicado
        With frmReport
            .iniciar
            Select Case lstPaquetes.ListItems(lstPaquetes.selectedItem.Index).SubItems(columnas.TIPO_)
            Case TOBJETO_SC_DETERMINACIONES
                .informe = "\SC\rptSCPaquetes_Solicitud_Analisis"
            Case TOBJETO_SC_EFICACIA
                .informe = "\SC\rptSCPaquetes_Solicitud_Analisis_CE"
            Case TOBJETO_SC_GENERICA, TOBJETO_SC_PEACH
                .informe = "\SC\rptSCPaquetes_Solicitud_Analisis_GEN"
            End Select
            .criterio = "{sc_paquetes.ID_PAQUETE}=" & ID_PAQUETE & " and {sc_paquetes.EDICION}=" & ID_EDICION & " and {fp.CODIGO} = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & " and {vencimiento.CODIGO} = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_VENCIMIENTOS
            .imprimir = False
            .generar
            .visible = True
        End With
        ' JGM-F : Cambio codigo duplicado
        Me.MousePointer = vbNormal
    'M1147-F
    End If
    frmReport.pdf = ""
    Me.MousePointer = vbNormal
    log ("Final impresion de documentos sc para envíos de paquetes")
    Else
        MsgBox "El paquete aún no ha sido aprobado para su trámite", vbOKOnly + vbInformation, App.Title
    End If
    Exit Sub
    
fallo:
    Me.MousePointer = vbNormal
    MsgBox "Error al generar los documentos de solicitud de análisis. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
' -------------------

' ----------------- Funciones auxiliares del formulario ----------------
Public Sub actualizar_lista(ID As Long, EDICION As Long)
    Dim rs As ADODB.Recordset
    Dim oSC_Paquete As New clsSC_Paquetes
    Dim color As String
    Dim ESTADO As String
    Dim indice As Integer
    Dim diff As Currency
    
    Set rs = oSC_Paquete.recuperarModificado(ID, EDICION)
    If rs.RecordCount <> 0 Then
        If CInt(rs(10)) <> TOBJETO_SC_GENERICA Or (CInt(rs(10)) = TOBJETO_SC_GENERICA And USUARIO.getPER_SCG = True) Then
            With lstPaquetes.selectedItem
                Dim oDeco As New clsDecodificadora
                color = rs(14)
                ESTADO = rs(15)
                .SubItems(columnas.SUBCONTRATA) = rs(1)
                .SubItems(columnas.PRESUPUESTO) = Replace(moneda(rs(2)), "€", rs(19))
                .SubItems(columnas.F_PETICION) = rs(3)
                .SubItems(columnas.USR_PETICION) = rs(4)
                .SubItems(columnas.TRAMITE) = ESTADO
                If Not IsNull(rs(8)) Then
                  .SubItems(columnas.F_TRAMITE) = rs(8)
                End If
                If rs(9) <> 0 Then
                  .SubItems(columnas.USR_TRAMITE) = rs(9)
                Else
                  .SubItems(columnas.USR_TRAMITE) = "--"
                End If
                .SubItems(columnas.ID) = rs(5)
                .SubItems(columnas.ID_CONTRATA) = rs(6)
                .SubItems(columnas.TIPO_) = rs(10)
                .SubItems(columnas.TIPO_MUESTRA) = " -- "
                diff = 0
                If IsNumeric(rs(2)) And IsNumeric(rs(12)) Then
                    diff = rs(2) - rs(12)
                    .SubItems(columnas.FACTURADO) = moneda(rs(12))
                Else
                    diff = 1
                    .SubItems(columnas.FACTURADO) = moneda(0)
                End If
                If diff < 0 Then
                    .SmallIcon = 3
                    .bold = True
                Else
                    If diff = 0 Then
                        .SmallIcon = 4
                    End If
                End If

                If Not IsNull(rs(16)) Then
                    .SubItems(columnas.TIPO_MUESTRA) = rs(16)
                Else
                    .SubItems(columnas.TIPO_MUESTRA) = "--"
                End If

                If Not IsNull(rs(17)) Then
                    .SubItems(columnas.subtipo) = rs(17)
                Else
                    .SubItems(columnas.subtipo) = "--"
                End If

                If IsNumeric(rs(13)) Then
                    .SubItems(columnas.EDICION) = rs(13)
                Else
                    .SubItems(columnas.EDICION) = 1
                End If
                .SubItems(columnas.ESTADO) = rs(7)
                If Not IsNull(rs(18)) Then
                    .SubItems(columnas.Centro) = rs(18)
                End If
                For indice = 1 To lstPaquetes.ColumnHeaders.Count - 1
                    If Trim(color) <> "" Then
                        .ListSubItems(indice).ForeColor = color
                    End If
                Next indice
            End With
        End If
    End If
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oSC_Paquete As New clsSC_Paquetes
     Dim color As String
    Dim ESTADO As String
    Dim indice As Integer
    Dim diff As Currency
    lstPaquetes.ListItems.Clear
    Dim tipo As Long
    Dim subtipo As Long
    If cmbTipo.ListIndex >= 0 Then
       tipo = CLng(cmbTipo.ItemData(cmbTipo.ListIndex))
    Else
       tipo = 0
    End If
    subtipo = CLng(cmbSubtipo.getPK_SALIDA)
    Dim Centro As Integer
    If cmbCentro.BoundText = "" Then
        Centro = 0
    Else
        Centro = cmbCentro.BoundText
    End If
    Set rs = oSC_Paquete.Listado(txtFiltro(1), txtFiltro(2), cmbSubcontratas.getTEXTO, Format(datFechaDesde, "yyyy-mm-dd"), Format(datFechaHasta, "yyyy-mm-dd"), 0, Format(Date, "yyyy-mm-dd"), Format(Date, "yyyy-mm-dd"), "", cmbEstado.getPK_SALIDA, tipo, subtipo, Centro)
    If rs.RecordCount <> 0 Then
        Do
            If CInt(rs(10)) <> TOBJETO_SC_GENERICA Or (CInt(rs(10)) = TOBJETO_SC_GENERICA And USUARIO.getPER_SCG = True) Then
                With lstPaquetes.ListItems.Add(, , rs(0))
                    Dim oDeco As New clsDecodificadora
                    color = rs(14)
                    ESTADO = rs(15)
                    .SubItems(columnas.SUBCONTRATA) = rs(1)
                    .SubItems(columnas.PRESUPUESTO) = Replace(moneda(rs(2)), "€", rs(19))
                    .SubItems(columnas.F_PETICION) = rs(3)
                    .SubItems(columnas.USR_PETICION) = rs(4)
                    .SubItems(columnas.TRAMITE) = ESTADO
                    If Not IsNull(rs(8)) Then
                      .SubItems(columnas.F_TRAMITE) = rs(8)
                    End If
                    If rs(9) <> 0 Then
                      .SubItems(columnas.USR_TRAMITE) = rs(9)
                    Else
                      .SubItems(columnas.USR_TRAMITE) = "--"
                    End If
                    .SubItems(columnas.ID) = rs(5)
                    .SubItems(columnas.ID_CONTRATA) = rs(6)
                    .SubItems(columnas.TIPO_) = rs(10)
                    .SubItems(columnas.TIPO_MUESTRA) = " -- "
                    diff = 0
                    If IsNumeric(rs(2)) And IsNumeric(rs(12)) Then
                        diff = rs(2) - rs(12)
                        .SubItems(columnas.FACTURADO) = moneda(rs(12))
                    Else
                        diff = 1 'Si dejáramos un cero aquí se marcaría en verde la fila
                                 'Tampoco es válido RS(2) porque tendríamos anomalías con presupuestos no numéricos
                        .SubItems(columnas.FACTURADO) = moneda(0)
                    End If
                    If diff < 0 Then
                        .SmallIcon = 3
                        .bold = True
                    Else
                        If diff = 0 Then
                            .SmallIcon = 4
                        End If
                    End If
                    If Not IsNull(rs(16)) Then
                        .SubItems(columnas.TIPO_MUESTRA) = rs(16)
                    Else
                        .SubItems(columnas.TIPO_MUESTRA) = "--"
                    End If
                    If Not IsNull(rs(17)) Then
                        .SubItems(columnas.subtipo) = rs(17)
                    Else
                        .SubItems(columnas.subtipo) = "--"
                    End If
                    If IsNumeric(rs(13)) Then
                        .SubItems(columnas.EDICION) = rs(13)
                    Else
                        .SubItems(columnas.EDICION) = 1
                    End If
                    .SubItems(columnas.ESTADO) = rs(7)
                    If Not IsNull(rs(18)) Then
                        .SubItems(columnas.Centro) = rs(18) ' CENTRO
                    Else
                        .SubItems(columnas.Centro) = ""
                    End If
                    For indice = 1 To lstPaquetes.ColumnHeaders.Count - 1
                        If Trim(color) <> "" Then
                            .ListSubItems(indice).ForeColor = color
                        End If
                    Next indice
                End With
            End If
            rs.MoveNext
        Loop Until rs.EOF
        lstPaquetes_Click
    End If
    
    lblsubtitulo = "Número de paquetes mostrados : " & rs.RecordCount
End Sub

Public Function alguno_seleccionado() As Boolean
    Dim booAlgunoSeleccionado As Boolean
    Dim i As Long
    
    alguno_seleccionado = True
    
    booAlgunoSeleccionado = False
    For i = 1 To lstPaquetes.ListItems.Count
        If lstPaquetes.ListItems(i).Checked = True Then
            booAlgunoSeleccionado = True
        End If
    Next i
    If Not booAlgunoSeleccionado Then
        alguno_seleccionado = False
        MsgBox "Debe seleccionar al menos un paquete.", vbOKOnly + vbInformation, App.Title
        Exit Function
    End If
End Function

Private Sub cargar_combos()
    llenar_combo cmbSubcontratas, New clsProveedor, 0, frmProveedores_Detalle, " ES_SUBCONTRATA = 1 "
'M1257-I
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbSubtipo, DECODIFICADORA.SC_SUBTIPOS
    Set oDeco = Nothing
    cmbTipo.Clear
    cmbTipo.AddItem "TODOS"
    cmbTipo.ItemData(cmbTipo.NewIndex) = 0
    cmbTipo.AddItem "DETERMINACIONES"
    cmbTipo.ItemData(cmbTipo.NewIndex) = TOBJETO.TOBJETO_SC_DETERMINACIONES
    cmbTipo.AddItem "C.EFICACIA"
    cmbTipo.ItemData(cmbTipo.NewIndex) = TOBJETO.TOBJETO_SC_EFICACIA
    cmbTipo.AddItem "GENÉRICO"
    cmbTipo.ItemData(cmbTipo.NewIndex) = TOBJETO.TOBJETO_SC_GENERICA
    cmbTipo.AddItem "PEACH"
    cmbTipo.ItemData(cmbTipo.NewIndex) = TOBJETO.TOBJETO_SC_PEACH
'M1257-F
    cargar_combo cmbCentro, New clsCentros
End Sub

'M0957-I
Private Sub permisos()
    ' Permiso tramitación
    If Not USUARIO.getPER_TRAMITACION_CONTRATA Then
        cmdTramitar.Enabled = False
        cmdMail.Enabled = False
    Else
        cmdTramitar.Enabled = True
        cmdMail.Enabled = True
    End If
    
    If Not USUARIO.getPER_FACTURACION Then
        ANCHO_CAMPO = 600
    Else
        ANCHO_CAMPO = 2800
    End If
End Sub

Private Sub cargaEstados()
    cmbEstado.limpiar
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbEstado, DECODIFICADORA.SC_ESTADOS
    Set oDeco = Nothing
End Sub
Private Sub botonTramitar()
    'JGM-I
    If Not USUARIO.getPER_TRAMITACION_CONTRATA Then
        cmdTramitar.Enabled = False
        Exit Sub
    End If
    'JGM-F
    cmdTramitar.Enabled = True
    If lstPaquetes.ListItems.Count = 0 Then Exit Sub
    If Trim(lstPaquetes.selectedItem.SubItems(columnas.TRAMITE)) = "" Then
        Exit Sub
    End If
    
    Select Case Trim(lstPaquetes.selectedItem.SubItems(columnas.TRAMITE))
    Case SC_PENDIENTE
        cmdTramitar.Caption = "Tramitar"
    Case SC_TRAMITADO
        cmdTramitar.Caption = "Recibir"
    Case SC_RECIBIDO
        cmdTramitar.Enabled = False
    Case SC_HISTORICO
        cmdTramitar.Enabled = False
    End Select
End Sub
'M0957-F
'M1274-I
Private Sub restoBotones()
    Select Case lstPaquetes.selectedItem.SubItems(columnas.ESTADO)
    Case SC_ESTADO_PENDIENTE
        cmdModificar.Caption = "Modificar"
        cmdModificar.Enabled = True
        cmdTramitar.Enabled = True
        cmdEliminar.Enabled = True
        cmdCrearDocumentos.Enabled = False
        cmdEtiqueta.Enabled = False
        cmdMail.Enabled = False
    Case SC_ESTADO_TRAMITADO
        cmdModificar.Caption = "Modificar"
        cmdModificar.Enabled = True
        cmdTramitar.Enabled = True
        cmdEliminar.Enabled = True
        cmdCrearDocumentos.Enabled = True
        cmdEtiqueta.Enabled = True
        cmdMail.Enabled = True
    Case SC_ESTADO_RECIBIDO
        cmdModificar.Caption = "Consultar"
        cmdModificar.Enabled = True
        cmdTramitar.Enabled = False
        cmdEliminar.Enabled = True
        cmdCrearDocumentos.Enabled = True
        cmdEtiqueta.Enabled = True
        cmdMail.Enabled = True
    Case SC_ESTADO_HISTORICO
        cmdModificar.Caption = "Consultar"
        cmdModificar.Enabled = True
        cmdTramitar.Enabled = False
        cmdEliminar.Enabled = False
        cmdCrearDocumentos.Enabled = False
        cmdEtiqueta.Enabled = False
        cmdMail.Enabled = False
    End Select
End Sub
'M1274-F
