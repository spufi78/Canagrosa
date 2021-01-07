VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmProveedores_Facturas_Listado_Modal 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Búsqueda de Facturas de Proveedores"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15465
   Icon            =   "frmProveedores_Facturas_Listado_Modal.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   15465
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   13275
      Style           =   1  'Graphical
      TabIndex        =   43
      Top             =   8010
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   8055
      Width           =   1050
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   4275
      TabIndex        =   29
      Top             =   4320
      Visible         =   0   'False
      Width           =   6450
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Generando documento EXCEL. Por favor, espere."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   330
         Index           =   1
         Left            =   540
         TabIndex        =   30
         Top             =   270
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Pdtes. Pago"
      Height          =   870
      Left            =   4320
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   8055
      Visible         =   0   'False
      Width           =   1050
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
      Height          =   1770
      Left            =   45
      TabIndex        =   12
      Top             =   360
      Width           =   15360
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   780
         Left            =   14220
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   540
         Width           =   1050
      End
      Begin VB.TextBox txtimportehasta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   3030
         TabIndex        =   39
         Top             =   990
         Width           =   1320
      End
      Begin VB.TextBox txtImporteDesde 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         TabIndex        =   37
         Top             =   990
         Width           =   1320
      End
      Begin VB.CheckBox chkVencidas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar las vencidas"
         Height          =   240
         Left            =   4500
         TabIndex        =   36
         Top             =   1260
         Width           =   1905
      End
      Begin VB.CheckBox chkPagoPrevisto 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Pago Previsto"
         Height          =   240
         Left            =   4500
         TabIndex        =   35
         Top             =   1485
         Width           =   1995
      End
      Begin VB.CheckBox chkIncidencias 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar INCIDENCIAS"
         Height          =   240
         Left            =   4500
         TabIndex        =   34
         Top             =   1035
         Width           =   2040
      End
      Begin VB.TextBox txtconcepto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1050
         TabIndex        =   32
         Top             =   1350
         Width           =   3300
      End
      Begin VB.CheckBox chkNoEnviadas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo no enviadas"
         Height          =   240
         Left            =   4500
         TabIndex        =   31
         Top             =   810
         Width           =   2220
      End
      Begin VB.CheckBox chkPendientesPago 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo pdtes. pago"
         Height          =   285
         Left            =   4500
         TabIndex        =   5
         Top             =   540
         Value           =   1  'Checked
         Width           =   2130
      End
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   345
         Left            =   1050
         TabIndex        =   0
         Top             =   225
         Width           =   13080
         _ExtentX        =   23072
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbFamilia 
         Height          =   345
         Left            =   8235
         TabIndex        =   1
         Top             =   585
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbGasto 
         Height          =   345
         Left            =   8235
         TabIndex        =   2
         Top             =   945
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbPago 
         Height          =   345
         Left            =   8235
         TabIndex        =   6
         Top             =   1305
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1050
         TabIndex        =   3
         Top             =   585
         Width           =   1365
         _ExtentX        =   2408
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
         Format          =   51642369
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3030
         TabIndex        =   4
         Top             =   585
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   51642369
         CurrentDate     =   38002
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   195
         Index           =   3
         Left            =   2655
         TabIndex        =   40
         Top             =   1035
         Width           =   345
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   38
         Top             =   1035
         Width           =   750
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concepto"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   33
         Top             =   1395
         Width           =   750
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Factura"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   27
         Top             =   675
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   2520
         TabIndex        =   26
         Top             =   675
         Width           =   405
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcuenta Pago"
         Height          =   195
         Index           =   2
         Left            =   6840
         TabIndex        =   17
         Top             =   1350
         Width           =   1200
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcuenta Gasto"
         Height          =   195
         Index           =   1
         Left            =   6840
         TabIndex        =   16
         Top             =   990
         Width           =   1245
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   0
         Left            =   6840
         TabIndex        =   15
         Top             =   630
         Width           =   480
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   14
         Top             =   270
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2175
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1130
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   14355
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8010
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5730
      Left            =   45
      TabIndex        =   7
      Top             =   2160
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   10107
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
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
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8055
      Width           =   1050
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Index           =   2
      Left            =   9765
      TabIndex        =   25
      Top             =   8730
      Width           =   825
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10980
      TabIndex        =   24
      Top             =   8730
      Width           =   2085
   End
   Begin VB.Label lblIVA 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10980
      TabIndex        =   23
      Top             =   8190
      Width           =   2085
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "IVA"
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
      Index           =   0
      Left            =   9765
      TabIndex        =   22
      Top             =   8190
      Width           =   645
   End
   Begin VB.Label lblBase 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10980
      TabIndex        =   21
      Top             =   7920
      Width           =   2085
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Base"
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
      Index           =   1
      Left            =   9765
      TabIndex        =   20
      Top             =   7920
      Width           =   870
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Retención"
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
      Index           =   3
      Left            =   9765
      TabIndex        =   19
      Top             =   8460
      Width           =   1275
   End
   Begin VB.Label lblRetencion 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   10980
      TabIndex        =   18
      Top             =   8460
      Width           =   2085
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "BUSQUEDA DE FACTURAS DE PROVEEDORES"
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
      Left            =   45
      TabIndex        =   13
      Top             =   45
      Width           =   15390
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   330
      Left            =   0
      Top             =   0
      Width           =   15705
   End
End
Attribute VB_Name = "frmProveedores_Facturas_Listado_Modal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LISTA_SEL As String
Private Enum COLS
    C_ID = 0
    C_PROVEEDOR = 1
    C_fecha = 2
    C_concepto = 3
    C_NUMERO = 4
    C_FAMILIA = 5
    C_SUBCUENTA = 6
    C_BASE = 7
    C_IVA_PORCENTAJE = 8
    C_IVA = 9
    C_RETENCION = 10
    C_total = 11
    C_FP = 12
    C_vencimiento = 13
    C_PAGO = 14
    C_TOBJETO = 15
    C_cOBJETO = 16
    C_IDPROVEEDOR = 17
'M1335-I
    C_CUENTA = 18
'M1335-F
    C_ENVIADA = 19
End Enum


Private Sub cmdAnadir_Click()
    With frmProveedores_Facturas
        .PK = 0
        .PK_FACTURA_ID = 0
        .TOBJETO = 0
        .COBJETO = 0
        .Show 1
    End With
    cargar_lista
End Sub

Private Sub chkIncidencias_Click()
    cargar_lista
End Sub

Private Sub chkNoEnviadas_Click()
    cargar_lista
End Sub

Private Sub chkPagoPrevisto_Click()
    cargar_lista
End Sub

Private Sub chkPendientesPago_Click()
    cargar_lista
End Sub

Private Sub chkVencidas_Click()
    cargar_lista
End Sub

Private Sub cmbfamilia_Change()
cargar_lista
End Sub

Private Sub cmbGasto_change()
cargar_lista
End Sub

Private Sub cmbPago_change()
cargar_lista
End Sub

Private Sub cmbProveedor_change()
    cargar_lista
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    LISTA_SEL = ""
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR la factura de proveedor " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oPF As New clsProveedores_Facturas
        If oPF.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
            MsgBox "Factura eliminada correctamente.", vbInformation, App.Title
            cargar_lista
        End If
        Set oPF = Nothing
    End If
End Sub

Private Sub cmdImprimir_Click()
       Me.MousePointer = vbHourglass
       Frame3.Visible = True
       Dim rs As New ADODB.Recordset
       Dim fecha As String
      
       rs.Fields.Append "c1", adChar, 10, adFldUpdatable    'NUMERO
       rs.Fields.Append "c2", adChar, 150, adFldUpdatable   'Proveedor
       rs.Fields.Append "c3", adChar, 20, adFldUpdatable    'Fecha
       rs.Fields.Append "c4", adChar, 50, adFldUpdatable    'Concepto
       rs.Fields.Append "c5", adChar, 100, adFldUpdatable   'Numero
       rs.Fields.Append "c6", adChar, 20, adFldUpdatable   'Base
       rs.Fields.Append "c7", adChar, 20, adFldUpdatable   'IVA
       rs.Fields.Append "c8", adChar, 20, adFldUpdatable   'Retención
       rs.Fields.Append "c9", adChar, 20, adFldUpdatable    'Total
       rs.Fields.Append "c10", adChar, 100, adFldUpdatable   'FormaPago
       rs.Fields.Append "c11", adChar, 20, adFldUpdatable    'F.Vencimiento
       rs.Fields.Append "c12", adChar, 100, adFldUpdatable   'F.Pago
       rs.Fields.Append "c13", adChar, 35, adFldUpdatable    'Cuenta Bancaria
       rs.Open
       
       Dim i As Integer

       For i = 1 To lista.ListItems.Count
'            If lista.ListItems(i).SubItems(C_PAGO) <> "" Then
                rs.AddNew
                rs("c1") = lista.ListItems(i).Text
                rs("c2") = lista.ListItems(i).SubItems(C_PROVEEDOR)
                rs("c3") = lista.ListItems(i).SubItems(C_fecha)
                rs("c4") = lista.ListItems(i).SubItems(C_concepto)
                rs("c5") = lista.ListItems(i).SubItems(C_NUMERO)
                rs("c6") = lista.ListItems(i).SubItems(C_BASE)
                rs("c7") = lista.ListItems(i).SubItems(C_IVA)
                rs("c8") = lista.ListItems(i).SubItems(C_RETENCION)
                rs("c9") = lista.ListItems(i).SubItems(C_total)
                rs("c10") = lista.ListItems(i).SubItems(C_FP)
                rs("c11") = lista.ListItems(i).SubItems(C_vencimiento)
                rs("c12") = lista.ListItems(i).SubItems(C_PAGO)
                rs("c13") = lista.ListItems(i).SubItems(C_CUENTA)
                rs.Update
'            End If
        Next i
        
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Listado de Facturas"
 
        'Cabecera
        With XLS.Range("A1:M1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With XLS.Range("A1:M1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:M1").Borders
            .LineStyle = vbSolid
        End With
        XLS.Range("A1:A1").ColumnWidth = 12
        XLS.Range("B1:B1").ColumnWidth = 55
        XLS.Range("C1:C1").ColumnWidth = 12
        XLS.Range("D1:D1").ColumnWidth = 25
        XLS.Range("E1:E1").ColumnWidth = 25
        XLS.Range("F1:F1").ColumnWidth = 12
        XLS.Range("G1:G1").ColumnWidth = 12
        XLS.Range("H1:H1").ColumnWidth = 12
        XLS.Range("I1:I1").ColumnWidth = 12
        XLS.Range("J1:J1").ColumnWidth = 20
        XLS.Range("K1:K1").ColumnWidth = 15
        XLS.Range("L1:L1").ColumnWidth = 15
        XLS.Range("M1:M1").ColumnWidth = 30

        XLS.Cells(1, 1) = "NºAsiento"
        XLS.Cells(1, 2) = "Proveedor"
        XLS.Cells(1, 3) = "Fecha"
        XLS.Cells(1, 4) = "Concepto"
        XLS.Cells(1, 5) = "Número"
        XLS.Cells(1, 6) = "Base"
        XLS.Cells(1, 7) = "IVA"
        XLS.Cells(1, 8) = "Retención"
        XLS.Cells(1, 9) = "Total"
        XLS.Cells(1, 10) = "Forma Pago"
        XLS.Cells(1, 11) = "F.Vencimiento"
        XLS.Cells(1, 12) = "F.Pago"
        XLS.Cells(1, 13) = "Cuenta Bancaria"
        
        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Cells(i, 1) = CLng(rs("c1")) ' Asiento
            XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True) ' Proveedor
            XLS.Cells(i, 3) = CDate(Trim(rs("c3"))) ' Fecha
            XLS.Cells(i, 4) = CStr(ClrStr(rs("c4"), False, True, True)) ' Concepto
            XLS.Cells(i, 5) = CStr(ClrStr(rs("c5"), False, True, True)) ' Numero
            XLS.Cells(i, 6) = CDbl(rs("c6")) 'Base
            XLS.Cells(i, 7) = CDbl(rs("c7")) ' Iva
            XLS.Cells(i, 8) = CDbl(rs("c8")) ' Retencion
            XLS.Cells(i, 9) = CDbl(rs("c9")) ' Total
            XLS.Cells(i, 10) = rs("C10") ' Forma Pago
            XLS.Cells(i, 11) = CDate(Trim(rs("C11")))      ' F.Vencimiento
            If Trim(rs("c12")) <> "" Then
                XLS.Cells(i, 12) = CDate(Trim(rs("c12"))) ' F.Pago
            End If
            XLS.Cells(i, 13) = rs("c13") ' Cuenta bancaria
            i = i + 1
             
            XLS.Range("A" & i).EntireRow.Insert
            rs.MoveNext
          Loop Until rs.EOF
        End If
        Frame3.Visible = False
        Me.MousePointer = vbNormal
        XLA.Visible = True
        Set rs = Nothing

End Sub

Private Sub cmdListado_Click()
'Listado de facturas pendientes de cobro
       Me.MousePointer = vbHourglass
       Frame3.Visible = True
       Dim rs As New ADODB.Recordset
       Dim fecha As String
      
       rs.Fields.Append "c1", adChar, 10, adFldUpdatable    'ID
       rs.Fields.Append "c2", adChar, 150, adFldUpdatable   'Proveedor
       rs.Fields.Append "c3", adChar, 35, adFldUpdatable    'Cuenta Bancaria
       rs.Fields.Append "c4", adChar, 20, adFldUpdatable    'Total
       rs.Fields.Append "c5", adChar, 20, adFldUpdatable    'Fecha
       rs.Fields.Append "c6", adChar, 50, adFldUpdatable   'Concepto
       rs.Fields.Append "c7", adChar, 100, adFldUpdatable    'Numero
       rs.Fields.Append "c8", adChar, 100, adFldUpdatable    'Familia
       rs.Fields.Append "c9", adChar, 150, adFldUpdatable    'Subcuenta
       rs.Fields.Append "c10", adChar, 20, adFldUpdatable   'Base
       rs.Fields.Append "c11", adChar, 10, adFldUpdatable   'IVA Porcentaje
       rs.Fields.Append "c12", adChar, 20, adFldUpdatable   'IVA
       rs.Open
       
       Dim i As Integer

       For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).SubItems(C_PAGO) <> "" Then
                rs.AddNew
                rs("c1") = lista.ListItems(i).Text
                rs("c2") = lista.ListItems(i).SubItems(C_PROVEEDOR)
                rs("c3") = lista.ListItems(i).SubItems(C_CUENTA)
                rs("c4") = lista.ListItems(i).SubItems(C_total)
                rs("c5") = lista.ListItems(i).SubItems(C_fecha)
                rs("c6") = lista.ListItems(i).SubItems(C_concepto)
                rs("c7") = lista.ListItems(i).SubItems(C_NUMERO)
                rs("c8") = lista.ListItems(i).SubItems(C_FAMILIA)
                rs("c9") = lista.ListItems(i).SubItems(C_SUBCUENTA)
                rs("c10") = lista.ListItems(i).SubItems(C_BASE)
                rs("c11") = lista.ListItems(i).SubItems(C_IVA_PORCENTAJE)
                rs("c12") = lista.ListItems(i).SubItems(C_IVA)
                rs.Update
            End If
        Next i
        
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Facturas pendientes de pago"
 
        'Cabecera
        With XLS.Range("A1:L1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With XLS.Range("A1:L1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:L1").Borders
            .LineStyle = vbSolid
        End With
        XLS.Range("A1:A1").ColumnWidth = 12
        XLS.Range("B1:B1").ColumnWidth = 55
        XLS.Range("C1:C1").ColumnWidth = 55
        XLS.Range("D1:D1").ColumnWidth = 20
        XLS.Range("E1:E1").ColumnWidth = 12
        XLS.Range("F1:F1").ColumnWidth = 12
        XLS.Range("G1:G1").ColumnWidth = 12
        XLS.Range("H1:H1").ColumnWidth = 10
        XLS.Range("I1:I1").ColumnWidth = 5
        XLS.Range("J1:J1").ColumnWidth = 30
        XLS.Range("K1:K1").ColumnWidth = 30
        XLS.Range("L1:L1").ColumnWidth = 30

        XLS.Cells(1, 1) = "ID"
        XLS.Cells(1, 2) = "Proveedor"
        XLS.Cells(1, 3) = "Cuenta Bancaria"
        XLS.Cells(1, 4) = "Total a pagar"
        XLS.Cells(1, 5) = "Fecha"
        XLS.Cells(1, 6) = "Concepto"
        XLS.Cells(1, 7) = "Número"
        XLS.Cells(1, 8) = "Familia"
        XLS.Cells(1, 9) = "Subcuenta"
        XLS.Cells(1, 10) = "Base"
        XLS.Cells(1, 11) = "IVA %"
        XLS.Cells(1, 13) = "IVA"
        
        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Cells(i, 1) = CLng(rs("c1"))
            XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True)
            XLS.Cells(i, 3) = rs("c3")
            XLS.Cells(i, 4) = CDbl(rs("c4"))
            XLS.Cells(i, 5) = Trim(rs("c5"))
            XLS.Cells(i, 6) = rs("c6")
            XLS.Cells(i, 7) = rs("c7")
            XLS.Cells(i, 8) = rs("C8")
            XLS.Cells(i, 9) = rs("C9")
            XLS.Cells(i, 10) = CDbl(rs("c10"))
            XLS.Cells(i, 11) = rs("c11")
            XLS.Cells(i, 12) = CDbl(rs("c12"))
            i = i + 1
             
            XLS.Range("A" & i).EntireRow.Insert
            rs.MoveNext
          Loop Until rs.EOF
        End If
        Frame3.Visible = False
        Me.MousePointer = vbNormal
        XLA.Visible = True
        Set rs = Nothing
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    With frmProveedores_Facturas
        .PK = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_IDPROVEEDOR)
        .PK_FACTURA_ID = lista.ListItems(lista.selectedItem.Index).Text
        .TOBJETO = 0
        .COBJETO = 0
        .Show 1
    End With
'    actualizar_lista
End Sub

Private Sub cmdok_Click()
    Dim i As Integer
    Dim salida As String
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If salida <> "" Then
                salida = salida & ","
            End If
            salida = salida & lista.ListItems(i).Text
        End If
    Next
    LISTA_SEL = salida
    Unload Me
End Sub

Private Sub fdesde_Change()
    cargar_lista
End Sub

Private Sub fhasta_Change()
    cargar_lista
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    fdesde = Date - 365
    fhasta = Date
    cargarCombos
    permisos
    cabecera
    cargar_lista
End Sub
Private Sub cargarCombos()
    cargarProveedores
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbGasto, DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_GASTOS
    oDeco.cargar_mi_combo cmbPago, DECODIFICADORA.DECODIFICADORA_CONTABILIDAD_SUBCUENTAS_PAGOS
    llenar_combo cmbFamilia, New clsFamilias, 0, Me, ""
    Set oDeco = Nothing
End Sub
Private Sub cargarProveedores()
    Dim consulta As String
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        consulta = "SELECT DISTINCT P.ID_PROVEEDOR,P.NOMBRE " & _
                   "  FROM PROVEEDORES AS P, PROVEEDORES_FACTURAS AS PF " & _
                   " WHERE P.ID_PROVEEDOR = PF.PROVEEDOR_ID "
        With cmbProveedor
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "PROVEEDORES"
            .setDESCRIPCION = "Proveedores"
            .setPK = "ID_PROVEEDOR"
            .setCAMPO = "NOMBRE"
            .setQUERY = consulta
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmProveedores_Detalle
        End With
    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Nº", 1100, lvwColumnLeft
        .Add , , "Proveedor", 2100, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Concepto", 1900, lvwColumnCenter
        .Add , , "Numero", 1000, lvwColumnCenter
        .Add , , "Familia", 1, lvwColumnLeft
        .Add , , "Subcuenta", 1, lvwColumnLeft
        .Add , , "Base", 1050, lvwColumnRight
        .Add , , "Iva %", 1, lvwColumnCenter
        .Add , , "Iva", 1000, lvwColumnRight
        .Add , , "Retención", 1000, lvwColumnRight
        .Add , , "Total", 1050, lvwColumnRight
        .Add , , "Forma Pago", 1200, lvwColumnCenter
        .Add , , "Fecha Vencimiento", 1050, lvwColumnCenter
        .Add , , "Fecha Pago", 1050, lvwColumnCenter
        .Add , , "TOBJETO", 1, lvwColumnLeft
        .Add , , "COBJETO", 1, lvwColumnLeft
        .Add , , "ID_PROVEEDOR", 1, lvwColumnLeft
        'M1335-I
        .Add , , "CUENTA_BANCARIA", 1, lvwColumnLeft
        'M1335-F
        .Add , , "Env", 350, lvwColumnLeft
    End With
End Sub
Private Sub permisos()
    If USUARIO.getPER_TESORERIA_FP = False Then
    End If
End Sub
Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPF As New clsProveedores_Facturas
    Dim ID As Long
   On Error GoTo cargar_lista_Error

    ID = 0
    If cmbProveedor.getTEXTO <> "" Then
        ID = cmbProveedor.getPK_SALIDA
    End If
    Dim familiaid As Long
    Dim subcuentagasto As Long
    Dim subcuentapago As Long
    If cmbFamilia.getTEXTO <> "" Then
        familiaid = cmbFamilia.getPK_SALIDA
    End If
    If cmbGasto.getTEXTO <> "" Then
        subcuentagasto = cmbGasto.getPK_SALIDA
    End If
    If cmbPago.getTEXTO <> "" Then
        subcuentapago = cmbPago.getPK_SALIDA
    End If
    Me.MousePointer = 11
    Dim revision As Integer
    revision = 3
    Set rs = oPF.ListadoCompleto(ID, chkPendientesPago.Value, fdesde, fhasta, familiaid, subcuentagasto, subcuentapago, chkNoEnviadas.Value, txtconcepto, chkVencidas.Value, chkPagoPrevisto.Value, chkIncidencias.Value, txtImporteDesde, txtimportehasta, False, "", "", False, "", "", "", False, "", revision)
    Dim BASE As Currency
    Dim IVA As Currency
    Dim retencion As Currency
    Dim total As Currency
    BASE = 0
    IVA = 0
    retencion = 0
    total = 0
    lista.ListItems.Clear
'    lblsubtitulo = "Se han detectado " & rs.RecordCount & " registros."
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "000000")) ' ID
           .SubItems(COLS.C_PROVEEDOR) = rs(17)
           .SubItems(COLS.C_IDPROVEEDOR) = rs(18)
            .SubItems(COLS.C_fecha) = Format(rs(1), "dd/mm/yyyy")  ' Fecha
            If Not IsNull(rs(2)) Then
                .SubItems(COLS.C_concepto) = rs(2)  ' Concepto
            End If
            If Not IsNull(rs(3)) Then
                .SubItems(COLS.C_NUMERO) = rs(3)  ' Numero
            End If
            If Not IsNull(rs(4)) Then
                .SubItems(COLS.C_FAMILIA) = rs(4)  ' Familia
            End If
            If Not IsNull(rs(5)) Then
                .SubItems(COLS.C_SUBCUENTA) = rs(5)  ' Subcuenta
            End If
            .SubItems(COLS.C_BASE) = Format(rs(6), "currency")  ' BI
            .SubItems(COLS.C_IVA_PORCENTAJE) = rs(7)  ' IVA PORCENTAJE
            .SubItems(COLS.C_IVA) = Format(rs(8), "currency")  ' IVA
            .SubItems(COLS.C_total) = Format(rs(9), "currency")  ' TOTAL
            BASE = BASE + rs(6)
            IVA = IVA + rs(8)
            retencion = retencion + rs(16)
            total = total + rs(9)
            If Not IsNull(rs(10)) Then
                .SubItems(COLS.C_FP) = rs(10)  ' FP
            End If
            If Not IsNull(rs(11)) Then
                .SubItems(COLS.C_vencimiento) = rs(11)  ' F.Vencimiento
            End If
            If Not IsNull(rs(13)) Then
                .SubItems(COLS.C_TOBJETO) = rs(13)  ' Tobjeto
            End If
            If Not IsNull(rs(14)) Then
                .SubItems(COLS.C_cOBJETO) = rs(14)  ' Cobjeto
            End If
            If Not IsNull(rs(16)) Then
                .SubItems(COLS.C_RETENCION) = Format(rs(16), "currency") ' RETENCION
            End If
            If Not IsNull(rs(12)) Then
                .SubItems(COLS.C_PAGO) = rs(12)
            End If
            If Not IsNull(rs(19)) Then
                .SubItems(COLS.C_CUENTA) = rs(19)
            End If
            If rs(20) = 0 Then
                .SubItems(COLS.C_ENVIADA) = "N"
            Else
                .SubItems(COLS.C_ENVIADA) = "S"
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    lblBase = Format(BASE, "currency")
    lblIVA = Format(IVA, "currency")
    lblRetencion = Format(retencion, "currency")
    lbltotal = Format(total, "currency")
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmProveedores_Facturas_Listado_Modal"

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
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
Private Sub txtconcepto_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cargar_lista
    End If
End Sub

Private Sub txtImporteDesde_GotFocus()
    txtImporteDesde.SelStart = 0
    txtImporteDesde.SelLength = Len(txtImporteDesde)
End Sub

Private Sub txtImporteDesde_LostFocus()
    If txtImporteDesde <> "" Then
        txtImporteDesde = moneda(txtImporteDesde)
        txtimportehasta = txtImporteDesde
    End If
End Sub

Private Sub txtimportehasta_GotFocus()
    txtimportehasta.SelStart = 0
    txtimportehasta.SelLength = Len(txtimportehasta)
End Sub

Private Sub txtimportehasta_LostFocus()
    If txtimportehasta <> "" Then
        txtimportehasta = moneda(txtimportehasta)
    End If
End Sub
