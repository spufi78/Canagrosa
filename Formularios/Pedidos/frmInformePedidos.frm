VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmInformePedidos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de pedidos realizados"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   16650
   Icon            =   "frmInformePedidos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   16650
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   4950
      TabIndex        =   16
      Top             =   8370
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
         Left            =   675
         TabIndex        =   17
         Top             =   225
         Width           =   5415
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1125
      Left            =   45
      TabIndex        =   2
      Top             =   540
      Width           =   16515
      Begin VB.TextBox txtMotivo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6660
         TabIndex        =   18
         Top             =   675
         Width           =   3390
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   840
         Left            =   15345
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1800
         TabIndex        =   4
         Top             =   675
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3825
         TabIndex        =   5
         Top             =   675
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   330
         Left            =   1815
         TabIndex        =   12
         Top             =   270
         Width           =   9300
         _ExtentX        =   16404
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmInformePedidos.frx":6852
         Height          =   315
         Left            =   11880
         TabIndex        =   20
         Top             =   675
         Width           =   2235
         _ExtentX        =   3942
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
         Left            =   11250
         TabIndex        =   21
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Motivo"
         Height          =   195
         Index           =   2
         Left            =   5940
         TabIndex        =   19
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   13
         Top             =   330
         Width           =   735
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   4
         Left            =   3240
         TabIndex        =   7
         Top             =   705
         Width           =   405
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de tramitación"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   6
         Top             =   720
         Width           =   1515
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   15480
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8370
      Width           =   1095
   End
   Begin VB.CommandButton cmdVerExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8370
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5730
      Left            =   45
      TabIndex        =   8
      Top             =   1980
      Width           =   16575
      _ExtentX        =   29236
      _ExtentY        =   10107
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
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
      Left            =   7920
      Top             =   8415
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
            Picture         =   "frmInformePedidos.frx":6898
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformePedidos.frx":7172
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformePedidos.frx":7A4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformePedidos.frx":8326
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformePedidos.frx":8C00
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformePedidos.frx":94DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   45
      TabIndex        =   15
      Top             =   7695
      Width           =   16575
   End
   Begin VB.Label lblSubtitulo 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   14
      Top             =   8010
      Width           =   16575
   End
   Begin VB.Image imagen 
      Height          =   480
      Index           =   1
      Left            =   16020
      Picture         =   "frmInformePedidos.frx":FD3C
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informe de pedidos realizados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   0
      Left            =   5940
      TabIndex        =   11
      Top             =   90
      Width           =   3645
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   16590
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione Criterio."
      BeginProperty Font 
         Name            =   "Verdana"
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
      TabIndex        =   10
      Top             =   1665
      Width           =   16515
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "MUESTRAS FUERA DE PLAZO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   180
      TabIndex        =   9
      Top             =   90
      Width           =   3855
   End
   Begin VB.Image imagen 
      Height          =   480
      Index           =   0
      Left            =   13050
      Picture         =   "frmInformePedidos.frx":10606
      Top             =   45
      Width           =   480
   End
End
Attribute VB_Name = "frmInformePedidos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************'
'*************** Fecha de creación del formulario: 10/03/2014 ****************'
'****************                MANTIS: 1305           ***************'
'*****************************************************************************'

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    fdesde = "01/01/" & Year(Date)
    fhasta = Date
    cabecera
    cargar_combo cmbCentro, New clsCentros
    llenar_combo cmbProveedor, New clsProveedor, 0, frmProveedores_Detalle, ""
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID_PEDIDO", 1, lvwColumnLeft
        .Add , , "TIPO", 1, lvwColumnLeft
        .Add , , "CENTRO", 850, lvwColumnCenter
        .Add , , "C.PEDIDO", 750, lvwColumnCenter
        .Add , , "FECHA", 1050, lvwColumnCenter
        .Add , , "F.PEDIDO", 1050, lvwColumnCenter
        .Add , , "PROVEEDOR", 2600, lvwColumnLeft
        .Add , , "DETALLE", 2600, lvwColumnLeft
        .Add , , "REFERENCIA", 1400, lvwColumnCenter
        .Add , , "MOTIVO", 1700, lvwColumnLeft
        .Add , , "CANTIDAD", 900, lvwColumnCenter
        .Add , , "TOTAL", 900, lvwColumnCenter
        .Add , , "PRECIO (Ud.)", 1200, lvwColumnRight
        .Add , , "IMPORTE", 1200, lvwColumnRight
    End With
End Sub

Private Sub cmdBuscar_Click()
    Call cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oPedido As New clsPedidos_bote_ex
    Dim codProveedor As Long
    Dim fechaI As String
    Dim fechaF As String
    Dim Suma As Double
    
    codProveedor = cmbProveedor.getPK_SALIDA
    fechaI = fdesde
    fechaF = fhasta
    
    Suma = 0
    
    Set rs = oPedido.ListadoPedidosRealizados(codProveedor, fechaI, fechaF, txtmotivo, cmbCentro.BoundText)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = rs(6)
             .SubItems(7) = rs(7)
             .SubItems(8) = rs(8)
             .SubItems(9) = rs(9)
             .SubItems(10) = rs(10)
             .SubItems(11) = rs(11)
             .SubItems(12) = moneda(rs(12))
             .SubItems(13) = moneda(rs(13))
             Suma = Suma + CDbl(rs(13))
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    lblsubtitulo = "Líneas mostradas : " & lista.ListItems.Count
    lblImporte = "Importe Total: " & moneda(str(Suma)) & "      "
    Set ocli = Nothing
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdVerExcel_Click()
   On Error GoTo cmdVerExcel_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
       Frame3.visible = True
       Me.MousePointer = vbHourglass
       Dim i As Integer
       Dim XLA As excel.Application
       Dim XLW As excel.Workbook
       Dim XLS As excel.Worksheet
       
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Informe de pedidos realizados"
 
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
        XLS.Range("A1:A1").ColumnWidth = 10
        XLS.Range("B1:B1").ColumnWidth = 10
        XLS.Range("C1:C1").ColumnWidth = 10
        XLS.Range("D1:D1").ColumnWidth = 10
        XLS.Range("E1:E1").ColumnWidth = 50
        XLS.Range("F1:F1").ColumnWidth = 50
        XLS.Range("G1:G1").ColumnWidth = 50
        XLS.Range("H1:H1").ColumnWidth = 50
        XLS.Range("I1:I1").ColumnWidth = 10
        XLS.Range("J1:J1").ColumnWidth = 10
        XLS.Range("J1:K1").ColumnWidth = 10
        XLS.Range("L1:L1").ColumnWidth = 10
        
        XLS.Cells(1, 1) = "Centro"
        XLS.Cells(1, 2) = "Cod.Pedido"
        XLS.Cells(1, 3) = "Fecha"
        XLS.Cells(1, 4) = "Fecha.Pedido"
        XLS.Cells(1, 5) = "Proveedor"
        XLS.Cells(1, 6) = "Detalle"
        XLS.Cells(1, 7) = "Referencia"
        XLS.Cells(1, 8) = "Motivo"
        XLS.Cells(1, 9) = "Cantidad"
        XLS.Cells(1, 10) = "Uds."
        XLS.Cells(1, 11) = "Precio"
        XLS.Cells(1, 12) = "Importe"
        
        For i = 1 To lista.ListItems.Count
            XLS.Cells(i + 1, 1) = lista.ListItems(i).SubItems(2)
            XLS.Cells(i + 1, 2) = lista.ListItems(i).SubItems(3)
            XLS.Cells(i + 1, 3) = lista.ListItems(i).SubItems(4)
            XLS.Cells(i + 1, 4) = lista.ListItems(i).SubItems(5)
            XLS.Cells(i + 1, 5) = ClrStr(lista.ListItems(i).SubItems(6), False, True, True)
            XLS.Cells(i + 1, 6) = ClrStr(lista.ListItems(i).SubItems(7), False, True, True)
            XLS.Cells(i + 1, 7) = ClrStr(lista.ListItems(i).SubItems(8), False, True, True)
            XLS.Cells(i + 1, 8) = ClrStr(lista.ListItems(i).SubItems(9), False, True, True)
            XLS.Cells(i + 1, 9) = lista.ListItems(i).SubItems(10)
            XLS.Cells(i + 1, 10) = CDbl(lista.ListItems(i).SubItems(11))
            XLS.Cells(i + 1, 11) = CDbl(lista.ListItems(i).SubItems(12))
            XLS.Cells(i + 1, 12) = CDbl(lista.ListItems(i).SubItems(13))
             
'            XLS.Range("A" & i + 1).EntireRow.Insert
        Next
        Frame3.visible = False
        Me.MousePointer = vbNormal
        XLA.visible = True
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cmdVerExcel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdVerExcel_Click of Formulario frmInformePedidos"
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        frmREX_Pedidos_Detalle.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmREX_Pedidos_Detalle.Show 1
    End If
End Sub
