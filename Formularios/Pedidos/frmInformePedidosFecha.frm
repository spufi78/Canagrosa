VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmInformePedidosFecha 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de pedidos por Fecha Prevista"
   ClientHeight    =   9270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15255
   Icon            =   "frmInformePedidosFecha.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9270
   ScaleWidth      =   15255
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   4275
      TabIndex        =   14
      Top             =   8370
      Visible         =   0   'False
      Width           =   6765
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
         Left            =   585
         TabIndex        =   15
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
      Width           =   15165
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   840
         Left            =   13995
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1485
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
         Left            =   3510
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
         Left            =   1485
         TabIndex        =   11
         Top             =   270
         Width           =   12090
         _ExtentX        =   21325
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmInformePedidosFecha.frx":6852
         Height          =   315
         Left            =   10305
         TabIndex        =   16
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
         Left            =   9675
         TabIndex        =   17
         Top             =   720
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   12
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
         Left            =   2925
         TabIndex        =   7
         Top             =   705
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Pedido"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   6
         Top             =   720
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   14085
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8370
      Width           =   1095
   End
   Begin VB.CommandButton cmdVerExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   13005
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8370
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6045
      Left            =   45
      TabIndex        =   8
      Top             =   1980
      Width           =   15180
      _ExtentX        =   26776
      _ExtentY        =   10663
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2160
      Top             =   8370
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformePedidosFecha.frx":6898
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformePedidosFecha.frx":6E26
            Key             =   ""
         EndProperty
      EndProperty
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
      TabIndex        =   13
      Top             =   8010
      Width           =   15180
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informe de pedidos por Fecha Prevista"
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
      Left            =   45
      TabIndex        =   10
      Top             =   90
      Width           =   15150
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   15195
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
      TabIndex        =   9
      Top             =   1665
      Width           =   15165
   End
   Begin VB.Image imagen 
      Height          =   480
      Index           =   0
      Left            =   13050
      Picture         =   "frmInformePedidosFecha.frx":7371
      Top             =   45
      Width           =   480
   End
End
Attribute VB_Name = "frmInformePedidosFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Enum COLS
    COL_CENTRO = 1
    COL_PROVEEDOR = 2
    COL_REACTIVO = 3
    COL_COD_PEDIDO = 4
    COL_F_PEDIDO = 5
    COL_F_PREVISTA = 6
    COL_DIAS_PROVEEDOR = 7
    COL_F_ESTIMADA = 8
    COL_F_RECEPCION = 9
    COL_DIAS_RETRASO = 10
    COL_ATRASO = 11
End Enum
    
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
        .Add , , "", 300, lvwColumnLeft
        .Add , , "CENTRO", 800, lvwColumnLeft
        .Add , , "PROVEEDOR", 3000, lvwColumnLeft
        .Add , , "REACTIVO", 3800, lvwColumnLeft
        .Add , , "COD.PEDIDO", 1000, lvwColumnCenter
        .Add , , "F.PEDIDO", 1100, lvwColumnCenter
        .Add , , "F.PREVISTA", 1100, lvwColumnCenter
        .Add , , "DIAS PROVEEDOR", 800, lvwColumnCenter
        .Add , , "F.ESTIMADA", 1100, lvwColumnCenter
        .Add , , "F.RECEPCION", 1100, lvwColumnCenter
        .Add , , "DIAS ATRASO", 800, lvwColumnCenter
        .Add , , "ATRASADO", 0, lvwColumnCenter
    End With
End Sub

Private Sub cmdBuscar_Click()
    Call cargar_lista
End Sub

Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oPedido As New clsPedidos_bote_ex
    Dim codProveedor As Long
    Dim fechaI As String
    Dim fechaF As String
   On Error GoTo cargar_lista_Error

    codProveedor = cmbProveedor.getPK_SALIDA
    fechaI = fdesde
    fechaF = fhasta
    Dim atrasos As Integer
    atrasos = 0
    Set rs = oPedido.ListadoPedidosFechaPrevista(codProveedor, fechaI, fechaF, cmbCentro.BoundText)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000"))
             .SubItems(COLS.COL_CENTRO) = rs(1)
             .SubItems(COLS.COL_PROVEEDOR) = rs(2)
             .SubItems(COLS.COL_REACTIVO) = rs(3)
             .SubItems(COLS.COL_COD_PEDIDO) = rs(4)
             .SubItems(COLS.COL_F_PEDIDO) = rs(5)
             .SubItems(COLS.COL_F_PREVISTA) = rs(6)
             .SubItems(COLS.COL_DIAS_PROVEEDOR) = rs(7)
             .SubItems(COLS.COL_F_ESTIMADA) = rs(8)
             .SubItems(COLS.COL_F_RECEPCION) = Format(rs(9), "dd/mm/yyyy")  ' Fecha recepcion
             ' Evaluar dias retraso
             Dim f1 As Date
             Dim f2 As Date
             ' Si la fecha prevista esta informada, la usamos sino, tomamos la estimada
             If .SubItems(COLS.COL_F_PREVISTA) = "" Then
                If IsDate(.SubItems(COLS.COL_F_ESTIMADA)) Then
                    f1 = .SubItems(COLS.COL_F_ESTIMADA)  'Fecha estimada
                End If
             Else
                If IsDate(.SubItems(COLS.COL_F_PREVISTA)) Then
                    f1 = .SubItems(COLS.COL_F_PREVISTA) 'Fecha prevista
                End If
             End If
             ' Si la fecha recepcion informada, la usamos, si no, la del dia
             If .SubItems(COLS.COL_F_RECEPCION) = "" Then
                f2 = Date
             Else
                f2 = .SubItems(COLS.COL_F_RECEPCION)
             End If
             ' Calculamos el numero de dias entre dos fecha
             dias = DateDiff("d", f1, f2)
             .SubItems(COLS.COL_DIAS_RETRASO) = dias
             i = lista.ListItems.Count
             If dias > 0 Then
                 atrasos = atrasos + 1
                 .SubItems(COLS.COL_ATRASO) = "X"
                 lista.ListItems(i).SmallIcon = 1
             Else
                 .SubItems(COLS.COL_ATRASO) = ""
                 lista.ListItems(i).SmallIcon = 2
             End If
             rs.MoveNext
            End With
        Loop Until rs.EOF
    End If
    lblsubtitulo = "Total Pedidos : " & lista.ListItems.Count & "     ->     Total Atrasos : " & atrasos
    Set oPedido = Nothing

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmInformePedidosFecha"
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
        XLW.Worksheets(1).Name = "Informe de pedidos"
 
        'Cabecera
        With XLS.Range("A1:K1")
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
        With XLS.Range("A1:K1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:K1").Borders
            .LineStyle = vbSolid
        End With
        XLS.Range("A1:A1").ColumnWidth = 10
        XLS.Range("B1:B1").ColumnWidth = 50
        XLS.Range("C1:C1").ColumnWidth = 50
        XLS.Range("D1:D1").ColumnWidth = 10
        XLS.Range("E1:E1").ColumnWidth = 10
        XLS.Range("F1:F1").ColumnWidth = 10
        XLS.Range("G1:G1").ColumnWidth = 10
        XLS.Range("H1:H1").ColumnWidth = 10
        XLS.Range("I1:I1").ColumnWidth = 10
        XLS.Range("J1:J1").ColumnWidth = 10
        XLS.Range("K1:K1").ColumnWidth = 10
        
        XLS.Cells(1, 1) = "Centro"
        XLS.Cells(1, 2) = "Proveedor"
        XLS.Cells(1, 3) = "Reactivo"
        XLS.Cells(1, 4) = "Cod.Pedido"
        XLS.Cells(1, 5) = "F.Pedido"
        XLS.Cells(1, 6) = "F.Prevista"
        XLS.Cells(1, 7) = "Dias Proveedor"
        XLS.Cells(1, 8) = "F.Estimada"
        XLS.Cells(1, 9) = "F.Recepcion"
        XLS.Cells(1, 10) = "Dias Atraso"
        XLS.Cells(1, 11) = "Atrasado"
        
        For i = 1 To lista.ListItems.Count
            XLS.Cells(i + 1, COLS.COL_CENTRO) = lista.ListItems(i).SubItems(COLS.COL_CENTRO)
            XLS.Cells(i + 1, COLS.COL_PROVEEDOR) = lista.ListItems(i).SubItems(COLS.COL_PROVEEDOR)
            XLS.Cells(i + 1, COLS.COL_REACTIVO) = lista.ListItems(i).SubItems(COLS.COL_REACTIVO)
            XLS.Cells(i + 1, COLS.COL_COD_PEDIDO) = lista.ListItems(i).SubItems(COLS.COL_COD_PEDIDO)
'            XLS.Cells(i + 1, COLS.COL_F_PEDIDO) = ClrStr(lista.ListItems(i).SubItems(COLS.COL_F_PEDIDO), False, True, True)
'            XLS.Cells(i + 1, COLS.COL_F_PREVISTA) = ClrStr(lista.ListItems(i).SubItems(COLS.COL_F_PREVISTA), False, True, True)
            XLS.Cells(i + 1, COLS.COL_F_PEDIDO) = ClrStr(Format(lista.ListItems(i).SubItems(COLS.COL_F_PEDIDO), "yyyy-mm-dd"), False, True, True)
            XLS.Cells(i + 1, COLS.COL_F_PREVISTA) = ClrStr(Format(lista.ListItems(i).SubItems(COLS.COL_F_PREVISTA), "yyyy-mm-dd"), False, True, True)
            
            XLS.Cells(i + 1, COLS.COL_DIAS_PROVEEDOR) = ClrStr(lista.ListItems(i).SubItems(COLS.COL_DIAS_PROVEEDOR), False, True, True)
'            XLS.Cells(i + 1, COLS.COL_F_ESTIMADA) = ClrStr(lista.ListItems(i).SubItems(COLS.COL_F_ESTIMADA), False, True, True)
'            XLS.Cells(i + 1, COLS.COL_F_RECEPCION) = ClrStr(lista.ListItems(i).SubItems(COLS.COL_F_RECEPCION), False, True, True)
            XLS.Cells(i + 1, COLS.COL_F_ESTIMADA) = ClrStr(Format(lista.ListItems(i).SubItems(COLS.COL_F_ESTIMADA), "yyyy-mm-dd"), False, True, True)
            XLS.Cells(i + 1, COLS.COL_F_RECEPCION) = ClrStr(Format(lista.ListItems(i).SubItems(COLS.COL_F_RECEPCION), "yyyy-mm-dd"), False, True, True)
            XLS.Cells(i + 1, COLS.COL_DIAS_RETRASO) = ClrStr(lista.ListItems(i).SubItems(COLS.COL_DIAS_RETRASO), False, True, True)
            XLS.Cells(i + 1, COLS.COL_ATRASO) = ClrStr(lista.ListItems(i).SubItems(COLS.COL_ATRASO), False, True, True)
        Next
        Frame3.visible = False
        Me.MousePointer = vbNormal
        XLA.visible = True
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cmdVerExcel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdVerExcel_Click of Formulario frmInformePedidosFecha"
End Sub
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        frmREX_Pedidos_Detalle.PK = lista.ListItems(lista.selectedItem.Index)
        frmREX_Pedidos_Detalle.Show 1
    End If
End Sub
