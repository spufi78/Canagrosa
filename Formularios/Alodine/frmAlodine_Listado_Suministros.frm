VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAlodine_Listado_Suministros 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Alodine Suministrado"
   ClientHeight    =   10260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11385
   Icon            =   "frmAlodine_Listado_Suministros.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   11385
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   2025
      TabIndex        =   17
      Top             =   9360
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
         Left            =   675
         TabIndex        =   18
         Top             =   225
         Width           =   5415
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   600
      Left            =   45
      TabIndex        =   13
      Top             =   8640
      Width           =   11265
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1395
         TabIndex        =   16
         Top             =   225
         Width           =   4965
      End
      Begin VB.Label lblSuma 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   9810
         TabIndex        =   15
         Top             =   180
         Width           =   1230
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "Suma Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   240
         Left            =   8280
         TabIndex        =   14
         Top             =   270
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdVerExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9315
      Width           =   1290
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10035
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9315
      Width           =   1290
   End
   Begin VB.Frame Frame1 
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
      Height          =   960
      Left            =   0
      TabIndex        =   6
      Top             =   585
      Width           =   11310
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   735
         Left            =   10215
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   1005
      End
      Begin MSDataListLib.DataCombo cmbProducto 
         Height          =   330
         Left            =   1035
         TabIndex        =   0
         Top             =   315
         Width           =   4245
         _ExtentX        =   7488
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker fecha_i 
         Height          =   330
         Left            =   6300
         TabIndex        =   1
         Top             =   315
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha_f 
         Height          =   330
         Left            =   8505
         TabIndex        =   2
         Top             =   315
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde el"
         Height          =   240
         Index           =   4
         Left            =   5535
         TabIndex        =   12
         Top             =   360
         Width           =   675
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta el"
         Height          =   240
         Index           =   6
         Left            =   7785
         TabIndex        =   11
         Top             =   360
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   405
         Width           =   645
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4725
      Left            =   45
      TabIndex        =   8
      Top             =   1575
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   8334
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
   Begin MSComctlLib.ListView listaFacturas 
      Height          =   2025
      Left            =   45
      TabIndex        =   19
      Top             =   6570
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   3572
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Facturas asociadas al Alodine"
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
      Left            =   45
      TabIndex        =   20
      Top             =   6300
      Width           =   11310
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Especifique los datos necesarios para localizar los lotes"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   3690
      TabIndex        =   10
      Top             =   315
      Width           =   3900
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Alodine Suministrado"
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
      Index           =   2
      Left            =   4005
      TabIndex        =   9
      Top             =   0
      Width           =   3375
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10800
      Picture         =   "frmAlodine_Listado_Suministros.frx":08CA
      Top             =   45
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   11340
   End
End
Attribute VB_Name = "frmAlodine_Listado_Suministros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' --------------------------------------------------------------------------------------------------------------
' MANTIS: 1290
' --------------------------------------------------------------------------------------------------------------
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    cabecera
    cargar_combo cmbProducto, New clsAlodine
    cargar_fechas
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1000, lvwColumnLeft
        .Add , , "Producto", 5500, lvwColumnCenter
        .Add , , "Descripción", 1800, lvwColumnCenter
        .Add , , "Precio", 1500, lvwColumnCenter
        .Add , , "Total (uds.)", 1200, lvwColumnCenter
        
        .Add , , "DOCS", 1, lvwColumnCenter
    End With
    
    With listaFacturas.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "NºDOC", 1000, lvwColumnCenter
        .Add , , "Cliente", 5300, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnCenter
        .Add , , "Pedido", 2500, lvwColumnCentter
    End With
    
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmdLimpiar_Click()
    cmbProducto.BoundText = ""
    cmbProducto.Text = ""
    cargar_fechas
    cargar_lista
End Sub
Public Sub cargar_fechas()
    fecha_i = "01/" & Month(Date) & "/" & Year(Date)
    fecha_f = Date
End Sub
Private Sub fecha_i_LostFocus()
    cargar_lista
End Sub
Private Sub fecha_f_LostFocus()
    cargar_lista
End Sub
Private Sub cmbproducto_Change()
    cargar_lista
End Sub
Public Sub cargar_lista()
    If validar Then
        Dim rs As New ADODB.Recordset
        Dim oAlodine As New clsAlodine
        Dim tipo As Long
        Dim Suma As Long
        
        lista.ListItems.Clear
        If cmbProducto.BoundText <> "" Then
            tipo = cmbProducto.BoundText
        Else
            tipo = 0
        End If
        
        Set rs = oAlodine.ListadoSuministrados(tipo, Format(fecha_i, "yyyy-mm-dd"), Format(fecha_f, "yyyy-mm-dd"))
        If rs.RecordCount <> 0 Then
            Suma = 0
            Do
                With lista.ListItems.Add(, , Format(rs("ID_ALODINE"), "0000"))
                 .SubItems(1) = rs("PRODUCTO")
                 .SubItems(2) = rs("DESCRIPCION")
                 .SubItems(3) = moneda(rs("PRECIO"))
                 .SubItems(4) = rs("SUMA")
                 .SubItems(5) = rs("DOCS")
                 Suma = Suma + CLng(rs("SUMA"))
                End With
                rs.MoveNext
            Loop Until rs.EOF
            Label1.Visible = True
            Label1.Caption = "Número de resultados encontrados: " & rs.RecordCount
            lblSuma.Caption = Suma & " Uds."
        Else
            Label1.Visible = False
            lblSuma.Caption = 0
        End If
        Set oAlodine = Nothing
        Set rs = Nothing
    End If
End Sub

Private Function validar() As Boolean
    validar = True
    If fecha_i > fecha_f Then
        MsgBox "La fecha de inicio del intervalo no es correcta", vbInformation
        validar = False
        fecha_i.SetFocus
    End If
End Function

Private Sub cmdVerExcel_Click()
  Dim cadena As String
    Me.MousePointer = vbHourglass
       Frame3.Visible = True
       Dim rs As New ADODB.Recordset
       rs.Fields.Append "c1", adChar, 10, adFldUpdatable
       rs.Fields.Append "c2", adChar, 500, adFldUpdatable
       rs.Fields.Append "c3", adChar, 350, adFldUpdatable
       rs.Fields.Append "c4", adChar, 20, adFldUpdatable
       rs.Fields.Append "c5", adChar, 20, adFldUpdatable
       rs.Open
       
       Dim i As Integer

       For i = 1 To lista.ListItems.Count
            rs.AddNew
            rs("c1") = lista.ListItems(i).Text
            rs("c2") = lista.ListItems(i).SubItems(1)
            rs("c3") = lista.ListItems(i).SubItems(2)
            rs("c4") = lista.ListItems(i).SubItems(3)
            rs("c5") = lista.ListItems(i).SubItems(4)
            rs.Update
        Next i
        
        Dim XLA As Excel.Application
        Dim XLW As Excel.Workbook
        Dim XLS As Excel.Worksheet
        
        Set XLA = New Excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Alodine_" & Format(fecha_i, "yyyy_mm_dd")
        'Cabecera
 
        With XLS.Range("A1:E1")
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
        With XLS.Range("A1:E1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:E1").Borders
            .LineStyle = vbSolid
        End With
        XLS.Range("A1:A1").ColumnWidth = 5
        XLS.Range("B1:B1").ColumnWidth = 60
        XLS.Range("C1:C1").ColumnWidth = 12
        XLS.Range("D1:D1").ColumnWidth = 10
        XLS.Range("E1:E1").ColumnWidth = 10
        XLS.Cells(1, 1) = "Tipo"
        XLS.Cells(1, 2) = "Producto"
        XLS.Cells(1, 3) = "Descripción"
        XLS.Cells(1, 4) = "Precio (€)"
        XLS.Cells(1, 5) = "Total (Uds.)"

        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Cells(i, 1) = rs("c1")
            XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True)
            XLS.Cells(i, 3) = ClrStr(rs("c3"), False, True, True)
            XLS.Cells(i, 4) = CDbl(rs("c4"))
            XLS.Cells(i, 5) = CDbl(rs("c5"))
             
            i = i + 1
             
            XLS.Range("A" & i).EntireRow.Insert
            rs.MoveNext
          Loop Until rs.EOF
        End If
        Me.MousePointer = vbNormal
        Frame3.Visible = False
        XLA.Visible = True
        
    Set rs = Nothing
End Sub

Private Sub lista_Click()
    listaFacturas.ListItems.Clear
    If lista.ListItems.Count > 0 Then
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT A.ID_DOC,A.NUMERO, B.NOMBRE,A.FECHA_FACTURA,C.CODIGO " & _
                   "  FROM DOCS_PAGO A " & _
                   "  LEFT JOIN CLIENTES B ON A.CLIENTE_ID = B.ID_CLIENTE " & _
                   "  LEFT JOIN CLIENTES_PEDIDOS C ON A.PEDIDO_ID = C.ID_PEDIDO " & _
                   " WHERE A.ID_DOC IN (" & lista.ListItems(lista.selectedItem.Index).SubItems(5) & ")" & _
                   " ORDER BY A.ID_DOC "
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            Do
                With listaFacturas.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(1)
                 .SubItems(2) = rs(2)
                 .SubItems(3) = Format(rs(3), "dd-mm-yyyy")
                 .SubItems(4) = rs(4)
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = Nothing
    End If
End Sub

Private Sub listaFacturas_DblClick()
    If listaFacturas.ListItems.Count = 0 Then Exit Sub
    If USUARIO.getPER_FACTURACION = 0 Then
        MsgBox "No tiene permisos para ver la facturacion.", vbExclamation, App.Title
    Else
        gdoc = listaFacturas.ListItems(listaFacturas.selectedItem.Index).Text
        frmListadoDocPago.Show
    End If
End Sub
