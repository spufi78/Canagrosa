VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmInformeFacturacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de facturación"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10500
   Icon            =   "frmInformeFacturacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   10500
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informe de total facturado en conceptos entre dos fechas "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   45
      TabIndex        =   12
      Top             =   1800
      Width           =   10395
      Begin VB.CheckBox chkTodos2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   7965
         TabIndex        =   14
         Top             =   315
         Width           =   870
      End
      Begin VB.CommandButton cmdGenerar2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar"
         Height          =   1005
         Left            =   9000
         Picture         =   "frmInformeFacturacion.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   1320
      End
      Begin MSComCtl2.DTPicker fhasta2 
         Height          =   330
         Left            =   3285
         TabIndex        =   15
         Top             =   720
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   60293121
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fdesde2 
         Height          =   330
         Left            =   900
         TabIndex        =   16
         Top             =   720
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   60293121
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbclientes2 
         Height          =   330
         Left            =   900
         TabIndex        =   21
         Top             =   270
         Width           =   6990
         _ExtentX        =   12330
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   90
         TabIndex        =   19
         Top             =   345
         Width           =   675
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   18
         Top             =   765
         Width           =   585
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   2610
         TabIndex        =   17
         Top             =   765
         Width           =   555
      End
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   90
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   960
      Left            =   9045
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3240
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informe de total facturado por tipo de muestra y tipo de análisis entre dos fechas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1350
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   10395
      Begin VB.CommandButton cmdExcel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar"
         Height          =   1005
         Left            =   9000
         Picture         =   "frmInformeFacturacion.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1320
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   7965
         TabIndex        =   1
         Top             =   315
         Width           =   870
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3285
         TabIndex        =   3
         Top             =   720
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   60293121
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   900
         TabIndex        =   2
         Top             =   720
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   60293121
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   900
         TabIndex        =   20
         Top             =   270
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   2610
         TabIndex        =   6
         Top             =   765
         Width           =   555
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   5
         Top             =   765
         Width           =   585
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   4
         Top             =   345
         Width           =   675
      End
   End
   Begin VB.Label lbltitulo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Generando Informe.......... espere"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   3375
      Visible         =   0   'False
      Width           =   2850
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Informes de Facturación"
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
      Index           =   4
      Left            =   45
      TabIndex        =   7
      Top             =   45
      Width           =   10380
   End
End
Attribute VB_Name = "frmInformeFacturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTodos_Click()
    If chkTodos.Value = Checked Then
        cmbclientes.limpiar
        cmbclientes.activar
    Else
        cmbclientes.desactivar
    End If
End Sub
Private Sub chkTodos2_Click()
    If chkTodos2.Value = Checked Then
        cmbclientes2.limpiar
        cmbclientes2.activar
    Else
        cmbclientes2.desactivar
    End If
End Sub

Private Sub cmdExcel_Click()
    Dim rs As ADODB.Recordset
    Dim cliente As String
   On Error GoTo cmdExcel_Click_Error

    If cmbclientes.getPK_SALIDA <> 0 Then
        cliente = "   and m.cliente_id = " & cmbclientes.getPK_SALIDA
    End If
    consulta = "select tm.nombre,ta.nombre,c.nombre,m.tipo_muestra_id,m.tipo_analisis_id,m.cliente_id,count(*) " & _
               "  from tipos_muestra tm, tipos_analisis ta, clientes c, muestras m " & _
               " where m.fecha_recepcion >= '" & Format(fdesde, "yyyy-mm-dd") & "' " & _
               "   and m.fecha_recepcion <= '" & Format(fhasta, "yyyy-mm-dd") & "' " & _
               "   and m.anulada = 0 " & _
               "   and m.documento_pago <> 0 " & _
               cliente & _
               "   and m.cliente_id = c.id_cliente " & _
               "   and m.tipo_muestra_id = tm.id_tipo_muestra " & _
               "   and m.tipo_analisis_id = ta.id_tipo_analisis " & _
               " group by tm.nombre,ta.nombre,c.nombre"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        MsgBox "No existen registros para la selección.", vbInformation, App.Title
        Exit Sub
    Else
        Me.MousePointer = 11
        lbltitulo.visible = True
        pb.visible = True
        pb.min = 0
        pb.Max = rs.RecordCount
        pb.Value = 1
        Dim i As Integer
        Dim rs_total As ADODB.Recordset
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
'        XLA.Visible = True
        XLS.Range("1:1").HorizontalAlignment = xlCenter
        XLS.Range("1:1").VerticalAlignment = xlCenter
        XLS.Range("A1:C1").ColumnWidth = 28
        XLS.Range("D1:E1").ColumnWidth = 10
        XLS.Range("1:1").WrapText = True
        'Cabecera
        XLS.Cells(1, 1) = "Tipo de Muestra"
        XLS.Cells(1, 2) = "Tipo de Análisis"
        XLS.Cells(1, 3) = "Cliente"
        XLS.Cells(1, 4) = "Muestras"
        XLS.Cells(1, 5) = "Importe"
        XLS.Range("A1:E1").Interior.ColorIndex = 6
        ' Datos
        i = 2
        Do
            XLS.Cells(i, 1) = rs(0)
            XLS.Cells(i, 2) = rs(1)
            XLS.Cells(i, 3) = rs(2)
            XLS.Cells(i, 4) = rs(6)
'            XLS.Cells(i, 6) = Format(rs(5), "dd/mm/yyyy")
'            XLS.Range(XLS.Cells(i, 6), XLS.Cells(i, 6)).HorizontalAlignment = xlRight
           ' Importe total
           'MDET
            consulta = "select sum(dm.precio) " & _
                        "  from muestras m left join docs_pago_muestras dm " & _
                        "    on m.id_muestra = dm.muestra_id and dm.determinacion_id = 0 " & _
                        " where m.tipo_muestra_id = " & rs(3) & _
                        "   and m.tipo_analisis_id = " & rs(4) & _
                        "   and m.cliente_id = " & rs(5) & _
                        "   and m.fecha_recepcion >= '" & Format(fdesde, "yyyy-mm-dd") & "' " & _
                        "   and m.fecha_recepcion <= '" & Format(fhasta, "yyyy-mm-dd") & "' " & _
                        "   and m.anulada = 0 " & _
                        "   and m.documento_pago <> 0 "
            Set rs_total = datos_bd(consulta)
            If Not IsNull(rs_total(0)) Then
                XLS.Range(XLS.Cells(i, 5), XLS.Cells(i, 5)).NumberFormat = "0.00"
                XLS.Cells(i, 5) = CSng(Replace(rs_total(0), ".", ","))
            End If
            If pb.Value < pb.Max Then
                pb = pb + 1
            End If
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
        XLS.Range("1:1").AutoFilter
        Me.MousePointer = 0
        pb.visible = False
        lbltitulo.visible = False
        XLA.visible = True
   End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cmdExcel_Click_Error:

    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExcel_Click of Formulario frmInformeFacturacion")
End Sub

Private Sub cmdGenerar2_Click()
    Dim rs As ADODB.Recordset
    Dim cliente As String
   On Error GoTo cmdExcel_Click_Error

    If cmbclientes2.getPK_SALIDA <> 0 Then
        cliente = "   and d.cliente_id = " & cmbclientes2.getPK_SALIDA
    End If
    consulta = "select d.id_doc,d.numero,year(d.fecha_factura),d.fecha_factura,c.nombre,sum(dc.precio) " & _
               "  from docs_pago_conceptos dc, docs_pago d, clientes c " & _
               " where dc.fecha >= '" & Format(fdesde2, "yyyy-mm-dd") & "' " & _
               "   and dc.fecha <= '" & Format(fhasta2, "yyyy-mm-dd") & "' " & _
               "   and dc.doc_id = d.id_doc " & _
               cliente & _
               "   and d.cliente_id = c.id_cliente " & _
               "   and d.anulado = 0 " & _
               " group by d.id_doc,d.numero,year(d.fecha_factura),d.fecha_factura,c.nombre"
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        MsgBox "No existen registros para la selección.", vbInformation, App.Title
        Exit Sub
    Else
        Me.MousePointer = 11
        lbltitulo.visible = True
        pb.visible = True
        pb.min = 1
        pb.Max = rs.RecordCount
        pb.Value = 1
        Dim i As Integer
        Dim rs_total As ADODB.Recordset
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
'        XLA.Visible = True
        XLS.Range("1:1").HorizontalAlignment = xlCenter
        XLS.Range("1:1").VerticalAlignment = xlCenter
        XLS.Range("A1:B1").ColumnWidth = 10
        XLS.Range("C1:C1").ColumnWidth = 10
        XLS.Range("D1:D1").ColumnWidth = 50
        XLS.Range("E1:E1").ColumnWidth = 50
        XLS.Range("F1:F1").ColumnWidth = 15
        XLS.Range("G1:G1").ColumnWidth = 15
        XLS.Range("1:1").WrapText = True
        'Cabecera
        XLS.Cells(1, 1) = "Número"
        XLS.Cells(1, 2) = "Año"
        XLS.Cells(1, 3) = "Fecha"
        XLS.Cells(1, 4) = "Cliente"
        XLS.Cells(1, 5) = "Concepto"
        XLS.Cells(1, 6) = "Imp.Concepto"
        XLS.Cells(1, 7) = "Imp.Total"
        XLS.Range("A1:G1").Interior.ColorIndex = 6
        ' Datos
        i = 2
        Do
            XLS.Cells(i, 1) = rs(1)
            XLS.Cells(i, 2) = rs(2)
            XLS.Range(XLS.Cells(i, 3), XLS.Cells(i, 3)).HorizontalAlignment = xlRight
            XLS.Range(XLS.Cells(i, 3), XLS.Cells(i, 3)).NumberFormat = "dd-mm-yyyy"
            XLS.Cells(i, 3) = Format(rs(3), "yyyy-mm-dd")
            XLS.Cells(i, 4) = rs(4)
            XLS.Cells(i, 5) = "Total Factura"
            XLS.Range(XLS.Cells(i, 7), XLS.Cells(i, 7)).NumberFormat = "0.00"
            XLS.Cells(i, 7) = CSng(Replace(rs(5), ".", ","))
            XLS.Range(XLS.Cells(i, 1), XLS.Cells(i, 7)).Font.bold = True
            ' Detalle de los conceptos
            Dim rs_conceptos As ADODB.Recordset
            Dim oDoc_Conceptos As New clsDocs_pago_conceptos
            Set rs_conceptos = oDoc_Conceptos.ConceptosDocumento(rs(0))
            If rs_conceptos.RecordCount > 0 Then
                Do
                    i = i + 1
                    XLS.Cells(i, 1) = rs(1)
                    XLS.Cells(i, 2) = rs(2)
                    XLS.Range(XLS.Cells(i, 3), XLS.Cells(i, 3)).HorizontalAlignment = xlRight
                    XLS.Range(XLS.Cells(i, 3), XLS.Cells(i, 3)).NumberFormat = "dd-mm-yyyy"
                    XLS.Cells(i, 3) = Format(rs_conceptos("FECHA"), "yyyy-mm-dd")
                    XLS.Cells(i, 4) = rs(4)
                    XLS.Cells(i, 5) = rs_conceptos("DESCRIPCION")
                    XLS.Range(XLS.Cells(i, 6), XLS.Cells(i, 6)).NumberFormat = "0.00"
                    XLS.Cells(i, 6) = CSng(Replace(rs_conceptos("PRECIO"), ".", ","))
                    rs_conceptos.MoveNext
                Loop Until rs_conceptos.EOF
            End If
            If pb.Value < pb.Max Then
                pb = pb + 1
            End If
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
        XLS.Range("1:1").AutoFilter
        Me.MousePointer = 0
        pb.visible = False
        lbltitulo.visible = False
        XLA.visible = True
   End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cmdExcel_Click_Error:

    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExcel_Click of Formulario frmInformeFacturacion")

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 20
    Me.top = 20
    cargar_clientes
    fdesde = Date
    fhasta = Date
    fdesde2 = Date
    fhasta2 = Date
End Sub
Public Sub cargar_clientes()
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbclientes2, New clsCliente, 0, frmClientes, ""
'    cargar_combo cmbClientes, New clsCliente
'    cargar_combo cmbclientes2, New clsCliente
End Sub

