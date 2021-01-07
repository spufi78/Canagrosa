VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmFacturacionInformeMuestrasConceptos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de facturación de muestras y conceptos"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   17610
   Icon            =   "frmFacturacionInformeMuestrasConceptos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8940
   ScaleWidth      =   17610
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   5175
      TabIndex        =   13
      Top             =   3960
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
         TabIndex        =   14
         Top             =   225
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exportar a Excel"
      Height          =   870
      Left            =   60
      Picture         =   "frmFacturacionInformeMuestrasConceptos.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8010
      Width           =   1590
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   16485
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   1050
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
      Height          =   1080
      Left            =   45
      TabIndex        =   4
      Top             =   360
      Width           =   17505
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   180
         TabIndex        =   17
         Top             =   540
         Width           =   2355
         Begin VB.OptionButton opT 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Facturas"
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   19
            Top             =   180
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton opT 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Albaranes"
            Height          =   195
            Index           =   1
            Left            =   1170
            TabIndex        =   18
            Top             =   180
            Width           =   1590
         End
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Conceptos"
         Height          =   195
         Index           =   1
         Left            =   1665
         TabIndex        =   8
         Top             =   270
         Width           =   1590
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Muestras"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   270
         Value           =   -1  'True
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   4830
         TabIndex        =   9
         Top             =   450
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
         Format          =   51642369
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   6675
         TabIndex        =   10
         Top             =   450
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
         Format          =   51642369
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbclientesFact 
         Height          =   345
         Left            =   9405
         TabIndex        =   15
         Top             =   225
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   609
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   885
         Left            =   16290
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   135
         Width           =   1155
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   9405
         TabIndex        =   20
         Top             =   585
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   8550
         TabIndex        =   21
         Top             =   615
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clien. Fact."
         Height          =   195
         Index           =   0
         Left            =   8550
         TabIndex        =   16
         Top             =   255
         Width           =   825
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha desde"
         Height          =   195
         Index           =   6
         Left            =   3780
         TabIndex        =   12
         Top             =   510
         Width           =   930
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   4
         Left            =   6195
         TabIndex        =   11
         Top             =   495
         Width           =   405
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6195
      Left            =   45
      TabIndex        =   1
      Top             =   1755
      Width           =   17490
      _ExtentX        =   30850
      _ExtentY        =   10927
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informe de facturación de muestras y conceptos"
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
      Height          =   330
      Index           =   4
      Left            =   45
      TabIndex        =   6
      Top             =   0
      Width           =   17520
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
      TabIndex        =   5
      Top             =   1440
      Width           =   17460
   End
End
Attribute VB_Name = "frmFacturacionInformeMuestrasConceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkbanos_Click()
    cabecera
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdMarcar_Click()
   On Error GoTo cmdMarcar_Click_Error

    If lista.ListItems.Count = 0 Then
        MsgBox "No existen registros en la lista.", vbInformation, App.Title
        Exit Sub
    Else
        Frame6.visible = True
        Me.MousePointer = 11
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLA.visible = False
        XLS.Range("1:1").HorizontalAlignment = xlCenter
        XLS.Range("1:1").VerticalAlignment = xlCenter
        'Cabecera
        Dim i As Integer
        For i = 2 To lista.ColumnHeaders.Count
            XLS.Cells(1, i - 1) = lista.ColumnHeaders(i).Text
        Next
        If opTipo(0).Value = True Then
            XLS.Range(XLS.Cells(1, 3), XLS.Cells(1, 6)).ColumnWidth = 35
        Else
            XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 3)).ColumnWidth = 35
        End If
'        XLS.Range(XLS.Cells(1, 4), XLS.Cells(1, 4)).ColumnWidth = 14
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, lista.ColumnHeaders.Count - 1)).Interior.ColorIndex = 6
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, lista.ColumnHeaders.Count - 1)).Interior.Pattern = xlSolid
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, lista.ColumnHeaders.Count - 1)).Font.ColorIndex = 3
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, lista.ColumnHeaders.Count - 1)).AutoFilter
        ' Datos
        Dim j As Integer
        For i = 1 To lista.ListItems.Count
            For j = 2 To lista.ColumnHeaders.Count
                If j = lista.ColumnHeaders.Count Then
                    XLS.Cells(i + 1, j - 1) = CSng(lista.ListItems(i).SubItems(j - 1))
                    XLS.Cells(i + 1, j - 1).Style = "currency"
                Else
                    If InStr(1, lista.ColumnHeaders(j).Text, "Fecha") Then
                        XLS.Cells(i + 1, j - 1) = Format(lista.ListItems(i).SubItems(j - 1), "mm/dd/yyyy")
                    Else
                        XLS.Cells(i + 1, j - 1) = lista.ListItems(i).SubItems(j - 1)
                    End If
                End If
            Next
        Next
        Frame6.visible = False
        XLA.visible = True
        Me.MousePointer = 0
    End If

   On Error GoTo 0
   Exit Sub

cmdMarcar_Click_Error:
    Frame6.visible = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMarcar_Click of Formulario frmFacturacionInformeMuestrasConceptos", vbCritical, App.Title
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    fdesde = "01/01/" & Year(Date)
    fhasta = Date
    cabecera
    llenar_combo cmbclientesFact, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
    opTipo(0).Value = True
End Sub
Private Sub cabecera()
    lista.ListItems.Clear
    lista.ColumnHeaders.Clear
    If opTipo(0).Value = True Then
        With lista.ColumnHeaders
            .Add , , "ID_MUESTRA", 1, lvwColumnLeft
            .Add , , "General", 800, lvwColumnCenter
            .Add , , "Año", 800, lvwColumnCenter
            .Add , , "Familia", 2300, lvwColumnLeft
            .Add , , "Tipo Muestra", 2300, lvwColumnLeft
            .Add , , "Cliente Fact", 2300, lvwColumnLeft
            .Add , , "Cliente", 2300, lvwColumnLeft
            .Add , , "Fecha Recepción", 1050, lvwColumnCenter
            .Add , , "Fecha Cierre", 1050, lvwColumnCenter
            .Add , , "Num.Factura", 1050, lvwColumnCenter
            .Add , , "Tipo", 1000, lvwColumnCenter
            .Add , , "Fecha Factura", 1050, lvwColumnCenter
            .Add , , "Importe", 1100, lvwColumnRight
        End With
    Else
        With lista.ColumnHeaders
            .Add , , "ID_DOC", 1, lvwColumnLeft
            .Add , , "Familia", 4500, lvwColumnLeft
            .Add , , "Cliente Fact", 3000, lvwColumnLeft
            .Add , , "Cliente", 3000, lvwColumnLeft
            .Add , , "Fecha", 1100, lvwColumnCenter
            .Add , , "Num.Factura", 1100, lvwColumnCenter
            .Add , , "Tipo", 1100, lvwColumnCenter
            .Add , , "Fecha Factura", 1100, lvwColumnCenter
            .Add , , "Total", 1200, lvwColumnRight
        End With
    End If
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim consulta As String
    lista.ListItems.Clear
    Dim filtro As String
    filtro = " AND c.FECHA_FACTURA >= '" & Format(fdesde, "yyyy-mm-dd") & "' AND c.FECHA_FACTURA <= '" & Format(fhasta, "yyyy-mm-dd") & "'"
    'M1373-F
    Dim tipo As String
    If opT(0).Value = True Then
        tipo = "  and c.TIPO > 1 "
    Else
        tipo = "  and c.TIPO = 1 "
    End If
    If cmbclientesFact.getTEXTO <> "" Then
        If opTipo(1).Value = True And opT(1).Value = True Then
            filtro = filtro & " and dd.CLIENTE_ID_FACTURA = " & cmbclientesFact.getPK_SALIDA
        Else
            filtro = filtro & " and c.CLIENTE_ID_FACTURA = " & cmbclientesFact.getPK_SALIDA
        End If
    End If
    If cmbclientes.getTEXTO <> "" Then
        If opTipo(1).Value = True And opT(1).Value = True Then
            filtro = filtro & " and dd.CLIENTE_ID = " & cmbclientesFact.getPK_SALIDA
        Else
            filtro = filtro & " and c.CLIENTE_ID = " & cmbclientesFact.getPK_SALIDA
        End If
    End If
    
    If opTipo(0).Value = True Then
        'MDET
        consulta = "select a.ID_MUESTRA, a.ID_GENERAL,a.ANNO, f.NOMBRE,tm.NOMBRE, cli.NOMBRE, clim.NOMBRE, a.FECHA_RECEPCION,a.FECHA_CIERRE, " & _
                   "       concat(c.NUMERO,'/',year(c.FECHA_FACTURA)) as NUMERO_FACTURA, " & _
                   "       CASE c.TIPO WHEN 1 THEN 'ALBARAN' WHEN 2 THEN 'FACTURA' ELSE 'ABONO' END  AS TIPO, " & _
                   "       c.FECHA_FACTURA , b.PRECIO " & _
                   "  from muestras a, docs_pago_muestras b, docs_pago c, tipos_muestra tm, familias f, clientes cli, clientes clim " & _
                   " where c.ANULADO = 0 " & _
                   "  and a.ID_MUESTRA = b.MUESTRA_ID " & _
                   "  and b.MUESTRA_ID <> 0 AND b.DETERMINACION_ID = 0 " & _
                   "  and b.DOC_ID = c.ID_DOC " & _
                   "  and a.TIPO_MUESTRA_ID = tm.ID_TIPO_MUESTRA " & _
                   "  and tm.FAMILIA_ID = f.ID_FAMILIA " & _
                   "  and c.CLIENTE_ID_FACTURA  = cli.ID_CLIENTE " & _
                   "  and c.CLIENTE_ID  = clim.ID_CLIENTE " & _
                   filtro & _
                   tipo & _
                   " order by a.ANNO,a.ID_GENERAL"
    Else
        consulta = "select c.ID_DOC,f.NOMBRE,cli.NOMBRE, clim.NOMBRE, b.FECHA, " & _
                    " concat(c.NUMERO,' (',year(c.FECHA_FACTURA),')') as NUMERO_FACTURA, " & _
                    " CASE c.TIPO WHEN 1 THEN 'ALBARAN' WHEN 2 THEN 'FACTURA' ELSE 'ABONO' END  AS TIPO, " & _
                    " c.FECHA_FACTURA , b.total " & _
                    " from docs_pago_conceptos b, docs_pago c, familias f, clientes cli, clientes clim " & _
                    IIf(opT(1).Value = True, ", docs_pago dd", "") & _
                    " where c.ANULADO = 0 " & _
                    " and b.DOC_ID = c.ID_DOC " & _
                    " and b.FAMILIA_ID = f.ID_FAMILIA " & _
                    " and c.CLIENTE_ID_FACTURA  = cli.ID_CLIENTE " & _
                   "  and c.CLIENTE_ID  = clim.ID_CLIENTE " & _
                   IIf(opT(1).Value = True, " and c.PAGADO = dd.ID_DOC and c.PAGADO <> 0 ", "") & _
                   filtro & _
                   tipo & _
                    " order by c.FECHA_FACTURA"
    End If
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    Dim total As Integer
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
            If opTipo(0).Value = True Then
                With lista.ListItems.Add(, , rs.Fields(0))
                    .SubItems(1) = rs.Fields(1)
                    .SubItems(2) = rs.Fields(2)
                    .SubItems(3) = rs.Fields(3)
                    .SubItems(4) = rs.Fields(4)
                    .SubItems(5) = rs.Fields(5) ' Cliente Fact
                    .SubItems(6) = rs.Fields(6) ' Cliente
                    .SubItems(7) = rs.Fields(7) '
                    If Not IsNull(rs(8)) Then
                        .SubItems(8) = rs.Fields(8) '
                    End If
                    .SubItems(9) = rs.Fields(9) '
                    .SubItems(10) = rs.Fields(10) '
                    .SubItems(11) = rs.Fields(11)
                    .SubItems(12) = moneda(rs.Fields(12)) ' precio
                End With
            Else
                With lista.ListItems.Add(, , rs.Fields(0)) ' fam
                    .SubItems(1) = rs.Fields(1) ' cli_fact
                    .SubItems(2) = rs.Fields(2) ' cli
                    .SubItems(3) = rs.Fields(3) ' fecha
                    .SubItems(4) = rs.Fields(4) ' n.factura
                    .SubItems(5) = rs.Fields(5) ' tipo
                    .SubItems(6) = rs.Fields(6) ' f.factura
                    .SubItems(7) = rs.Fields(7) ' f.factura
                    .SubItems(8) = moneda(rs.Fields(8)) ' total
                End With
            End If
            rs.MoveNext
        Wend
        lblMsg.Caption = "Listado entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy")
    Else
        lblMsg.Caption = "No existen registros con esos criterios."
    End If
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras.", vbCritical, Err.Description
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

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        If opTipo(0).Value = True Then
            gmuestra = lista.ListItems(lista.selectedItem.Index).Text
            frmVerMuestra.Show 1
        End If
    End If
End Sub

Private Sub opTipo_Click(Index As Integer)
    cabecera
End Sub
