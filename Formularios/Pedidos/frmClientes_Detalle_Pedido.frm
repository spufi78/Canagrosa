VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form frmClientes_Detalle_Pedido 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de documentos de pago asociados al pedido :"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13275
   Icon            =   "frmClientes_Detalle_Pedido.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   13275
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   12060
      TabIndex        =   18
      Top             =   1485
      Width           =   1185
      Begin VB.CommandButton cmdListado2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Listado"
         Height          =   600
         Left            =   90
         Picture         =   "frmClientes_Detalle_Pedido.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   900
         Width           =   1020
      End
      Begin VB.CommandButton cmdImprimir2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Imprimir"
         Height          =   645
         Left            =   90
         Picture         =   "frmClientes_Detalle_Pedido.frx":6B5C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   225
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del pedido"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1110
      Left            =   0
      TabIndex        =   9
      Top             =   360
      Width           =   13230
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   7335
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   11
         Top             =   315
         Width           =   5730
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   7335
         Locked          =   -1  'True
         MaxLength       =   100
         TabIndex        =   10
         Top             =   675
         Width           =   5730
      End
      Begin MSComCtl2.DTPicker txtfecha 
         Height          =   330
         Left            =   1215
         TabIndex        =   12
         Top             =   675
         Width           =   1305
         _ExtentX        =   2302
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
         CalendarTitleBackColor=   12632256
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker txtbaja 
         Height          =   330
         Left            =   4860
         TabIndex        =   13
         Top             =   675
         Width           =   1305
         _ExtentX        =   2302
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
         CalendarTitleBackColor=   12632256
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   1215
         TabIndex        =   21
         Top             =   315
         Width           =   4950
         _ExtentX        =   8731
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   22
         Top             =   360
         Width           =   420
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   240
         Index           =   3
         Left            =   6390
         TabIndex        =   17
         Top             =   360
         Width           =   945
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Pedido"
         Height          =   240
         Index           =   2
         Left            =   180
         TabIndex        =   16
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Index           =   4
         Left            =   6390
         TabIndex        =   15
         Top             =   720
         Width           =   915
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Baja"
         Height          =   240
         Index           =   0
         Left            =   3915
         TabIndex        =   14
         Top             =   720
         Width           =   825
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   12015
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7380
      Width           =   1200
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5790
      Left            =   30
      TabIndex        =   2
      Top             =   1485
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   10213
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
      NumItems        =   0
   End
   Begin VB.Shape Shape1 
      Height          =   1005
      Left            =   6975
      Top             =   7335
      Width           =   5010
   End
   Begin VB.Image imagen 
      Height          =   240
      Left            =   12960
      Picture         =   "frmClientes_Detalle_Pedido.frx":D3AE
      Top             =   45
      Width           =   240
   End
   Begin VB.Label lblrestan 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   0
      Left            =   9315
      TabIndex        =   8
      Top             =   8010
      Width           =   2610
   End
   Begin VB.Label lblrestan 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Restan"
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
      Index           =   1
      Left            =   7020
      TabIndex        =   7
      Top             =   8010
      Width           =   2265
   End
   Begin VB.Label lblpedido 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   9315
      TabIndex        =   6
      Top             =   7380
      Width           =   2610
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Pedido"
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
      Left            =   7020
      TabIndex        =   5
      Top             =   7380
      Width           =   2265
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Facturado"
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
      Index           =   2
      Left            =   7020
      TabIndex        =   4
      Top             =   7695
      Width           =   2265
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Left            =   9315
      TabIndex        =   3
      Top             =   7695
      Width           =   2610
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de documentos de Pago asociados al pedido : "
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
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13230
   End
End
Attribute VB_Name = "frmClientes_Detalle_Pedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cargar_lista()
    Dim oclientes_pedidos As New clsClientes_pedidos
    oclientes_pedidos.Carga (PK)
    cmbTipo.BoundText = oclientes_pedidos.getTIPO_ID
    lbltitulo = lbltitulo & oclientes_pedidos.getCODIGO
    txtDatos(1) = oclientes_pedidos.getCODIGO
    txtDatos(2) = oclientes_pedidos.getDESCRIPCION
    txtFecha = oclientes_pedidos.getFECHA_PEDIDO
    txtbaja = oclientes_pedidos.getFECHA_BAJA
    
    lblpedido.Caption = Format(Replace(oclientes_pedidos.getIMPORTE, ".", ","), "currency")
    Me.Caption = lbltitulo
    Dim IMPORTE As Currency
    Dim BASE As Currency
    Dim IVA As Currency
    
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    Dim oDoc As New clsDocs_pago
    Me.MousePointer = 11
    Set rs = oDoc.Documentos_por_pedido(PK)
    desactivar_controles
    If rs.RecordCount <> 0 Then
        Dim total As Currency
        total = 0
        Do
            With lista.ListItems.Add(, , rs.Fields(1))
                .SubItems(1) = rs.Fields(2)
                .SubItems(2) = rs.Fields(3)
                .SubItems(9) = rs.Fields(0)
                IMPORTE = rs.Fields(8)
                If IsNull(rs.Fields("descuento")) Or rs.Fields("descuento") = "0" Then
                    BASE = IMPORTE
                Else
                    BASE = IMPORTE - ((IMPORTE * rs.Fields("descuento")) / 100)
                End If
                IVA = (BASE * rs.Fields("iva")) / 100
                .SubItems(3) = Format(IMPORTE, "currency")
                .SubItems(4) = Format(rs.Fields("descuento"), "Standard")
                .SubItems(5) = Format(BASE, "currency")
                .SubItems(6) = rs.Fields("iva")
                .SubItems(7) = Format(IVA, "currency")
                .SubItems(8) = Format(BASE + IVA, "currency")
                total = total + .SubItems(5)
                .SubItems(10) = rs(6)
                .SubItems(11) = rs(7)
                .SubItems(12) = rs(10) ' FP_ID
            End With
            rs.MoveNext
        Loop Until rs.EOF
        lblTotal = Format(total, "currency")
        lista_Click
    End If
    lblrestan(0).Caption = Format(lblpedido - lblTotal, "currency")
    Me.MousePointer = 0
    Set oMuestra = Nothing
    Set oDoc = Nothing
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al buscar los documentos del cliente.", vbCritical, Err.Description
End Sub

Private Sub cmdImprimir2_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim cont As Integer
    Dim oDoc_pago As New clsDocs_pago
    If oDoc_pago.validar_previos_documento(lista.ListItems(lista.selectedItem.Index).SubItems(9)) Then
        oDoc_pago.generar_factura lista.ListItems(lista.selectedItem.Index).SubItems(9), False, "", "rptFactura"
    End If
End Sub

Private Sub cmdListado2_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    On Error GoTo fallo
    generar_excel_listado
    Exit Sub
fallo:
    MsgBox "Error al generar el listado de documentos.", vbCritical, Err.Description

End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdsalir_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, DECODIFICADORA.PEDIDOS_CLIENTES_TIPOS
    cabecera
    If PK > 0 Then
        cargar_lista
    End If
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "NºDoc", 700, lvwColumnLeft)
        .Tag = "NºDoc"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 3300, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1300, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Importe", 1200, lvwColumnRight)
        .Tag = "Importe"
    End With
    With lista.ColumnHeaders.Add(, , "Dto. %", 800, lvwColumnCenter)
        .Tag = "Dto. %"
    End With
    With lista.ColumnHeaders.Add(, , "Base", 1200, lvwColumnRight)
        .Tag = "Base"
    End With
    With lista.ColumnHeaders.Add(, , "I.V.A.%", 800, lvwColumnRight)
        .Tag = "I.V.A.%"
    End With
    With lista.ColumnHeaders.Add(, , "Cuota I.V.A.", 1200, lvwColumnRight)
        .Tag = "Cuota I.V.A."
    End With
    With lista.ColumnHeaders.Add(, , "Total", 1200, lvwColumnRight)
        .Tag = "Total"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "TIPO", 1, lvwColumnCenter)
        .Tag = "TIPO"
    End With
    With lista.ColumnHeaders.Add(, , "PAGADO", 1, lvwColumnCenter)
        .Tag = "PAGADO"
    End With
    With lista.ColumnHeaders.Add(, , "FP", 1, lvwColumnCenter)
        .Tag = "FP"
    End With
    
End Sub

Public Sub desactivar_controles()
'    chkLogo.Enabled = True
End Sub

Private Sub lista_Click()
    desactivar_controles
End Sub
Private Sub generar_excel_listado()
'    Dim IMPORTE As Currency
'    Dim base As Currency
    Dim t1 As Currency
    Dim t2 As Currency
    Dim t3 As Currency
'    Dim IVA As Currency
    Dim rs As New ADODB.Recordset
            
            
        rs.Fields.Append "c1", adChar, 1, adFldUpdatable
        rs.Fields.Append "c2", adChar, 10, adFldUpdatable
        rs.Fields.Append "c3", adChar, 50, adFldUpdatable
        rs.Fields.Append "c4", adChar, 15, adFldUpdatable
        rs.Fields.Append "c5", adChar, 15, adFldUpdatable
        rs.Fields.Append "c6", adChar, 10, adFldUpdatable
        rs.Fields.Append "c7", adChar, 10, adFldUpdatable
        rs.Fields.Append "c8", adChar, 15, adFldUpdatable
        rs.Fields.Append "c9", adChar, 15, adFldUpdatable
        rs.Fields.Append "c10", adChar, 1, adFldUpdatable
        rs.Fields.Append "c11", adChar, 15, adFldUpdatable ' Fecha vencimiento
        rs.Open
        Dim i As Integer
        Dim oFP As New clsFP
        t1 = t2 = t3 = 0
        For i = 1 To lista.ListItems.Count
            rs.AddNew
            rs("c1") = "F"
            rs("c2") = lista.ListItems(i).Text
            rs("c3") = Left(lista.ListItems(i).SubItems(1), 50)
            rs("c4") = Format(lista.ListItems(i).SubItems(2), "dd/mm/yyyy")
            rs("c5") = lista.ListItems(i).SubItems(5)
            rs("c6") = lista.ListItems(i).SubItems(4)
            rs("c7") = lista.ListItems(i).SubItems(6)
            rs("c8") = lista.ListItems(i).SubItems(7)
            rs("c9") = lista.ListItems(i).SubItems(8)
            If lista.ListItems(i).SubItems(11) = 0 Then
                rs("c10") = "N"
            Else
                rs("c10") = "S"
            End If
            ' Si es factura,no esta pagada y FP esta informada, calculo el vencimiento
            rs("c11") = ""
            If CInt(lista.ListItems(i).SubItems(12)) <> 0 Then
                 oFP.CARGAR lista.ListItems(i).SubItems(12)
                 rs("c11") = Format(CDate(rs("c4")) + oFP.getDIAS, "dd/mm/yyyy")
            Else
                 rs("c11") = rs("c4")
            End If
            t1 = t1 + lista.ListItems(i).SubItems(5)
            t2 = t2 + lista.ListItems(i).SubItems(7)
            rs.Update
        Next
        Dim XLA As Excel.Application
        Dim XLW As Excel.Workbook
        Dim XLS As Excel.Worksheet
        
        Set XLA = New Excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Listado de facturas"
        XLA.Visible = True
        XLS.Range("1:1").HorizontalAlignment = xlCenter
        XLS.Range("1:1").VerticalAlignment = xlCenter
        XLS.Range("1:1").RowHeight = 30
        XLS.Range("1:1").WrapText = True
        'Cabecera
        XLS.Cells(1, 1) = "Tipo"
        XLS.Cells(1, 2) = "Documento"
        XLS.Cells(1, 3) = "Cliente"
        XLS.Cells(1, 4) = "Fecha"
        XLS.Cells(1, 5) = "Vencim."
        XLS.Cells(1, 6) = "Base"
        XLS.Cells(1, 7) = "Dto"
        XLS.Cells(1, 8) = "Iva"
        XLS.Cells(1, 9) = "Imp.Iva"
        XLS.Cells(1, 10) = "Total"
        XLS.Cells(1, 11) = "Pagada"
        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Range(XLS.Cells(i, 6), XLS.Cells(i, 10)).NumberFormat = "0.00"
            XLS.Cells(i, 1) = rs("c1")
            XLS.Cells(i, 2) = rs("c2")
            XLS.Cells(i, 3) = rs("c3")
            XLS.Cells(i, 4) = rs("c4")
            XLS.Cells(i, 5) = rs("c11")
            XLS.Cells(i, 6) = CSng(rs("c5"))
            XLS.Cells(i, 7) = CSng(rs("c6"))
            XLS.Cells(i, 8) = CSng(rs("c7"))
            XLS.Cells(i, 9) = CSng(rs("c8"))
            XLS.Cells(i, 10) = CSng(rs("c9"))
            XLS.Cells(i, 11) = rs("c10")
            i = i + 1
            rs.MoveNext
          Loop Until rs.EOF
        End If
        
    Set rs = Nothing
End Sub


