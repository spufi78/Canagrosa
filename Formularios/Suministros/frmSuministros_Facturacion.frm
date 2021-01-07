VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmSuministros_Facturacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lotes de Suministros pendientes de facturación"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15465
   Icon            =   "frmSuministros_Facturacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   15465
   Begin VB.Frame frmBusqueda 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      TabIndex        =   17
      Top             =   360
      Width           =   15390
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   900
         TabIndex        =   18
         Top             =   270
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   10170
         TabIndex        =   20
         Top             =   270
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
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   12285
         TabIndex        =   21
         Top             =   285
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
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta el"
         Height          =   195
         Index           =   4
         Left            =   11610
         TabIndex        =   23
         Top             =   315
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde el"
         Height          =   195
         Index           =   2
         Left            =   9360
         TabIndex        =   22
         Top             =   345
         Width           =   645
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clientes"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   19
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.Frame frmFacturaManual 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informar Número de Factura en los lotes marcados"
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
      Height          =   885
      Left            =   3870
      TabIndex        =   9
      Top             =   3870
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txtnumero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   855
         TabIndex        =   12
         Top             =   330
         Width           =   1185
      End
      Begin XtremeSuiteControls.PushButton cmdCerrarFacturaManual 
         Height          =   435
         Left            =   6210
         TabIndex        =   10
         Top             =   270
         Width           =   1410
         _Version        =   851970
         _ExtentX        =   2487
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   5
         Picture         =   "frmSuministros_Facturacion.frx":030A
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   435
         Left            =   4770
         TabIndex        =   11
         Top             =   270
         Width           =   1410
         _Version        =   851970
         _ExtentX        =   2487
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Informar"
         Appearance      =   5
         Picture         =   "frmSuministros_Facturacion.frx":6B6C
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   360
         Left            =   3721
         TabIndex        =   14
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   2004
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196612
         OrigLeft        =   3510
         OrigTop         =   315
         OrigRight       =   3750
         OrigBottom      =   660
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   405
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   225
         Index           =   1
         Left            =   2295
         TabIndex        =   15
         Top             =   390
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdCrear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear Albaran"
      Height          =   915
      Index           =   1
      Left            =   11205
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSuministros_Facturacion.frx":D3CE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8970
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Cliente"
      Height          =   330
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8970
      Width           =   1410
   End
   Begin VB.CommandButton cmdCrear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear &Factura"
      Height          =   915
      Index           =   2
      Left            =   12645
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmSuministros_Facturacion.frx":E010
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8970
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin VB.CommandButton cmdMarcar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   330
      Left            =   1470
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8970
      Width           =   1410
   End
   Begin VB.CommandButton cmdDesmarcar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   330
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8970
      Width           =   1410
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   915
      Left            =   14085
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8970
      Width           =   1365
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7800
      Left            =   45
      TabIndex        =   1
      Top             =   1125
      Width           =   15405
      _ExtentX        =   27173
      _ExtentY        =   13758
      View            =   3
      LabelEdit       =   1
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
   Begin XtremeSuiteControls.PushButton cmdNoFacturable 
      Height          =   480
      Left            =   6255
      TabIndex        =   24
      Top             =   9270
      Width           =   2535
      _Version        =   851970
      _ExtentX        =   4471
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Marcar como Factura Manual"
      Appearance      =   5
      Picture         =   "frmSuministros_Facturacion.frx":EC52
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre el lote para ver el detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   5700
      TabIndex        =   2
      Top             =   8970
      Width           =   3735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Lotes de Suministros pendientes de facturación"
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
      TabIndex        =   0
      Top             =   0
      Width           =   15405
   End
End
Attribute VB_Name = "frmSuministros_Facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClientes_change()
    cargar_lista
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdCerrarFacturaManual_Click()
    frmFacturaManual.visible = False
End Sub

Private Sub cmdCrear_Click(Index As Integer)
    Dim strcadena As String
    If contar_marcados = 0 Then
        MsgBox "Debe seleccionar algún lote para facturar.", vbInformation, App.Title
        Exit Sub
    End If
    Dim TIPO As String
    If Index = 1 Then
        TIPO = "Albaranes"
    Else
        TIPO = "Facturas"
    End If
    strcadena = "Va a generar " & TIPO & " para  " & contar_marcados & " lotes. ¿Desea continuar?"
    If MsgBox(strcadena, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        generar_documentos (Index)
        Me.MousePointer = 0
    End If
    Call cargar_lista
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdNoFacturable_Click()
    frmFacturaManual.visible = Not frmFacturaManual.visible
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    Dim LOTE As Integer
    If lista.ListItems.Count > 0 Then
        LOTE = lista.ListItems(lista.selectedItem.Index).SubItems(6)
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).SubItems(6) = LOTE Then
                lista.ListItems(i).Checked = True
                lista.ListItems(i).Selected = True
                lista.ListItems(i).EnsureVisible
            End If
        Next
    End If
End Sub

Private Sub fdesde_Change()
    cargar_lista
End Sub

Private Sub fhasta_Change()
    cargar_lista
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    txtanno = Year(Date)
    cambiar.Max = Year(Date)
    cambiar.min = 2004
    fdesde = "01/01/2004"
    fhasta = Date
    cabecera
    rellenar_clientes
    cargar_lista
End Sub
Private Sub rellenar_clientes()
    Dim consulta As String
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        consulta = "SELECT DISTINCT C.ID_CLIENTE,C.NOMBRE " & _
                   "  FROM SUMINISTROS_LOTES_CLIENTES SL, CLIENTES C " & _
                   " WHERE SL.CLIENTE_ID = C.ID_CLIENTE " & _
                   "   AND SL.DOC_ID = 0 "
        With cmbclientes
            .setCONN = conn
            .setQUERY = consulta
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "CLIENTES"
            .setDESCRIPCION = "Clientes"
            .setPK = "C.ID_CLIENTE"
            .setFILTRO = ""
            .setCAMPO = "C.NOMBRE"
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmClientes
        End With
    End If
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
End Sub

Private Function inlote() As String
    Dim s As String
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            s = s & lista.ListItems(i).SubItems(5) & ","
        End If
    Next
    s = Left(s, Len(s) - 1)
    inlote = s
End Function
'Private Function incliente() As String
'    Dim s As String
'    Dim i As Integer
'    For i = 1 To lista.ListItems.Count
'        If lista.ListItems(i).Checked = True Then
'            s = s & lista.ListItems(i).SubItems(7) & ","
'        End If
'    Next
'    s = Left(s, Len(s) - 1)
'    incliente = s
'End Function
Private Sub generar_documentos(TIPO_DOCUMENTO As Integer)
   On Error GoTo generar_documentos_Error

    log ("Comiento generación de facturación de suministros")
    Dim num_doc As Long
    Dim cliente_ant As Long
    Dim pedido_ant As Long
    Dim total_doc As Integer
    total_doc = 0
    cliente_ant = 0
    pedido_ant = 0
    'cIVA
'    Dim oParametros As New clsParametros
'    Dim IVA As Integer
'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
    ' Planificacion
    Dim oSL As New clsSuministros_lotes
    Dim oSLC As New clsSuministros_lotes_clientes
    Dim oDocPago As New clsDocs_pago
    Dim oConcepto As New clsDocs_pago_conceptos
    Dim rs As ADODB.Recordset
    Dim oCliente As New clsCliente
    With oSL
        Set rs = .Listado_Para_Factura(inlote)
        If rs.RecordCount = 0 Then
            Exit Sub
        Else
            Do
                If rs(0) <> cliente_ant Or rs(1) <> pedido_ant Then
                    If cliente_ant <> 0 Then
                        oDocPago.Informar_total_factura (num_doc)
                        oSLC.doc_pago_pedido inlote, cliente_ant, pedido_ant, CLng(num_doc)
                    End If
                    With oDocPago
                         .setTIPO = TIPO_DOCUMENTO
                         .setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
                         .setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
                         .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                         .setCLIENTE_ID = rs(0)
                         .setCLIENTE_ID_FACTURA = rs(0)
                         .setTOTAL = "0.00"
                         .setDESCUENTO = "0.00"
'                         If TIPO_DOCUMENTO = 2 Then
'                             .setIVA = IVA
'                         Else
'                             .setIVA = 0
'                         End If
                         .setPAGADO = 0
                         .setANULADO = 0
                         .setFACTURA_CONCEPTOS = 1
                         .setPEDIDO_ID = rs(1)
                         oCliente.CargaCliente rs(0)
                         oDocPago.setFP_ID = oCliente.getFP_ID
                         ' Insertamos el documento de pago
                         num_doc = .InsertarDocPago
                         If num_doc = 0 Then
                             MsgBox "Error al generar las facturas, contacte con mantenimiento.", vbCritical, App.Title
                             Exit Sub
                         End If
                         ' Informar el documento de pago en alodine_planificación
                         total_doc = total_doc + 1
                    End With
                    cliente_ant = rs(0)
                    pedido_ant = rs(1)
                End If
                ' Insertamos los conceptos
                With oConcepto
                        .setDOC_ID = num_doc
                        .setDESCRIPCION = rs(3)
                        .setFECHA = Format(rs(2), "yyyy-mm-dd")
                        .setFAMILIA_ID = familia.suministros ' Familia de alodine por defecto
                        .setAPARTADO = 0
                        .setDTO = 0
                        
                        .setCANTIDAD = rs(4)
                        .setPRECIO = Replace(rs(5), ",", ".")
                        .setSUBTOTAL = Replace(rs(6), ",", ".")
                        .setTOTAL = Replace(rs(6), ",", ".")
                        
                        If .Insertar = False Then
                           Exit Sub
                        End If
                End With
                rs.MoveNext
            Loop Until rs.EOF = True
            If cliente_ant <> 0 Then
                oDocPago.Informar_total_factura (num_doc)
                oSLC.doc_pago_pedido inlote, cliente_ant, pedido_ant, CLng(num_doc)
            End If
        End If
    End With
    Set oDocPago = Nothing
    Dim sTIPO As String
    If TIPO_DOCUMENTO = 1 Then
        sTIPO = "Albaran"
    Else
        sTIPO = "Factura"
    End If
    log ("Final generación de facturación suministros")
    If total_doc = 1 Then
        MsgBox "Se ha registrado 1 " & sTIPO & ".", vbOKOnly + vbInformation, App.Title
    Else
        MsgBox "Se han registrado " & total_doc & " " & sTIPO & "s.", vbOKOnly + vbInformation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

generar_documentos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_documentos of Formulario frmSuministros_Facturacion"
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
Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Cliente", 3600, lvwColumnLeft
        .Add , , "Producto", 3600, lvwColumnLeft
        .Add , , "Lote", 900, lvwColumnCenter
        .Add , , "F.Alta", 1050, lvwColumnCenter
        .Add , , "Caducidad", 1050, lvwColumnCenter
        .Add , , "ID", 600, lvwColumnCenter
        .Add , , "CLIENTE_ID", 1, lvwColumnCenter
        .Add , , "Num.Botes", 700, lvwColumnCenter
        .Add , , "Precio", 900, lvwColumnCenter
        .Add , , "Pedido", 2500, lvwColumnLeft
        .Add , , "PEDIDO_ID", 1, lvwColumnLeft
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oSM As New clsSuministros_lotes
    lista.ListItems.Clear
    Set rs = oSM.Listado_Pendientes_Facturar(IIf(cmbclientes.getTEXTO = "", 0, cmbclientes.getPK_SALIDA), fdesde, fhasta)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = Format(rs(5), "0000")
             .SubItems(6) = Format(rs(6), "0000")
             
             .SubItems(7) = rs(7)
             .SubItems(8) = moneda(rs(8))
             If Not IsNull(rs(9)) Then
                .SubItems(9) = Trim(rs(9))
             End If
             .SubItems(10) = rs(10)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oSM = Nothing
End Sub
Private Sub actualizarLista()
    Dim rs As New ADODB.Recordset
    Dim oSM As New clsSuministros_lotes
    Set rs = oSM.Listado_Pendientes_Facturar_ID(lista.ListItems(lista.selectedItem.Index).SubItems(5), lista.ListItems(lista.selectedItem.Index).SubItems(6))
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems(lista.selectedItem.Index)
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = Format(rs(5), "0000")
             .SubItems(6) = Format(rs(6), "0000")
             
             .SubItems(7) = rs(7)
             .SubItems(8) = moneda(rs(8))
             .SubItems(9) = Trim(rs(9))
             .SubItems(10) = rs(10)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oSM = Nothing
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        frmSuministros_Lote.PK = lista.ListItems(lista.selectedItem.Index).SubItems(5)
        frmSuministros_Lote.Show 1
        actualizarLista
    End If
End Sub
Public Function contar_marcados() As Integer
    Dim i As Integer
    contar_marcados = 0
    For i = 1 To lista.ListItems.Count
       If lista.ListItems(i).Checked = True Then
        contar_marcados = contar_marcados + 1
      End If
    Next
End Function
Private Sub PushButton2_Click()
    Dim strcadena As String
   On Error GoTo PushButton2_Click_Error

    If contar_marcados = 0 Then
        MsgBox "Debe seleccionar algún lote.", vbCritical, App.Title
        Exit Sub
    End If
    If txtnumero = "" Then
        MsgBox "Debe informar el número de factura.", vbCritical, App.Title
        Exit Sub
    End If
    If txtanno = "" Then
        MsgBox "Debe informar el año de la factura.", vbCritical, App.Title
        Exit Sub
    End If
    Dim oDocPago As New clsDocs_pago
    Dim ID As Long
    ID = oDocPago.recuperarIdPorNumero(txtnumero, txtanno)
    
    If ID = 0 Then
        MsgBox "El número de factura indicado NO EXISTE.", vbCritical, App.Title
        Exit Sub
    Else
        Dim oSLC As New clsSuministros_lotes_clientes
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                oSLC.doc_pago lista.ListItems(i).SubItems(5), lista.ListItems(i).SubItems(6), ID
            End If
        Next
        MsgBox "Lotes actualizados correctamente.", vbInformation, App.Title
        frmFacturaManual.visible = False
        cargar_lista
    End If

   On Error GoTo 0
   Exit Sub

PushButton2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PushButton2_Click of Formulario frmSuministros_Facturacion"
End Sub
