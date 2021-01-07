VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmListadoCompras 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Solicitudes de Oferta"
   ClientHeight    =   8070
   ClientLeft      =   2070
   ClientTop       =   1395
   ClientWidth     =   13305
   Icon            =   "frmListadoCompras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   13305
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Imprimir"
      Height          =   870
      Left            =   5700
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7140
      Width           =   1050
   End
   Begin VB.CheckBox chkLogo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Imprimir sin logotipo y pie "
      Height          =   195
      Left            =   3450
      TabIndex        =   11
      Top             =   7455
      Width           =   2160
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7140
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2265
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7140
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7140
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12210
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7155
      Width           =   1050
   End
   Begin VB.Frame Frame1 
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
      Height          =   1050
      Left            =   45
      TabIndex        =   1
      Top             =   315
      Width           =   13230
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10845
         TabIndex        =   5
         Top             =   450
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   780
         Left            =   11880
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cmbTipoSolicitud 
         Height          =   315
         Left            =   1380
         TabIndex        =   3
         Top             =   405
         Width           =   9165
         _ExtentX        =   16166
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Sol. Oferta"
         Height          =   195
         Index           =   0
         Left            =   105
         TabIndex        =   4
         Top             =   450
         Width           =   1125
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5700
      Left            =   45
      TabIndex        =   7
      Top             =   1395
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   10054
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
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Listado de Solicitudes de Oferta"
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
      Index           =   4
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   13245
   End
End
Attribute VB_Name = "frmListadoCompras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public lineas_factura As Integer
Public ancho_linea As Integer
Private Sub chkTodos_Click()
    If chkTodos.value = Checked Then
        cmbTipoSolicitud.Text = ""
        cmbTipoSolicitud.Enabled = False
    Else
        cmbTipoSolicitud.Enabled = True
    End If
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub

Private Sub buscar()
    Dim IMPORTE As Currency
    Dim base As Currency
    Dim IVA As Currency
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As ADODB.RecordSet
    Dim odoc As New clsDocs_pago
    Me.MousePointer = 11
    If chkTodos.value = Unchecked Then
        If cmbTipoSolicitud.BoundText = "" Then
            MsgBox "Seleccione un cliente.", vbInformation, App.Title
            Me.MousePointer = 0
            Exit Sub
        Else
            Set rs = odoc.AlbaranesDeUnCliente(CLng(cmbTipoSolicitud.BoundText))
        End If
    Else
        Set rs = odoc.Albaranes()
    End If
    If rs.RecordCount <> 0 Then
'        Dim omuestra As New clsDocs_pago_muestras
'        Dim oConcepto As New clsDocs_pago_conceptos
        Do
            With lista.ListItems.Add(, , rs.Fields(1))
                    .SubItems(1) = rs.Fields(2)
                    .SubItems(2) = rs.Fields(3)
                    .SubItems(9) = rs.Fields(0)
'                    IMPORTE = odoc.ImporteTotalDocumento(rs.Fields(0))
                    IMPORTE = rs.Fields(8)
                    If IsNull(rs.Fields("descuento")) Or rs.Fields("descuento") = "0" Then
                        base = IMPORTE
                    Else
                        base = IMPORTE - ((IMPORTE * rs.Fields("descuento")) / 100)
                    End If
                    IVA = (base * rs.Fields("iva")) / 100
                    .SubItems(3) = Format(IMPORTE, "currency")
                    .SubItems(4) = Format(rs.Fields("descuento"), "Standard")
                    .SubItems(5) = Format(base, "currency")
                    .SubItems(6) = rs.Fields("iva")
                    .SubItems(7) = Format(IVA, "currency")
                    .SubItems(8) = Format(base + IVA, "currency")
                    .SubItems(10) = rs(9)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Me.MousePointer = 0
'    Set omuestra = Nothing
'    Set oConcepto = Nothing
    Set odoc = Nothing
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al buscar los albaranes.", vbCritical, Err.Description
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdFacturar_Click()
   On Error GoTo cmdFacturar_Click_Error

    If contar_marcados = 0 Then
        MsgBox "Seleccione algún documento para facturar.", vbExclamation, App.Title
        Exit Sub
    End If
    ' cIVA
    Dim oParametros As New clsParametros
    Dim IVA As Integer
    IVA = recuperaIVA()
    If IVA = 0 Then
        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
        Exit Sub
    End If
    ' Creamos el documento de pago
    Dim ocliente As New clsCliente
    Dim oDocPago As New clsDocs_pago
    If ocliente.CargaCliente(cmbTipoSolicitud.BoundText) = False Then
        MsgBox "Error al cargar los datos del cliente.", vbCritical, App.Title
        Exit Sub
    End If
    Dim i As Integer
    Dim num_doc As Integer
    Dim PEDIDO As Integer
    Dim oConceptos As New clsDocs_pago_conceptos
    Dim omuestras As New clsDocs_pago_muestras
    Dim omuestra As New clsMuestra
    Dim rs As ADODB.RecordSet
    PEDIDO = -1
    ' Verificamos de que tipo vamos a generar la factura
    Dim FMUESTRAS As Boolean
    Dim FCONCEPTOS As Boolean
    Dim FMIXTA As Boolean
    FMUESTRAS = False
    FCONCEPTOS = False
    FMIXTA = False
    Dim tipo As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If lista.ListItems(i).SubItems(10) = 0 Then
                FMUESTRAS = True
            End If
            If lista.ListItems(i).SubItems(10) = 1 Then
                FCONCEPTOS = True
            End If
            If lista.ListItems(i).SubItems(10) = 2 Then
                FMIXTA = True
            End If
        End If
    Next
    If FMIXTA Or (FMUESTRAS And FCONCEPTOS) Then
        tipo = 2
    Else
        If FCONCEPTOS Then
            tipo = 1
        Else
            tipo = 0
        End If
    End If
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            oDocPago.CargarDocumento (lista.ListItems(i).SubItems(9))
            If PEDIDO <> oDocPago.getPEDIDO_ID Then
                PEDIDO = oDocPago.getPEDIDO_ID
                ' Insertamos el detalle de factura por pedido
                With oDocPago
                    .setTIPO = 2 ' FACTURA
                    .setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
                    .setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
                    .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                    .setCLIENTE_ID = cmbTipoSolicitud.BoundText
                    .setTOTAL = "0.00"
                    .setDESCUENTO = Replace(Format("0", "0.00"), ",", ".")
                    .setIVA = IVA
                    .setPAGADO = 0
                    .setANULADO = 0
                    .setFP_ID = ocliente.getFP_ID
                    .setPEDIDO_ID = PEDIDO
                    .setFACTURA_CONCEPTOS = tipo
                    ' Insertamos el documento de pago
                    num_doc = oDocPago.InsertarDocPago
                    If num_doc = 0 Then
                        Exit Sub
                    End If
                End With
            End If
            ' Modificamos el posible albaran de alodine
            execute_bd "UPDATE ALODINE_PLANIFICACION SET DOC_ID = " & num_doc & " WHERE DOC_ID = " & lista.ListItems(i).SubItems(9)
            ' Insertamos el detalle de la factura de conceptos
            Set rs = oConceptos.ConceptosDocumento(lista.ListItems(i).SubItems(9))
            If rs.RecordCount > 0 Then
                Do
                    With oConceptos
                        .setDOC_ID = num_doc
                        .setDESCRIPCION = rs("DESCRIPCION")
                        .setFECHA = Format(rs("FECHA"), "yyyy-mm-dd")
                        .setPRECIO = Replace(Format(rs("precio"), "0.00"), ",", ".")
                        If .Insertar = False Then
                            Exit Sub
                        End If
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            ' Insertamos el detalle de la factura de muestras
            Set rs = omuestras.MuestrasDocumento(lista.ListItems(i).SubItems(9))
            If rs.RecordCount > 0 Then
                Do
                    With omuestras
                        .setDOC_ID = num_doc
                        .setMUESTRA_ID = rs(6)
                        .setORDEN = rs(8)
                        .setCODIGO = rs(9)
                        .setFECHA = Format(rs(2), "yyyy-mm-dd")
                        .setTIPO_ANALISIS = rs(3)
                        .setREFERENCIA_CLIENTE = rs(4)
'                        .setPRECIO = rs(5)
                        .setPRECIO = Replace(Format(rs(5), "0.00"), ",", ".")
                        If .Insertar_doc_pago_muestra(0) = -1 Then
                            MsgBox "Error al insertar en doc_pago_muestra", vbCritical, App.Title
                            Exit Sub
                        End If
                        ' Modificar el documento de pago de la muestra
                        If omuestra.Informar_Documento_Pago(lista.ListItems(i).SubItems(9), 2) = False Then
                            MsgBox "Error al informar el documento de pago, contacte con mantenimiento.", vbCritical, App.Title
                            Exit Sub
                        End If
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            ' Marcamos el albaran como facturado
            oDocPago.Facturar_Albaran lista.ListItems(i).SubItems(9), num_doc
            ' Informamos el total de la factura
            oDocPago.Informar_total_factura (num_doc)
        End If
    Next
    MsgBox "Factura registrada correctamente.", vbInformation, App.Title
    cmdBuscar_Click
   On Error GoTo 0
   Exit Sub

cmdFacturar_Click_Error:

    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure cmdFacturar_Click of Formulario frmListadoCompras")
End Sub

Private Sub cmdImprimir_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If contar_marcados = 0 Then
        Dim oDoc_pago As New clsDocs_pago
        oDoc_pago.generar_factura lista.ListItems(lista.SelectedItem.Index).SubItems(9), chkLogo.value, False, ""
    Else
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                oDoc_pago.generar_factura lista.ListItems(i).SubItems(9), chkLogo.value, True, ""
            End If
        Next
    End If
End Sub
Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Me.SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.Top = 50
    cabecera
    rellenar_clientes
    permisos
End Sub
Public Sub rellenar_clientes()
    Dim odoc As New clsDocs_pago
    Set rsFacturas = odoc.ClientesConAlbaranesPendientes
    Set cmbTipoSolicitud.RowSource = rsFacturas
    cmbTipoSolicitud.ListField = "nombre"
    cmbTipoSolicitud.DataField = "id_cliente"
    cmbTipoSolicitud.BoundColumn = "id_cliente"
    Set odoc = Nothing
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Nº Oferta", 800, lvwColumnLeft)
        .Tag = "Nº Oferta"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo Solicitud", 3700, lvwColumnLeft)
        .Tag = "Tipo Solicitud"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1300, lvwColumnCenter)
        .Tag = "Fecha Solicitud"
    End With
    With lista.ColumnHeaders.Add(, , "Responsable Tecnico", 1200, lvwColumnRight)
        .Tag = "Responsable Técnico"
    End With
End Sub
Private Sub lista_DblClick()
    cmdImprimir_Click
End Sub
Public Function contar_marcados() As Integer
    Dim i As Integer
    Dim cont As Integer
    cont = 0
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            cont = cont + 1
        End If
    Next
    contar_marcados = cont
End Function
Public Sub permisos()
End Sub
