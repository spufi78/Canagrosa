VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmListadoAlbaranes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Albaranes pendientes de facturar"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmListadoAlbaranes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   13635
   Begin VB.Frame frmGenera 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3705
      Left            =   1980
      TabIndex        =   16
      Top             =   2430
      Visible         =   0   'False
      Width           =   10320
      Begin VB.TextBox txtDto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6120
         TabIndex        =   35
         Text            =   "0"
         Top             =   2160
         Width           =   735
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturar agrupando por CLIENTE - ALBARAN - MUESTRA. DTO % "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   27
         Top             =   2205
         Width           =   6090
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturar agrupando por CLIENTE - ""CA"" - FAMILIA. TCT-SKY"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   25
         Top             =   1890
         Width           =   10140
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturar agrupando por Cliente. Genera una sola factura para todos los albaranes marcados. 1 Línea por Cliente."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   450
         Visible         =   0   'False
         Width           =   10140
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturar agrupando por Albaran. Genera una sola factura para todos los albaranes marcados.1 Línea por Albaran."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   1575
         Width           =   10185
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturar los albaranes normalmente. Genera una factura por cada albaran marcado."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   1260
         Value           =   -1  'True
         Width           =   8790
      End
      Begin VB.CommandButton cmdCancelarFactura 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cancelar"
         Height          =   735
         Left            =   8280
         Picture         =   "frmListadoAlbaranes.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2790
         Width           =   1860
      End
      Begin VB.CommandButton cmdGenerarFactura 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar Factura"
         Height          =   735
         Left            =   6345
         Picture         =   "frmListadoAlbaranes.frx":6B5C
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2790
         Width           =   1860
      End
      Begin pryCombo.miCombo cmbClienteFactura 
         Height          =   345
         Left            =   810
         TabIndex        =   17
         Top             =   765
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   609
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H000080FF&
         Caption         =   "INDIQUE LOS DATOS PARA GENERAR LA FACTURA"
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
         Index           =   0
         Left            =   360
         TabIndex        =   21
         Top             =   90
         Width           =   9420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   810
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7515
      Width           =   1050
   End
   Begin VB.Frame frmOpciones 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Opciones"
      Height          =   1035
      Left            =   45
      TabIndex        =   4
      Top             =   7470
      Width           =   12405
      Begin VB.CommandButton cmdEditar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Editar"
         Height          =   870
         Left            =   2880
         Picture         =   "frmListadoAlbaranes.frx":D3AE
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   45
         Width           =   1050
      End
      Begin VB.CommandButton cmdMarcarFacturada 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marcar como Facturada"
         Height          =   870
         Left            =   10485
         Picture         =   "frmListadoAlbaranes.frx":DC78
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   45
         Width           =   1860
      End
      Begin VB.CommandButton cmdFacturar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Facturar"
         Height          =   870
         Left            =   8595
         Picture         =   "frmListadoAlbaranes.frx":E542
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   45
         Width           =   1860
      End
      Begin VB.CommandButton cmdDesmarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desmarcar Todas"
         Height          =   870
         Left            =   1440
         Picture         =   "frmListadoAlbaranes.frx":EE0C
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   45
         Width           =   1410
      End
      Begin VB.CommandButton cmdMarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marcar Todas"
         Height          =   870
         Left            =   90
         Picture         =   "frmListadoAlbaranes.frx":1565E
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   45
         Width           =   1320
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Imprimir"
         Height          =   870
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   45
         Width           =   1050
      End
      Begin VB.Shape Shape1 
         Height          =   465
         Left            =   5535
         Top             =   225
         Width           =   2400
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   6120
         TabIndex        =   34
         Top             =   315
         Width           =   1755
      End
      Begin VB.Label lblrestan 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Total"
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
         Height          =   285
         Index           =   1
         Left            =   5580
         TabIndex        =   33
         Top             =   315
         Width           =   510
      End
   End
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
      Height          =   1050
      Left            =   45
      TabIndex        =   1
      Top             =   315
      Width           =   13545
      Begin VB.CheckBox chkVencimiento 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   135
         TabIndex        =   28
         Top             =   675
         Width           =   285
      End
      Begin VB.CheckBox chkIberia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Iberia"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10980
         TabIndex        =   26
         Top             =   720
         Width           =   780
      End
      Begin VB.CheckBox chkAirbus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Airbus"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10980
         TabIndex        =   13
         Top             =   495
         Width           =   780
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   10980
         TabIndex        =   6
         Top             =   270
         Width           =   780
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   780
         Left            =   12285
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   1095
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   900
         TabIndex        =   14
         Top             =   270
         Width           =   9780
         _ExtentX        =   17251
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesdev 
         Height          =   330
         Left            =   1575
         TabIndex        =   29
         Top             =   630
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   52166657
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhastav 
         Height          =   330
         Left            =   3645
         TabIndex        =   30
         Top             =   630
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   52166657
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha desde "
         Height          =   195
         Index           =   4
         Left            =   465
         TabIndex        =   32
         Top             =   705
         Width           =   975
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   5
         Left            =   3060
         TabIndex        =   31
         Top             =   675
         Width           =   405
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clientes"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Width           =   585
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6015
      Left            =   45
      TabIndex        =   10
      Top             =   1395
      Width           =   13515
      _ExtentX        =   23839
      _ExtentY        =   10610
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
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Albaranes pendientes de facturar"
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
      Width           =   13560
   End
End
Attribute VB_Name = "frmListadoAlbaranes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTodos_Click()
    If chkTodos.Value = Checked Then
        cmbclientes.limpiar
        cmbclientes.desactivar
        chkAirbus.Enabled = True
        chkIberia.Enabled = True
    Else
        cmbclientes.activar
        chkAirbus.Enabled = False
        chkIberia.Enabled = False
    End If
End Sub

Private Sub chkVencimiento_Click()
    If chkVencimiento.Value = Checked Then
        fdesdev.Enabled = True
        fhastav.Enabled = True
    Else
        fdesdev.Enabled = False
        fhastav.Enabled = False
    End If
End Sub

Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim IMPORTE As Currency
    Dim BASE As Currency
    Dim IVA As Currency
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As ADODB.Recordset
    Dim oDoc As New clsDocs_pago
    Me.MousePointer = 11
    Dim cliente As Long
    cliente = 0
    If chkTodos.Value = Unchecked Then
        If cmbclientes.getTEXTO = "" Then
            MsgBox "Seleccione un cliente.", vbInformation, App.Title
            Me.MousePointer = 0
            Exit Sub
        Else
            cliente = cmbclientes.getPK_SALIDA
        End If
    End If
    Dim fdesde As String
    Dim fhasta As String
    fdesde = ""
    fhasta = ""
    If chkVencimiento.Value = Checked Then
        fdesde = fdesdev.Value
        fhasta = fhastav.Value
    End If
    Set rs = oDoc.Albaranes(cliente, chkAirbus.Value, chkIberia.Value, fdesde, fhasta)
    Label1(4) = "Listado de Albaranes pendientes de facturar. Total registros : " & rs.RecordCount
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs.Fields(1), "0000"))
                .SubItems(1) = rs.Fields(2)
                .SubItems(2) = rs.Fields(3)
                .SubItems(9) = rs.Fields(0) ' ID
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
                .SubItems(10) = rs(9) ' TIPO
                .SubItems(11) = rs(11) 'PEDIDO
                .SubItems(12) = rs(10) 'ID_PEDIDO
                .SubItems(13) = rs(12) 'ID_CLIENTE
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    calcularTotal
    Me.MousePointer = 0
    Set oDoc = Nothing
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al buscar los albaranes.", vbCritical, Err.Description
End Sub
Private Sub calcularTotal()
    Dim i As Integer
    Dim T_TOTAL As Currency
    T_TOTAL = 0
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            T_TOTAL = T_TOTAL + CCur(lista.ListItems(i).SubItems(8))
        End If
    Next
    lbltotal = moneda(CStr(T_TOTAL))
End Sub
Private Sub cmdCancelarFactura_Click()
    frmGenera.visible = False
    frmBusqueda.Enabled = True
    frmOpciones.Enabled = True
    lista.Enabled = True
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
calcularTotal
End Sub

Private Sub cmdEditar_Click()
    If lista.ListItems.Count > 0 Then
     If CInt(lista.ListItems(lista.selectedItem.Index).SubItems(10)) <> 1 Then
        frmDocumento_Edicion.PK_DOCUMENTO = lista.ListItems(lista.selectedItem.Index).SubItems(9)
        frmDocumento_Edicion.Show 1
'        actualizar_lista CLng(lista.ListItems(lista.selectedItem.Index).SubItems(9)), lista.selectedItem.Index
     Else
        gdoc = lista.ListItems(lista.selectedItem.Index).SubItems(9)
        frmFacturaConceptos.Show 1
     End If
    End If
End Sub

Private Sub facturaAgrupada()
    ' cIVA
'    Dim oParametros As New clsParametros
'    Dim IVA As Integer
   On Error GoTo facturaAgrupada_Error

'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
    ' Creamos el documento de pago
    Dim oCliente As New clsCliente
    If oCliente.CargaCliente(cmbClienteFactura.getPK_SALIDA) = False Then
        MsgBox "Error al cargar los datos del cliente.", vbCritical, App.Title
        Exit Sub
    End If
    Dim i As Integer
    Dim pedido As Integer
    pedido = 0
'   1. VERIFICAR QUE LOS ALBARANES MARCADOS SON DEL MISMO PEDIDO
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If pedido <> lista.ListItems(i).SubItems(12) And pedido <> 0 Then
                MsgBox "Sólo se pueden facturar albaranes que son del mismo PEDIDO.", vbCritical, App.Title
                Exit Sub
            End If
            pedido = lista.ListItems(i).SubItems(12)
        End If
    Next
'   2. INSERTAR EL DOCUMENTO DE PAGO
    Dim num_doc As Long
    Dim oDocPago As New clsDocs_pago
    With oDocPago
        .setTIPO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA
        .setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
        .setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        
        .setCLIENTE_ID = cmbClienteFactura.getPK_SALIDA
        .setCLIENTE_ID_FACTURA = cmbClienteFactura.getPK_SALIDA
        
        .setTOTAL = moneda_bd("0")
        .setDESCUENTO = moneda_bd("0")
'        .setIVA = IVA
        .setPAGADO = 0
        .setANULADO = 0
        .setFP_ID = oCliente.getFP_ID
        .setPEDIDO_ID = pedido
        .setFACTURA_CONCEPTOS = 1 ' FACTURA SOLO DE CONCEPTOS
        ' Insertamos el documento de pago
        num_doc = oDocPago.InsertarDocPago
        If num_doc = 0 Then
            Exit Sub
        End If
    End With
'   3. PARA CADA ALBARAN, INSERTAMOS UN CONCEPTO CON LOS DATOS DEL ALBARAN
    Dim oConceptos As New clsDocs_pago_conceptos
    Dim concepto As String
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            oDocPago.CargarDocumento (lista.ListItems(i).SubItems(9))
            'M1537-I
            oCliente.CargaCliente oDocPago.getCLIENTE_ID
            'M1537-F
            With oConceptos
                .setDOC_ID = num_doc
                .setALBARAN_ID = oDocPago.getID_DOC
                'M1537-I
                '.setDESCRIPCION = "ALBARAN Nº" & oDocPago.getNUMERO & "/" & Year(oDocPago.getFECHA_FACTURA) & " FECHA : " & Format(oDocPago.getFECHA_FACTURA, "dd-mm-yyyy")
                concepto = "ALBARAN Nº" & oDocPago.getNUMERO & "/" & Year(oDocPago.getFECHA_FACTURA) & " FECHA : " & Format(oDocPago.getFECHA_FACTURA, "dd-mm-yyyy") & " " & _
                           oCliente.getNOMBRE
                .setDESCRIPCION = concepto
                'M1537-F
                .setFECHA = Format(oDocPago.getFECHA_FACTURA, "yyyy-mm-dd")
                .setPRECIO = moneda_bd(oDocPago.getTOTAL - ((oDocPago.getTOTAL * oDocPago.getDESCUENTO) / 100))
                .setFAMILIA_ID = 0
                        .setCANTIDAD = 1
                        .setAPARTADO = 0
                        .setDTO = 0
                        .setSUBTOTAL = moneda_bd(oDocPago.getTOTAL - ((oDocPago.getTOTAL * oDocPago.getDESCUENTO) / 100))
                        .setTOTAL = moneda_bd(oDocPago.getTOTAL - ((oDocPago.getTOTAL * oDocPago.getDESCUENTO) / 100))
                
                If .Insertar = False Then
                    Exit Sub
                End If
            End With
            ' Marcamos el albaran como facturado
            oDocPago.Facturar_Albaran lista.ListItems(i).SubItems(9), num_doc
        End If
    Next
    ' Informamos el total de la factura
    oDocPago.Informar_total_factura (num_doc)
    Me.MousePointer = 0
    MsgBox "Factura registrada correctamente.", vbInformation, App.Title
    cmdBuscar_Click

   On Error GoTo 0
   Exit Sub

facturaAgrupada_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure facturaAgrupada of Formulario frmListadoAlbaranes"
End Sub
Private Sub facturaTCT()
    ' cIVA
'    Dim oParametros As New clsParametros
'    Dim IVA As Integer
   On Error GoTo facturaAgrupada_Error

'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
    ' Creamos el documento de pago
    Dim oCliente As New clsCliente
    If oCliente.CargaCliente(cmbClienteFactura.getPK_SALIDA) = False Then
        MsgBox "Error al cargar los datos del cliente.", vbCritical, App.Title
        Exit Sub
    End If
    Dim i As Integer
    Dim pedido As Integer
    pedido = 0
'   1. VERIFICAR QUE LOS ALBARANES MARCADOS SON DEL MISMO PEDIDO
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If pedido <> lista.ListItems(i).SubItems(12) And pedido <> 0 Then
                MsgBox "Sólo se pueden facturar albaranes que son del mismo PEDIDO.", vbCritical, App.Title
                Exit Sub
            End If
            pedido = lista.ListItems(i).SubItems(12)
        End If
    Next
'   2. INSERTAR EL DOCUMENTO DE PAGO
    Dim num_doc As Long
    Dim oDocPago As New clsDocs_pago
    With oDocPago
        .setTIPO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA
        .setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
        .setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        
        .setCLIENTE_ID = cmbClienteFactura.getPK_SALIDA
        .setCLIENTE_ID_FACTURA = cmbClienteFactura.getPK_SALIDA
        
        .setTOTAL = moneda_bd("0")
        .setDESCUENTO = moneda_bd("0")
'        .setIVA = IVA
        .setPAGADO = 0
        .setANULADO = 0
        .setFP_ID = oCliente.getFP_ID
        .setPEDIDO_ID = pedido
        .setFACTURA_CONCEPTOS = 1 ' FACTURA SOLO DE CONCEPTOS
        ' Insertamos el documento de pago
        num_doc = oDocPago.InsertarDocPago
        If num_doc = 0 Then
            Exit Sub
        End If
    End With
'   3. PARA CADA ALBARAN, INSERTAMOS UN CONCEPTO CON LOS DATOS DEL ALBARAN
    Dim oConceptos As New clsDocs_pago_conceptos
    Dim rsConceptos As ADODB.Recordset
    Dim concepto As String
    Dim consulta As String
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            oDocPago.CargarDocumento (lista.ListItems(i).SubItems(9))
            oCliente.CargaCliente oDocPago.getCLIENTE_ID
            'MDET
            consulta = " SELECT AA.C2 AS FAMILIA, SUM(AA.C3) FROM ( " & _
                       "   SELECT DP.NUMERO AS C4, F.NOMBRE AS C2, SUM(DPM.PRECIO) - ((sum(DPM.PRECIO) * DP.descuento) / 100) AS C3 " & _
                       "     FROM MUESTRAS M, TIPOS_MUESTRA TM, DOCS_PAGO_MUESTRAS DPM, DOCS_PAGO DP, FAMILIAS F " & _
                       "    WHERE m.TIPO_MUESTRA_ID = tm.id_tipo_muestra " & _
                       "      AND TM.FAMILIA_ID = F.ID_FAMILIA " & _
                       "      AND DPM.MUESTRA_ID = M.ID_MUESTRA " & _
                       "      AND DP.ID_DOC = DPM.DOC_ID " & _
                       "      AND DP.ID_DOC = " & lista.ListItems(i).SubItems(9) & _
                       "      AND DPM.MUESTRA_ID <> 0 AND DPM.DETERMINACION_ID = 0 " & _
                       "    GROUP BY  DP.NUMERO,F.NOMBRE "
            consulta = consulta & " UNION " & _
                       "   select a.NUMERO AS C4, c.NOMBRE AS C2, sum(b.total) AS C3 " & _
                       "     from docs_pago a, docs_pago_conceptos b, familias c " & _
                       "    where a.ID_DOC = b.DOC_ID " & _
                       "      and b.FAMILIA_ID  = c.ID_FAMILIA " & _
                       "      and a.ID_DOC = " & lista.ListItems(i).SubItems(9) & _
                       "    group by a.NUMERO, c.NOMBRE " & _
                       " ) AS AA " & _
                       " GROUP BY AA.C2"
'            "   select a.NUMERO AS C4, c.NOMBRE AS C2, sum(b.PRECIO) - ((sum(b.PRECIO) * a.descuento) / 100) AS C3 "

            Set rsConceptos = datos_bd(consulta)
            If rsConceptos.RecordCount > 0 Then
                Do
                    With oConceptos
                        .setDOC_ID = num_doc
                        .setALBARAN_ID = oDocPago.getID_DOC
                        concepto = oCliente.getNOMBRE & "-" & "CA" & "-" & rsConceptos(0)
                        .setDESCRIPCION = concepto
                        .setFECHA = Format(oDocPago.getFECHA_FACTURA, "yyyy-mm-dd")
                        .setPRECIO = moneda_bd(rsConceptos(1))
                        .setFAMILIA_ID = 0
                        
                        .setCANTIDAD = 1
                        .setAPARTADO = 0
                        .setDTO = 0
                        .setSUBTOTAL = moneda_bd(rsConceptos(1))
                        .setTOTAL = moneda_bd(rsConceptos(1))
                        
                        If .Insertar = False Then
                            Exit Sub
                        End If
                    End With
                    rsConceptos.MoveNext
                Loop Until rsConceptos.EOF
            End If
            ' Marcamos el albaran como facturado
            oDocPago.Facturar_Albaran lista.ListItems(i).SubItems(9), num_doc
        End If
    Next
    ' Informamos el total de la factura
    oDocPago.Informar_total_factura (num_doc)
    Me.MousePointer = 0
    MsgBox "Factura registrada correctamente.", vbInformation, App.Title
    cmdBuscar_Click

   On Error GoTo 0
   Exit Sub

facturaAgrupada_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure facturaAgrupada of Formulario frmListadoAlbaranes"
End Sub

Private Sub facturaCliente()
    ' cIVA
'    Dim oParametros As New clsParametros
'    Dim IVA As Integer
'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
    ' Creamos el documento de pago
    Dim oCliente As New clsCliente
    If oCliente.CargaCliente(cmbClienteFactura.getPK_SALIDA) = False Then
        MsgBox "Error al cargar los datos del cliente.", vbCritical, App.Title
        Exit Sub
    End If
    Dim i As Integer
    Dim pedido As Integer
    pedido = 0
'   1. VERIFICAR QUE LOS ALBARANES MARCADOS SON DEL MISMO PEDIDO
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If pedido <> lista.ListItems(i).SubItems(12) And pedido <> 0 Then
                MsgBox "Sólo se pueden facturar albaranes que son del mismo PEDIDO.", vbCritical, App.Title
                Exit Sub
            End If
            pedido = lista.ListItems(i).SubItems(12)
        End If
    Next
'   2. INSERTAR EL DOCUMENTO DE PAGO
    Dim num_doc As Long
    Dim oDocPago As New clsDocs_pago
    With oDocPago
        .setTIPO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA
        .setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
        .setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        
        .setCLIENTE_ID = cmbClienteFactura.getPK_SALIDA
        .setCLIENTE_ID_FACTURA = cmbClienteFactura.getPK_SALIDA
        
        .setTOTAL = moneda_bd("0")
        .setDESCUENTO = moneda_bd("0")
'        .setIVA = IVA
        .setPAGADO = 0
        .setANULADO = 0
        .setFP_ID = oCliente.getFP_ID
        .setPEDIDO_ID = pedido
        .setFACTURA_CONCEPTOS = 1 ' FACTURA SOLO DE CONCEPTOS
        ' Insertamos el documento de pago
        num_doc = oDocPago.InsertarDocPago
        If num_doc = 0 Then
            Exit Sub
        End If
    End With
'   3. PARA CADA ALBARAN, INSERTAMOS UN CONCEPTO CON LOS DATOS DEL ALBARAN
    Dim oConceptos As New clsDocs_pago_conceptos
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            oDocPago.CargarDocumento (lista.ListItems(i).SubItems(9))
            With oConceptos
                .setDOC_ID = num_doc
                .setALBARAN_ID = oDocPago.getID_DOC
                .setDESCRIPCION = "ALBARAN Nº" & oDocPago.getNUMERO & "/" & Year(oDocPago.getFECHA_FACTURA) & " FECHA : " & Format(oDocPago.getFECHA_FACTURA, "dd-mm-yyyy")
                .setFECHA = Format(oDocPago.getFECHA_FACTURA, "yyyy-mm-dd")
                .setPRECIO = moneda_bd(oDocPago.getTOTAL - ((oDocPago.getTOTAL * oDocPago.getDESCUENTO) / 100))
                .setFAMILIA_ID = 0
                        
                .setCANTIDAD = 1
                .setAPARTADO = 0
                .setDTO = 0
                .setSUBTOTAL = moneda_bd(oDocPago.getTOTAL - ((oDocPago.getTOTAL * oDocPago.getDESCUENTO) / 100))
                .setTOTAL = moneda_bd(oDocPago.getTOTAL - ((oDocPago.getTOTAL * oDocPago.getDESCUENTO) / 100))
                
                If .Insertar = False Then
                    Exit Sub
                End If
            End With
            ' Marcamos el albaran como facturado
            oDocPago.Facturar_Albaran lista.ListItems(i).SubItems(9), num_doc
        End If
    Next
    ' Informamos el total de la factura
    oDocPago.Informar_total_factura (num_doc)
    Me.MousePointer = 0
    MsgBox "Factura registrada correctamente.", vbInformation, App.Title
    cmdBuscar_Click
End Sub

Private Sub facturaNormal()
    ' cIVA
'    Dim oParametros As New clsParametros
'    Dim IVA As Integer
   On Error GoTo facturaNormal_Error

'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
    ' Creamos el documento de pago
    Dim oCliente As New clsCliente
    Dim oDocPago As New clsDocs_pago
    If oCliente.CargaCliente(cmbClienteFactura.getPK_SALIDA) = False Then
        MsgBox "Error al cargar los datos del cliente.", vbCritical, App.Title
        Exit Sub
    End If
    Dim i As Integer
    Dim num_doc As Long
    Dim pedido As Integer
    Dim oConceptos As New clsDocs_pago_conceptos
    Dim omuestras As New clsDocs_pago_muestras
    Dim oMuestra As New clsMuestra
    Dim rs As ADODB.Recordset
    pedido = -1
    ' Verificamos de que tipo vamos a generar la factura
    Dim FMUESTRAS As Boolean
    Dim FCONCEPTOS As Boolean
    Dim FMIXTA As Boolean
    FMUESTRAS = False
    FCONCEPTOS = False
    FMIXTA = False
    Dim TIPO As Integer
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
        TIPO = 2
    Else
        If FCONCEPTOS Then
            TIPO = 1
        Else
            TIPO = 0
        End If
    End If
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            oDocPago.CargarDocumento (lista.ListItems(i).SubItems(9))
            If pedido <> oDocPago.getPEDIDO_ID Then
                pedido = oDocPago.getPEDIDO_ID
                ' Insertamos el detalle de factura por pedido
                With oDocPago
                    .setTIPO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA
                    .setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
                    .setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
                    .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                    .setCLIENTE_ID = cmbClienteFactura.getPK_SALIDA
                    .setCLIENTE_ID_FACTURA = cmbClienteFactura.getPK_SALIDA
                    .setTOTAL = moneda_bd("0")
                    .setDESCUENTO = moneda_bd("0")
'                    .setIVA = IVA
                    .setPAGADO = 0
                    .setANULADO = 0
                    .setFP_ID = oCliente.getFP_ID
                    .setPEDIDO_ID = pedido
                    .setFACTURA_CONCEPTOS = TIPO
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
                        .setPRECIO = moneda_bd(rs("precio"))
                        'M0524-I
                        .setFAMILIA_ID = rs("FAMILIA_ID")
                        'M0524-F
                        .setCANTIDAD = rs("cantidad")
                        .setAPARTADO = rs("apartado")
                        .setDTO = rs("dto")
                        .setSUBTOTAL = moneda_bd(rs("subtotal"))
                        .setTOTAL = moneda_bd(rs("total"))
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
                        .setPRECIO = moneda_bd(rs(5))
                        If .Insertar_doc_pago_muestra(0) = -1 Then
                            MsgBox "Error al insertar en doc_pago_muestra", vbCritical, App.Title
                            Exit Sub
                        End If
                        If oMuestra.Informar_Documento_Pago(rs(6), 2) = False Then
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
    Me.MousePointer = 0
    MsgBox "Factura registrada correctamente.", vbInformation, App.Title
    cmdBuscar_Click

   On Error GoTo 0
   Exit Sub

facturaNormal_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure facturaNormal of Formulario frmListadoAlbaranes"
End Sub

Private Sub cmdFacturar_Click()
   On Error GoTo cmdFacturar_Click_Error

    If contar_marcados = 0 Then
        MsgBox "Seleccione algún documento para facturar.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim i As Integer
    Dim cliente As Long
    Dim primerCliente As Boolean
    Dim masDeUnCliente As Boolean
    primerCliente = True
    masDeUnCliente = False
    cliente = 0
    cmbClienteFactura.limpiar
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If primerCliente Then
                cliente = lista.ListItems(i).SubItems(13)
                primerCliente = False
            End If
            If cliente <> lista.ListItems(i).SubItems(13) Then
                masDeUnCliente = True
            End If
            
        End If
    Next
    If Not masDeUnCliente And cliente <> 0 Then
        cmbClienteFactura.MostrarElemento cliente
    End If
    frmGenera.visible = True
    frmBusqueda.Enabled = False
    frmOpciones.Enabled = False
    lista.Enabled = False

   On Error GoTo 0
   Exit Sub

cmdFacturar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdFacturar_Click of Formulario frmListadoAlbaranes"
End Sub

Private Sub cmdGenerarFactura_Click()
    If cmbClienteFactura.getTEXTO = "" Then
        MsgBox "Debe indicar el cliente al que generar la factura.", vbCritical, App.Title
        Exit Sub
    End If
    Me.MousePointer = 11
    If opTipo(0).Value = True Then
        facturaNormal
    ElseIf opTipo(1).Value = True Then
        facturaAgrupada
    ElseIf opTipo(2).Value = True Then
        facturaCliente
    ElseIf opTipo(3).Value = True Then
        facturaTCT
    ElseIf opTipo(4).Value = True Then
        facturaIBERIA
    End If
    Me.MousePointer = 0
    frmGenera.visible = False
    frmBusqueda.Enabled = True
    frmOpciones.Enabled = True
    lista.Enabled = True
End Sub

Private Sub cmdImprimir_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oDoc_pago As New clsDocs_pago
    oDoc_pago.generar_factura lista.ListItems(lista.selectedItem.Index).SubItems(9), False, "", "rptFactura"
    Set oDoc_pago = Nothing
End Sub
Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
calcularTotal
End Sub

Private Sub cmdMarcarFacturada_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If contar_marcados = 0 Then
        MsgBox "Marque las facturas que desea marcar como facturadas.", vbExclamation, App.Title
    Else
        Dim i As Integer
        Dim oDoc_pago As New clsDocs_pago
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                oDoc_pago.MARCAR_FACTURADA lista.ListItems(i).SubItems(9)
            End If
        Next
        Set oDoc_pago = Nothing
        buscar
    End If
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
    Me.top = 50
    cabecera
    rellenar_clientes
    fdesdev = Date - 31
    fhastav = Date
    permisos
End Sub
Private Sub rellenar_clientes()
    Dim consulta As String
    consulta = "SELECT DISTINCT A.ID_CLIENTE,A.NOMBRE " & _
               "  FROM CLIENTES A,DOCS_PAGO B" & _
               " WHERE B.TIPO=1 AND A.ANULADO=0 AND B.PAGADO=0 AND B.ANULADO = 0 " & _
               "   AND A.ID_CLIENTE = B.CLIENTE_ID "
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbclientes
            .setCONN = conn
            .setQUERY = consulta
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "CLIENTES"
            .setDESCRIPCION = "Clientes"
            .setPK = "A.ID_CLIENTE"
            .setFILTRO = ""
            .setCAMPO = "A.NOMBRE"
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmClientes
        End With
    End If
    llenar_combo cmbClienteFactura, New clsCliente, 0, frmClientes, ""
    
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "NºDoc", 800, lvwColumnLeft
        .Add , , "Cliente", 3200, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Importe", 1100, lvwColumnRight
        .Add , , "Dto. %", 800, lvwColumnCenter
        .Add , , "Base", 1100, lvwColumnRight
        .Add , , "I.V.A.%", 800, lvwColumnRight
        .Add , , "Cuota I.V.A.", 1200, lvwColumnRight
        .Add , , "Total", 1100, lvwColumnRight
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "TIPO", 1, lvwColumnCenter
        .Add , , "Pedido", 2000, lvwColumnLeft
        .Add , , "ID_PEDIDO", 1, lvwColumnLeft
        .Add , , "ID_CLIENTE", 1, lvwColumnLeft
    End With
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
    cmdImprimir_Click
End Sub
Private Function contar_marcados() As Integer
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
Private Sub permisos()
End Sub

Private Sub facturaIBERIA()
    ' cIVA
'    Dim oParametros As New clsParametros
'    Dim IVA As Integer
   On Error GoTo facturaIBERIA_Error

'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
    ' Creamos el documento de pago
    Dim oCliente As New clsCliente
    If oCliente.CargaCliente(cmbClienteFactura.getPK_SALIDA) = False Then
        MsgBox "Error al cargar los datos del cliente.", vbCritical, App.Title
        Exit Sub
    End If
    Dim i As Integer
    Dim pedido As Integer
    pedido = 0
'   1. VERIFICAR QUE LOS ALBARANES MARCADOS SON DEL MISMO PEDIDO
'    For i = 1 To lista.ListItems.Count
'        If lista.ListItems(i).Checked = True Then
'            If PEDIDO <> lista.ListItems(i).SubItems(12) And PEDIDO <> 0 Then
'                MsgBox "Sólo se pueden facturar albaranes que son del mismo PEDIDO.", vbCritical, App.Title
'                Exit Sub
'            End If
'            PEDIDO = lista.ListItems(i).SubItems(12)
'        End If
'    Next
'   2. INSERTAR EL DOCUMENTO DE PAGO
    Dim num_doc As Long
    Dim total As Single
    Dim oDocPago As New clsDocs_pago
    With oDocPago
        .setTIPO = C_TIPOS_DOCS_PAGO.C_TIPOS_DOCS_PAGO_FACTURA
        .setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
        .setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        
        .setCLIENTE_ID = cmbClienteFactura.getPK_SALIDA
        .setCLIENTE_ID_FACTURA = cmbClienteFactura.getPK_SALIDA
        
        .setTOTAL = moneda_bd("0")
        .setDESCUENTO = moneda_bd("0")
'        .setIVA = IVA
        .setPAGADO = 0
        .setANULADO = 0
        .setFP_ID = oCliente.getFP_ID
        .setPEDIDO_ID = pedido
        .setFACTURA_CONCEPTOS = 1 ' FACTURA SOLO DE CONCEPTOS
        ' Insertamos el documento de pago
        num_doc = oDocPago.InsertarDocPago
        If num_doc = 0 Then
            Exit Sub
        End If
    End With
'   3. PARA CADA ALBARAN, INSERTAMOS UN CONCEPTO CON LOS DATOS DEL ALBARAN
    Dim oConceptos As New clsDocs_pago_conceptos
    Dim rsConceptos As ADODB.Recordset
    Dim concepto As String
    Dim consulta As String
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            oDocPago.CargarDocumento (lista.ListItems(i).SubItems(9))
            oCliente.CargaCliente oDocPago.getCLIENTE_ID
            ' Cabecera del concepto
            With oConceptos
                .setDOC_ID = num_doc
                .setALBARAN_ID = oDocPago.getID_DOC
                concepto = oCliente.getNOMBRE & "-" & "ALBARAN Nº" & oDocPago.getNUMERO & "/" & Year(oDocPago.getFECHA_FACTURA)
                .setDESCRIPCION = concepto
                .setFECHA = Format(oDocPago.getFECHA_FACTURA, "yyyy-mm-dd")
                .setPRECIO = moneda_bd(oDocPago.getTOTAL)
                .setFAMILIA_ID = 0
                .setCANTIDAD = 1
                .setAPARTADO = 0
                .setDTO = txtDto
                total = oDocPago.getTOTAL - ((oDocPago.getTOTAL * CSng(Replace(txtDto, ".", ","))) / 100)
                .setSUBTOTAL = moneda_bd(oDocPago.getTOTAL)
                .setTOTAL = moneda_bd(CStr(total))
                If .Insertar = False Then
                    Exit Sub
                End If
            End With
            Dim odpm As New clsDocs_pago_muestras
            Dim rs2 As ADODB.Recordset
            Set rs2 = odpm.lineas_factura_sin_determinaciones(oDocPago.getID_DOC)
            If rs2.RecordCount > 0 Then
                Do
                    With oConceptos
                        .setDOC_ID = num_doc
                        .setALBARAN_ID = oDocPago.getID_DOC
                        concepto = rs2(1) & " - " & rs2(4) & " - " & rs2(5)
                        .setDESCRIPCION = concepto
                        .setFECHA = Format(rs2(2), "yyyy-mm-dd")
                        .setPRECIO = moneda_bd(rs2(7))
                        .setFAMILIA_ID = 0

                        .setCANTIDAD = 1
                        .setAPARTADO = 1
                        .setDTO = txtDto
                        total = rs2(7) - ((rs2(7) * CSng(Replace(txtDto, ".", ","))) / 100)
                        .setSUBTOTAL = moneda_bd(rs2(7))
                        .setTOTAL = moneda_bd(CStr(total))

                        If .Insertar = False Then
                            Exit Sub
                        End If
                    End With
                    
                    rs2.MoveNext
                Loop Until rs2.EOF
            End If
            ' Insertamos el detalle de la factura de conceptos
            Set rs2 = oConceptos.ConceptosDocumento(lista.ListItems(i).SubItems(9))
            If rs2.RecordCount > 0 Then
                Do
                    With oConceptos
                        .setDOC_ID = num_doc
                        .setDESCRIPCION = rs2("DESCRIPCION")
                        .setFECHA = Format(rs2("FECHA"), "yyyy-mm-dd")
                        .setPRECIO = moneda_bd(rs2("precio"))
                        .setFAMILIA_ID = rs2("FAMILIA_ID")
                        .setCANTIDAD = Replace(rs2("cantidad"), ",", ".")
                        .setAPARTADO = 1
                        .setDTO = txtDto
                        total = rs2("precio") - ((rs2("precio") * CSng(Replace(txtDto, ".", ","))) / 100)
                        .setSUBTOTAL = moneda_bd(rs2("precio"))
                        .setTOTAL = moneda_bd(CStr(total))
                        If .Insertar = False Then
                            Exit Sub
                        End If
                    End With
                    rs2.MoveNext
                Loop Until rs2.EOF
            End If
            ' Marcamos el albaran como facturado
            oDocPago.Facturar_Albaran lista.ListItems(i).SubItems(9), num_doc
        End If
    Next
    ' Informamos el total de la factura
    oDocPago.Informar_total_factura (num_doc)
    Me.MousePointer = 0
    MsgBox "Factura registrada correctamente.", vbInformation, App.Title
    cmdBuscar_Click

   On Error GoTo 0
   Exit Sub

facturaIBERIA_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure facturaIBERIA of Formulario frmListadoAlbaranes"

End Sub


Private Sub lista_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    calcularTotal
End Sub

