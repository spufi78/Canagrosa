VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPresupuesto 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Presupuesto"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9360
   Icon            =   "frmPresupuesto.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   9360
   Begin VB.CommandButton cmdAnalisis 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir Análisis"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   60
      Picture         =   "frmPresupuesto.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8100
      Visible         =   0   'False
      Width           =   1545
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   7170
      Picture         =   "frmPresupuesto.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8100
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1560
      Left            =   60
      TabIndex        =   20
      Top             =   6480
      Width           =   9240
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar Concepto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   8010
         Picture         =   "frmPresupuesto.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdNuevo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3030
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1050
         Width           =   975
      End
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4110
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1050
         Width           =   975
      End
      Begin VB.TextBox txtprecio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1080
         TabIndex        =   7
         Top             =   1080
         Width           =   1635
      End
      Begin VB.TextBox txtdes 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   1080
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   180
         Width           =   6765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   5
         Left            =   90
         TabIndex        =   22
         Top             =   1140
         Width           =   630
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   4
         Left            =   90
         TabIndex        =   21
         Top             =   510
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   8280
      Picture         =   "frmPresupuesto.frx":11E8
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8100
      Width           =   1035
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del presupuesto"
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
      Height          =   1185
      Left            =   45
      TabIndex        =   14
      Top             =   360
      Width           =   9195
      Begin VB.CommandButton cmdMas 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8760
         TabIndex        =   25
         Top             =   300
         Width           =   255
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7650
         TabIndex        =   4
         Top             =   750
         Width           =   1005
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4680
         TabIndex        =   3
         Top             =   720
         Width           =   885
      End
      Begin VB.TextBox txtdescuento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3000
         TabIndex        =   2
         Top             =   720
         Width           =   975
      End
      Begin MSDataListLib.DataCombo cmbClientes 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   300
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
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
      Begin MSComCtl2.DTPicker ffactura 
         Height          =   330
         Left            =   900
         TabIndex        =   1
         Top             =   690
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   23592961
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recogida de Muestras"
         Height          =   195
         Index           =   8
         Left            =   5880
         TabIndex        =   24
         Top             =   780
         Width           =   1605
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Iva"
         Height          =   195
         Index           =   7
         Left            =   4320
         TabIndex        =   23
         Top             =   780
         Width           =   225
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dto (%)"
         Height          =   195
         Index           =   2
         Left            =   2370
         TabIndex        =   18
         Top             =   780
         Width           =   510
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   360
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha "
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   780
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4560
      Left            =   60
      TabIndex        =   5
      Top             =   1890
      Width           =   9240
      _ExtentX        =   16298
      _ExtentY        =   8043
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      Enabled         =   0   'False
      NumItems        =   0
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del presupuesto"
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
      Left            =   60
      TabIndex        =   19
      Top             =   1560
      Width           =   9210
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Presupuesto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   4
      Left            =   75
      TabIndex        =   17
      Top             =   30
      Width           =   9180
   End
End
Attribute VB_Name = "frmPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdaceptar_Click()
    If cmbClientes.Text = "" Then
        MsgBox "Seleccione algún cliente.", vbInformation, App.Title
        cmbClientes.SetFocus
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        Dim oDocPago As New clsDocs_pago
        oDocPago.setFECHA_FACTURA = Format(ffactura, "yyyy-mm-dd")
        oDocPago.setCLIENTE_ID = cmbClientes.BoundText
        oDocPago.setDESCUENTO = Replace(Format(txtdescuento, "0.00"), ",", ".")
        oDocPago.setFP_ID = cmbfp.BoundText
        If cmbPedido.BoundText = "" Then
            oDocPago.setPEDIDO_ID = 0
        Else
            oDocPago.setPEDIDO_ID = cmbPedido.BoundText
        End If
    
        ' Insertamos el documento de pago
        If oDocPago.Modificar(gdoc) = False Then
             Exit Sub
        End If
        ' Borramos los conceptos anteriores
        Dim oConcepto As New clsDocs_pago_conceptos
        oConcepto.EliminarConceptos (gdoc)
        ' Insertamos los conceptos
        For i = 1 To lista.ListItems.Count
            oConcepto.setDOC_ID = gdoc
            oConcepto.setDESCRIPCION = lista.ListItems(i).SubItems(1)
            oConcepto.setFECHA = Format(lista.ListItems(i), "yyyy-mm-dd")
            oConcepto.setPRECIO = lista.ListItems(i).SubItems(2)
            If oConcepto.Insertar = False Then
                Exit Sub
            End If
        Next
        MsgBox "Documento modificado correctamente.", vbOKOnly + vbInformation, App.Title
        Unload Me
    Else
        MsgBox "Necesita algún concepto para la factura.", vbInformation, App.Title
    End If
End Sub
Private Sub cmdBuscar_Click()
    frmDocs_Pago_conceptos_guardados.Show 1
    If gid_concepto > 0 Then
        Dim odpc As New clsDocs_pago_conceptos_guardados
        With odpc
            .Carga (gid_concepto)
            txtdes = .getTEXTO
            txtprecio = Replace(.getPRECIO, ".", ",")
        End With
    End If
End Sub
Private Sub cmdEliminar_Click()
    If lista.SelectedItem.Index > 0 Then
     lista.ListItems.Remove (lista.SelectedItem.Index)
     cmdEliminar.Enabled = False
    End If
End Sub
Private Sub cmdMas_Click()
    frmClientes.Show 1
    cargar_clientes
End Sub
Private Sub cmdNuevo_Click()
    If valida_datos Then
        ' Añadimos el concepto
        With lista.ListItems.Add(, , lista.ListItems.Count + 1)
            .SubItems(1) = txtdes
            .SubItems(2) = Replace(Format(txtprecio, "0.00"), ",", ".")
        End With
        borrar_campos
    End If
End Sub

Private Sub cmdSalir_Click()
    log ("Cerrando modificación factura de conceptos")
    If lista.ListItems.Count > 0 Then
       If MsgBox("Existen conceptos. ¿Esta seguro de querer salir?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Unload Me
       End If
    Else
       Unload Me
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' Esc
            cmdSalir_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 100
    Me.Top = 100
    cargar_clientes
    cabecera_grid
    ffactura = Now
    txtfecha = Now
    If gdoc <> 0 Then
        Me.Left = 200
        Me.Top = 500
        cargar_documento
    End If
End Sub
Public Sub cabecera_grid()
    With lista.ColumnHeaders
        .Add , , "Orden", 1, lvwColumnLeft
        .Add , , "Descripción", 6700, lvwColumnLeft
        .Add , , "Precio", 1500, lvwColumnRight
    End With
End Sub
Public Sub cargar_clientes()
    Dim ocliente As New clsCliente
    Dim rsClientes As New ADODB.Recordset
    Set rsClientes = ocliente.Listado("", "", "")
    Set cmbClientes.RowSource = rsClientes
    cmbClientes.ListField = "nombre"
    cmbClientes.DataField = "id_cliente"
    cmbClientes.BoundColumn = "id_cliente"
    Set ocliente = Nothing
End Sub
Public Sub borrar_campos()
    txtdes = ""
    txtprecio = ""
    txtdes.SetFocus
End Sub

Public Function valida_datos() As Boolean
    valida_datos = True
    If txtdes = "" Then
        MsgBox "El concepto esta vacio.", vbInformation, App.Title
        txtdes.SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtprecio = "" Then
        MsgBox "El campo precio esta vacio.", vbInformation, App.Title
        txtprecio.SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        cmdEliminar.Enabled = True
    Else
        cmdEliminar.Enabled = False
    End If
End Sub
Private Sub txtdes_GotFocus()
    txtdes.BackColor = &H80C0FF
    txtdes.SelStart = 0
    txtdes.SelLength = Len(txtdes)
End Sub

Private Sub txtdes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdes_LostFocus()
    txtdes.BackColor = &HFFFFFF
End Sub

Private Sub txtprecio_LostFocus()
    txtprecio.BackColor = &HFFFFFF
    If txtprecio <> "" Then
        If Not IsNumeric(txtprecio) Then
            MsgBox "El precio debe ser numérico.", vbInformation, App.Title
            txtprecio = ""
            txtprecio.SetFocus
        End If
    End If
End Sub
Private Sub txtprecio_GotFocus()
    txtprecio.BackColor = &H80C0FF
    txtprecio.SelStart = 0
    txtprecio.SelLength = Len(txtprecio)
End Sub

Private Sub txtprecio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
    ' Escribir ',' al pulsar '.'
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub
Public Sub generar_documento(Tipo_documento As Integer)
    Dim i As Integer
    Dim num_doc As Integer
    Dim ocliente As New clsCliente
    Dim oDocPago As New clsDocs_pago
   On Error GoTo generar_documento_Error

    If ocliente.CargaCliente(cmbClientes.BoundText) = False Then
        Exit Sub
    End If
    oDocPago.setTIPO = Tipo_documento
    oDocPago.setFECHA_FACTURA = Format(ffactura, "yyyy-mm-dd")
    oDocPago.setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
    oDocPago.setEMPLEADO_ID = USUARIO.getID_EMPLEADO
    oDocPago.setCLIENTE_ID = cmbClientes.BoundText
    oDocPago.setTOTAL = "0.00"
    oDocPago.setDESCUENTO = Replace(Format(txtdescuento, "0.00"), ",", ".")
    If Tipo_documento = 2 Then
         oDocPago.setIVA = 16
    Else
         oDocPago.setIVA = 0
    End If
    oDocPago.setPAGADO = 0
    oDocPago.setANULADO = 0
    oDocPago.setFP_ID = cmbfp.BoundText
    If cmbPedido.BoundText = "" Then
        oDocPago.setPEDIDO_ID = 0
    Else
        oDocPago.setPEDIDO_ID = cmbPedido.BoundText
    End If
    oDocPago.setFACTURA_CONCEPTOS = 1
    ' Insertamos el documento de pago
    num_doc = oDocPago.InsertarDocPago
    If num_doc = 0 Then
         Exit Sub
    End If
    ' Insertamos los conceptos
    Dim oConcepto As New clsDocs_pago_conceptos
    Dim omuestra As New clsMuestra
    For i = 1 To lista.ListItems.Count
        oConcepto.setDOC_ID = num_doc
        oConcepto.setDESCRIPCION = lista.ListItems(i).SubItems(1)
        oConcepto.setFECHA = Format(lista.ListItems(i), "yyyy-mm-dd")
        oConcepto.setPRECIO = lista.ListItems(i).SubItems(2)
        If oConcepto.Insertar = False Then
            Exit Sub
        End If
        If EMPRESA.getID_EMPRESA = 3 Then
            If lista.ListItems(i).SubItems(3) <> "" Then
                omuestra.Informar_Documento_Pago_general lista.ListItems(i).SubItems(3), 5
            End If
            ' Insertamos el conceptos en los guardados
            Dim odpg As New clsDocs_pago_conceptos_guardados
            With odpg
                .setTEXTO = lista.ListItems(i).SubItems(1)
                .setPRECIO = lista.ListItems(i).SubItems(2)
                .Insertar
            End With
        End If
    Next
    Dim stipo As String
    If Tipo_documento = 1 Then
        MsgBox "Albaran registrado correctamente.", vbOKOnly + vbInformation, App.Title
    Else
        MsgBox "Factura registrada correctamente.", vbOKOnly + vbInformation, App.Title
    End If
    Unload Me
   On Error GoTo 0
   Exit Sub

generar_documento_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_documento of Formulario frmPresupuesto"
End Sub
Public Sub cargar_documento()
    Label1(4).BackColor = &H80C0FF
    Label1(4) = "Modificación de Factura de Conceptos"
    cmdAceptar.Visible = True
    cmdFactura.Enabled = False
    cmdAlbaran.Enabled = False
    lista.Enabled = True
    ' Documento
    Dim oDoc_pago As New clsDocs_pago
    oDoc_pago.CargarDocumento (gdoc)
    cmbClientes.BoundText = oDoc_pago.getCLIENTE_ID
    ffactura = oDoc_pago.getFECHA_FACTURA
    txtdescuento = oDoc_pago.getDESCUENTO
    cmbfp.BoundText = oDoc_pago.getFP_ID
    cmbPedido.BoundText = oDoc_pago.getPEDIDO_ID
    ' Conceptos
    Dim oDoc_pago_conceptos As New clsDocs_pago_conceptos
    Dim rs As ADODB.Recordset
    Set rs = oDoc_pago_conceptos.ConceptosDocumento(gdoc)
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs("fecha"), "dd/mm/yyyy"))
                .SubItems(1) = rs("descripcion")
                .SubItems(2) = Replace(Format(rs("precio"), "0.00"), ",", ".")
                If EMPRESA.getID_EMPRESA = 3 Then
                    .SubItems(3) = ""
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub
