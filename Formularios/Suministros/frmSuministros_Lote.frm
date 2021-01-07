VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmSuministros_Lote 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Gestión de nuevo Lote de Suministro"
   ClientHeight    =   11685
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10425
   Icon            =   "frmSuministros_Lote.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11685
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFactura 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Factura"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   27
      Top             =   10755
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1740
      Left            =   45
      TabIndex        =   7
      Top             =   8955
      Width           =   9480
      Begin VB.TextBox txtPrecio 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3240
         TabIndex        =   11
         Top             =   945
         Width           =   1680
      End
      Begin VB.TextBox txtbotes 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1170
         TabIndex        =   10
         Top             =   945
         Width           =   960
      End
      Begin MSDataListLib.DataCombo cmbCapacidad 
         Height          =   330
         Left            =   1170
         TabIndex        =   9
         Top             =   585
         Width           =   7025
         _ExtentX        =   12383
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
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1170
         TabIndex        =   8
         Top             =   225
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbPedidos 
         Height          =   330
         Left            =   1170
         TabIndex        =   12
         Top             =   1305
         Width           =   8070
         _ExtentX        =   14235
         _ExtentY        =   582
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   90
         TabIndex        =   26
         Top             =   1350
         Width           =   690
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio Bote"
         Height          =   240
         Index           =   4
         Left            =   2295
         TabIndex        =   25
         Top             =   990
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   19
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Capacidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Capacidad"
         Height          =   240
         Index           =   3
         Left            =   90
         TabIndex        =   18
         Top             =   675
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero Botes"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   17
         Top             =   990
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdInsertacliente 
      BackColor       =   &H00E0E0E0&
      Height          =   690
      Left            =   9585
      Picture         =   "frmSuministros_Lote.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   9180
      Width           =   735
   End
   Begin VB.CommandButton cmdEliminacliente 
      BackColor       =   &H00E0E0E0&
      Height          =   690
      Left            =   9585
      Picture         =   "frmSuministros_Lote.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7065
      Width           =   735
   End
   Begin VB.CommandButton cmdmodificarcliente 
      BackColor       =   &H00E0E0E0&
      Height          =   690
      Left            =   9585
      Picture         =   "frmSuministros_Lote.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   9855
      Width           =   735
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9330
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   10755
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   8235
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   10755
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   45
      TabIndex        =   1
      Top             =   630
      Width           =   10335
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   660
         Width           =   1545
      End
      Begin pryCombo.miCombo cmbProducto 
         Height          =   330
         Left            =   1080
         TabIndex        =   24
         Top             =   270
         Width           =   9105
         _ExtentX        =   16060
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Width           =   945
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Num.Lote"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   2
         Top             =   675
         Width           =   1005
      End
   End
   Begin MSComctlLib.ListView listaReactivos 
      Height          =   3630
      Left            =   45
      TabIndex        =   20
      Top             =   2025
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   6403
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
   Begin MSComctlLib.ListView clientes 
      Height          =   2985
      Left            =   45
      TabIndex        =   21
      Top             =   5940
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5265
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
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Marque el reactivo Interno a suministrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   45
      TabIndex        =   23
      Top             =   1755
      Width           =   10320
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Clientes y Capacidades"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   45
      TabIndex        =   22
      Top             =   5670
      Width           =   10320
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo Lote de Suministro"
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
      Left            =   90
      TabIndex        =   5
      Top             =   30
      Width           =   2670
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9765
      Picture         =   "frmSuministros_Lote.frx":2328
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del Tipo de Suministro"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   330
      Width           =   2100
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   10380
   End
End
Attribute VB_Name = "frmSuministros_Lote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub clientes_Click()
    cmdEliminacliente.Enabled = True
    cmdInsertacliente.Enabled = True
    cmdmodificarcliente.Enabled = True
    cmdFactura.visible = False
    If clientes.ListItems.Count > 0 Then
        cmbClientes.MostrarElemento clientes.ListItems(clientes.selectedItem.Index).SubItems(3)
        cmbCapacidad.BoundText = clientes.ListItems(clientes.selectedItem.Index).SubItems(4)
        txtbotes = clientes.ListItems(clientes.selectedItem.Index).SubItems(2)
        txtPrecio = clientes.ListItems(clientes.selectedItem.Index).SubItems(5)
        cmbPedidos.MostrarElemento clientes.ListItems(clientes.selectedItem.Index).SubItems(7)
        
        If CLng(clientes.ListItems(clientes.selectedItem.Index).SubItems(8)) <> 0 Then
            cmdEliminacliente.Enabled = False
            cmdInsertacliente.Enabled = False
            cmdmodificarcliente.Enabled = False
            cmdFactura.visible = True
        End If
        
    End If
End Sub

Private Sub cmbClientes_change()
    If cmbClientes.getTEXTO <> "" Then
        pedidos cmbClientes.getPK_SALIDA
    Else
        cmbClientes.limpiar
    End If
End Sub

Private Sub cmbproducto_Change()
    If PK = 0 Then
        cargar_producto
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Public Function validar_cliente() As Boolean
    validar_cliente = True
    If cmbClientes.getTEXTO = "" Then
        MsgBox "Seleccione un cliente.", vbInformation, App.Title
        cmbClientes.SetFocus
        validar_cliente = False
        Exit Function
    End If
    If cmbCapacidad.BoundText = "" Then
        MsgBox "Seleccione un tipo de capacidad.", vbInformation, App.Title
        cmbCapacidad.SetFocus
        validar_cliente = False
        Exit Function
    End If
    If Trim(txtbotes) = "" Then
        MsgBox "Debe introducir el numero de botes.", vbInformation, App.Title
        txtbotes.SetFocus
        validar_cliente = False
        Exit Function
    End If
    If Trim(txtPrecio) = "" Then
        MsgBox "Debe introducir el precio del Bote.", vbInformation, App.Title
        txtPrecio.SetFocus
        validar_cliente = False
        Exit Function
    End If
End Function

Private Sub cmdFactura_Click()
   On Error GoTo cmdfactura_Click_Error

    If clientes.ListItems.Count > 0 Then
        Dim oDP As New clsDocs_pago
        oDP.generar_factura clientes.ListItems(clientes.selectedItem.Index).SubItems(8), False, "", "rptFactura"
        Set oDP = Nothing
    End If

   On Error GoTo 0
   Exit Sub

cmdfactura_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdFactura_Click of Formulario frmSuministros_Lote"
End Sub

Private Sub cmdInsertacliente_Click()
    If validar_cliente Then
        With clientes.ListItems.Add(, , cmbClientes.getTEXTO)
            .SubItems(1) = cmbCapacidad.Text
            .SubItems(2) = txtbotes
            .SubItems(3) = cmbClientes.getPK_SALIDA
            .SubItems(4) = cmbCapacidad.BoundText
            .SubItems(5) = txtPrecio
            .SubItems(6) = cmbPedidos.getTEXTO
            .SubItems(7) = cmbPedidos.getPK_SALIDA
        End With
        clientes.ListItems(clientes.ListItems.Count).EnsureVisible
    End If
End Sub

Private Sub cmdmodificarcliente_Click()
    If clientes.ListItems.Count = 0 Then
        Exit Sub
    End If
    If validar_cliente Then
        clientes.ListItems(clientes.selectedItem.Index).Text = cmbClientes.getTEXTO
        clientes.ListItems(clientes.selectedItem.Index).SubItems(1) = cmbCapacidad.Text
        clientes.ListItems(clientes.selectedItem.Index).SubItems(2) = txtbotes
        clientes.ListItems(clientes.selectedItem.Index).SubItems(3) = cmbClientes.getPK_SALIDA
        clientes.ListItems(clientes.selectedItem.Index).SubItems(4) = cmbCapacidad.BoundText
        clientes.ListItems(clientes.selectedItem.Index).SubItems(5) = txtPrecio
        clientes.ListItems(clientes.selectedItem.Index).SubItems(6) = cmbPedidos.getTEXTO
        clientes.ListItems(clientes.selectedItem.Index).SubItems(7) = cmbPedidos.getPK_SALIDA
        clientes.ListItems(clientes.ListItems.Count).EnsureVisible
    End If

End Sub
Private Sub cmdEliminacliente_Click()
    If clientes.ListItems.Count > 0 Then
        clientes.ListItems.Remove clientes.selectedItem.Index
    End If
End Sub
Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = True Then
      Dim oSumLote As New clsSuministros_lotes
      Dim i As Integer
      Dim LOTE As Long
      With oSumLote
'        .setTIPO_SUMINISTRO_ID = cmbProducto.BoundText
        .setTIPO_SUMINISTRO_ID = cmbProducto.getPK_SALIDA
        .setNUMERO_LOTE = txtDatos(0)
        For i = 1 To listaReactivos.ListItems.Count
          If listaReactivos.ListItems(i).Selected = True Then
            .setBOTE_PR_ID = listaReactivos.ListItems(i).Text
            .setFECHA_FABRICACION = Format(listaReactivos.ListItems(i).SubItems(4), "yyyy-mm-dd")
            .setFECHA_CADUCIDAD = Format(listaReactivos.ListItems(i).SubItems(5), "yyyy-mm-dd")
          End If
        Next
        If PK = 0 Then
            LOTE = .Insertar
        Else
            .Modificar (PK)
            LOTE = PK
        End If
        If LOTE <> 0 Then
            'Clientes
            Dim oSumClientes As New clsSuministros_lotes_clientes
            oSumClientes.Eliminar LOTE
            For i = 1 To clientes.ListItems.Count
              With oSumClientes
                .setLOTE_ID = LOTE
                .setORDEN = i
                .setCLIENTE_ID = clientes.ListItems(i).SubItems(3)
                .setNUMERO_BOTES = clientes.ListItems(i).SubItems(2)
                .setCAPACIDAD_ID = clientes.ListItems(i).SubItems(4)
                .setPRECIO = moneda_bd(clientes.ListItems(i).SubItems(5))
                If Trim(clientes.ListItems(i).SubItems(7)) = "" Then
                    .setPEDIDO_ID = 0
                Else
                    .setPEDIDO_ID = clientes.ListItems(i).SubItems(7)
                End If
                .Insertar
              End With
            Next
           Else
           Exit Sub
        End If
      End With
      If PK = 0 Then
          MsgBox "El lote se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title '      Else
      Else
          MsgBox "El lote se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
    Exit Sub
fallo:
    error_grave (Err.Description)
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    Call cargar_combos
    If PK <> 0 Then
'        cmbProducto.Enabled = False
        cmbProducto.desactivar
        cargar_lote
    End If
End Sub

Private Sub cabecera()
    With listaReactivos.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Numero", 1000, lvwColumnLeft
        .Add , , "Codigo", 2100, lvwColumnCenter
        .Add , , "Reactivo", 3400, lvwColumnCenter
        .Add , , "F.Fabricacion", 1100, lvwColumnCenter
        .Add , , "F.Caducidad", 1100, lvwColumnCenter
        .Add , , "Volumen", 1100, lvwColumnCenter
    End With
    With clientes.ColumnHeaders
        .Add , , "Cliente", 3000, lvwColumnLeft
        .Add , , "Capacidad", 1200, lvwColumnCenter
        .Add , , "Num.Botes", 900, lvwColumnCenter
        .Add , , "ID_CLIENTE", 0, lvwColumnCenter
        .Add , , "ID_CAPACIDAD", 0, lvwColumnCenter
        .Add , , "Precio Bote", 1100, lvwColumnCenter
        .Add , , "Pedido", 3000, lvwColumnLeft
        .Add , , "PEDIDO_ID", 0, lvwColumnLeft
        .Add , , "DOC_ID", 0, lvwColumnLeft
    End With
End Sub



Private Sub listaReactivos_DblClick()
    If listaReactivos.ListItems.Count > 0 Then
        frmRPR_Bote.PK = listaReactivos.ListItems(listaReactivos.selectedItem.Index).Text
        frmRPR_Bote.Show 1
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargar_producto()
   On Error GoTo cargar_producto_Error

    On Error GoTo fallo
'    If cmbProducto.BoundText <> "" Then
    If cmbProducto.getTEXTO <> "" Then
        listaReactivos.ListItems.Clear
        clientes.ListItems.Clear
        ' Número de Lote
        Dim oSumLote As New clsSuministros_lotes
        With oSumLote
'            txtDatos(0) = .Proximo_NUMERO_LOTE(cmbProducto.BoundText, Year(Date))
            txtDatos(0) = .Proximo_NUMERO_LOTE(cmbProducto.getPK_SALIDA, Year(Date))
        End With
        ' Reactivos propios a suministrar
        Dim oSumTipo As New clsSuministros_tipos
        oSumTipo.Carga cmbProducto.getPK_SALIDA
        cargar_reactivos oSumTipo.getID_REACTIVO_PR
        ' Clientes
        Dim oSumClientes As New clsSuministros_clientes
        Dim rs As ADODB.Recordset
        Set rs = oSumClientes.Listado(cmbProducto.getPK_SALIDA)
        If rs.RecordCount > 0 Then
            Do
                With clientes.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(1)
                    .SubItems(2) = rs(2)
                    .SubItems(3) = rs(3)
                    .SubItems(4) = rs(4)
                    .SubItems(5) = moneda(rs(5)) ' PRECIO
                    
                    .SubItems(6) = "" 'Pedido
                    .SubItems(7) = "0" 'Pedido_id
                    .SubItems(8) = "0" ' Doc_id
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        
    End If
    Exit Sub
fallo:
    error_grave (Err.Description)

   On Error GoTo 0
   Exit Sub

cargar_producto_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_producto of Formulario frmSuministros_Lote"
End Sub
Private Function validar() As Boolean
    validar = True
'    If cmbProducto.BoundText = "" Then
    If cmbProducto.getTEXTO = "" Then
        MsgBox "Debe seleccionar el producto a suministrar.", vbInformation, App.Title
        cmbProducto.SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(0) = "" Then
        MsgBox "Debe especificar el número de lote.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    Else
        If Not IsNumeric(txtDatos(0)) Then
            MsgBox "El número de lote debe ser numérico.", vbInformation, App.Title
            txtDatos(0).SetFocus
            validar = False
            Exit Function
        End If
    End If
    If listaReactivos.ListItems.Count = 0 Then
        MsgBox "Debe indicar el reactivo propio que se utiliza.", vbInformation, App.Title
        listaReactivos.SetFocus
        validar = False
    End If
    If clientes.ListItems.Count = 0 Then
        MsgBox "Debe indicar algún cliente.", vbInformation, App.Title
        clientes.SetFocus
        validar = False
    End If
End Function

Private Sub cargar_combos()
'    cargar_combo cmbProducto, New clsSuministros_tipos
    llenar_combo cmbProducto, New clsSuministros_tipos, 0, frmSuministros_Tipos, ""
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    cargar_combo cmbCapacidad, New clsAlodine_capacidad
    pedidos 0
End Sub
Private Sub cargar_reactivos(Reactivo As Long)
    Dim rs As ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT A.ID_BOTE_PR, " & _
               "       A.NUMERO, " & _
               "       B.CODIGO, " & _
               "       B.NOMBRE, " & _
               "       A.FECHA_FABRICACION, " & _
               "       A.FECHA_CADUCIDAD, " & _
               "       A.VOLUMEN " & _
               " FROM RPR_BOTES A, " & _
               "      RPR_TIPOS B " & _
               " WHERE A.TIPO_REACTIVO_PR_ID = B.ID_TIPO_REACTIVO_PR "
    If PK = 0 Then
       consulta = consulta & "   AND A.TIPO_REACTIVO_PR_ID=" & Reactivo & _
                             "   AND A.TIPO_ID = 2"
    Else
       consulta = consulta & "   AND A.ID_BOTE_PR=" & Reactivo
    End If
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            With listaReactivos.ListItems.Add(, , Format(rs(0), "00000"))
                .SubItems(1) = Format(rs(0), "00000")
                .SubItems(2) = rs(2) & "-" & Format(rs(1), "000")
                .SubItems(3) = rs(3)
                .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
                .SubItems(5) = Format(rs(5), "dd-mm-yyyy")
                .SubItems(6) = rs(6)
            End With
            rs.MoveNext
        Wend
    End If
End Sub

Private Sub cargar_lote()
    Dim oSumLote As New clsSuministros_lotes
    With oSumLote
        .Carga PK
'        cmbProducto.BoundText = .getTIPO_SUMINISTRO_ID
        cmbProducto.MostrarElemento .getTIPO_SUMINISTRO_ID
        txtDatos(0) = .getNUMERO_LOTE
        
        cargar_reactivos .getBOTE_PR_ID
        ' Clientes
        Dim oSumClientes As New clsSuministros_lotes_clientes
        Dim rs As ADODB.Recordset
        Set rs = oSumClientes.Listado(PK)
        If rs.RecordCount > 0 Then
            Do
                With clientes.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(1)
                    .SubItems(2) = rs(2)
                    .SubItems(3) = rs(3)
                    .SubItems(4) = rs(4)
                    .SubItems(5) = moneda(rs(5)) ' PRECIO
                    .SubItems(6) = Trim(rs(6)) ' PEDIDO
                    .SubItems(7) = rs(7) ' PEDIDO_ID
                    .SubItems(8) = rs(8) ' DOC_ID
                End With
                rs.MoveNext
            Loop Until rs.EOF
            clientes_Click
        End If
    End With
    Set oSumLote = Nothing
End Sub
Private Sub pedidos(ID As Long)
    Dim filtro As String
    If ID <> 0 Then
        If listaReactivos.ListItems.Count > 0 Then
            filtro = " AND CLIENTE_ID = " & ID & " AND FECHA_BAJA >= '" & Format(listaReactivos.ListItems(listaReactivos.selectedItem.Index).SubItems(4), "YYYY-MM-DD") & "'"
        Else
            filtro = " AND CLIENTE_ID = " & ID
        End If
    End If
    cmbPedidos.limpiar
    llenar_combo cmbPedidos, New clsClientes_pedidos, 0, frmClientes_Pedidos, filtro
End Sub

Private Sub txtprecio_LostFocus()
    If txtPrecio <> "" Then
        txtPrecio = moneda(txtPrecio)
    End If

End Sub
