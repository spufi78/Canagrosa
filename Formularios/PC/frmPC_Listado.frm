VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmPC_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Pedidos de Productos Controlados"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13305
   Icon            =   "frmPC_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   13305
   Begin VB.CommandButton cmdAlbaran 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Albaran"
      Height          =   870
      Left            =   4545
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8775
      Width           =   1185
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
      Height          =   1425
      Left            =   45
      TabIndex        =   13
      Top             =   675
      Width           =   13230
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   795
         Left            =   12195
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   390
         Width           =   915
      End
      Begin pryCombo.miCombo cmbProducto 
         Height          =   330
         Left            =   1050
         TabIndex        =   0
         Top             =   240
         Width           =   11070
         _ExtentX        =   19526
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1050
         TabIndex        =   1
         Top             =   630
         Width           =   11070
         _ExtentX        =   19526
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1035
         TabIndex        =   2
         Top             =   990
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
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3150
         TabIndex        =   3
         Top             =   1005
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
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde el"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   1065
         Width           =   645
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta el"
         Height          =   195
         Index           =   4
         Left            =   2475
         TabIndex        =   18
         Top             =   1035
         Width           =   585
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   17
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   300
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdetiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiquetas"
      Height          =   870
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8775
      Width           =   1185
   End
   Begin VB.CommandButton cmdCertificado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Certificado"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8775
      Width           =   1185
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12255
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8760
      Width           =   1020
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8775
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8775
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8775
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6555
      Left            =   60
      TabIndex        =   12
      Top             =   2130
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   11562
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Pedidos de Productos Controlados"
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
      Left            =   90
      TabIndex        =   15
      Top             =   60
      Width           =   4800
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Especifique los datos necesarios para localizar un Pedido"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   360
      Width           =   4050
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   45
      Top             =   0
      Width           =   13275
   End
End
Attribute VB_Name = "frmPC_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbClientes_change()
    cargar_lista
End Sub
Private Sub cmdAlbaran_Click()
    If lista.ListItems.Count > 0 Then
        Dim oPC As New clsPc_pedido
        oPC.ImprimirAlbaran lista.ListItems(lista.selectedItem.Index).Text
        Set oPC = Nothing
    End If
End Sub

Private Sub cmdCertificado_Click()
    If lista.ListItems.Count > 0 Then
        Dim oPC As New clsPc_pedido
        oPC.ImprimirCertificado lista.ListItems(lista.selectedItem.Index).Text
        Set oPC = Nothing
    End If
End Sub

Private Sub cmdetiqueta_Click()
    If lista.ListItems.Count > 0 Then
        Dim oPC As New clsPc_pedido
        oPC.ImprimirEtiquetas lista.ListItems(lista.selectedItem.Index).Text
        Set oPC = Nothing
    End If
End Sub

Private Sub cmdLimpiar_Click()
    cmbProducto.Limpiar
    cmbclientes.Limpiar
    cargar_lista
End Sub
Private Sub cmbproducto_Change()
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    frmPC_Detalle.PK = 0
    frmPC_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el Lote del pedido número: " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oPC As New clsPc_pedido
            If oPC.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
            Set oPC = Nothing
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmPC_Detalle.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmPC_Detalle.Show 1
        actualizar_lista
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub fdesde_Change()
    cargar_lista
End Sub
Private Sub fhasta_Change()
    cargar_lista
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    txtanno = Year(Date)
    cabecera
    cargar_reactivos
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    fdesde = "01/01/" & Year(Date)
    fhasta = Date
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Número", 600, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Cliente", 2800, lvwColumnLeft
        .Add , , "Cantidad", 800, lvwColumnCenter
        .Add , , "Producto", 3300, lvwColumnLeft
        .Add , , "Identificación", 1200, lvwColumnCenter
        .Add , , "P.Bote", 900, lvwColumnCenter
        .Add , , "Pedido", 2100, lvwColumnCenter
        .Add , , "PEDIDO_ID", 1, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPC As New clsPc_pedido
    lista.ListItems.Clear
    Dim producto As String
    Dim cliente As String
    If cmbProducto.getTEXTO <> "" Then
        producto = cmbProducto.getPK_SALIDA
    End If
    If cmbclientes.getTEXTO <> "" Then
        cliente = cmbclientes.getPK_SALIDA
    End If
    Set rs = oPC.Listado(producto, cliente, fdesde, fhasta)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000")) ' Numero
             .SubItems(1) = Format(rs(1), "dd-mm-yyyy") ' Fecha
             .SubItems(2) = rs(2) ' Cliente
             .SubItems(3) = rs(3) ' Cantidad
             .SubItems(4) = rs(4) ' Producto
             .SubItems(5) = rs(5) ' Lote (Identificacion)
             .SubItems(6) = moneda(rs(6)) ' PRECIO
             If Not IsNull(rs(7)) Then
                .SubItems(7) = rs(7) ' PEDIDO
             End If
             .SubItems(8) = rs(8) ' PEDIDO_ID
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oPC = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
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
Private Sub actualizar_lista()
    Dim oPC As New clsPc_pedido
    Dim rs As ADODB.Recordset
    Set rs = oPC.Listado_por_ID(lista.ListItems(lista.selectedItem.Index).Text)
    If rs.RecordCount > 0 Then
        With lista.ListItems(lista.selectedItem.Index)
         .SubItems(1) = Format(rs(1), "dd-mm-yyyy") ' Fecha
         .SubItems(2) = rs(2) ' Cliente
         .SubItems(3) = rs(3) ' Cantidad
         .SubItems(4) = rs(4) ' Producto
         .SubItems(5) = rs(5) ' Lote (Identificacion)
         .SubItems(6) = moneda(rs(6)) ' PRECIO
         .SubItems(7) = rs(7) ' PEDIDO
         .SubItems(8) = rs(8) ' PEDIDO_ID
        End With
    End If
    Set oPC = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
Private Sub cargar_reactivos()
    cmbProducto.Limpiar
    Dim CONSULTA As String
    CONSULTA = " SELECT DISTINCT TB.ID_TIPO_BOTE_EX,T.NOMBRE " & _
               "   FROM TIPOS_REACTIVO_EX T, TIPOS_BOTE_EX TB " & _
               "  WHERE T.ID_TIPO_REACTIVO_EX = TB.TIPO_REACTIVO_EX_ID " & _
               "    AND TB.TIPO_M_REFERENCIA_ID = 7"
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbProducto
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "TIPOS_BOTE_EX"
            .setDESCRIPCION = "Producto a suministrar"
            .setPK = "TB.ID_TIPO_BOTE_EX"
            .setCAMPO = "T.NOMBRE"
            .setFILTRO = ""
            .setQUERY = CONSULTA
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmREX_Bote
        End With
    End If
End Sub


