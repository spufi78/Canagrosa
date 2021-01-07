VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmInformePedidosClientes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Pedidos de Clientes"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15000
   Icon            =   "frmInformePedidosClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   15000
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   13875
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8055
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
      Height          =   1515
      Left            =   45
      TabIndex        =   10
      Top             =   360
      Width           =   14895
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
         Left            =   10425
         TabIndex        =   1
         Top             =   270
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox txtFiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   0
         Left            =   5130
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1080
         Width           =   2085
      End
      Begin VB.TextBox txtFiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   1
         Left            =   9000
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1080
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   855
         TabIndex        =   2
         Top             =   660
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
         Format          =   60555265
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   2970
         TabIndex        =   3
         Top             =   660
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
         Format          =   60555265
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbFiltroTipo 
         Height          =   315
         Left            =   855
         TabIndex        =   4
         Top             =   1080
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   855
         TabIndex        =   0
         Top             =   270
         Width           =   9330
         _ExtentX        =   16457
         _ExtentY        =   582
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   975
         Left            =   13365
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   315
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   19
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   240
         Index           =   9
         Left            =   4500
         TabIndex        =   18
         Top             =   1125
         Width           =   630
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Index           =   8
         Left            =   7785
         TabIndex        =   17
         Top             =   1125
         Width           =   1185
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   16
         Top             =   1125
         Width           =   420
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Alta"
         Height          =   195
         Index           =   6
         Left            =   165
         TabIndex        =   15
         Top             =   750
         Width           =   465
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   4
         Left            =   2400
         TabIndex        =   14
         Top             =   705
         Width           =   405
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5820
      Left            =   45
      TabIndex        =   8
      Top             =   2205
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   10266
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
   Begin VB.Label lblBase 
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
      Left            =   11745
      TabIndex        =   21
      Top             =   8055
      Width           =   1980
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Importe"
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
      Index           =   0
      Left            =   10710
      TabIndex        =   20
      Top             =   8055
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Informes de Pedidos de Clientes"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
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
      TabIndex        =   13
      Top             =   0
      Width           =   14910
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
      Left            =   60
      TabIndex        =   12
      Top             =   1890
      Width           =   14895
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre el análisis para ver el detalle"
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
      Left            =   5310
      TabIndex        =   11
      Top             =   8055
      Width           =   4095
   End
End
Attribute VB_Name = "frmInformePedidosClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTodas_Click()
    If chkTodas.value = Checked Then
        cmbTiposMuestra.Limpiar
        cmbTiposMuestra.desactivar
    Else
        cmbTiposMuestra.activar
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.value = Checked Then
        cmbClientes.Limpiar
        cmbClientes.desactivar
    Else
        cmbClientes.activar
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
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
    fdesde = Date - 30
    fhasta = Date
    cabecera
    cargar_combos
End Sub
Private Sub cargar_combos()
    cmbClientes.desactivar
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbFiltroTipo, DECODIFICADORA.PEDIDOS_CLIENTES_TIPOS
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Cliente", 2000, lvwColumnLeft
        .Add , , "Tipo", 2000, lvwColumnLeft
        .Add , , "Código", 2000, lvwColumnCenter
        .Add , , "Descripción", 3000, lvwColumnLeft
        .Add , , "F.Alta", 1050, lvwColumnCenter
        .Add , , "F.Pedido", 1050, lvwColumnCenter
        .Add , , "F.Baja", 1050, lvwColumnCenter
        .Add , , "Importe", 1050, lvwColumnRight
        .Add , , "ID_TIPO", 1, lvwColumnRight
        .Add , , "Restan", 1200, lvwColumnRight
    End With
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim oCliente_Pedidos As New clsClientes_pedidos
    Dim rs As ADODB.Recordset
    Dim idCliente As Long
    idCliente = 0
    If cmbClientes.getTEXTO <> "" Then
        idCliente = cmbClientes.getPK_SALIDA
    End If
    Set rs = oCliente_Pedidos.ListadoCompleto(idCliente, cmbFiltroTipo.BoundText, txtFiltro(0), txtFiltro(1), fdesde, fhasta)
    Dim oDoc As New clsDocs_pago
    lista.ListItems.Clear
    Dim IMPORTE As Currency
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(10) ' CLIENTE
                .SubItems(2) = rs(6) ' TIPO
                .SubItems(3) = rs(1)
                .SubItems(4) = rs(2)
                .SubItems(5) = Format(rs(8), "dd-mm-yyyy") ' F.Alta
                .SubItems(6) = Format(rs(3), "dd-mm-yyyy") ' F.Pedido
                .SubItems(7) = Format(rs(4), "dd-mm-yyyy") ' F.Baja
                .SubItems(8) = Format(rs(5), "currency")
                .SubItems(9) = rs(7) ' TIPO_ID
                .SubItems(10) = moneda(rs(9))
                IMPORTE = IMPORTE + CCur(rs(5))
            End With
            rs.MoveNext
        Loop Until rs.EOF
'        lista_Click
        lblBase = moneda(CStr(IMPORTE))
    End If
    Set oCliente_Pedidos = Nothing
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
        frmClientes_Detalle_Pedido.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmClientes_Detalle_Pedido.Show 1
    End If
End Sub
