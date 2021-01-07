VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmGastos_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Otros Gastos"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14115
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmGastos_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   14115
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por"
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
      Height          =   1410
      Left            =   45
      TabIndex        =   9
      Top             =   405
      Width           =   14055
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   8955
         TabIndex        =   0
         Top             =   180
         Width           =   3255
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   825
         Left            =   12870
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   1050
      End
      Begin pryCombo.miCombo cmbProveedor 
         Height          =   345
         Left            =   1035
         TabIndex        =   11
         Top             =   225
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbBanco 
         Height          =   345
         Left            =   1035
         TabIndex        =   12
         Top             =   585
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbTipo 
         Height          =   345
         Left            =   1035
         TabIndex        =   15
         Top             =   945
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   8955
         TabIndex        =   17
         Top             =   585
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
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   10935
         TabIndex        =   18
         Top             =   585
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
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   0
         Left            =   10395
         TabIndex        =   20
         Top             =   675
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   7965
         TabIndex        =   19
         Top             =   675
         Width           =   450
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Gasto"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   990
         Width           =   780
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   14
         Top             =   270
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Banco"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   630
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   7965
         TabIndex        =   10
         Top             =   225
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13035
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8235
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8235
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8235
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8235
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6315
      Left            =   45
      TabIndex        =   6
      Top             =   1875
      Width           =   14055
      _ExtentX        =   24791
      _ExtentY        =   11139
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
   Begin VB.Label lblTotal 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total :"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   12510
      TabIndex        =   8
      Top             =   90
      Width           =   1260
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Otros Gastos"
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
      TabIndex        =   7
      Top             =   30
      Width           =   2520
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   14200
   End
End
Attribute VB_Name = "frmGastos_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbBanco_change()
    cargar_lista
End Sub

Private Sub cmbProveedor_change()
    cargar_lista
End Sub

Private Sub cmbTipo_change()
    cargar_lista
End Sub

Private Sub cmdLimpiar_Click()
    txtfiltro(0) = ""
    cmbProveedor.Limpiar
    cmbBanco.Limpiar
    cmbTipo.Limpiar
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    frmGastos_Detalle.PK = 0
    frmGastos_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el Gasto : " & lista.ListItems(lista.selectedItem.Index).SubItems(3), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oGasto As New clsGastos
            oGasto.Eliminar (lista.ListItems(lista.selectedItem.Index).Text)
            Set oGasto = Nothing
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmGastos_Detalle.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmGastos_Detalle.Show 1
        modificar_ensayo
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

Private Sub Form_Activate()
    Me.SetFocus
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.top = 100
    Me.Left = 100
    cargar_botones Me
    fdesde = Date - 60
    fhasta = Date
    cabecera
    cargar_combo
    cargar_lista
End Sub
Private Sub cargar_combo()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbTipo, DECODIFICADORA.DECODIFICADORA_GASTOS_TIPOS
    llenar_combo cmbProveedor, New clsProveedor, 0, Me, " ANULADO = 0 "
    llenar_combo cmbBanco, New clsBancos, 0, Me, ""
    Set oDeco = Nothing
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 750, lvwColumnLeft
        .Add , , "Tipo", 1500, lvwColumnCenter
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Descripción", 3500, lvwColumnLeft
        .Add , , "Proveedor", 2400, lvwColumnLeft
        .Add , , "Banco", 1500, lvwColumnCenter
        .Add , , "F.Pago", 1500, lvwColumnCenter
        .Add , , "Importe", 1200, lvwColumnRight
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oGastos As New clsGastos
    Dim proveedor As Long
    If cmbProveedor.getTEXTO = "" Then
        proveedor = 0
    Else
        proveedor = cmbProveedor.getPK_SALIDA
    End If
    Dim banco As Long
    If cmbBanco.getTEXTO = "" Then
        banco = 0
    Else
        banco = cmbBanco.getPK_SALIDA
    End If
    Dim tipo As Long
    If cmbTipo.getTEXTO = "" Then
        tipo = 0
    Else
        tipo = cmbTipo.getPK_SALIDA
    End If
    Set rs = oGastos.Listado(proveedor, banco, tipo, txtfiltro(0), fdesde, fhasta)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "00000"))
                .SubItems(1) = rs(1) ' tipo
                .SubItems(2) = Format(rs(2), "dd-mm-yyyy") ' fecha
                .SubItems(3) = rs(3) ' des
                .SubItems(4) = rs(4) ' proveedor
                .SubItems(5) = rs(5) ' banco
                .SubItems(6) = rs(6) ' fp
                .SubItems(7) = moneda(rs(7)) ' importe
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oGastos = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    lblTotal = "Total : " & lista.ListItems.Count
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
Private Sub modificar_ensayo()
    Dim oGasto As New clsGastos
    Dim rs As ADODB.Recordset
    Set rs = oGasto.ListadoID(lista.ListItems(lista.selectedItem.Index))
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1) 'tipo
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = Format(rs(2), "dd-mm-yyyy") ' fecha
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs(3) 'des
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs(4) 'prov
        lista.ListItems(lista.selectedItem.Index).SubItems(5) = rs(5) 'banco
        lista.ListItems(lista.selectedItem.Index).SubItems(6) = rs(6) 'f.p.
        lista.ListItems(lista.selectedItem.Index).SubItems(7) = moneda(rs(7)) ' importe
    End If
    Set rs = Nothing
    Set oBanco = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
