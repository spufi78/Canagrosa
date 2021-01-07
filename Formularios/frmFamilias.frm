VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFamilias 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Familias"
   ClientHeight    =   9930
   ClientLeft      =   2520
   ClientTop       =   2025
   ClientWidth     =   10770
   Icon            =   "frmFamilias.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   10770
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9000
      Width           =   1080
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9630
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   9000
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9000
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9000
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1875
      Left            =   45
      TabIndex        =   11
      Top             =   7065
      Width           =   10650
      Begin VB.CheckBox chkPedido 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incluir en los Pedidos a Proveedor"
         Height          =   285
         Left            =   180
         TabIndex        =   4
         Top             =   1530
         Width           =   3075
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   8145
         TabIndex        =   3
         Top             =   1080
         Width           =   2235
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1575
         TabIndex        =   2
         Top             =   1080
         Width           =   2820
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1575
         TabIndex        =   0
         Top             =   270
         Width           =   8805
      End
      Begin MSDataListLib.DataCombo cmbTipos 
         Bindings        =   "frmFamilias.frx":1272
         Height          =   360
         Left            =   1575
         TabIndex        =   1
         Top             =   675
         Width           =   8805
         _ExtentX        =   15531
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proyecto Contaplus"
         Height          =   195
         Index           =   3
         Left            =   6615
         TabIndex        =   15
         Top             =   1125
         Width           =   1380
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cuenta Contable"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   14
         Top             =   1125
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   13
         Top             =   315
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sector"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   12
         Top             =   720
         Width           =   465
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6675
      Left            =   60
      TabIndex        =   9
      Top             =   360
      Width           =   10650
      _ExtentX        =   18785
      _ExtentY        =   11774
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
      Caption         =   "Mantenimiento de Familias"
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
      Height          =   330
      Index           =   3
      Left            =   60
      TabIndex        =   10
      Top             =   15
      Width           =   10635
   End
End
Attribute VB_Name = "frmFamilias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If MsgBox("Va a insertar la familia. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim ofam As New clsFamilias
        ofam.setNOMBRE = txtDatos(0)
        If cmbTipos.Text <> "" Then
            ofam.setSECTOR_ID = cmbTipos.BoundText
        End If
        ofam.setCC = txtDatos(1)
        ofam.setCODIGO_CONTAPLUS = txtDatos(2)
        ofam.setPEDIDO = chkPedido.value
        ofam.Insertar
        cargar_lista
        txtDatos(0) = ""
        txtDatos(0).SetFocus
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a eliminar la familia : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oTIPO As New clsFamilias
        If oTIPO.Eliminar(CInt(lista.ListItems(lista.selectedItem.Index).SubItems(4))) Then
            cargar_lista
        End If
        Set oTIPO = Nothing
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a modificar la familia. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim ofam As New clsFamilias
            ofam.setNOMBRE = txtDatos(0)
            If cmbTipos.Text <> "" Then
                ofam.setSECTOR_ID = cmbTipos.BoundText
            End If
            ofam.setCC = txtDatos(1)
            ofam.setCODIGO_CONTAPLUS = txtDatos(2)
            ofam.setPEDIDO = chkPedido.value
            ofam.Modificar (lista.ListItems(lista.selectedItem.Index).SubItems(4))
            cargar_lista
            txtDatos(0) = ""
            cmbTipos.Text = ""
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.top = 200
    Me.Left = 200
    cargar_botones Me
    With lista.ColumnHeaders
        .Add , , "Familia", 3400, lvwColumnLeft
        .Add , , "Sector", 3400, lvwColumnLeft
        .Add , , "C. Contable", 1250, lvwColumnCenter
        .Add , , "P. Contaplus", 1250, lvwColumnCenter
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "Pedido", 800, lvwColumnCenter
    End With
    cargar_tipos
    cargar_lista
End Sub
Private Sub cargar_tipos()
    Dim otipos As New clsSectores
    Set cmbTipos.RowSource = otipos.Listado
    cmbTipos.ListField = "nombre"
    cmbTipos.DataField = "id_sector" 'campo asociado
    cmbTipos.BoundColumn = "id_sector" 'lo que realmente
    Set otipos = Nothing
End Sub
Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ofam As New clsFamilias
    Dim oTIPO As New clsSectores
    Set rs = ofam.Listado_completo
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("nombre"))
            oTIPO.CARGAR (rs("sector_id"))
            .SubItems(1) = oTIPO.getNOMBRE
            .SubItems(2) = rs("CC")
            .SubItems(3) = rs("CODIGO_CONTAPLUS")
            .SubItems(4) = rs("id_familia")
            If rs("pedido") = 0 Then
                .SubItems(5) = ""
            Else
                .SubItems(5) = "X"
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oMuestra = Nothing
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtDatos(0).Text = lista.ListItems(lista.selectedItem.Index).Text
        cmbTipos.Text = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        txtDatos(1).Text = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        txtDatos(2).Text = lista.ListItems(lista.selectedItem.Index).SubItems(3)
        If lista.ListItems(lista.selectedItem.Index).SubItems(5) = "" Then
            chkPedido.value = Unchecked
        Else
            chkPedido.value = Checked
        End If
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
