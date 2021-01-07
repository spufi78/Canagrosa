VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmSE_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Sellantes"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13800
   Icon            =   "frmSE_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   13800
   Begin VB.CommandButton cmdanular 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   870
      Left            =   3330
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7965
      Width           =   1050
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
      Height          =   1095
      Left            =   45
      TabIndex        =   13
      Top             =   585
      Width           =   13650
      Begin VB.CheckBox chkAnulados 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Anulados"
         Height          =   195
         Left            =   180
         TabIndex        =   19
         Top             =   810
         Width           =   2490
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   10530
         MaxLength       =   75
         TabIndex        =   3
         Top             =   360
         Width           =   1545
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   945
         MaxLength       =   75
         TabIndex        =   0
         Top             =   360
         Width           =   1680
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   7785
         MaxLength       =   75
         TabIndex        =   2
         Top             =   360
         Width           =   1545
      End
      Begin VB.CommandButton cmdLimpiarCampos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar Campos"
         Height          =   915
         Left            =   12240
         Picture         =   "frmSE_Listado.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   135
         Width           =   1320
      End
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   315
         Left            =   3555
         TabIndex        =   1
         Top             =   360
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   195
         Index           =   3
         Left            =   9720
         TabIndex        =   17
         Top             =   405
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   405
         Width           =   525
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   2790
         TabIndex        =   15
         Top             =   405
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
         Height          =   195
         Index           =   2
         Left            =   7020
         TabIndex        =   14
         Top             =   405
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7965
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12690
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7965
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6210
      Left            =   45
      TabIndex        =   10
      Top             =   1710
      Width           =   13680
      _ExtentX        =   24130
      _ExtentY        =   10954
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
      Caption         =   "Listado de Sellantes"
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
      TabIndex        =   12
      Top             =   45
      Width           =   2145
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13185
      Picture         =   "frmSE_Listado.frx":711C
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "En la lista existen un total de 0 registros"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   11
      Top             =   270
      Width           =   2775
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   13770
   End
End
Attribute VB_Name = "frmSE_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAnulados_Click()
    cargar_lista
End Sub
Private Sub cmbCliente_Change()
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    gSE_Sellante = 0
    frmSE_Detalle.Show 1
    cargar_lista
End Sub
Private Sub cmdanular_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ANULAR el tipo de sellante : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oSellante As New clsSellantes
            If oSellante.Anular(lista.ListItems(lista.selectedItem.Index)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdduplicar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a duplicar el sellante. ¿esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oSellante As New clsSellantes
            Dim oSellante_Nuevo As New clsSellantes
            Dim SELLANTE As Long
            If oSellante.Carga(lista.ListItems(lista.selectedItem.Index)) Then
                With oSellante_Nuevo
                    .setENSAYO = oSellante.getENSAYO & " (Duplicado)"
                    .setENSAYO_INGLES = oSellante.getENSAYO_INGLES
                    .setCLIENTE_ID = oSellante.getCLIENTE_ID
                    .setPROCESO = oSellante.getPROCESO
                    .setPROCESO_INGLES = oSellante.getPROCESO_INGLES
                    .setINSTALACION = oSellante.getINSTALACION
                    .setINSTALACION_INGLES = oSellante.getINSTALACION_INGLES
                    .setPREPARACION = oSellante.getPREPARACION
                    .setPREPARACION_INGLES = oSellante.getPREPARACION_INGLES
                    .setPRODUCTO = oSellante.getPRODUCTO
                    .setOBSERVACIONES = oSellante.getOBSERVACIONES
                    SELLANTE = .Insertar
                End With
                Dim rs As ADODB.Recordset
                Dim oSe_ensayos As New clsSellantes_ensayos
                Set rs = oSe_ensayos.Listado(lista.ListItems(lista.selectedItem.Index))
                If rs.RecordCount > 0 Then
                    Do
                        With oSe_ensayos
                            .setSELLANTE_ID = SELLANTE
                            .setORDEN = rs("ORDEN")
                            .setENSAYO = rs("ENSAYO")
                            .setENSAYO_INGLES = rs("ENSAYO_INGLES")
                            .setRANGO_INFERIOR = rs("RANGO_INFERIOR")
                            .setRANGO_SUPERIOR = rs("RANGO_SUPERIOR")
                            .setUNIDAD_ID = rs("UNIDAD_ID")
                            .setTIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
                            .setREFERENCIA = rs("REFERENCIA")
                            .Insertar
                        End With
                        rs.MoveNext
                    Loop Until rs.EOF
                End If
                cargar_lista
                MsgBox "Sellante duplicado correctamente.", vbInformation + vbOKOnly, App.Title
                
            End If
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el tipo de sellante : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oSellante As New clsSellantes
            If oSellante.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdLimpiarCampos_Click()
    txtdatos(0) = ""
    txtdatos(1) = ""
    txtdatos(2) = ""
    cmbCliente.Text = ""
    cmbCliente.BoundText = ""
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        gSE_Sellante = lista.ListItems(lista.selectedItem.Index)
        frmSE_Detalle.Show 1
        actualizar_lista
        gSE_Sellante = 0
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_combos
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Ensayo", 3200, lvwColumnLeft
        .Add , , "Cliente", 3200, lvwColumnLeft
        .Add , , "Proceso", 3200, lvwColumnLeft
        .Add , , "Producto", 3000, lvwColumnLeft
        .Add , , "Anulado", 800, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oSellante As New clsSellantes
    lista.ListItems.Clear
    Set rs = oSellante.Listado(txtdatos(1), cmbCliente.BoundText, txtdatos(0), txtdatos(2), chkAnulados.Value)
    lbltitulo(1) = "En la lista existen un total de " & rs.RecordCount & " registros"
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             If rs(5) <> 0 Then
                .SubItems(5) = "Anulado"
             End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oSellante = Nothing
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
    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim oSellante As New clsSellantes
    With oSellante
        If .Carga(lista.ListItems(lista.selectedItem.Index)) = True Then
            Dim oCliente As New clsCliente
            oCliente.CargaCliente .getCLIENTE_ID
            lista.ListItems(lista.selectedItem.Index).SubItems(1) = .getENSAYO
            lista.ListItems(lista.selectedItem.Index).SubItems(2) = oCliente.getNOMBRE
            lista.ListItems(lista.selectedItem.Index).SubItems(3) = .getPROCESO
            lista.ListItems(lista.selectedItem.Index).SubItems(4) = .getPRODUCTO
            If .getANULADO <> 0 Then
                lista.ListItems(lista.selectedItem.Index).SubItems(5) = "Anulado"
            End If
        End If
    End With
    Set oSellante = Nothing
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub cargar_combos()
    Dim oSellante As New clsSellantes
    Set cmbCliente.RowSource = oSellante.Listado_Combo_Clientes
    cmbCliente.ListField = "C2"
    cmbCliente.DataField = "C1" 'campo asociado
    cmbCliente.BoundColumn = "C1" 'lo que realmente
    Set oSellante = Nothing
End Sub
