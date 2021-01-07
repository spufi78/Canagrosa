VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFluidos_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Fluidos"
   ClientHeight    =   8805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13800
   Icon            =   "frmFluidos_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   13800
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
      Height          =   1065
      Left            =   30
      TabIndex        =   8
      Top             =   750
      Width           =   13725
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1050
         TabIndex        =   10
         Top             =   240
         Width           =   4515
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   795
         Left            =   12720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   210
         Width           =   915
      End
      Begin MSDataListLib.DataCombo cmbBano 
         Height          =   330
         Left            =   6780
         TabIndex        =   11
         Top             =   240
         Width           =   5820
         _ExtentX        =   10266
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
      Begin MSDataListLib.DataCombo cmbCliente 
         Height          =   330
         Left            =   6780
         TabIndex        =   14
         Top             =   630
         Width           =   5820
         _ExtentX        =   10266
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
      Begin MSDataListLib.DataCombo cmbClasificacion 
         Height          =   330
         Left            =   1050
         TabIndex        =   16
         Top             =   630
         Width           =   4500
         _ExtentX        =   7938
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clasificación"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   17
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   5970
         TabIndex        =   15
         Top             =   690
         Width           =   480
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fluido"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         Height          =   195
         Index           =   0
         Left            =   5970
         TabIndex        =   12
         Top             =   300
         Width           =   375
      End
   End
   Begin VB.CommandButton cmdNormas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Normas"
      Height          =   870
      Left            =   3330
      Picture         =   "frmFluidos_Listado.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7905
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7905
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7905
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7905
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12690
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7905
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5985
      Left            =   45
      TabIndex        =   0
      Top             =   1875
      Width           =   13680
      _ExtentX        =   24130
      _ExtentY        =   10557
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
      Caption         =   "Listado de Fluidos"
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
      TabIndex        =   6
      Top             =   60
      Width           =   1935
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13095
      Picture         =   "frmFluidos_Listado.frx":1B3C
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmFluidos_Listado.frx":208C
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   90
      TabIndex        =   5
      Top             =   420
      Width           =   12705
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   705
      Left            =   0
      Top             =   0
      Width           =   13770
   End
End
Attribute VB_Name = "frmFluidos_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbBano_Change()
    cargar_lista
End Sub
Private Sub cmbClasificacion_Change()
    cargar_lista
End Sub

Private Sub cmbCliente_Change()
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    frmFluidos_Detalle.PK = 0
    frmFluidos_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar la ficha de fluido : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oFluido As New clsFluidos_ficha
            If oFluido.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdLimpiar_Click()
    txtfiltro(0) = ""
    cmbClasificacion.Text = ""
    cmbBano.Text = ""
    cmbCliente.Text = ""
    cargar_lista
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmFluidos_Detalle.PK = lista.ListItems(lista.selectedItem.Index)
        frmFluidos_Detalle.Show 1
        actualizar_lista
'        gSE_Sellante = lista.ListItems(lista.SelectedItem.Index)
'        frmSE_Detalle.Show 1
'        actualizar_lista
'        gSE_Sellante = 0
    End If
End Sub

Private Sub cmdNormas_Click()
    frmFluidos_Normas.Show 1
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    cabecera
    cargar_combos
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Fluido", 3400, lvwColumnLeft
        .Add , , "Clasificación (Tipo Muestra)", 2500, lvwColumnLeft
        .Add , , "SubClasificación", 2000, lvwColumnLeft
        .Add , , "Baño", 2800, lvwColumnLeft
        .Add , , "Cliente", 2700, lvwColumnLeft
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oFicha As New clsFluidos_ficha
    lista.ListItems.Clear
    Set rs = oFicha.Listado(txtfiltro(0), cmbClasificacion.BoundText, cmbBano.BoundText, cmbCliente.BoundText)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oFicha = Nothing
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

Private Sub actualizar_lista()
    Dim rs As ADODB.Recordset
    Dim oFicha As New clsFluidos_ficha
    Set rs = oFicha.Listado_por_ID(lista.ListItems(lista.selectedItem.Index).Text)
    If rs.RecordCount <> 0 Then
        With lista.ListItems(lista.selectedItem.Index)
         .SubItems(1) = rs(1)
         .SubItems(2) = rs(2)
         .SubItems(3) = rs(3)
         .SubItems(4) = rs(4)
         .SubItems(5) = rs(5)
        End With
    End If
    Set oFicha = Nothing
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub cargar_combos()
    ' CLASIFICACION
    Dim otm As New clsTipos_muestra
    Set cmbClasificacion.RowSource = otm.Listado_Fluidos
    cmbClasificacion.ListField = "nombre"
    cmbClasificacion.BoundColumn = "id_tipo_muestra"
    ' BANO
    Dim oFD As New clsFluidos_ficha
    Set cmbBano.RowSource = oFD.Listado_Banos
    cmbBano.ListField = "NOMBRE"
    cmbBano.BoundColumn = "ID_BANO"
    ' CLIENTE
    Set cmbCliente.RowSource = oFD.Listado_Clientes
    cmbCliente.ListField = "NOMBRE"
    cmbCliente.BoundColumn = "ID_CLIENTE"
    
End Sub
