VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmNC_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Incidencias"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12930
   Icon            =   "frmNC_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   12930
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7155
      Width           =   1020
   End
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
      TabIndex        =   15
      Top             =   315
      Width           =   12840
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Incidencias No procedentes"
         Height          =   195
         Left            =   5445
         TabIndex        =   5
         Top             =   1080
         Width           =   3345
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1140
         MaxLength       =   255
         TabIndex        =   4
         Top             =   1035
         Width           =   3930
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   1050
         Left            =   10485
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1050
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   1050
         Left            =   11655
         Picture         =   "frmNC_Listado.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo cmbtipos 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   225
         Width           =   4050
         _ExtentX        =   7144
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
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   630
         Width           =   4035
         _ExtentX        =   7117
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
      Begin MSDataListLib.DataCombo cmborigen 
         Height          =   315
         Left            =   6345
         TabIndex        =   1
         Top             =   225
         Width           =   3915
         _ExtentX        =   6906
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
      Begin MSDataListLib.DataCombo cmbafectado 
         Height          =   315
         Left            =   6345
         TabIndex        =   3
         Top             =   630
         Width           =   3915
         _ExtentX        =   6906
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Afectado"
         Height          =   195
         Index           =   3
         Left            =   5445
         TabIndex        =   20
         Top             =   690
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   1080
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Origen"
         Height          =   195
         Index           =   5
         Left            =   5445
         TabIndex        =   18
         Top             =   285
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Hecho"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   17
         Top             =   285
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   675
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2310
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7155
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7155
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7155
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11835
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7155
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5310
      Left            =   45
      TabIndex        =   8
      Top             =   1740
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   9366
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
      BackColor       =   &H00C0C0FF&
      Caption         =   "Listado de Incidencias"
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
      Height          =   270
      Index           =   3
      Left            =   45
      TabIndex        =   14
      Top             =   30
      Width           =   13185
   End
End
Attribute VB_Name = "frmNC_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmborigen_Change()
    cargar_lista
End Sub

Private Sub cmdImprimir_Click()
    Dim strIncidencias As String
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        strIncidencias = strIncidencias & CLng(lista.ListItems(i).Text) & ","
    Next
'            .criterio = "{botes_ex.ID_BOTE_EX} in [" & LISTA_REACTIVOS & "]"
    
    With frmReport
        .iniciar
        .informe = "rptNC_Listado"
        .criterio = "{decodificadora.codigo}=11 and {nc.ID_NC} in [" & Left(strIncidencias, Len(strIncidencias) - 1) & "]"
        .imprimir = False
        .generar
        .Visible = True
    End With
End Sub
Private Sub cmbestados_Change()
    cmdBuscar_Click
End Sub

Private Sub cmbtipos_Change()
    cmdBuscar_Click
End Sub
Private Sub cmdAnadir_Click()
    frmNC_Detalle.PK = 0
    frmNC_Detalle.Show 1
    cargar_lista
End Sub
Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar la NO CONFORMIDAD : " & lista.ListItems(lista.selectedItem.Index).SubItems(3), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oNC As New clsNc
            If oNC.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub
Private Sub cmdLimpiar_Click()
    cmbtipos.Text = ""
    cmbestados.Text = ""
    cmbOrigen.Text = ""
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmNC_Detalle.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmNC_Detalle.Show 1
        actualizar_lista
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 100
    Me.Left = 100
    cabecera
    cargar_botones Me
    cargar_combos
    cargar_lista
    permisos
'    If USUARIO.getUSUARIO = "julio" Then
'        cmdCargar.Visible = True
'    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "NºGeneral", 900, lvwColumnLeft
        .Add , , "NºParticular", 900, lvwColumnCenter
        .Add , , "Tipo Hecho", 2300, lvwColumnCenter
        .Add , , "Origen", 2300, lvwColumnCenter
        .Add , , "D. Afectado", 2300, lvwColumnLeft
        .Add , , "F.Alta", 1100, lvwColumnCenter
        .Add , , "F.Cierre", 1100, lvwColumnCenter
        .Add , , "Estado", 1500, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim RS As ADODB.Recordset
    Dim oNC As New clsNc
    lista.ListItems.Clear
    Dim tipo As String
    Dim origen As String
    Dim ESTADO As String
    Dim nombre As String
    Dim afectado As String
    If cmbtipos.Text = "" Then
        tipo = 0
    Else
        tipo = cmbtipos.BoundText
    End If
    If cmbOrigen.Text = "" Then
        origen = 0
    Else
        origen = cmbOrigen.BoundText
    End If
    If cmbestados.Text = "" Then
        ESTADO = 0
    Else
        ESTADO = cmbestados.BoundText
    End If
    If cmbafectado.Text = "" Then
        afectado = 0
    Else
        afectado = cmbafectado.BoundText
    End If
    nombre = txtDatos(1)
    Set RS = oNC.Listado(tipo, origen, ESTADO, nombre, afectado)
    If RS.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(RS(0), "0000"))
             .SubItems(1) = Format(RS(1), "0000")
             .SubItems(2) = RS(2)
             .SubItems(3) = RS(3)
             .SubItems(4) = RS(8)
             If IsDate(RS(5)) Then
                .SubItems(5) = Format(RS(5), "dd-mm-yyyy")
             Else
                .SubItems(5) = ""
             End If
             If IsDate(RS(6)) Then
                 .SubItems(6) = Format(RS(6), "dd-mm-yyyy")
             Else
                .SubItems(6) = ""
             End If
             .SubItems(7) = RS(7)
            End With
            RS.MoveNext
        Loop Until RS.EOF
    End If
    Set oNC = Nothing
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
    Dim RS As ADODB.Recordset
    Dim oNC As New clsNc
    Set RS = oNC.Listado_por_Codigo(lista.ListItems(lista.selectedItem.Index))
    If RS.RecordCount <> 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = Format(RS(1), "0000")
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = RS(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = RS(3)
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = RS(8)
        If IsDate(RS(5)) Then
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = Format(RS(5), "dd-mm-yyyy")
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(5) = ""
        End If
        If IsDate(RS(6)) Then
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = Format(RS(6), "dd-mm-yyyy")
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(6) = ""
        End If
        lista.ListItems(lista.selectedItem.Index).SubItems(7) = RS(7)
    End If
    Set oNorma = Nothing
End Sub

Public Sub cargar_combos()
    Dim oDecodificadora As New clsDecodificadora
    oDecodificadora.cargar_combo cmbtipos, DECODIFICADORA.NC_TIPOS_HECHOS
    oDecodificadora.cargar_combo cmbOrigen, DECODIFICADORA.NC_ORIGENES
    oDecodificadora.cargar_combo cmbestados, DECODIFICADORA.NC_ESTADOS
'J001-I
    oDecodificadora.cargar_combo cmbafectado, DECODIFICADORA.NC_AFECTADO
'J001-F
End Sub
Public Sub permisos()
'    If Not USUARIO.getPER_DOCUMENTACION_CALIDAD Then
'        cmdAnadir.Enabled = False
'        cmdModificar.Enabled = False
'        cmdEliminar.Enabled = False
'    End If
End Sub
Private Sub txtDatos_Change(Index As Integer)
    cmdBuscar_Click
End Sub
