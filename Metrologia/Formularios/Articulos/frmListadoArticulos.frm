VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmListadoArticulos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Articulos"
   ClientHeight    =   9135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   Icon            =   "frmListadoArticulos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9135
   ScaleWidth      =   12585
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   885
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8190
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8190
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8190
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8190
      Width           =   1155
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8190
      Width           =   1155
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar listado de artículos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1200
      Left            =   90
      TabIndex        =   3
      Top             =   390
      Width           =   12465
      Begin VB.CommandButton cmdRestaurar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar Campos"
         Height          =   915
         Left            =   11385
         Picture         =   "frmListadoArticulos.frx":09EA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   180
         Width           =   960
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   915
         Left            =   10395
         Picture         =   "frmListadoArticulos.frx":12B4
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   960
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
         Height          =   315
         Index           =   0
         Left            =   1620
         TabIndex        =   0
         Top             =   495
         Width           =   3090
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   5580
         TabIndex        =   9
         Top             =   495
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5010
         TabIndex        =   10
         Top             =   555
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   555
         Width           =   990
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6495
      Left            =   75
      TabIndex        =   2
      Top             =   1635
      Width           =   12480
      _ExtentX        =   22013
      _ExtentY        =   11456
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14609914
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
      Caption         =   "Listado de Artículos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   3
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13545
   End
End
Attribute VB_Name = "frmListadoArticulos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdImprimir_Click()
    Dim FILTRO As String
    If txtdes(0) <> "" Then
        FILTRO = FILTRO & " {articulos.DESCRIPCION} like '*" & txtdes(0) & "*'"
    End If
    If cmbTipo.Text <> "" Then
        If FILTRO <> "" Then
            FILTRO = FILTRO & " AND "
        End If
        FILTRO = FILTRO & " {articulos.TIPO_ARTICULO_ID} = " & cmbTipo.BoundText
    End If
    With frmReport
        .iniciar
        .CRITERIO = FILTRO
        .informe = "rptARTICULOS_LISTADO"
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport
End Sub
Private Sub cmbTipo_Change()
    cargar_lista
End Sub

Private Sub cmdRestaurar_Click()
    cmbTipo.Text = ""
    txtdes(0) = ""
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    frmAnadirArticulo.pk = 0
    frmAnadirArticulo.Show 1
    cargar_lista
    lista.SetFocus
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        Dim pos As Integer
        Dim cliente As Integer
        If MsgBox("Va a ELIMINAR el Artículo " & lista.ListItems(lista.SelectedItem.Index).SubItems(2) & ". ¿Esta seguro?", vbQuestion + vbYesNo, "Informacion") = vbYes Then
            Dim oArt As New clsArticulos
            If oArt.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
                cargar_lista
            End If
            Set oArt = Nothing
        End If
        lista.SetFocus
    End If
End Sub
Private Sub cmdListado_Click()
    cargar_lista
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmAnadirArticulo.pk = lista.ListItems(lista.SelectedItem.Index)
        frmAnadirArticulo.Show 1
        actualizar_lista
        lista.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 100
    Me.Left = 100
    cargar_botones Me
    cabecera
    Cargar_Combo cmbTipo, New clsArticulos_Tipos
    cargar_lista
End Sub

Public Sub cargar_lista()
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim oArticulo As New clsArticulos
    Set rs = oArticulo.Listado(txtdes(0), cmbTipo.BoundText)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "0000"))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
            .SubItems(4) = moneda(rs(4))
            .SubItems(5) = moneda(rs(5))
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oArticulo = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    Exit Sub
fallo:
    MsgBox "Error al cargar la lista: " & Err.Description, vbCritical, App.Title
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
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems(lista.SelectedItem.Index) <> "" Then
      cmdmodificar.Enabled = True
      cmdeliminar.Enabled = True
    End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
Public Sub actualizar_lista()
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim oArticulo As New clsArticulos
    Set rs = oArticulo.Listado_PK(lista.ListItems(lista.SelectedItem.Index).Text)
    If rs.RecordCount <> 0 Then
           With lista.ListItems(lista.SelectedItem.Index)
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
                .SubItems(4) = moneda(rs(4))
                .SubItems(5) = moneda(rs(5))
           End With
    End If
    Exit Sub
fallo:
    MsgBox "Error al cargar la lista: " & Err.Description, vbCritical, App.Title
End Sub
Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    lista_Click
End Sub
Private Sub txtdes_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub txtdes_GotFocus(Index As Integer)
    txtdes(Index).BackColor = &H80C0FF
    txtdes(Index).SelStart = 0
    txtdes(Index).SelLength = Len(txtdes(Index))
End Sub
Private Sub txtdes_LostFocus(Index As Integer)
    txtdes(Index).BackColor = vbWhite
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Codigo", 1500, lvwColumnLeft
        .Add , , "Tipo", 2200, lvwColumnCenter
        .Add , , "Descripcion", 4600, lvwColumnLeft
        .Add , , "Proveedor", 1700, lvwColumnLeft
        .Add , , "Precio", 1100, lvwColumnRight
        .Add , , "Comisión", 1100, lvwColumnRight
    End With
End Sub
