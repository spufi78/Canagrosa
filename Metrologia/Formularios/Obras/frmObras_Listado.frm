VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmObras_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Obras"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13110
   Icon            =   "frmObras_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8985
   ScaleWidth      =   13110
   Begin VB.CommandButton cmdEstadistica 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Estadística"
      Height          =   885
      Left            =   5970
      Picture         =   "frmObras_Listado.frx":09EA
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8040
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   885
      Left            =   3594
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8040
      Width           =   1155
   End
   Begin VB.CommandButton cmdTarifa 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tarifa"
      Height          =   885
      Left            =   4782
      Picture         =   "frmObras_Listado.frx":12B4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8040
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro de búsqueda"
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
      Height          =   690
      Left            =   45
      TabIndex        =   8
      Top             =   405
      Width           =   13005
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   720
         TabIndex        =   0
         Top             =   270
         Width           =   1425
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbCliente 
         Height          =   375
         Left            =   5040
         TabIndex        =   13
         Top             =   270
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   661
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   2760
         TabIndex        =   1
         Top             =   270
         Width           =   1425
      End
      Begin MSDataListLib.DataCombo cmbTipoObra 
         Height          =   315
         Left            =   10890
         TabIndex        =   2
         Top             =   270
         Width           =   1935
         _ExtentX        =   3413
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
         Caption         =   "Código"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   16
         Top             =   330
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Obra"
         Height          =   195
         Left            =   10080
         TabIndex        =   11
         Top             =   330
         Width           =   705
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   4380
         TabIndex        =   10
         Top             =   330
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obra"
         Height          =   195
         Index           =   1
         Left            =   2280
         TabIndex        =   9
         Top             =   330
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2406
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8040
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1218
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8040
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   11910
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8040
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8040
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6840
      Left            =   60
      TabIndex        =   7
      Top             =   1140
      Width           =   13005
      _ExtentX        =   22939
      _ExtentY        =   12065
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
      Caption         =   "Listado de Obras"
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
      TabIndex        =   17
      Top             =   0
      Width           =   13035
   End
End
Attribute VB_Name = "frmObras_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK_CLIENTE As Long


Private Sub cmdEstadistica_Click()
    gFECHA_DESDE = ""
    gFECHA_HASTA = ""
    frmFechas.Show 1
    If gFECHA_DESDE <> "" Then
            Dim FILTRO As String
            If txtDatos(0) <> "" Then
                Dim i As Integer
                Dim s As String
                For i = 1 To lista.ListItems.Count
                    s = s & lista.ListItems(i).Text & ","
                Next
                FILTRO = FILTRO & " {obras.ID_OBRA} in [" & Left(s, Len(s) - 1) & "]"
            End If
            If txtDatos(1) <> "" Then
                If FILTRO <> "" Then
                    FILTRO = FILTRO & " AND "
                End If
                
                FILTRO = FILTRO & " {obras.NOMBRE} like '*" & txtDatos(1) & "*'"
            End If
            If cmbCliente.getTEXTO <> "" Then
                If FILTRO <> "" Then
                    FILTRO = FILTRO & " AND "
                End If
                FILTRO = FILTRO & " {clientes.ID_CLIENTE} = " & cmbCliente.getPK_SALIDA
            End If
            If cmbTipoObra.Text <> "" Then
                If FILTRO <> "" Then
                    FILTRO = FILTRO & " AND "
                End If
                FILTRO = FILTRO & " {obras.TIPO_OBRA_ID} = " & cmbTipoObra.BoundText
            End If
            
            If FILTRO <> "" Then
                FILTRO = FILTRO & " AND "
            End If
            FILTRO = FILTRO & " {documentos_detalle.CANTIDAD} <> 0.00"
            
            Dim p1() As String
            Dim p2() As String
            ReDim p1(2) As String
            ReDim p2(2) As String
            p1(1) = "FECHA_DESDE"
            p1(2) = "FECHA_HASTA"
            
            p2(1) = gFECHA_DESDE
            p2(2) = gFECHA_HASTA
            With frmReport
                .iniciar
                .CRITERIO = FILTRO
                .informe = "rptobras_estadistica"
                .ParametrosNombre = p1
                .ParametrosValores = p2
                .imprimir = False
                .generar
                .Show 1
            End With
            Unload frmReport
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim FILTRO As String
    If txtDatos(0) <> "" Then
        FILTRO = FILTRO & " {obras.ID_OBRA} = " & txtDatos(0)
    End If
    If txtDatos(1) <> "" Then
        If FILTRO <> "" Then
            FILTRO = FILTRO & " AND "
        End If
        
        FILTRO = FILTRO & " {obras.NOMBRE} like '*" & txtDatos(1) & "*'"
    End If
    If cmbCliente.getTEXTO <> "" Then
        If FILTRO <> "" Then
            FILTRO = FILTRO & " AND "
        End If
        FILTRO = FILTRO & " {clientes.ID_CLIENTE} = " & cmbCliente.getPK_SALIDA
    End If
    If cmbTipoObra.Text <> "" Then
        If FILTRO <> "" Then
            FILTRO = FILTRO & " AND "
        End If
        FILTRO = FILTRO & " {obras.TIPO_OBRA_ID} = " & cmbTipoObra.BoundText
    End If
    
    With frmReport
        .iniciar
        .consulta = ""
        .CRITERIO = FILTRO
        .informe = "rptObras_Listado"
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport
End Sub

Private Sub cmbTipoObra_Change()
    cargar_lista
End Sub
'Private Sub chkanulados_Click()
'    cargar_lista
'End Sub

Private Sub cmbCliente_change()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmObras.pk = 0
    frmObras.Show 1
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
        If MsgBox("Va a ELIMINAR la Obra " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
            Dim oObra As New clsObras
            If oObra.Eliminar(lista.ListItems(lista.SelectedItem.Index).Text) = True Then
                cargar_lista
            End If
            Set oObra = Nothing
        End If
        lista.SetFocus
    End If
End Sub

Private Sub cmdModificar_Click()
    If USUARIO.getPER_3 = 0 Then
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        frmObras.pk = lista.ListItems(lista.SelectedItem.Index)
        frmObras.Show 1
        actualizar_lista
        lista.SetFocus
    End If
End Sub

Private Sub cmdTarifa_Click()
    If lista.ListItems.Count > 0 Then
        frmObras_Tarifas.pk = lista.ListItems(lista.SelectedItem.Index).Text
        frmObras_Tarifas.Show 1
    End If
End Sub

Private Sub Form_Activate()
    Me.SetFocus
    txtDatos(0).SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 100
    Me.Top = 100
    cabecera
    cargar_combos
    If PK_CLIENTE <> 0 Then
        Me.Left = 400
        Me.Top = 400
        cmbCliente.MostrarElemento PK_CLIENTE
    End If
    cargar_lista
    permisos
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oObra As New clsObras
    If cmbCliente.getTEXTO = "" Then
        Set rs = oObra.Listado(txtDatos(0), txtDatos(1), 0, cmbTipoObra.BoundText)
    Else
        Set rs = oObra.Listado(txtDatos(0), txtDatos(1), cmbCliente.getPK_SALIDA, cmbTipoObra.BoundText)
    End If
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(0))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
            .SubItems(4) = rs(4)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oObra = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
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
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index) <> "" Then
          cmdmodificar.Enabled = True
          cmdeliminar.Enabled = True
        End If
        permisos
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim oObra As New clsObras
    With oObra
        If .Carga(lista.ListItems(lista.SelectedItem.Index).Text) = True Then
            lista.ListItems(lista.SelectedItem.Index).SubItems(1) = .getNOMBRE
            Dim ocliente As New clsCliente
            ocliente.CargaCliente .getCLIENTE_ID
            lista.ListItems(lista.SelectedItem.Index).SubItems(2) = ocliente.getNOMBRE
            lista.ListItems(lista.SelectedItem.Index).SubItems(3) = .getDIRECCION
            lista.ListItems(lista.SelectedItem.Index).SubItems(4) = .getTelefono
        End If
    End With
    Set oObra = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       cmdModificar_Click
    End If
End Sub
Public Sub permisos()
    If USUARIO.getPER_1 = 0 Then
'        cmdImprimir.Enabled = False
    End If
    If USUARIO.getPER_2 = 0 Then
        cmdanadir.Enabled = False
    End If
    If USUARIO.getPER_3 = 0 Then
        cmdmodificar.Enabled = False
    End If
    If USUARIO.getPER_4 = 0 Then
        cmdeliminar.Enabled = False
    End If
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Codigo", 1000, lvwColumnLeft
        .Add , , "Obra", 3500, lvwColumnLeft
        .Add , , "Cliente", 3900, lvwColumnLeft
        .Add , , "Dirección", 2200, lvwColumnLeft
        .Add , , "Teléfono", 2100, lvwColumnCenter
    End With
End Sub

Private Sub cargar_combos()
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    Dim oDeco As New clsDecodificadora
    oDeco.Cargar_Combo cmbTipoObra, DECODIFICADORA.D_TIPOS_OBRAS
End Sub
