VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmAlbaranesNoValorados 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Albaranes no Valorados"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmAlbaranesNoValorados.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   13635
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro de Selección de Albaranes"
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
      Height          =   1425
      Left            =   60
      TabIndex        =   3
      Top             =   390
      Width           =   13485
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   1050
         Left            =   12150
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1185
      End
      Begin MSDataListLib.DataCombo cmbTipoFacturacion 
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Top             =   990
         Width           =   2925
         _ExtentX        =   5159
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
      Begin vb6projectpryComboBCA.miComboBCA cmbCliente 
         Height          =   345
         Left            =   1380
         TabIndex        =   9
         Top             =   240
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   609
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbObra 
         Height          =   345
         Left            =   1380
         TabIndex        =   10
         Top             =   600
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   609
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Facturación"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1020
         Width           =   1200
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obra"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   630
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   12390
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6390
      Left            =   60
      TabIndex        =   0
      Top             =   1830
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   11271
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
      Caption         =   "Listado de Albaranes No Valorados"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   30
      TabIndex        =   2
      Top             =   30
      Width           =   13545
   End
End
Attribute VB_Name = "frmAlbaranesNoValorados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbCliente_change()
    cargar_lista
End Sub

Private Sub cmbObra_change()
    cargar_lista
End Sub

Private Sub cmbTipoFacturacion_Change()
    cargar_lista
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
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
    cabecera_lista
    cargar_combos
    
    cargar_lista
End Sub
Public Sub cargar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim cliente As String
    Dim obra As String
    Dim numero As String
    Dim anno As String
    Dim ESTADO As String
'    If cmbTipo.Text <> "" Then
        tipo = " AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.ALBARAN
        ESTADO = " AND FACTURADO = 0 AND DD.TOTAL = 0 "
'    End If
    If cmbCliente.getTEXTO <> "" Then
        cliente = " AND O.CLIENTE_ID = " & cmbCliente.getPK_SALIDA
    End If
    If cmbObra.getTEXTO <> "" Then
        obra = " AND D.OBRA_ID = " & cmbObra.getPK_SALIDA
    End If
    If cmbTipoFacturacion.Text <> "" Then
        numero = " AND O.TIPO_FACTURACION = " & cmbTipoFacturacion.BoundText
    End If
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT DISTINCT D.FECHA,D.NUMERO,C.NOMBRE,O.NOMBRE,D.TOTAL,D.PORTES,D.ID_DOCUMENTO,TD.ID_TIPO_DOCUMENTO " & _
               "  FROM DOCUMENTOS D " & _
               " INNER JOIN  DOCUMENTOS_TIPOS TD ON TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               " INNER JOIN OBRAS O ON D.OBRA_ID = O.ID_OBRA " & _
               " INNER JOIN CLIENTES C ON O.CLIENTE_ID = C.ID_CLIENTE " & _
               "  LEFT JOIN DOCUMENTOS_DETALLE DD ON D.ID_DOCUMENTO = DD.DOCUMENTO_ID " & _
               " WHERE 1 = 1 " & _
               tipo & cliente & obra & numero & anno & ESTADO & _
               " ORDER BY D.TIPO_DOCUMENTO_ID, D.NUMERO DESC"
    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim total As Currency
        total = 0
        While Not rs.EOF
            With lista.ListItems.Add(, , Format(rs.Fields(0), "yyyy-mm-dd"))
                .SubItems(1) = rs.Fields(1)
                If Not IsNull(rs.Fields(2)) Then
                 .SubItems(2) = rs.Fields(2) ' Numero de factura
                End If
                .SubItems(3) = rs.Fields(3)
                .SubItems(4) = Format(Replace(rs.Fields(4), ".", ","), "currency")
                .SubItems(5) = Format(Replace(rs.Fields(5), ".", ","), "currency")
                .SubItems(6) = rs.Fields(6)
                .SubItems(7) = rs.Fields(7)
            End With
            rs.MoveNext
        Wend
'        lista.SetFocus
    Else
        MsgBox "No existen albaranes pendientes de facturar.", vbInformation, App.Title
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Documentos : " & Err.Description, vbCritical, Err.Description
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

Private Sub cabecera_lista()
    ' Pendientes
    With lista.ColumnHeaders
        .Add , , "Fecha", 1300, lvwColumnLeft
        .Add , , "Numero", 1400, lvwColumnCenter
        .Add , , "Cliente", 3500, lvwColumnLeft
        .Add , , "Obra", 3500, lvwColumnLeft
        .Add , , "Base", 1600, lvwColumnRight
        .Add , , "Portes", 1600, lvwColumnRight
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "TIPO_ID", 1, lvwColumnCenter
    End With
End Sub
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
            frmDocumento.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
            frmDocumento.Show 1
            actualizar_lista
    End If
End Sub
Private Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
   On Error GoTo actualizar_lista_Error

    consulta = "SELECT D.FECHA,D.NUMERO,C.NOMBRE,O.NOMBRE,D.TOTAL,D.PORTES,D.ID_DOCUMENTO,TD.ID_TIPO_DOCUMENTO " & _
               "  FROM DOCUMENTOS D " & _
               " INNER JOIN  DOCUMENTOS_TIPOS TD ON TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               " INNER JOIN OBRAS O ON D.OBRA_ID = O.ID_OBRA " & _
               " INNER JOIN CLIENTES C ON O.CLIENTE_ID = C.ID_CLIENTE " & _
               "  LEFT JOIN DOCUMENTOS_DETALLE DD ON D.ID_DOCUMENTO = DD.DOCUMENTO_ID " & _
               " WHERE TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.ALBARAN & _
               "   AND FACTURADO = 0 AND DD.TOTAL = 0 " & _
               "   AND D.ID_DOCUMENTO = " & lista.ListItems(lista.SelectedItem.Index).SubItems(6)
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        With lista.ListItems(lista.SelectedItem.Index)
            .Text = Format(rs.Fields(0), "yyyy-mm-dd")
            .SubItems(1) = rs.Fields(1)
            If Not IsNull(rs.Fields(2)) Then
             .SubItems(2) = rs.Fields(2) ' Numero de factura
            End If
            .SubItems(3) = rs.Fields(3)
            .SubItems(4) = Format(Replace(rs.Fields(4), ".", ","), "currency")
            .SubItems(5) = Format(Replace(rs.Fields(5), ".", ","), "currency")
            .SubItems(6) = rs.Fields(6)
            .SubItems(7) = rs.Fields(7)
        End With
    Else
        lista.ListItems.Remove lista.SelectedItem.Index
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

actualizar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_lista of Formulario frmAlbaranesNoValorados"
End Sub
Private Sub cargar_combos()
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    llenar_combo cmbObra, New clsObras, 0, frmObras, ""
    Dim oDeco As New clsDecodificadora
    oDeco.Cargar_Combo cmbTipoFacturacion, DECODIFICADORA.D_TIPOS_FACTURACION
End Sub

