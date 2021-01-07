VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmFacturas_Listado_Cobrar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas pendientes de Cobro"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmFacturas_Listado_Cobrar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   13635
   Begin VB.OptionButton opListado 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ordenar por número de factura"
      Height          =   195
      Index           =   1
      Left            =   2385
      TabIndex        =   21
      Top             =   8775
      Width           =   2940
   End
   Begin VB.OptionButton opListado 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ordenar por cliente"
      Height          =   195
      Index           =   0
      Left            =   2385
      TabIndex        =   20
      Top             =   8415
      Value           =   -1  'True
      Width           =   2940
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Listado"
      Height          =   885
      Index           =   0
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8280
      Width           =   2205
   End
   Begin VB.CommandButton cmdCobrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cobrar Facturas Seleccionada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   6660
      Picture         =   "frmFacturas_Listado_Cobrar.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8280
      Width           =   2865
   End
   Begin VB.CommandButton cmdCobrar2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar Factura Seleccionada"
      Height          =   885
      Left            =   9570
      Picture         =   "frmFacturas_Listado_Cobrar.frx":1F14
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Width           =   2835
   End
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
      Height          =   1725
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
      Begin MSDataListLib.DataCombo cmbEstado 
         Height          =   315
         Left            =   1380
         TabIndex        =   7
         Top             =   960
         Width           =   3510
         _ExtentX        =   6191
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
      Begin MSDataListLib.DataCombo cmbAgente 
         Height          =   315
         Left            =   7380
         TabIndex        =   12
         Top             =   960
         Width           =   3315
         _ExtentX        =   5847
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
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1395
         TabIndex        =   16
         Top             =   1305
         Width           =   1350
         _ExtentX        =   2381
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
         CalendarTitleBackColor=   12632256
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3585
         TabIndex        =   17
         Top             =   1305
         Width           =   1305
         _ExtentX        =   2302
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
         CalendarTitleBackColor=   12632256
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   19
         Top             =   1365
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2955
         TabIndex        =   18
         Top             =   1395
         Width           =   420
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Agente"
         Height          =   195
         Left            =   6720
         TabIndex        =   13
         Top             =   1020
         Width           =   510
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   990
         Width           =   495
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
      Left            =   12450
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6090
      Left            =   60
      TabIndex        =   0
      Top             =   2130
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   10742
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
      Caption         =   "Facturas Pendientes de Cobro"
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
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   13545
   End
End
Attribute VB_Name = "frmFacturas_Listado_Cobrar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdImprimir_Click(Index As Integer)
    Dim FILTRO As String
   On Error GoTo cmdImprimir_Click_Error

    FILTRO = " {documentos.ANULADO} = 0 AND {documentos.TIPO_DOCUMENTO_ID}=" & ENUM_TIPOS_DOCUMENTOS.factura
    If cmbCliente.getTEXTO <> "" Then
        FILTRO = FILTRO & " AND {clientes.ID_CLIENTE} = " & cmbCliente.getPK_SALIDA
    End If
    If cmbObra.getTEXTO <> "" Then
        FILTRO = FILTRO & " AND {obras.ID_OBRA} = " & cmbObra.getPK_SALIDA
    End If
    FILTRO = FILTRO & " AND {documentos.FECHA} in Date (" & Year(fdesde) & "," & Month(fdesde) & "," & Day(fdesde) & ") to Date (" & Year(fhasta) & "," & Month(fhasta) & "," & Day(fhasta) & ")"
    
    If cmbEstado.Text <> "" Then
        FILTRO = FILTRO & " AND {documentos.ESTADO_ID} = " & cmbEstado.BoundText
    End If
    If cmbAgente.Text <> "" Then
        FILTRO = FILTRO & "  AND {obras.COMERCIAL_ID} = " & cmbAgente.BoundText
    End If
    
    Me.MousePointer = 11
    Dim p1() As String
    Dim p2() As String
    ReDim p1(4) As String
    ReDim p2(4) As String
    p1(1) = "FECHA_DESDE"
    p1(2) = "FECHA_HASTA"
    p1(3) = "AGENTE"
    p1(4) = "ESTADO"
    
    p2(1) = fdesde
    p2(2) = fhasta
    If cmbAgente.Text = "" Then
        p2(3) = "Todos"
    Else
        p2(3) = cmbAgente.Text
    End If
    If cmbEstado.Text = "" Then
        p2(4) = "Todos"
    Else
        p2(4) = cmbEstado.Text
    End If
    With frmReport
        .iniciar
        .CRITERIO = FILTRO
        If Index = 0 Then
            .informe = "rptfacturas_listado_agente"
        End If
        .ParametrosNombre = p1
        .ParametrosValores = p2
        .imprimir = False
        If opListado(0).Value = True Then
            .ordenacion = "{clientes.NOMBRE}"
        Else
            .ordenacion = "{documentos.NUMERO}"
        End If
        .generar
        .Show 1
    End With
    Unload frmReport
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cmdImprimir_Click_Error:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmAlbaranes_Listado"

End Sub


Private Sub cmbAgente_Change()
    cargar_lista
End Sub
Private Sub cmbCliente_change()
    cargar_lista
End Sub

Private Sub cmbEstado_Change()
    cargar_lista
End Sub
Private Sub cmbObra_change()
    cargar_lista
End Sub
Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCobrar_Click()
    Dim i As Integer
    Dim algo As Boolean

    If lista.ListItems.Count > 0 Then
'        For i = 1 To lista.ListItems.Count
'            If lista.ListItems(i).Checked = True Then
'                algo = True
'            End If
'        Next
'        If Not algo Then
'            MsgBox "Marque las facturas que desea cobrar.", vbExclamation, App.Title
'            Exit Sub
'        End If
        frmDocumento_Cobro.pk = lista.ListItems(lista.SelectedItem.Index).Text
        frmDocumento_Cobro.Show 1
        actualizar_lista
        pasar_siguiente
    Else
        MsgBox "No existen facturas para cobrar.", vbExclamation, App.Title
        Exit Sub
    End If
    
End Sub
Private Sub pasar_siguiente()
        If lista.ListItems.Count > lista.SelectedItem.Index Then
            Set lista.SelectedItem = lista.ListItems(lista.SelectedItem.Index + 1)
        End If
End Sub

Private Sub cmdCobrar2_Click()
    If lista.ListItems.Count > 0 Then
        frmDocumento.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).Text
        frmDocumento.Show 1
        actualizar_lista
    End If

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
    fdesde = Date - 31
    fhasta = Date
    cabecera_lista
    cargar_combos
    cmbEstado.BoundText = 1
    cargar_lista
End Sub
Public Sub cargar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim cliente As String
    Dim OBRA As String
    Dim numero As String
    Dim anno As String
    Dim ESTADO As String
    Dim agente As String
    tipo = " AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.factura
    
    If cmbCliente.getTEXTO <> "" Then
        cliente = " AND O.CLIENTE_ID = " & cmbCliente.getPK_SALIDA
    End If
    If cmbObra.getTEXTO <> "" Then
        OBRA = " AND D.OBRA_ID = " & cmbObra.getPK_SALIDA
    End If
    If cmbEstado.Text <> "" Then
        ESTADO = " AND D.ESTADO_ID = " & cmbEstado.BoundText
    End If
    If cmbAgente.Text <> "" Then
        agente = " AND O.COMERCIAL_ID = " & cmbAgente.BoundText
    End If
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT D.ID_DOCUMENTO,D.NUMERO,D.FECHA,O.ID_OBRA,O.NOMBRE,D.TOTAL ,DC.FECHA,COMER.NOMBRE,D.IVA " & _
               "  FROM DOCUMENTOS D " & _
               "  LEFT JOIN DOCUMENTOS_TIPOS TD ON TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "  LEFT JOIN OBRAS O ON D.OBRA_ID = O.ID_OBRA " & _
               "  LEFT JOIN CLIENTES C ON C.ID_CLIENTE = O.CLIENTE_ID " & _
               "  LEFT JOIN COMERCIALES COMER ON O.COMERCIAL_ID = COMER.ID_COMERCIAL " & _
               "  LEFT JOIN DOCUMENTOS_COBROS DC ON D.ID_DOCUMENTO = DC.DOCUMENTO_ID " & _
               " WHERE 1 = 1 " & _
               "   AND D.FECHA >= '" & Format(fdesde, "YYYY-MM-DD") & "'" & _
               "   AND D.FECHA <= '" & Format(fhasta, "YYYY-MM-DD") & "'" & _
               tipo & cliente & OBRA & numero & anno & ESTADO & agente & _
               " ORDER BY D.NUMERO ASC, DC.VENCIMIENTO DESC "
    lista.ListItems.Clear
    Me.MousePointer = 11
    Dim ID As Long
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim total As Currency
        total = 0
        While Not rs.EOF
        
            If ID <> rs(0) Then ' Para evitar los duplicados en los vencimientos de varios recibos (DOCUMENTOS_COBROS)
                With lista.ListItems.Add(, , rs(0))
                    .SubItems(1) = Format(rs.Fields(1), "0000")
                    .SubItems(2) = Format(rs.Fields(2), "dd-mm-yyyy")
                    .SubItems(3) = rs.Fields(3)
                    .SubItems(4) = rs(4)
                    .SubItems(5) = moneda((rs(5) + ((rs(5) * rs(8)) / 100)))
                    If Not IsNull(rs(6)) Then
                        .SubItems(6) = Format(rs.Fields(6), "dd-mm-yyyy")
                    End If
                    If Not IsNull(rs(7)) Then
                        .SubItems(7) = rs(7)
                    End If
                End With
                ID = rs(0)
            End If
            rs.MoveNext
        Wend
'        lista.SetFocus
'    Else
'        MsgBox "No existen facturas con esos criterios.", vbInformation, App.Title
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
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "NºFactura", 800, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Cod.Obra", 1500, lvwColumnCenter
        .Add , , "Obra", 4300, lvwColumnLeft
        .Add , , "Importe", 1200, lvwColumnRight
        .Add , , "F.Cobro", 1200, lvwColumnCenter
        .Add , , "Agente", 2500, lvwColumnCenter
    End With
End Sub
Private Function grupo(L As ListView) As String
    Dim s As String
    Dim i As Integer
    For i = 1 To L.ListItems.Count
        If L.ListItems(i).Checked = True Then
            s = s & L.ListItems(i).SubItems(6) & ","
        End If
    Next
    If Len(s) > 0 Then
        s = Left(s, Len(s) - 1)
    End If
    grupo = s
End Function
Private Sub lista_DblClick()
    cmdCobrar_Click
End Sub

Private Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
   On Error GoTo actualizar_lista_Error

    consulta = "SELECT D.ID_DOCUMENTO,D.NUMERO,D.FECHA,O.ID_OBRA,O.NOMBRE,D.TOTAL,DC.FECHA,COMER.NOMBRE, D.IVA " & _
               "  FROM DOCUMENTOS D " & _
               "  LEFT JOIN DOCUMENTOS_TIPOS TD ON TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "  LEFT JOIN OBRAS O ON D.OBRA_ID = O.ID_OBRA " & _
               "  LEFT JOIN CLIENTES C ON C.ID_CLIENTE = O.CLIENTE_ID " & _
               "  LEFT JOIN COMERCIALES COMER ON O.COMERCIAL_ID = COMER.ID_COMERCIAL " & _
               "  LEFT JOIN DOCUMENTOS_COBROS DC ON D.ID_DOCUMENTO = DC.DOCUMENTO_ID " & _
               " WHERE 1 = 1 " & _
               "   AND D.ID_DOCUMENTO = " & lista.ListItems(lista.SelectedItem.Index).Text & _
               " ORDER BY D.NUMERO ASC, DC.VENCIMIENTO DESC"

    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
            With lista.ListItems(lista.SelectedItem.Index)
                .SubItems(1) = Format(rs.Fields(1), "0000")
                .SubItems(2) = Format(rs.Fields(2), "dd-mm-yyyy")
                .SubItems(3) = rs.Fields(3)
                .SubItems(4) = rs(4)
                .SubItems(5) = moneda((rs(5) + ((rs(5) * rs(8)) / 100)))
                If Not IsNull(rs(6)) Then
                    .SubItems(6) = Format(rs.Fields(6), "dd-mm-yyyy")
                End If
                If Not IsNull(rs(7)) Then
                    .SubItems(7) = rs(7)
                End If
            End With
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

actualizar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_lista of Formulario frmFacturas_Listado_Cobrar"
End Sub
Private Sub cargar_combos()
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    llenar_combo cmbObra, New clsObras, 0, frmObras, ""
    Dim oDeco As New clsDecodificadora
    oDeco.Cargar_Combo cmbEstado, DECODIFICADORA.D_DOCUMENTOS_ESTADOS
    Cargar_Combo cmbAgente, New clsComercial
End Sub

