VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmBuscarDocumento 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localizar Documento"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13215
   Icon            =   "frmBuscarDocumento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9075
   ScaleWidth      =   13215
   Begin VB.CommandButton cmdCobrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cobro"
      Height          =   885
      Left            =   6090
      Picture         =   "frmBuscarDocumento.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   8100
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   885
      Left            =   3690
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8100
      Width           =   1155
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   10830
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8100
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8100
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   12045
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8100
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1290
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8100
      Width           =   1155
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   885
      Left            =   2490
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8100
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criterios de búsqueda de Documento"
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
      Height          =   1830
      Left            =   45
      TabIndex        =   15
      Top             =   360
      Width           =   13155
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
         Height          =   330
         Index           =   1
         Left            =   8670
         TabIndex        =   2
         Top             =   345
         Width           =   1785
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbObra 
         Height          =   345
         Left            =   930
         TabIndex        =   4
         Top             =   1035
         Width           =   10590
         _ExtentX        =   18680
         _ExtentY        =   609
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbCliente 
         Height          =   345
         Left            =   930
         TabIndex        =   3
         Top             =   690
         Width           =   10590
         _ExtentX        =   18680
         _ExtentY        =   609
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   1050
         Left            =   11700
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   1185
      End
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6570
         TabIndex        =   1
         Top             =   345
         Width           =   885
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
         Height          =   330
         Index           =   0
         Left            =   4680
         TabIndex        =   0
         Top             =   345
         Width           =   1245
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   930
         TabIndex        =   6
         Top             =   345
         Width           =   2715
         _ExtentX        =   4789
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
      Begin MSComCtl2.UpDown cambiar 
         Height          =   330
         Left            =   7455
         TabIndex        =   20
         Top             =   345
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196619
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   945
         TabIndex        =   26
         Top             =   1395
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
         Format          =   51576833
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3135
         TabIndex        =   27
         Top             =   1395
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
         Format          =   51576833
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   7
         Left            =   2505
         TabIndex        =   29
         Top             =   1485
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   28
         Top             =   1455
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cod.Cliente"
         Height          =   195
         Index           =   4
         Left            =   7800
         TabIndex        =   25
         Top             =   405
         Width           =   810
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obra"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   23
         Top             =   1065
         Width           =   345
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   19
         Top             =   405
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   195
         Index           =   6
         Left            =   6165
         TabIndex        =   18
         Top             =   405
         Width           =   285
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         Height          =   195
         Index           =   0
         Left            =   3990
         TabIndex        =   17
         Top             =   405
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   16
         Top             =   720
         Width           =   480
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5820
      Left            =   60
      TabIndex        =   14
      Top             =   2190
      Width           =   13125
      _ExtentX        =   23151
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
   Begin VB.CommandButton cmdAlbaranes 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Albaranes"
      Height          =   885
      Left            =   4890
      Picture         =   "frmBuscarDocumento.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8100
      Width           =   1155
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Búsqueda de Documentos"
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
      TabIndex        =   30
      Top             =   0
      Width           =   13545
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "0,00"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   330
      Left            =   8370
      TabIndex        =   22
      Top             =   8490
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Index           =   2
      Left            =   8370
      TabIndex        =   21
      Top             =   8190
      Visible         =   0   'False
      Width           =   2265
   End
End
Attribute VB_Name = "frmBuscarDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TIPO_DOCUMENTO_ID As Long
Private Sub cmbCliente_change()
    cmdBuscar_Click
End Sub

Private Sub cmbObra_change()
    cmdBuscar_Click
End Sub

Private Sub cmbTipo_Change()
    cmdBuscar_Click
End Sub

Private Sub cmdAlbaranes_Click()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index).SubItems(7) = ENUM_TIPOS_DOCUMENTOS.factura Then
           frmListadoAlbaranesFactura.pk = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
           frmListadoAlbaranesFactura.Show 1
        Else
            MsgBox "Debe seleccionar una factura para ver sus albaranes.", vbExclamation, App.Title
        End If
    End If
End Sub

Private Sub cmdAnadir_Click()
    frmDocumento.PK_CLIENTE = 0
    frmDocumento.PK_DOCUMENTO = 0
    frmDocumento.Show 1
End Sub

Private Sub cmdBuscar_Click()
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim cliente As String
    Dim OBRA As String
    Dim numero As String
    Dim anno As String
    Dim CODCLIENTE As String
    Dim fecha As String
    If cmbTipo.Text <> "" Then
        tipo = " AND TIPO_DOCUMENTO_ID = " & cmbTipo.BoundText
    End If
    If cmbCliente.getTEXTO <> "" Then
        cliente = " AND O.CLIENTE_ID = " & cmbCliente.getPK_SALIDA
    End If
    If cmbObra.getTEXTO <> "" Then
        OBRA = " AND D.OBRA_ID = " & cmbObra.getPK_SALIDA
    End If
    ' LP001
    If txtDatos(0).Text <> "" Then
        numero = " AND numero = " & txtDatos(0)
    End If
    If txtanno.Text <> "" Then
        anno = " AND anno =" & txtanno
    End If
    If txtDatos(1).Text <> "" Then
        If IsNumeric(txtDatos(1)) Then
            CODCLIENTE = " AND C.ID_CLIENTE LIKE '%" & txtDatos(1) & "%'"
        End If
    End If
    If txtDatos(0) = "" Then
        fecha = " AND D.FECHA BETWEEN '" & Format(fdesde, "YYYY-MM-DD") & "' AND '" & Format(fhasta, "YYYY-MM-DD") & "'"
    End If
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT D.FECHA,TD.NOMBRE,D.NUMERO,C.NOMBRE,O.NOMBRE,D.TOTAL," & _
               "       D.ID_DOCUMENTO,TD.ID_TIPO_DOCUMENTO,D.ANULADO,DECO.DESCRIPCION,D.IVA,D.FACTURADO,D.DESCUENTO " & _
               "  FROM DOCUMENTOS D, DOCUMENTOS_TIPOS TD, OBRAS O, CLIENTES C, DECODIFICADORA DECO " & _
               " WHERE D.OBRA_ID = O.ID_OBRA " & _
               "   AND O.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "   AND DECO.CODIGO = " & DECODIFICADORA.D_DOCUMENTOS_ESTADOS & _
               "   AND DECO.VALOR = D.ESTADO_ID " & _
               fecha & tipo & CODCLIENTE & cliente & OBRA & numero & anno & ESTADO & _
               " ORDER BY D.TIPO_DOCUMENTO_ID, D.NUMERO DESC"
    lista.ListItems.Clear
    lbltotal = Format("0", "currency")
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
'        Dim total As Currency
'        total = 0
        While Not rs.EOF
            With lista.ListItems.Add(, , Format(rs.Fields(0), "yyyy-mm-dd"))
                .SubItems(1) = rs.Fields(1)
'                If Not IsNull(rs.Fields(2)) Then
                 .SubItems(2) = Format(rs.Fields(2), "0000") ' Numero de factura
'                End If
                .SubItems(3) = rs.Fields(3)
                .SubItems(4) = rs.Fields(4)
                If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                    .SubItems(5) = moneda(rs(5) - rs(12))
                Else
                    .SubItems(5) = moneda((rs(5) - rs(12)) + ((rs(5) - rs(12)) * rs(10) / 100)) ' TOTAL
                End If
                .SubItems(6) = rs.Fields(6)
                .SubItems(7) = rs.Fields(7)
                .SubItems(8) = rs.Fields(8)
                .SubItems(9) = rs.Fields(9)
                If rs(7) = ENUM_TIPOS_DOCUMENTOS.ALBARAN And rs(8) = 0 Then
                    If rs(11) = 0 Then
                        .SubItems(9) = "PENDIENTE"
                    Else
                        .SubItems(9) = "FACTURADO"
                    End If
                End If
                If rs(8) = 1 Then
                    colorear_anulado rs(8), lista.ListItems.Count
                End If
            End With
            rs.MoveNext
        Wend
'        lbltotal = Format(total, "currency")
        On Error Resume Next
        lista_Click
        lista.SetFocus
    Else
        MsgBox "No existen Documentos con esos criterios.", vbInformation, App.Title
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Documentos : " & Err.Description, vbCritical, Err.Description
End Sub

Private Sub cmdcancel_Click()
    gDocumento = 0
    Unload Me
End Sub

Private Sub cmdCobrar_Click()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index).SubItems(7) = ENUM_TIPOS_DOCUMENTOS.factura Then
           frmDocumento_Cobro.pk = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
           frmDocumento_Cobro.Show 1
           actualizar_lista
        Else
           MsgBox "Debe seleccionar una factura para poder Cobrarla.", vbExclamation, App.Title
        End If
    End If
End Sub

Private Sub cmdEliminar_Click()
'    If lista.ListItems.Count > 0 Then
'        If MsgBox("Va a eliminar el documento seleccionado. ¿Esta TOTALMENTE seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'            On Error GoTo fallo
'            Dim oDOCUMENTO As New clsDocumentos
'            If oDOCUMENTO.Eliminar(lista.ListItems(lista.SelectedItem.Index).SubItems(6)) = True Then
'                MsgBox "El documento se ha eliminado correctamente.", vbInformation, App.Title
'                cmdBuscar_Click
'            End If
'        End If
'    End If
    If lista.ListItems.Count > 0 Then
        frmDocumento_Anular.pk = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
        frmDocumento_Anular.Show 1
        actualizar_lista
    End If
    Exit Sub
fallo:
    MsgBox "Error al eliminar el documento : " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdImprimir_Click()
    If lista.ListItems.Count > 0 Then
'        Dim oDOCUMENTO As New clsDocumentos
'        oDOCUMENTO.imprimir lista.ListItems(lista.SelectedItem.Index).SubItems(6), False
        frmimprimir.pk = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
        frmimprimir.Show 1
'        actualizar_lista
'        gDocumento = 0
'        Set oDOCUMENTO = Nothing
    End If
End Sub

Private Sub cmdModificar_Click()
    lista_DblClick
End Sub

Private Sub cmdok_Click()
    If lista.ListItems.Count > 0 Then
            frmDocumento.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
            frmDocumento.Show 1
            actualizar_lista
    Else
        gDocumento = 0
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 27
        cmdcancel_Click
     Case 121 ' F10
        cmdok_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.Top = 50
    txtanno = Year(Date)
    fdesde = Date - 60
    fhasta = Date
    Call cabecera
    Call cargar_combos
    Call permisos
    cmbTipo.BoundText = TIPO_DOCUMENTO_ID
'    cmbTipo.BoundText = TIPOS_DOCUMENTOS.ALBARAN
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index).SubItems(7) = ENUM_TIPOS_DOCUMENTOS.factura Then
            cmdCobrar.Visible = True
            cmdAlbaranes.Visible = True
        Else
            cmdCobrar.Visible = False
            cmdAlbaranes.Visible = False
        End If
    End If
End Sub

Private Sub lista_DblClick()
    cmdok_Click
End Sub

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    lista_Click
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If txtDatos(Index) <> "" Then
        cmdBuscar_Click
    End If
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Fecha", 1100, lvwColumnLeft
        .Add , , "Tipo", 1100, lvwColumnCenter
        .Add , , "Numero", 1000, lvwColumnCenter
        .Add , , "Cliente", 3500, lvwColumnLeft
        .Add , , "Obra", 3500, lvwColumnLeft
        .Add , , "Total", 1400, lvwColumnRight
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "TIPO_ID", 1, lvwColumnCenter
        .Add , , "ANULADA", 1, lvwColumnCenter
        .Add , , "Estado", 1200, lvwColumnCenter
    End With
End Sub

Public Sub cargar_combos()
    Cargar_Combo cmbTipo, New clsDocumentos_tipos
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbObra, New clsObras, 0, frmObras, ""
End Sub
Public Sub permisos()
    If USUARIO.getPER_1 = 0 Then
        cmdImprimir.Enabled = False
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

Public Sub actualizar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT D.FECHA,TD.NOMBRE,D.NUMERO,C.NOMBRE,O.NOMBRE,D.TOTAL, " & _
               "       D.ID_DOCUMENTO,TD.ID_TIPO_DOCUMENTO,D.ANULADO, DECO.DESCRIPCION, D.IVA, D.FACTURADO, D.DESCUENTO " & _
               "  FROM DOCUMENTOS D, DOCUMENTOS_TIPOS TD, OBRAS O, CLIENTES C,DECODIFICADORA DECO " & _
               " WHERE D.OBRA_ID = O.ID_OBRA " & _
               "   AND O.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "   AND DECO.CODIGO = " & DECODIFICADORA.D_DOCUMENTOS_ESTADOS & _
               "   AND DECO.VALOR = D.ESTADO_ID " & _
               "   AND D.ID_DOCUMENTO=" & lista.ListItems(lista.SelectedItem.Index).SubItems(6)
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        With lista.ListItems(lista.SelectedItem.Index)
            .Text = Format(rs(0), "yyyy-mm-dd")
            .SubItems(1) = rs.Fields(1)
            .SubItems(2) = Format(rs.Fields(2), "0000")
            .SubItems(3) = rs.Fields(3)
            .SubItems(4) = rs.Fields(4)
            If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                .SubItems(5) = moneda(rs(5) - rs(12))
            Else
                .SubItems(5) = moneda((rs(5) - rs(12)) + ((rs(5) - rs(12)) * rs(10) / 100)) ' TOTAL
            End If
            .SubItems(6) = rs.Fields(6)
            .SubItems(7) = rs.Fields(7) ' TIPO_DOCUMENTO_ID
            .SubItems(8) = rs.Fields(8) ' ANULADO
            .SubItems(9) = rs(9) ' DESCRIPCION ESTADO
            If rs(7) = ENUM_TIPOS_DOCUMENTOS.ALBARAN And rs(8) = 0 Then
                If rs(11) = 0 Then
                    .SubItems(9) = "PENDIENTE"
                Else
                    .SubItems(9) = "FACTURADO"
                End If
            End If
            colorear_anulado rs(8), lista.SelectedItem.Index
        End With
        lista.SetFocus
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al actualziar el Documento : " & Err.Description, vbCritical, Err.Description
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
Private Sub colorear_anulado(anulado As Integer, fila As Long)
    If anulado = 1 Then
        lista.ListItems(fila).ForeColor = vbRed
        Dim i As Integer
        For i = 1 To lista.ColumnHeaders.Count - 1
            lista.ListItems(fila).ListSubItems(i).ForeColor = vbRed
        Next
    End If
End Sub
