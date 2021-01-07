VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmRemesas_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de la Remesa"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13605
   Icon            =   "frmRemesas_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   13605
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmEditar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Editar datos del Efecto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   2925
      TabIndex        =   12
      Top             =   1620
      Visible         =   0   'False
      Width           =   8385
      Begin VB.CommandButton cmdSalirEdicion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Salir"
         Height          =   885
         Left            =   7005
         Picture         =   "frmRemesas_Detalle.frx":164A
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1170
         Width           =   1155
      End
      Begin VB.CommandButton cmdAceptarEdicion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Height          =   885
         Left            =   5805
         Picture         =   "frmRemesas_Detalle.frx":1F14
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1170
         Width           =   1155
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   1  'Right Justify
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
         Index           =   4
         Left            =   1440
         TabIndex        =   14
         Top             =   675
         Width           =   1380
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
         Height          =   330
         Index           =   5
         Left            =   1440
         TabIndex        =   13
         Top             =   315
         Width           =   6735
      End
      Begin MSComCtl2.DTPicker fechaVencimiento 
         Height          =   330
         Left            =   1440
         TabIndex        =   15
         Top             =   1035
         Width           =   1395
         _ExtentX        =   2461
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
         Format          =   51773441
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   18
         Top             =   720
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Vencimiento"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   17
         Top             =   1125
         Width           =   1050
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   16
         Top             =   360
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdAceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   885
      Left            =   11250
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8280
      Width           =   1155
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
      Height          =   3690
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   6509
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
   Begin MSComctlLib.ListView efectos 
      Height          =   3150
      Left            =   30
      TabIndex        =   5
      Top             =   5100
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   5556
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
   Begin vb6projectpryComboBCA.miComboBCA cmbBanco 
      Height          =   345
      Left            =   660
      TabIndex        =   7
      Top             =   4170
      Width           =   7545
      _ExtentX        =   13309
      _ExtentY        =   609
   End
   Begin MSComCtl2.DTPicker fecha 
      Height          =   330
      Left            =   9450
      TabIndex        =   10
      Top             =   4170
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
      Format          =   51773441
      CurrentDate     =   38002
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
      Height          =   195
      Index           =   0
      Left            =   8820
      TabIndex        =   11
      Top             =   4230
      Width           =   450
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Banco"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   8
      Top             =   4230
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Efectos Pendientes (Doble click para añadir a la Remesa)"
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
      Index           =   0
      Left            =   30
      TabIndex        =   6
      Top             =   4740
      Width           =   13545
   End
   Begin VB.Label lblremesa 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Efectos de la Remesa"
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
      Left            =   30
      TabIndex        =   4
      Top             =   0
      Width           =   13545
   End
   Begin VB.Label lblalbaranes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total"
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
      Height          =   300
      Left            =   11010
      TabIndex        =   3
      Top             =   4080
      Width           =   2565
   End
   Begin VB.Label lbltotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
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
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   11010
      TabIndex        =   2
      Top             =   4350
      Width           =   2550
   End
End
Attribute VB_Name = "frmRemesas_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pk As Long


Private Sub cmdAceptar_Click()

   On Error GoTo cmdAceptar_Click_Error

    If validar Then
        If MsgBox("Va a grabar la remesa.¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            Exit Sub
        End If
        ' Remesa
        Dim remesa As Long
        Dim oRemesa As New clsRemesas
        With oRemesa
            .setBANCO_ID = cmbBanco.getPK_SALIDA
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            If pk = 0 Then
                remesa = .Insertar
            Else
                .Modificar pk
                remesa = pk
            End If
        End With
        ' Documentos
        Dim ORD As New clsRemesas_documentos
        Dim oDR As New clsDocumentos_Recibos
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            With ORD
                .setREMESA_ID = remesa
                .setDOCUMENTO_ID = lista.ListItems(i).SubItems(6)
                .setVENCIMIENTO = lista.ListItems(i).SubItems(7)
                .setCLIENTE_ID = lista.ListItems(i).SubItems(1)
                .setFECHA_VENCIMIENTO = Format(lista.ListItems(i).SubItems(5), "yyyy-mm-dd")
                .setDESCRIPCION = lista.ListItems(i).SubItems(3)
                .setIMPORTE = moneda_bd(lista.ListItems(i).SubItems(4))
                .setSITUACION = Left(lista.ListItems(i).SubItems(8), 1)
                .setCONTA = "S"
                If pk = 0 Then
                    .setID = 0
                    .Insertar
                Else
                    If lista.ListItems(i).Text = "--" Then
                        .setID = 0
                        .Insertar
                    Else
                        .Modificar lista.ListItems(i).Text
                    End If
                End If
                ' Marcar la tabla de recibos el estado
                If Left(lista.ListItems(i).SubItems(8), 1) = "D" Then
                    oDR.ESTADO lista.ListItems(i).SubItems(6), lista.ListItems(i).SubItems(7), ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_DESCUENTO
                Else
                    oDR.ESTADO lista.ListItems(i).SubItems(6), lista.ListItems(i).SubItems(7), ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_REMESA
                End If
            End With
        Next
        MsgBox "La Remesa se ha almacenado correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdAceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptar_Click of Formulario frmRemesas_Detalle"

End Sub
Private Function validar() As Boolean
    validar = True
    If cmbBanco.getTEXTO = "" Then
        MsgBox "El Banco debe estar informado.", vbExclamation, App.Title
        cmbBanco.SetFocus
        validar = False
    End If
End Function

Private Sub cmdAceptarEdicion_Click()
    lista.ListItems(lista.SelectedItem.Index).SubItems(3) = txtDatos(5)
    lista.ListItems(lista.SelectedItem.Index).SubItems(4) = txtDatos(4)
    lista.ListItems(lista.SelectedItem.Index).SubItems(5) = Format(fechaVencimiento, "yyyy-mm-dd")
    frmEditar.Visible = False
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdSalirEdicion_Click()
    frmEditar.Visible = False
End Sub

Private Sub efectos_DblClick()
    If efectos.ListItems.Count > 0 Then
        ' Añadir el efecto
'        Dim oRD As New clsRemesas_documentos
'        Dim apunte As Long
'        oRD.CrearID
'        apunte = oRD.getID
        With lista.ListItems.Add(, , "--")
            .SubItems(1) = efectos.ListItems(efectos.SelectedItem.Index).SubItems(10) ' Nº Cliente
            .SubItems(2) = efectos.ListItems(efectos.SelectedItem.Index).SubItems(2) ' Cliente
            .SubItems(3) = "FACTURA " & efectos.ListItems(efectos.SelectedItem.Index).SubItems(1) & "/" & Year(efectos.ListItems(efectos.SelectedItem.Index).SubItems(5))
            .SubItems(4) = efectos.ListItems(efectos.SelectedItem.Index).SubItems(7) ' Importe
            .SubItems(5) = Format(efectos.ListItems(efectos.SelectedItem.Index).SubItems(6), "yyyy-mm-dd") ' F.Vencimiento
            .SubItems(6) = efectos.ListItems(efectos.SelectedItem.Index).Text ' ID_DOC
            .SubItems(7) = efectos.ListItems(efectos.SelectedItem.Index).SubItems(4) ' VENCIMIENTO
            .SubItems(8) = "REMESA" ' SITUACION
        End With
        ' Eliminar
        efectos.ListItems.Remove efectos.SelectedItem.Index
        calcular_total
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
    cabecera_lista
    cargar_combos
    cargar_efectos 0, 0
    If pk = 0 Then
        Dim oRemesa As New clsRemesas
        oRemesa.CrearID
        lblremesa = "Creación de nueva remesa : " & oRemesa.getID_REMESA & " (Doble click para ELIMINAR - Boton derecho MODIFICAR)"
        fecha = Date
    Else
        cargar_remesa
    End If
End Sub
Private Sub cargar_remesa()
    On Error GoTo fallo
    Dim consulta As String
    ' Remesa
    Dim oRemesa As New clsRemesas
    oRemesa.Carga pk
    lblremesa = "Efectos de la Remesa : " & oRemesa.getID_REMESA & " (Doble click para ELIMINAR - Boton derecho MODIFICAR)"
    cmbBanco.MostrarElemento oRemesa.getBANCO_ID
    fecha = oRemesa.getFECHA
    ' Documentos
    Dim rs As New ADODB.Recordset
    consulta = "SELECT ID,A.CLIENTE_ID, C.NOMBRE,A.DESCRIPCION,A.IMPORTE,A.FECHA_VENCIMIENTO,A.DOCUMENTO_ID,A.VENCIMIENTO,A.SITUACION " & _
               "  FROM REMESAS_DOCUMENTOS A " & _
               "  LEFT JOIN DOCUMENTOS B ON A.DOCUMENTO_ID = B.ID_DOCUMENTO " & _
               "  LEFT JOIN CLIENTES C ON A.CLIENTE_ID = C.ID_CLIENTE " & _
               " WHERE REMESA_ID = " & pk & _
               " ORDER BY ID"
    
    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(0)) ' Apunte
                .SubItems(1) = rs.Fields(1) ' NºCliente
                .SubItems(2) = rs.Fields(2) ' Cliente
                .SubItems(3) = rs.Fields(3) ' Descripcion
                .SubItems(4) = moneda(rs.Fields(4)) ' Importe
                .SubItems(5) = Format(rs.Fields(5), "yyyy-mm-dd") ' Vencimiento
                .SubItems(6) = rs(6) ' ID_DOC
                .SubItems(7) = rs(7) ' VENCIMIENTO
                If rs(8) = "D" Then
                    .SubItems(8) = "DESCUENTO"
                Else
                    .SubItems(8) = "REMESA"
                End If
            End With
            rs.MoveNext
        Wend
    End If
    calcular_total
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Documentos : " & Err.Description, vbCritical, Err.Description
End Sub
Private Sub calcular_total()
    Dim i As Integer
    Dim total As Currency
    total = 0
    For i = 1 To lista.ListItems.Count
        total = total + lista.ListItems(i).SubItems(4)
    Next
    lblalbaranes = "Total (" & lista.ListItems.Count & " efectos)"
    lbltotal = Format(total, "currency")
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
Private Sub efectos_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If efectos.ListItems.Count > 0 Then
     efectos.SortKey = ColumnHeader.Index - 1
     If efectos.SortOrder = 0 Then
        efectos.SortOrder = 1
     Else
        efectos.SortOrder = 0
     End If
     efectos.Sorted = True
   End If
End Sub

Private Sub cabecera_lista()
    ' Listado
    With lista.ColumnHeaders
        .Add , , "Apunte", 900, lvwColumnLeft
        .Add , , "NºCliente", 1000, lvwColumnCenter
        .Add , , "Cliente", 3600, lvwColumnLeft
        .Add , , "Descripción", 3600, lvwColumnLeft
        .Add , , "Importe", 1400, lvwColumnRight
        .Add , , "F.Vencimiento", 1200, lvwColumnCenter
        .Add , , "ID_DOC", 1, lvwColumnCenter
        .Add , , "VENCIMIENTO", 1, lvwColumnCenter
        .Add , , "Estado(D/R)", 1200, lvwColumnCenter
    End With
    ' efectos
    With efectos.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "NºFactura", 800, lvwColumnCenter
        .Add , , "Cliente", 2900, lvwColumnCenter
        .Add , , "Obra", 2900, lvwColumnLeft
        .Add , , "Vencimiento", 800, lvwColumnCenter
        .Add , , "F.Factura", 1100, lvwColumnCenter
        .Add , , "F.Vencimiento", 1100, lvwColumnCenter
        .Add , , "Importe", 1100, lvwColumnRight
        .Add , , "Total Factura", 1100, lvwColumnRight
        .Add , , "Estado", 1400, lvwColumnCenter
        .Add , , "CLIENTE_ID", 1, lvwColumnCenter
    End With
End Sub
Private Sub cargar_efectos(ID As Long, VENCIMIENTO As Integer)
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim cliente As String
    Dim OBRA As String
    Dim numero As String
    Dim anno As String
    Dim ESTADO As String
    Dim agente As String
    
    If ID = 0 Then
        efectos.ListItems.Clear
        tipo = " AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.factura
        ESTADO = " AND DR.COBRADO = " & ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_PENDIENTE
    Else
        ESTADO = " AND DR.DOCUMENTO_ID = " & ID & " AND DR.VENCIMIENTO = " & VENCIMIENTO
    End If
    Dim rs As New ADODB.Recordset
    consulta = "SELECT DISTINCT D.ID_DOCUMENTO,D.NUMERO,C.NOMBRE,O.NOMBRE,DR.VENCIMIENTO, D.FECHA, " & _
               "                DR.FECHA,DR.IMPORTE,D.TOTAL - D.DESCUENTO,DECO.DESCRIPCION,C.ID_CLIENTE,D.IVA " & _
               "  FROM DOCUMENTOS D " & _
               "  LEFT JOIN DOCUMENTOS_TIPOS TD ON TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "  LEFT JOIN OBRAS O ON D.OBRA_ID = O.ID_OBRA " & _
               "  LEFT JOIN CLIENTES C ON O.CLIENTE_ID = C.ID_CLIENTE " & _
               " INNER JOIN DOCUMENTOS_RECIBOS DR ON D.ID_DOCUMENTO = DR.DOCUMENTO_ID " & _
               "  LEFT JOIN DECODIFICADORA DECO ON DECO.VALOR = DR.COBRADO " & _
               " WHERE 1 = 1 " & _
               "   AND DECO.CODIGO  = " & DECODIFICADORA.D_EFECTOS_ESTADOS & _
               "   AND D.ANULADO = 0 " & _
               tipo & ESTADO & _
               " ORDER BY D.NUMERO ASC, DR.VENCIMIENTO ASC "
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
                With efectos.ListItems.Add(, , rs(0))
                    .SubItems(1) = Format(rs.Fields(1), "0000")
                    .SubItems(2) = rs(2) ' CLIENTE
                    .SubItems(3) = rs.Fields(3) ' OBRA
                    .SubItems(4) = rs.Fields(4) ' NUMERO VENCIMIENTO
                    .SubItems(5) = Format(rs(5), "dd-mm-yyyy") ' Fecha factura
                    .SubItems(6) = Format(rs.Fields(6), "dd-mm-yyyy") ' F. Vencimiento
                    .SubItems(7) = moneda(rs(7) + (rs(7) * rs(11) / 100))  ' I. Vencimiento
                    .SubItems(8) = moneda(rs(8) + (rs(8) * rs(11) / 100)) ' Total
                    .SubItems(9) = rs(9) ' Estado efecto (DECO = 8)
                    .SubItems(10) = rs(10) ' ID_CLIENTE
                End With
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los Documentos : " & Err.Description, vbCritical, Err.Description
End Sub
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.Cargar_miCombo cmbBanco, DECODIFICADORA.D_BANCOS
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index).SubItems(6) <> 0 Then
            If lista.ListItems(lista.SelectedItem.Index).Text <> "--" Then
                MsgBox "No se puede eliminar. Ya se encuentra contabilizada la Remesa", vbExclamation, App.Title
                Exit Sub
            End If
            
            cargar_efectos lista.ListItems(lista.SelectedItem.Index).SubItems(6), lista.ListItems(lista.SelectedItem.Index).SubItems(7)
            lista.ListItems.Remove lista.SelectedItem.Index
            calcular_total
        Else
            MsgBox "No se puede eliminar el efecto. No tiene factura asociada.", vbExclamation, App.Title
            
        End If
    End If
End Sub

Private Sub lista_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = 2 Then ' Boton derecho
        If lista.ListItems.Count > 0 Then
            txtDatos(5) = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
            txtDatos(4) = lista.ListItems(lista.SelectedItem.Index).SubItems(4)
            fechaVencimiento = lista.ListItems(lista.SelectedItem.Index).SubItems(5)
            frmEditar.Visible = True
            frmEditar.Top = lista.ListItems(lista.SelectedItem.Index).Top + 600
        End If
    End If
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
   On Error GoTo txtDatos_LostFocus_Error

    If Index = 4 Then
        If txtDatos(Index) = "" Then
            txtDatos(Index) = moneda("0")
        Else
            txtDatos(Index) = moneda(txtDatos(Index))
        End If
    End If

   On Error GoTo 0
   Exit Sub

txtDatos_LostFocus_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure txtDatos_LostFocus of Formulario frmRemesas_Detalle"
End Sub
