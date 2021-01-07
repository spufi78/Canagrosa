VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmDescuentos_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del Descuento"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13605
   Icon            =   "frmDescuentos_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   13605
   StartUpPosition =   2  'CenterScreen
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
      Format          =   51183617
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
      Caption         =   "Listado de Efectos en Remesa (Doble click para añadir al Descuento)"
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
      Caption         =   "Efectos del Descuento"
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
Attribute VB_Name = "frmDescuentos_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pk As Long


Private Sub cmbBanco_change()
    cargar_efectos 0, 0
End Sub

Private Sub cmdAceptar_Click()

   On Error GoTo cmdAceptar_Click_Error

    If validar Then
        If MsgBox("Va a grabar el descuento.¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            Exit Sub
        End If
        ' Remesa
        Dim descuento As Long
        Dim oDTO As New clsDescuentos
        With oDTO
            .setBANCO_ID = cmbBanco.getPK_SALIDA
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setUSUARIO_ID = usuario.getID_EMPLEADO
            If pk = 0 Then
                descuento = .Insertar
            Else
                .Modificar pk
                descuento = pk
            End If
        End With
        ' Documentos
        Dim oDD As New clsDescuentos_documentos
        Dim oDR As New clsDocumentos_Recibos
        Dim i As Integer
        oDD.Eliminar descuento
        For i = 1 To lista.ListItems.Count
            With oDD
                .setDESCUENTO_ID = descuento
                .setAPUNTE_ID = lista.ListItems(i).Text
                .Insertar
                ' Marcar la tabla de recibos el estado
                oDR.ESTADO lista.ListItems(i).SubItems(6), lista.ListItems(i).SubItems(7), ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_DESCUENTO
                ' Marcar la tabla de remesas_documentos como apunte en DESCUENTO
                execute_bd " UPDATE REMESAS_DOCUMENTOS SET SITUACION = 'D' WHERE ID = " & lista.ListItems(i).Text
            End With
        Next
        MsgBox "El descuento se ha almacenado correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdAceptar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAceptar_Click of Formulario frmDescuentos_Detalle"

End Sub
Private Function validar() As Boolean
    validar = True
    If cmbBanco.getTEXTO = "" Then
        MsgBox "El Banco debe estar informado.", vbExclamation, App.Title
        cmbBanco.SetFocus
        validar = False
    End If
End Function
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub efectos_DblClick()
    If efectos.ListItems.Count > 0 Then
        With lista.ListItems.Add(, , efectos.ListItems(efectos.SelectedItem.Index).Text)
            .SubItems(1) = efectos.ListItems(efectos.SelectedItem.Index).SubItems(1)
            .SubItems(2) = efectos.ListItems(efectos.SelectedItem.Index).SubItems(2)
            .SubItems(3) = efectos.ListItems(efectos.SelectedItem.Index).SubItems(3)
            .SubItems(4) = efectos.ListItems(efectos.SelectedItem.Index).SubItems(4)
            .SubItems(5) = efectos.ListItems(efectos.SelectedItem.Index).SubItems(5)
            .SubItems(6) = efectos.ListItems(efectos.SelectedItem.Index).SubItems(6)
            .SubItems(7) = efectos.ListItems(efectos.SelectedItem.Index).SubItems(7)
            .SubItems(8) = "DESCUENTO"
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
        Dim oDTO As New clsDescuentos
        oDTO.CrearID
        lblremesa = "Creación de nuevo descuento : " & oDTO.getID_DESCUENTO
        fecha = Date
    Else
        cargar_remesa
    End If
End Sub
Private Sub cargar_remesa()
    On Error GoTo fallo
    Dim consulta As String
    ' Remesa
    Dim oDTO As New clsDescuentos
    oDTO.Carga pk
    lblremesa = "Efectos del Descuento : " & oDTO.getID_DESCUENTO & " (Doble click para eliminar el efecto)"
    cmbBanco.MostrarElemento oDTO.getBANCO_ID
    fecha = oDTO.getFECHA
    ' Documentos
    Dim rs As New ADODB.Recordset
    
    consulta = "SELECT B.APUNTE_ID,C.CLIENTE_ID,D.NOMBRE,C.DESCRIPCION,C.IMPORTE,C.FECHA_VENCIMIENTO,C.DOCUMENTO_ID,C.VENCIMIENTO,C.SITUACION " & _
                " FROM DESCUENTOS A " & _
                " LEFT JOIN DESCUENTOS_DOCUMENTOS B ON A.ID_DESCUENTO = B.DESCUENTO_ID " & _
                " LEFT JOIN REMESAS_DOCUMENTOS C ON C.ID = B.APUNTE_ID " & _
                " LEFT JOIN CLIENTES D ON C.CLIENTE_ID = D.ID_CLIENTE " & _
                " Where A.ID_DESCUENTO = " & pk & _
                " ORDER BY A.ID_DESCUENTO DESC"
    
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
End Sub
Private Sub cargar_efectos(ID As Long, VENCIMIENTO As Integer)
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim cliente As String
    Dim obra As String
    Dim numero As String
    Dim anno As String
    Dim ESTADO As String
    Dim agente As String
    
    If ID = 0 Then
        efectos.ListItems.Clear
        ESTADO = " AND C.SITUACION <> 'D' "
    Else
        ESTADO = " AND C.DOCUMENTO_ID = " & ID & " AND C.VENCIMIENTO = " & VENCIMIENTO
    End If
    Dim rs As New ADODB.Recordset
    consulta = " SELECT C.ID,C.CLIENTE_ID,D.NOMBRE,C.DESCRIPCION,C.IMPORTE,C.FECHA_VENCIMIENTO,C.DOCUMENTO_ID,C.VENCIMIENTO,C.SITUACION " & _
                "  FROM REMESAS_DOCUMENTOS C " & _
                " LEFT JOIN CLIENTES D ON C.CLIENTE_ID = D.ID_CLIENTE " & _
                " LEFT JOIN REMESAS R ON C.REMESA_ID = R.ID_REMESA " & _
                " WHERE 1 = 1 " & _
                "   AND R.BANCO_ID = " & cmbBanco.getPK_SALIDA & _
                ESTADO & _
                " ORDER BY C.ID ASC "
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
            With efectos.ListItems.Add(, , rs(0)) ' Apunte
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
'            If lista.ListItems(lista.SelectedItem.Index).Text <> "--" Then
'                MsgBox "No se puede eliminar. Ya se encuentra contabilizada la Remesa", vbExclamation, App.Title
'                Exit Sub
'            End If
            
            cargar_efectos lista.ListItems(lista.SelectedItem.Index).SubItems(6), lista.ListItems(lista.SelectedItem.Index).SubItems(7)
            lista.ListItems.Remove lista.SelectedItem.Index
            calcular_total
        Else
            MsgBox "No se puede eliminar el efecto. No tiene factura asociada.", vbExclamation, App.Title
            
        End If
    End If
End Sub
