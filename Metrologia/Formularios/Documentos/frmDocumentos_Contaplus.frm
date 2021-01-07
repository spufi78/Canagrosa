VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDocumentos_Contaplus 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Captura de pedidos de Contaplus"
   ClientHeight    =   7605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   Icon            =   "frmDocumentos_Contaplus.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7605
   ScaleWidth      =   11685
   Begin VB.CommandButton cmdCapturar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Capturar"
      Height          =   885
      Left            =   60
      Picture         =   "frmDocumentos_Contaplus.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6690
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   10500
      Picture         =   "frmDocumentos_Contaplus.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6690
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6270
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   11580
      _ExtentX        =   20426
      _ExtentY        =   11060
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
      Left            =   7890
      TabIndex        =   5
      Top             =   6780
      Width           =   2265
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
      Left            =   7890
      TabIndex        =   4
      Top             =   7080
      Width           =   2250
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Captura de pedidos de Contaplus"
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
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   11595
   End
End
Attribute VB_Name = "frmDocumentos_Contaplus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCapturar_Click()
   On Error GoTo cmdCapturar_Click_Error

    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a capturar los pedidos seleccionados. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        log (String(100, "*"))
        log ("CAPTURANDO DATOS DE CONTAPLUS")
        Dim conn2 As Connection
        Set conn2 = New ADODB.Connection
        Dim ruta As String
        ruta = ReadINI(App.Path + "\config.ini", "Documentos", "Contaplus")
        conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ruta & ";"
        conn2.Open
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Dim rs2 As New ADODB.Recordset
        Dim i As Integer
        Dim oPedido_contaplus As New clsPEDIDOS_CONTAPLUS
        Dim oArticulo As New clsArticulo
        Dim ocliente As New clsCliente
        Dim oDOCUMENTO As New clsDocumentos
        Dim oDocumento_Detalle As New clsDocumentos_detalle
        Dim documento As Long
        Dim cliente As Long
        rs2.ActiveConnection = conn2
        rs2.CursorLocation = adUseClient
        rs2.CursorType = adOpenForwardOnly
        rs2.LockType = adLockReadOnly
        Dim articulo As Long
        Dim j As Integer
        For i = 1 To lista.ListItems.Count
                log ("Procesando pedido : " & lista.ListItems(i))
                ' Insertamos el documento
                With oDOCUMENTO
                    .setNUMERO = lista.ListItems(i)
                    .setTIPO_DOCUMENTO_ID = 1 ' Pedido
                    .setANNO = Format(lista.ListItems(i).SubItems(1), "yyyy")
                    .setFECHA = Format(lista.ListItems(i).SubItems(1), "yyyy-mm-dd")
                    .setFACTURADO = 0
                    .setTIPO_ENTRADA_ID = 5 ' cONTAPLUS
                    If lista.ListItems(i).SubItems(6) = "S" Then ' Pedido ya albaranado, se descuentas los articulos del stock sin pasar por OM
                        .setESTADO_ID = 4
                    Else
                        .setESTADO_ID = 2 ' Confirmado
                    End If
                    .setDTO1 = lista.ListItems(i).SubItems(7)
                    .setDTO2 = lista.ListItems(i).SubItems(8)
                    .setCOMISION = 0
                    ' Buscamos el cliente o lo damos de alta
                    consulta = "select id_cliente from clientes where nombre like '%" & lista.ListItems(i).SubItems(2) & "%'"
                    Set rs = datos_bd(consulta)
                    If rs.RecordCount = 0 Then
                        ocliente.setID_CLIENTE = 0
                        ocliente.setNOMBRE = lista.ListItems(i).SubItems(2)
                        cliente = ocliente.insertar_cliente
                        log ("No existe cliente, se crea : " & cliente & "/" & lista.ListItems(i).SubItems(2))
                        .setCLIENTE_ID = cliente
                    Else
                        log ("Existe cliente : " & rs(0))
                        .setCLIENTE_ID = rs(0)
                    End If
                    .setTOTAL = lista.ListItems(i).SubItems(5)
                    documento = .Insertar
                    log ("Insertamos documento : " & documento)
                End With
                consulta = "select cref,cdetalle,ncanped,npreunit from pedclil " & _
                           " where nnumped = " & lista.ListItems(i).Text
                ' Lineas del documento
                rs2.Open consulta
                j = 1
                If rs2.RecordCount > 0 Then
                    Do
                        With oDocumento_Detalle
                            .setDOCUMENTO_ID = documento
                            .setORDEN = j
                            j = j + 1
                            ' Buscamos el articulo
                            consulta = "select * from articulos where ean = '" & rs2(0) & "'"
                            Set rs = datos_bd(consulta)
                            If rs.RecordCount <> 0 Then
                                log ("Articulo existente : " & rs("ID_ARTICULO"))
                                If lista.ListItems(i).SubItems(6) = "S" Then ' Pedido ya albaranado, se descuentas los articulos del stock sin pasar por OM
                                    oArticulo.descontar_unidades rs("ID_ARTICULO"), rs2(2)
                                    log ("Pedido en estado S, descontamos unidades")
                                Else
                                    .setARTICULO_ID = rs("ID_ARTICULO")
                                    .setEAN = rs("EAN")
                                    .setDESCRIPCION = rs("DESCRIPCION")
                                    .setCANTIDAD = rs2(2)
                                    .setPRECIO = rs2(3)
                                    .setTOTAL = rs2(2) * rs2(3)
                                    .Insertar
                                    log ("Insertamos linea de pedido.")
                                End If
                            Else
                                log ("Articulo NO existe EAN : " & rs2(0))
                                oArticulo.setID_ARTICULO = 0
                                If IsNull(rs2(0)) Then
                                    oArticulo.setEAN = ""
                                    .setEAN = ""
                                Else
                                    oArticulo.setEAN = rs2(0)
                                    .setEAN = rs2(0)
                                End If
                                If IsNull(rs2(1)) Then
                                    oArticulo.setDESCRIPCION = ""
                                    .setDESCRIPCION = ""
                                Else
                                    oArticulo.setDESCRIPCION = rs2(1)
                                    .setDESCRIPCION = rs2(1)
                                End If
                                If IsNull(rs2(3)) Then
                                    oArticulo.setPRECIO_VENTA = "0"
                                Else
                                    oArticulo.setPRECIO_VENTA = rs2(3)
                                End If
                                oArticulo.setFAMILIA_ID = 0
                                oArticulo.setSUBFAMILIA_ID = 0
                                oArticulo.setSTOCK = 0
                                oArticulo.setVOLUMEN = "0"
                                oArticulo.setTIPO_ARTICULO_ID = 6
                                articulo = oArticulo.Insertar
                                log ("Inserta el nuevo articulo : " & articulo)
                                .setARTICULO_ID = articulo
                                .setCANTIDAD = rs2(2)
                                .setPRECIO = rs2(3)
                                .setTOTAL = rs2(2) * rs2(3)
                                .Insertar
                                log ("Inserta la lina con el nuevo articulo : " & articulo)
'                                MsgBox "El pedido " & lista.ListItems(i).Text & " contiene el artículo " & rs2(0) & " inexistente.", vbInformation, App.Title
                            End If
                        End With
                        rs2.MoveNext
                    Loop Until rs2.EOF
                End If
                rs2.Close
            ' Insertamos el documento en la tabla de pedidos ya procesados
            With oPedido_contaplus
                .setNNUMPED = lista.ListItems(i).Text
                .setDFECPED = lista.ListItems(i).SubItems(1)
                .setCNOMCLI = lista.ListItems(i).SubItems(2)
                .setCCOMENT = lista.ListItems(i).SubItems(3)
                .setDOCUMENTO_ID = documento
                .Insertar
                log ("Insertamos el pedido en la tabla de CONTAPLUS : " & documento)
            End With
        Next
    End If
    Set conn2 = Nothing
    log ("PROCESO DE CAPTURA COMPLETADO")
    log (String(100, "*"))
    MsgBox "Pedidos capturados correctamente.", vbInformation, App.Title
    cargar_lista

   On Error GoTo 0
   Exit Sub

cmdCapturar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCapturar_Click of Formulario frmDocumentos_Contaplus"
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
    cabecera
    cargar_lista
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
Public Sub cargar_lista()
    Me.MousePointer = 11
    On Error GoTo fallo
    DoEvents
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    ' Conexion a la bd a copiar
    Dim oPedidos_Contaplus As New clsPEDIDOS_CONTAPLUS
    Dim Ultimo_pedido As Long
    Ultimo_pedido = oPedidos_Contaplus.Ultimo_pedido()
    Dim conn2 As Connection
    Set conn2 = New ADODB.Connection
    Dim ruta As String
    ruta = ReadINI(App.Path + "\config.ini", "Documentos", "Contaplus")
    conn2.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ruta & ";"
    conn2.Open
    Dim total As Currency
    total = 0
    Dim ESTADO As String
    If UCase(ReadINI(App.Path + "\config.ini", "Parametros", "Estado")) = "P" Then
        ESTADO = "   and b.cestado = ""P"""
    End If
    ' Consulta a la tabla de pedidos
    consulta = "select a.nnumped,b.dfecped,d.cnomcli,a.ndto,b.ndtoesp,b.cestado,sum(a.ncanped*a.npreunit) " & _
               " from pedclil a, pedclit b, clientes d " & _
               " where a.nnumped = b.nnumped " & _
               "   and b.ccodcli = d.ccodcli " & _
               ESTADO & _
               "   and b.nnumped > " & Ultimo_pedido & _
               " group by a.nnumped,b.dfecped,d.cnomcli,a.ndto,b.ndtoesp,b.cestado " & _
               " order by a.nnumped asc"
    log (consulta)
    rs.ActiveConnection = conn2
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
    rs.Open consulta
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "00000"))
            .SubItems(1) = Format(rs(1), "dd-mm-yyyy")
            .SubItems(2) = rs(2)
            .SubItems(3) = Format(rs(6), "currency")
            If rs(3) > 0 Then
                .SubItems(4) = Format(.SubItems(3) - (.SubItems(3) * rs(3) / 100), "currency")
            Else
                .SubItems(4) = .SubItems(3)
            End If
            .SubItems(5) = Format(.SubItems(4) - (.SubItems(4) * rs(4) / 100), "currency")
            .SubItems(6) = rs(5)
            .SubItems(7) = rs(3)
            .SubItems(8) = rs(4)
            total = total + .SubItems(5)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    lbltotal = Format(total, "currency")
    rs.Close
    Set conn2 = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    log ("Error al cargar la lista: " & Err.Description)
    MsgBox "Error al cargar la lista: " & Err.Description, vbCritical, App.Title
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Numero", 1300, lvwColumnLeft
        .Add , , "Fecha", 1300, lvwColumnCenter
        .Add , , "Cliente", 4000, lvwColumnLeft
        .Add , , "Importe", 1300, lvwColumnRight
        .Add , , "Dto.", 1300, lvwColumnRight
        .Add , , "P.Final", 1300, lvwColumnRight
        .Add , , "Estado", 800, lvwColumnCenter
        .Add , , "Dto.1", 1, lvwColumnRight
        .Add , , "Dto.2", 1, lvwColumnRight
    End With
End Sub
