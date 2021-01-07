VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmPP_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listado de Pedidos a Proveedor"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13440
   Icon            =   "frmPP_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exportar"
      Height          =   960
      Left            =   10125
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Exportar datos a impresora o excel"
      Top             =   8415
      Width           =   1140
   End
   Begin VB.CommandButton cmdEnviar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Enviar"
      Height          =   960
      Left            =   5085
      Picture         =   "frmPP_Listado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   28
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8415
      Width           =   1215
   End
   Begin VB.CommandButton cmdRecepcionar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recepcionar"
      Height          =   960
      Left            =   6345
      Picture         =   "frmPP_Listado.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   27
      ToolTipText     =   "Recepcionar pedido"
      Top             =   8415
      Width           =   1215
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   960
      Left            =   8865
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8415
      Width           =   1215
   End
   Begin VB.CommandButton cmdMail 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Correo al proveedor"
      Height          =   960
      Left            =   7605
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "crear etiqueta para envío de paquete"
      Top             =   8415
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   1680
      Left            =   0
      TabIndex        =   13
      Top             =   630
      Width           =   13425
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   855
         TabIndex        =   2
         Top             =   540
         Width           =   1455
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   855
         TabIndex        =   0
         Top             =   180
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker datFechaDesde 
         Height          =   315
         Left            =   855
         TabIndex        =   3
         Top             =   1260
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
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
         CalendarTitleBackColor=   14737632
         Format          =   61276161
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker datFechaHasta 
         Height          =   315
         Left            =   2385
         TabIndex        =   4
         Top             =   1260
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
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
         CalendarTitleBackColor=   14737632
         Format          =   61276161
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbSubcontratas 
         Height          =   330
         Left            =   4725
         TabIndex        =   1
         Top             =   180
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbFamilia 
         Height          =   330
         Left            =   4725
         TabIndex        =   21
         Top             =   900
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbDescripcion 
         Height          =   330
         Left            =   4725
         TabIndex        =   23
         Top             =   540
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmPP_Listado.frx":205E
         Height          =   315
         Left            =   855
         TabIndex        =   25
         Top             =   900
         Width           =   2820
         _ExtentX        =   4974
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
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   90
         TabIndex        =   26
         Top             =   945
         Width           =   465
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Left            =   3780
         TabIndex        =   24
         Top             =   585
         Width           =   915
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   16
         Left            =   3780
         TabIndex        =   22
         Top             =   945
         Width           =   480
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         Height          =   240
         Left            =   2250
         TabIndex        =   18
         Top             =   1305
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   240
         Left            =   90
         TabIndex        =   17
         Top             =   1305
         Width           =   600
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   240
         Index           =   4
         Left            =   90
         TabIndex        =   16
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Pedido"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   15
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         Height          =   240
         Left            =   3780
         TabIndex        =   14
         Top             =   225
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdCrearDocumentos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Pedido"
      Height          =   960
      Left            =   3825
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Crear documento de solicitud de análisis"
      Top             =   8415
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   960
      Left            =   1305
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Modificar paquete seleccionado"
      Top             =   8415
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC - Salir"
      Height          =   960
      Left            =   12195
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   8415
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo"
      Height          =   960
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Crear nuevo paquete"
      Top             =   8415
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   960
      Left            =   2565
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar paquete seleccionado"
      Top             =   8415
      Width           =   1215
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6030
      Left            =   0
      TabIndex        =   5
      Top             =   2340
      Width           =   13425
      _ExtentX        =   23680
      _ExtentY        =   10636
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
      Caption         =   "Listado de Pedidos a Proveedores"
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
      TabIndex        =   20
      Top             =   45
      Width           =   3645
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12870
      Picture         =   "frmPP_Listado.frx":20A4
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Proveedores"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   19
      Top             =   360
      Width           =   1680
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   13495
   End
End
Attribute VB_Name = "frmPP_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdImprimir_Click()
    If lista.ListItems.Count > 0 Then
        generar_excel_listado
    Else
        MsgBox "Para exportar debe existir algún registro en la lista", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmbCentro_Change()
    cargar_lista
End Sub

Private Sub cmbDescripcion_change()
    cargar_lista
End Sub

Private Sub cmbfamilia_Change()
    cargar_lista
End Sub

Private Sub cmdEnviar_Click()
   On Error GoTo cmdEnviar_Click_Error

    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a marcar el pedido como Enviado. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oPP As New clsPP
        If oPP.Enviar(lista.ListItems(lista.selectedItem.Index)) Then
            actualizarLista
            MsgBox "Pedido Enviado Correctamente.", vbInformation, App.Title
        End If
        Set oPP = Nothing
    End If

   On Error GoTo 0
   Exit Sub

cmdEnviar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEnviar_Click of Formulario frmPP_Listado"
End Sub

Private Sub cmdRecepcionar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    frmPP_Detalle_Recepcion.PK = lista.ListItems(lista.selectedItem.Index)
    frmPP_Detalle_Recepcion.Show 1
End Sub


Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a duplicar el pedido seleccionado. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
      Dim PP As Long
      Dim oPP As New clsPP
      Dim oPP_N As New clsPP
      If oPP.Carga(lista.ListItems(lista.selectedItem.Index).Text) = True Then
          With oPP_N
            .setPRESUPUESTO = oPP.getPRESUPUESTO
            .setOBSERVACIONES = oPP.getOBSERVACIONES
            .setCENTRO_ID = oPP.getCENTRO_ID
            .setPROVEEDOR_ID = oPP.getPROVEEDOR_ID
            .setFACTURA_RECIBIDA = 0
            .setFFACTURA = "0000-00-00"
            .setNFACTURA = 0
            .setFECHA_CREACION = Left(Format(Date, "yyyy-mm-dd hh:nn:ss"), 10)
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setESTADO = SC_ESTADO_PENDIENTE
            .setTIPO = oPP.getTIPO
            .setFECHA_ENVIO = "NULL"
            .setFECHA_RECEPCION = "NULL"
            
            PP = .Insertar
             
            If PP = 0 Then
                MsgBox "Error al insertar el pedido duplicado.", vbCritical, App.Title
               Exit Sub
            End If
          End With
          ' Detalle
          Dim oPP_Detalle As New clsPP_Detalle
          Dim rs As ADODB.Recordset
          Set rs = oPP_Detalle.Listado(lista.ListItems(lista.selectedItem.Index).Text)
          Do While Not rs.EOF
             With oPP_Detalle
                .setPP_ID = PP
                .setREFERENCIA = rs("REFERENCIA")
                .setFAMILIA_ID = rs("FAMILIA_ID")
                .setDESCRIPCION = rs("DESCRIPCION")
                .setUNIDADES = rs("UNIDADES")
                .setDESCUENTO = rs("DESCUENTO")
                .setPRECIO = moneda_bd(rs("PRECIO"))
                .setIMPORTE = moneda_bd(rs("IMPORTE"))
                
                If .Insertar = False Then
                    MsgBox "Error al insertar el detalle del pedido duplicado", vbCritical, App.Title
                    Exit Sub
                End If
             End With
            rs.MoveNext
          Loop
          MsgBox "El pedido se ha duplicado correctamente.", vbOKOnly + vbInformation, App.Title
          cargar_lista
          frmPP_Detalle.PK = PP
          frmPP_Detalle.Show 1
      End If
    End If
    Exit Sub
fallo:
    MsgBox "Error en el proceso de duplicación.", vbCritical, App.Title
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Número", 1000, lvwColumnLeft
        .Add , , "Proveedor", 4400, lvwColumnLeft
        .Add , , "Importe", 1100, lvwColumnCenter
        .Add , , "F.Alta", 1750, lvwColumnCenter
        .Add , , "F.Envío", 1150, lvwColumnCenter
        .Add , , "F.Recepción", 1150, lvwColumnCenter
        .Add , , "Centro", 1150, lvwColumnCenter
        .Add , , "Usuario Alta", 1150, lvwColumnCenter
        .Add , , "ID_Proveedor", 1, lvwColumnCenter
    End With
End Sub
Private Sub generar_excel_listado()
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    
   On Error GoTo generar_excel_listado_Error

    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Add
    Set XLS = XLW.Worksheets(1)
    Me.MousePointer = 11
    XLW.Worksheets(3).Delete
    XLW.Worksheets(2).Delete
    XLW.Worksheets(1).Name = "Listado de Pedidos a Proveedor"
    XLS.Range("1:1").HorizontalAlignment = xlCenter
    XLS.Range("1:1").VerticalAlignment = xlCenter
    XLS.Range("1:1").RowHeight = 30
    XLS.Range("1:1").WrapText = True
    'Cabecera
    XLS.Cells(1, 1) = "Número"
    XLS.Cells(1, 2) = "Proveedor"
    XLS.Cells(1, 3) = "Importe"
    XLS.Cells(1, 4) = "F.Alta"
    XLS.Cells(1, 5) = "F.Envío"
    XLS.Cells(1, 6) = "F.Recepción"
    XLS.Cells(1, 7) = "Centro"
    XLS.Cells(1, 8) = "Usuario Alta"
    
    Dim i As Integer
    i = 2
    ' Datos
    For i = 1 To lista.ListItems.Count
        XLS.Cells(i + 1, 1) = lista.ListItems(i).SubItems(1)
        XLS.Cells(i + 1, 2) = lista.ListItems(i).SubItems(2)
        XLS.Cells(i + 1, 3) = lista.ListItems(i).SubItems(3) ' Importe
        XLS.Cells(i + 1, 4) = Format(lista.ListItems(i).SubItems(4), "yyyy-mm-dd")
        XLS.Cells(i + 1, 5) = Format(lista.ListItems(i).SubItems(5), "yyyy-mm-dd")
        XLS.Cells(i + 1, 6) = Format(lista.ListItems(i).SubItems(6), "yyyy-mm-dd")
        XLS.Cells(i + 1, 7) = lista.ListItems(i).SubItems(7)
        XLS.Cells(i + 1, 8) = lista.ListItems(i).SubItems(8)
    Next
    For i = 1 To 8
        XLS.Columns(i).AutoFit
    Next
    XLS.Range("2:" & lista.ListItems.Count + 1).HorizontalAlignment = xlLeft
    
    Me.MousePointer = 0
    XLA.visible = True
'    Set XLS = Nothing
'    Set XLW = Nothing
'    Set XLA = Nothing
   On Error GoTo 0
   Exit Sub

generar_excel_listado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_excel_listado of Formulario frmEquipoListado"
    
End Sub


Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    cabecera
    cargar_combos
    datFechaDesde = DateAdd("m", -1, Date)
    datFechaHasta = Date
    cargar_lista
End Sub

Private Sub cmdCrearDocumentos_Click()
On Error GoTo fallo
    
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    Dim ID_PAQUETE As Long
    
    Me.MousePointer = vbHourglass
    log ("Comienzo impresion de documentos pedidos a proveedor")
    
    Me.MousePointer = vbHourglass
    ID_PAQUETE = CLng(lista.selectedItem.Text)
 
    With frmReport
        .iniciar
        .informe = "\Pedidos\rptPedidos_Proveedor"
        .criterio = "{PP.ID_PP}=" & ID_PAQUETE & " and {fp.CODIGO} = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_FP & " and {vencimiento.CODIGO} = " & DECODIFICADORA.DECODIFICADORA_PROVEEDORES_VENCIMIENTOS
        .imprimir = False
        .generar
        .visible = True
    End With
 
    Me.MousePointer = vbNormal
    
    frmReport.pdf = ""
    Me.MousePointer = vbNormal
    log ("Final impresion de documentos pedidos a proveedor")
    Exit Sub
  
fallo:
    Me.MousePointer = vbNormal
    MsgBox "Error al generar el documento " & Err.Description, vbCritical, App.Title
End Sub


Private Sub cmdmail_Click()
    envioCorreoProveedor
End Sub

Private Sub envioCorreoProveedor()
    Dim oPaquete As New clsPP
    Dim oProveedor As New clsProveedor
    Dim mail As String
    Dim ASUNTO As String
    Dim texto As String
    Dim pdf As String
    Dim code As String

    On Error Resume Next
    MkDir App.Path & "\tmp"

    code = Replace(lista.selectedItem.SubItems(1), "/", "_")
    pdf = App.Path & "\tmp\PedidoProveedor_" & code & ".pdf"
 
    On Error Resume Next
    Kill pdf

    On Error GoTo errorCorreo
    oPaquete.imprimir CLng(lista.selectedItem.Text), pdf

    If oProveedor.Carga(CLng(lista.selectedItem.SubItems(9))) Then
        'M1339-I
        'mail = oProveedor.getEMAIL
'        If Trim(oProveedor.getEMAIL_FACTURACION) = "" Then
            mail = oProveedor.getEMAIL
'        Else
'            mail = oProveedor.getEMAIL_FACTURACION
'        End If
        'M1339-F
    End If

    ASUNTO = "Pedido con código: " & lista.selectedItem.SubItems(1)
    texto = "Solicitamos el pedido del documento pdf adjunto."

    Call enviar_correo(mail, "", "", True, texto, ASUNTO, pdf)
    Exit Sub
    
errorCorreo:
    Me.MousePointer = vbNormal
    MsgBox "Error al generar / enviar el correo de pedido " & Err.Description, vbCritical, App.Title
End Sub
Private Sub datFechaDesde_Change()
    Call cargar_lista
End Sub

Private Sub datFechaHasta_Change()
    Call cargar_lista
End Sub
Private Sub datFechaFactura_Change()
    Call cargar_lista
End Sub
Private Sub datFechaFacturaF_Change()
    Call cargar_lista
End Sub
Private Sub txtfiltro_Change(Index As Integer)
    Call cargar_lista
End Sub
Private Sub cmbSubcontratas_Change()
    Call cargar_lista
End Sub
Private Sub txtfiltro_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case KeyAscii
        Case Asc("'"): ' no se permite introducir comillas simples
            KeyAscii = 0
    End Select
End Sub
' lista
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.selectedItem.SubItems(5) <> "" Then
      cmdModificar.Enabled = True
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
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
' -------------------

' botones
Private Sub cmdAnadir_Click()
     frmPP_Detalle.PK = 0
     frmPP_Detalle.Show 1
     cargar_lista
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmPP_Detalle.PK = lista.ListItems(lista.selectedItem.Index)
        frmPP_Detalle.Show 1
        actualizarLista
    Else
        MsgBox "Debe seleccionar el pedido que desea modificar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdEliminar_Click()
    If Not (lista.selectedItem Is Nothing) Then
        If MsgBox("Se va a eliminar el pedido a proveedor, nº : " & lista.selectedItem & vbCrLf & _
                  "¿Está seguro?", vbYesNo + vbInformation, App.Title) = vbYes Then
            Dim oPP As New clsPP
            If oPP.Eliminar(lista.ListItems(lista.selectedItem.Index)) Then
                MsgBox "El pedido se ha eliminado correctamente.", vbOKOnly + vbInformation, App.Title
            End If
            Call cargar_lista
            Set oPP = Nothing
        End If
    Else
        MsgBox "Debe seleccionar el pedido que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oPP As New clsPP
    lista.ListItems.Clear

    Set rs = oPP.Listado(txtfiltro(1), txtfiltro(2), cmbCentro.BoundText, cmbSubcontratas.getPK_SALIDA, Format(datFechaDesde, "yyyy-mm-dd 00:00:00"), Format(datFechaHasta, "yyyy-mm-dd 23:59:59"), cmbFamilia.getPK_SALIDA, cmbDescripcion.getTEXTO)

    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0)) ' ID
                .SubItems(1) = Format(rs(1), "0000") & "/" & rs(2) ' NUMERO
                .SubItems(2) = rs(3) ' PROVEEDOR
                .SubItems(3) = rs(4) ' PRESUPUESTO
                .SubItems(4) = rs(5) ' F.ALTA
                If Not IsNull(rs(6)) Then
                    .SubItems(5) = rs(6) ' F.ENVIO
                End If
                If Not IsNull(rs(7)) Then
                    .SubItems(6) = rs(7) ' F.RECEPCION
                End If
                .SubItems(7) = rs(8) ' CENTRO
                .SubItems(8) = rs(9) ' USUARIO
                .SubItems(9) = rs(10)
            End With
            rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
    lblsubtitulo = "Número de pedidos mostrados : " & rs.RecordCount
End Sub
Private Sub actualizarLista()
    Dim rs As ADODB.Recordset
    Dim oPP As New clsPP
    Set rs = oPP.ListadoId(lista.ListItems(lista.selectedItem.Index))
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems(lista.selectedItem.Index)
                .SubItems(1) = Format(rs(1), "0000") & "/" & rs(2) ' NUMERO
                .SubItems(2) = rs(3) ' PROVEEDOR
                .SubItems(3) = rs(4) ' PRESUPUESTO
                .SubItems(4) = rs(5) ' F.ALTA
                If Not IsNull(rs(6)) Then
                    .SubItems(5) = rs(6) ' F.ENVIO
                End If
                If Not IsNull(rs(7)) Then
                    .SubItems(6) = rs(7) ' F.RECEPCION
                End If
                .SubItems(7) = rs(8) 'CENTRO
                .SubItems(8) = rs(9) ' USUARIO
                .SubItems(9) = rs(10)
            End With
            rs.MoveNext
        Loop Until rs.EOF
        lista_Click
    End If
End Sub

Public Function alguno_seleccionado() As Boolean
    Dim booAlgunoSeleccionado As Boolean
    Dim i As Long
    
    alguno_seleccionado = True
    
    booAlgunoSeleccionado = False
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            booAlgunoSeleccionado = True
        End If
    Next i
    If Not booAlgunoSeleccionado Then
        alguno_seleccionado = False
        MsgBox "Debe seleccionar al menos un pedido.", vbOKOnly + vbInformation, App.Title
        Exit Function
    End If
    
End Function

Private Sub cargar_combos()
    cargar_combo cmbCentro, New clsCentros
    llenar_combo cmbFamilia, New clsFamilias, 0, Me, ""
    ' Proveedores que tienen pedidos
    Dim c As String
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        c = "SELECT DISTINCT A.ID_PROVEEDOR, A.NOMBRE " & _
            "  FROM PROVEEDORES A, PP B " & _
            " WHERE A.ID_PROVEEDOR = B.PROVEEDOR_ID "
        With cmbSubcontratas
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setFILTRO = ""
            .setTABLA = "PROVEEDORES"
            .setDESCRIPCION = "Proveedores"
            .setPK = "ID_PROVEEDOR"
            .setCAMPO = "NOMBRE"
            .setQUERY = c
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmProveedores_Detalle
        End With
        c = "SELECT DISTINCT A.ID, A.DESCRIPCION " & _
            "  FROM PP_DETALLE A " & _
            " WHERE 1 = 1 "
        With cmbDescripcion
            .setCONN = conn
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setFILTRO = ""
            .setTABLA = "PP_DETALLE"
            .setDESCRIPCION = "Detalle de Pedido"
            .setPK = "A.ID"
            .setCAMPO = "A.DESCRIPCION"
            .setQUERY = c
            .setMUESTRA_DETALLE = False
            Set .FORMULARIO = Me
        End With
    End If
End Sub

