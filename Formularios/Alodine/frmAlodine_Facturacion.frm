VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmAlodine_Facturacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alodine pendiente de facturación"
   ClientHeight    =   9930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15945
   Icon            =   "frmAlodine_Facturacion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9930
   ScaleWidth      =   15945
   Begin VB.Frame frmBusqueda 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   45
      TabIndex        =   18
      Top             =   360
      Width           =   15840
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   900
         TabIndex        =   19
         Top             =   270
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   9945
         TabIndex        =   21
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
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
         CalendarTitleBackColor=   14737632
         Format          =   51445761
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   12060
         TabIndex        =   22
         Top             =   285
         Width           =   1320
         _ExtentX        =   2328
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
         CalendarTitleBackColor=   14737632
         Format          =   51445761
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta el"
         Height          =   195
         Index           =   4
         Left            =   11385
         TabIndex        =   24
         Top             =   315
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde el"
         Height          =   195
         Index           =   2
         Left            =   9135
         TabIndex        =   23
         Top             =   345
         Width           =   645
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clientes"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   20
         Top             =   315
         Width           =   585
      End
   End
   Begin VB.Frame frmFacturaManual 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Informar Número de Factura en los alodines marcados"
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
      Height          =   885
      Left            =   4050
      TabIndex        =   9
      Top             =   3870
      Visible         =   0   'False
      Width           =   7695
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2805
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   330
         Width           =   915
      End
      Begin VB.TextBox txtnumero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   855
         TabIndex        =   13
         Top             =   330
         Width           =   1185
      End
      Begin XtremeSuiteControls.PushButton cmdCerrarFacturaManual 
         Height          =   435
         Left            =   6210
         TabIndex        =   10
         Top             =   270
         Width           =   1410
         _Version        =   851970
         _ExtentX        =   2487
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Cerrar"
         Appearance      =   5
         Picture         =   "frmAlodine_Facturacion.frx":030A
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   435
         Left            =   4770
         TabIndex        =   12
         Top             =   270
         Width           =   1410
         _Version        =   851970
         _ExtentX        =   2487
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Informar"
         Appearance      =   5
         Picture         =   "frmAlodine_Facturacion.frx":6B6C
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   360
         Left            =   3721
         TabIndex        =   15
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   2004
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196612
         OrigLeft        =   3510
         OrigTop         =   315
         OrigRight       =   3750
         OrigBottom      =   660
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   17
         Top             =   405
         Width           =   585
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   225
         Index           =   1
         Left            =   2295
         TabIndex        =   16
         Top             =   390
         Width           =   345
      End
   End
   Begin VB.CommandButton cmdAlbaran 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear Albaran"
      Height          =   915
      Left            =   11655
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAlodine_Facturacion.frx":D3CE
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8970
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Lote"
      Height          =   330
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8970
      Width           =   1410
   End
   Begin VB.CommandButton cmdCrear 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Crear &Factura"
      Height          =   915
      Index           =   2
      Left            =   13095
      MaskColor       =   &H00FFFFFF&
      Picture         =   "frmAlodine_Facturacion.frx":E010
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8970
      UseMaskColor    =   -1  'True
      Width           =   1365
   End
   Begin VB.CommandButton cmdMarcar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   330
      Left            =   1470
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8970
      Width           =   1410
   End
   Begin VB.CommandButton cmdDesmarcar 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   330
      Left            =   2910
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8970
      Width           =   1410
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   915
      Left            =   14535
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8970
      Width           =   1365
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7800
      Left            =   45
      TabIndex        =   1
      Top             =   1125
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   13758
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
   Begin XtremeSuiteControls.PushButton cmdNoFacturable 
      Height          =   480
      Left            =   4635
      TabIndex        =   11
      Top             =   9360
      Width           =   3075
      _Version        =   851970
      _ExtentX        =   5424
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Marcar como Factura Manual"
      Appearance      =   5
      Picture         =   "frmAlodine_Facturacion.frx":EC52
   End
   Begin XtremeSuiteControls.PushButton cmdNoFacturables 
      Height          =   480
      Left            =   7785
      TabIndex        =   25
      Top             =   9360
      Width           =   3210
      _Version        =   851970
      _ExtentX        =   5662
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Marcar como NO FACTURABLES"
      Appearance      =   5
      Picture         =   "frmAlodine_Facturacion.frx":154B4
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre el alodine para ver el detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   5895
      TabIndex        =   2
      Top             =   8970
      Width           =   4065
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Lotes de Alodine pendientes de facturación"
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
      Height          =   330
      Index           =   4
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   15855
   End
End
Attribute VB_Name = "frmAlodine_Facturacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClientes_change()
    cargar_lista
End Sub

Private Sub cmdAlbaran_Click()
    Dim strcadena As String
    If contar_marcados = 0 Then
        MsgBox "Debe seleccionar algún lote de alodine.", vbInformation, App.Title
        Exit Sub
    End If
    If Not validarPedidos Then Exit Sub
    If Not informacionFacturas("Albaran/es") Then Exit Sub
'    strcadena = "Va a generar albaranes para  " & contar_marcados & " lotes de alodine. ¿Desea continuar?"
'    If MsgBox(strcadena, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        generar_documentos_no_eads (1) ' Tipo 3 Factura de Alodine
        generar_documentos_eads (1) ' Tipo 3 Factura de Alodine
        Me.MousePointer = 0
'    End If
    Call cargar_lista
End Sub
Private Function informacionFacturas(TIPO As String) As Boolean
    Dim rs As ADODB.Recordset
    Dim salida As String
    Dim oAlodine_planificacion As New clsAlodine_planificacion
    Set rs = oAlodine_planificacion.Listado_Lotes_Informacion(inlote, incliente, inpedidos, 0) ' NO EADS
    If rs.RecordCount > 0 Then
        salida = salida & "Se van a generar " & rs.RecordCount & " " & TIPO & " para clientes NO AIRBUS : " & vbNewLine & vbNewLine
        Do
            salida = salida & " - " & rs(2)
            If rs(3) <> "" Then
                salida = salida & " Pedido : " & rs(3)
            End If
            salida = salida & vbNewLine
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = oAlodine_planificacion.Listado_Lotes_Informacion(inlote, incliente, inpedidos, 1) ' EADS
    If rs.RecordCount > 0 Then
        If salida <> "" Then
            salida = salida & vbNewLine
        End If
        salida = salida & "Se van a generar 1 factura con los siguientes clientes AIRBUS : " & vbNewLine & vbNewLine
'        salida = salida & "Se van a generar " & rs.RecordCount & " " & tipo & " para clientes AIRBUS : " & vbNewLine & vbNewLine
        Do
            salida = salida & " - " & rs(2)
            If rs(3) <> "" Then
                salida = salida & " Pedido : " & rs(3)
            End If
            salida = salida & vbNewLine
            rs.MoveNext
        Loop Until rs.EOF
    End If
    If salida <> "" Then
        If MsgBox(salida & vbNewLine & "¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            informacionFacturas = False
            Exit Function
        End If
    End If
    informacionFacturas = True
        

End Function
Private Function validarPedidos() As Boolean
    ' Validar pedidos
    Dim ERROR As String
    Dim e As Boolean
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        e = False
        If lista.ListItems(i).Checked = True And lista.ListItems(i).SubItems(9) <> "" Then
            Dim f As Date
            f = lista.ListItems(i).SubItems(4)
            If lista.ListItems(i).SubItems(9) <> "" And lista.ListItems(i).SubItems(8) <> "0" Then
                Dim oCP As New clsClientes_pedidos
                oCP.Carga lista.ListItems(i).SubItems(8)
                If f < oCP.getFECHA_PEDIDO Or f > oCP.getFECHA_BAJA Then
                    e = True
                End If
            End If
        End If
        If e Then
            ERROR = ERROR & vbNewLine & " - El pedido : " & lista.ListItems(i).SubItems(9) & " de " & lista.ListItems(i).Text & " no existe para la fecha " & Format(f, "dd-mm-yyyy")
        End If
    Next
    If ERROR <> "" Then
        If MsgBox("Se han detectado los siguientes errores en los pedidos. " & vbNewLine & ERROR & vbNewLine & vbNewLine & " ¿Desea continuar aunque los pedidos esten erroneos?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            validarPedidos = False
            Exit Function
        End If
    End If
    validarPedidos = True

End Function

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCerrarFacturaManual_Click()
    frmFacturaManual.visible = False
End Sub

Private Sub cmdCrear_Click(Index As Integer)
    Dim strcadena As String
    If contar_marcados = 0 Then
        MsgBox "Debe seleccionar algún lote de alodine.", vbInformation, App.Title
        Exit Sub
    End If
    If Not validarPedidos Then Exit Sub
    If Not informacionFacturas("Factura/s") Then Exit Sub
'    strcadena = "Va a generar documentos de pago para  " & contar_marcados & " lotes de alodine. ¿Desea continuar?"
'    If MsgBox(strcadena, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Me.MousePointer = 11
        generar_documentos_no_eads (2) ' Tipo 3 Factura de Alodine
        generar_documentos_eads (2) ' Tipo 3 Factura de Alodine
'        informar_facturados
        Me.MousePointer = 0
'    End If
    Call cargar_lista
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdNoFacturable_Click()
    frmFacturaManual.visible = Not frmFacturaManual.visible
End Sub

Private Sub cmdNoFacturables_Click()
   On Error GoTo cmdNoFacturables_Click_Error
    If contar_marcados = 0 Then
        MsgBox "Debe seleccionar algún lote de alodine.", vbCritical, App.Title
        Exit Sub
    End If
    If MsgBox("Se marcaran los lotes como NO FACTURABLES. ¿Estas completamente insegura?", vbQuestion + vbYesNo, App.Title) = vbNo Then Exit Sub
    Dim OAP As New clsAlodine_planificacion
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            OAP.doc_pago lista.ListItems(i).SubItems(6), lista.ListItems(i).SubItems(7), lista.ListItems(i).SubItems(8), -1
        End If
    Next
    MsgBox "Lotes actualizados correctamente.", vbInformation, App.Title
    cargar_lista

   On Error GoTo 0
   Exit Sub

cmdNoFacturables_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdNoFacturables_Click of Formulario frmAlodine_Facturacion"
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    Dim LOTE As Integer
    If lista.ListItems.Count > 0 Then
        LOTE = lista.ListItems(lista.selectedItem.Index).SubItems(6)
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).SubItems(6) = LOTE Then
                lista.ListItems(i).Checked = True
                lista.ListItems(i).Selected = True
                lista.ListItems(i).EnsureVisible
            End If
        Next
    End If
End Sub

Private Sub fdesde_Change()
    cargar_lista
End Sub

Private Sub fhasta_Change()
    cargar_lista
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    txtanno = Year(Date)
    cambiar.Max = Year(Date)
    cambiar.min = 2004
    fdesde = "01/01/2000"
    fhasta = Date
    cabecera
    rellenar_clientes
    cargar_lista
End Sub
Private Sub rellenar_clientes()
    Dim consulta As String
    consulta = "SELECT DISTINCT C.ID_CLIENTE,C.NOMBRE " & _
               "  FROM ALODINE_LOTES AL, CLIENTES C,ALODINE A,ALODINE_PLANIFICACION AP " & _
               " WHERE AL.ALODINE_ID = A.ID_ALODINE " & _
               "   AND AL.ID_LOTE = AP.LOTE_ID " & _
               "   AND AP.CLIENTE_ID = C.ID_CLIENTE " & _
               "   AND AP.DOC_ID = 0 "
    
    Dim conn As ADODB.Connection
    If CrearConexionGlobal(conn, "", "") = True Then ' CONECTAR EL RS
        With cmbclientes
            .setCONN = conn
            .setQUERY = consulta
            .setFK_CAMPO = ""
            .setFK_VALOR = 0
            .setTABLA = "CLIENTES"
            .setDESCRIPCION = "Clientes"
            .setPK = "C.ID_CLIENTE"
            .setFILTRO = ""
            .setCAMPO = "C.NOMBRE"
            .setMUESTRA_DETALLE = True
            Set .FORMULARIO = frmClientes
        End With
    End If
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
End Sub

Public Function inlote() As String
    Dim s As String
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            s = s & lista.ListItems(i).SubItems(6) & ","
        End If
    Next
    s = Left(s, Len(s) - 1)
    inlote = s
End Function
Public Function incliente() As String
    Dim s As String
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            s = s & lista.ListItems(i).SubItems(7) & ","
        End If
    Next
    s = Left(s, Len(s) - 1)
    incliente = s
End Function
Public Function inpedidos() As String
    Dim s As String
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            s = s & lista.ListItems(i).SubItems(8) & ","
        End If
    Next
    s = Left(s, Len(s) - 1)
    inpedidos = s
End Function

Public Sub generar_documentos_no_eads(TIPO_DOCUMENTO As Integer)
'    Dim i As Integer
    log ("Comiento generación de facturación alodine NO eads")
    Dim num_doc As Long
    Dim cliente_ant As Long
    Dim pedido_ant As Long
    Dim total_doc As Integer
    total_doc = 0
    cliente_ant = 0
    pedido_ant = 0
    'cIVA
    Dim oParametros As New clsParametros
'    Dim IVA As Integer
'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
    ' Planificacion
    Dim oAlodine_Lote As New clsAlodine_lotes
    Dim oAlodine_planificacion As New clsAlodine_planificacion
    Dim oDocPago As New clsDocs_pago
    Dim oCliente As New clsCliente
    Dim oConcepto As New clsDocs_pago_conceptos
    Dim rs As ADODB.Recordset
    Dim salida As String
    With oAlodine_planificacion
        Set rs = .Listado_Lotes_Grupo(inlote, incliente, inpedidos)
        If rs.RecordCount = 0 Then
            Exit Sub
        Else
            Do
                If cliente_ant <> rs(0) Or pedido_ant <> rs(10) Then
                    oDocPago.setTIPO = TIPO_DOCUMENTO
                    oDocPago.setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
                    oDocPago.setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
                    oDocPago.setEMPLEADO_ID = USUARIO.getID_EMPLEADO
                    oDocPago.setCLIENTE_ID = rs(0)
                    oDocPago.setCLIENTE_ID_FACTURA = rs(0)
                    oDocPago.setTOTAL = "0.00"
                    oDocPago.setDESCUENTO = "0.00"
'                    If TIPO_DOCUMENTO = 2 Then
'                        oDocPago.setIVA = IVA
'                    Else
'                        oDocPago.setIVA = 0
'                    End If
                    oDocPago.setPAGADO = 0
                    oDocPago.setANULADO = 0
                    oDocPago.setFACTURA_CONCEPTOS = 1
                    oCliente.CargaCliente rs(0)
                    oDocPago.setFP_ID = oCliente.getFP_ID
                    oDocPago.setPEDIDO_ID = rs(10)
                    ' Insertamos el documento de pago
                    num_doc = oDocPago.InsertarDocPago
                    If num_doc = 0 Then
                        MsgBox "Error al generar las facturas, contacte con mantenimiento.", vbCritical, App.Title
                        Exit Sub
                    End If
                    ' Informar el documento de pago en alodine_planificación
                    .doc_pago inlote, rs(0), rs(10), CLng(num_doc)
                    total_doc = total_doc + 1
                    cliente_ant = rs(0)
                    pedido_ant = rs(10)
                    
                    salida = salida & vbNewLine & " - " & oCliente.getNOMBRE & "  -> Factura : " & Format(oDocPago.getNUMERO, "0000") & "/" & Year(oDocPago.getFECHA_FACTURA)
                End If
                ' Insertamos los conceptos
                oConcepto.setDOC_ID = num_doc
                Dim oAL As New clsAlodine_lotes
                oAL.Carga rs(6)
                Dim producto As String
                Dim productoA() As String
                If InStr(1, rs(1), "#") > 0 Then
                    productoA = Split(rs(1), "#")
                    producto = productoA(1)
                Else
                    producto = rs(1)
                End If
                oConcepto.setDESCRIPCION = "Preparado de " & rs(3) & " botes de " & rs(2) & " del producto " & producto & _
                                           " (Lote:" & oAL.getNUMERO_LOTE & "/" & Year(oAL.getFECHA_ALTA) & ")"
                oConcepto.setFECHA = Format(rs(5), "yyyy-mm-dd")
                oConcepto.setAPARTADO = 0
                oConcepto.setDTO = 0
                
                oConcepto.setCANTIDAD = rs(3)
                oConcepto.setPRECIO = Replace(rs(7), ",", ".")
                oConcepto.setSUBTOTAL = Replace(rs(4), ",", ".")
                oConcepto.setTOTAL = Replace(rs(4), ",", ".")
                
                oConcepto.setFAMILIA_ID = familia.alodine  ' Alodine por defecto
                If oConcepto.Insertar = False Then
                   Exit Sub
                End If
                oDocPago.Informar_total_factura (num_doc)
                rs.MoveNext
            Loop Until rs.EOF = True
        End If
    End With
    Set oDocPago = Nothing
    Dim sTIPO As String
    If TIPO_DOCUMENTO = 1 Then
        sTIPO = "Albaran"
    Else
        sTIPO = "Factura"
    End If
    log ("Final generación de facturación alodine NO eads")
    If total_doc = 1 Then
'        MsgBox "Se ha registrado 1 " & sTIPO & " para clientes NO AIRBUS : " & vbNewLine & salida, vbOKOnly + vbInformation, App.Title
        MsgBox "Se ha registrado 1 " & sTIPO & " para clientes NO AIRBUS : " & vbNewLine & salida, vbOKOnly + vbInformation, App.Title
    Else
'        MsgBox "Se han registrado " & total_doc & " " & sTIPO & "s para clientes NO AIRBUS : " & vbNewLine & salida, vbOKOnly + vbInformation, App.Title
        MsgBox "Se han registrado " & total_doc & " " & sTIPO & "s para clientes NO AIRBUS : " & vbNewLine & salida, vbOKOnly + vbInformation, App.Title
    End If
End Sub
Public Sub generar_documentos_eads(TIPO_DOCUMENTO As Integer)
    log ("Comiento generación de facturación alodine SI eads")
'    Dim i As Integer
    Dim num_doc As Long
    Dim cliente_ant As Long
    Dim total_doc As Integer
    total_doc = 0
    cliente_ant = 0
    'cIVA
    Dim oParametros As New clsParametros
'    Dim IVA As Integer
    Dim cliente As Long
'    IVA = recuperaIVA()
'    If IVA = 0 Then
'        MsgBox "Error al recuperar el IVA. Imposible facturar.", vbCritical, App.Title
'        Exit Sub
'    End If
    oParametros.Carga parametros.ALODINE_CLIENTE_AIRBUS, ""
    cliente = CLng(oParametros.getVALOR)
    ' Planificacion
    Dim oAlodine_Lote As New clsAlodine_lotes
    Dim oAlodine_planificacion As New clsAlodine_planificacion
    Dim oDocPago As New clsDocs_pago
    Dim oConcepto As New clsDocs_pago_conceptos
    Dim rs As ADODB.Recordset
    Dim oCliente As New clsCliente
    Dim salida As String
    With oAlodine_planificacion
        Set rs = .Listado_Lotes_Grupo_EADS(inlote, incliente, inpedidos)
        If rs.RecordCount = 0 Then
            Exit Sub
        Else
            With oDocPago
                 .setTIPO = TIPO_DOCUMENTO
                 .setFECHA_FACTURA = Format(Date, "yyyy-mm-dd")
                 .setFECHA_GENERACION = Format(Date, "yyyy-mm-dd")
                 .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
'                 .setCLIENTE_ID = rs(0)
'                 .setCLIENTE_ID_FACTURA = rs(0)
                 .setCLIENTE_ID = cliente
                 .setCLIENTE_ID_FACTURA = cliente
                 .setTOTAL = "0.00"
                 .setDESCUENTO = "0.00"
'                 If TIPO_DOCUMENTO = 2 Then
'                     .setIVA = IVA
'                 Else
'                     .setIVA = 0
'                 End If
                 .setPAGADO = 0
                 .setANULADO = 0
                 .setFACTURA_CONCEPTOS = 1
'                 oCliente.CargaCliente rs(0)
                 oCliente.CargaCliente cliente
                 oDocPago.setFP_ID = oCliente.getFP_ID
                 ' Insertamos el documento de pago
                 num_doc = .InsertarDocPago
                 If num_doc = 0 Then
                     MsgBox "Error al generar las facturas, contacte con mantenimiento.", vbCritical, App.Title
                     Exit Sub
                 End If
                 ' Informar el documento de pago en alodine_planificación
                 oAlodine_planificacion.doc_pago_eads inlote, incliente, CLng(num_doc)
                 salida = salida & vbNewLine & " - " & oCliente.getNOMBRE & "  -> Factura : " & Format(oDocPago.getNUMERO, "0000") & "/" & Year(oDocPago.getFECHA_FACTURA)
                 total_doc = total_doc + 1
            End With
            Do
                ' Insertamos los conceptos
                With oConcepto
                    .setDOC_ID = num_doc
                    Dim oAL As New clsAlodine_lotes
                    oAL.Carga rs(7)
                    Dim producto As String
                    Dim productoA() As String
                    If InStr(1, rs(1), "#") > 0 Then
                        productoA = Split(rs(1), "#")
                        producto = productoA(1)
                    Else
                        producto = rs(1)
                    End If
                    .setDESCRIPCION = rs(6) & ". Preparado de " & rs(3) & " botes de " & rs(2) & " del producto " & producto & _
                                               " (Lote:" & oAL.getNUMERO_LOTE & "/" & Year(oAL.getFECHA_ALTA) & ")"
                    
                    .setFECHA = Format(rs(5), "yyyy-mm-dd")
                    .setFAMILIA_ID = familia.alodine  ' Familia de alodine por defecto
                    .setAPARTADO = 0
                    .setDTO = 0
                    
                    .setCANTIDAD = rs(3)
                    .setPRECIO = Replace(rs(8), ",", ".")
                    .setSUBTOTAL = Replace(rs(4), ",", ".")
                    .setTOTAL = Replace(rs(4), ",", ".")
                    
                    If .Insertar = False Then
                       Exit Sub
                    End If
                End With
                rs.MoveNext
            Loop Until rs.EOF = True
            oDocPago.Informar_total_factura (num_doc)
        End If
    End With
    Set oDocPago = Nothing
    Dim sTIPO As String
    If TIPO_DOCUMENTO = 1 Then
        sTIPO = "Albaran"
    Else
        sTIPO = "Factura"
    End If
    log ("Final generación de facturación alodine SI eads")
    If total_doc = 1 Then
'        MsgBox "Se ha registrado 1 " & sTIPO & " para AIRBUS.", vbOKOnly + vbInformation, App.Title
        MsgBox "Se ha registrado 1 " & sTIPO & " para clientes AIRBUS: " & vbNewLine & salida, vbOKOnly + vbInformation, App.Title
    Else
'        MsgBox "Se han registrado " & total_doc & " " & sTIPO & "s para AIRBUS.", vbOKOnly + vbInformation, App.Title
        MsgBox "Se han registrado " & total_doc & " " & sTIPO & "s para clientes AIRBUS: " & vbNewLine & salida, vbOKOnly + vbInformation, App.Title
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
Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Cliente", 4000, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Producto", 3400, lvwColumnLeft)
        .Tag = "Producto"
    End With
    With lista.ColumnHeaders.Add(, , "Codigo", 1700, lvwColumnCenter)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Numero", 900, lvwColumnCenter)
        .Tag = "Numero"
    End With
    With lista.ColumnHeaders.Add(, , "F.Alta", 1200, lvwColumnCenter)
        .Tag = "F.Alta"
    End With
    With lista.ColumnHeaders.Add(, , "Caducidad", 1200, lvwColumnCenter)
        .Tag = "Caducidad"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 600, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "CLIENTE_ID", 1, lvwColumnCenter)
        .Tag = "CLIENTE_ID"
    End With
    With lista.ColumnHeaders.Add(, , "PEDIDO_ID", 1, lvwColumnCenter)
        .Tag = "PEDIDO_ID"
    End With
    With lista.ColumnHeaders.Add(, , "PEDIDO", 1500, lvwColumnCenter)
        .Tag = "PEDIDO"
    End With
    With lista.ColumnHeaders.Add(, , "AIRBUS", 1000, lvwColumnCenter)
        .Tag = "AIRBUS"
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oAlodine_Lote As New clsAlodine_lotes
    lista.ListItems.Clear
    
    Set rs = oAlodine_Lote.Listado_Pendientes_Facturar(IIf(cmbclientes.getTEXTO = "", 0, cmbclientes.getPK_SALIDA), fdesde, fhasta)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = rs(6)
             .SubItems(7) = rs(8)
             .SubItems(8) = rs(9) 'PEDIDO_ID
             If Not IsNull(rs(10)) Then
                 .SubItems(9) = rs(10) 'PEDIDO
             End If
             If rs(11) = 0 Then
                .SubItems(10) = ""
             Else
                .SubItems(10) = "S"
             End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oAlodine_Lote = Nothing
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        glote = lista.ListItems(lista.selectedItem.Index).SubItems(6)
        frmAlodine_Lote.Show 1
        glote = 0
    End If
End Sub
Public Function contar_marcados() As Integer
    Dim i As Integer
    contar_marcados = 0
    For i = 1 To lista.ListItems.Count
       If lista.ListItems(i).Checked = True Then
        contar_marcados = contar_marcados + 1
      End If
    Next
End Function

'Public Sub informar_facturados()
'    Dim i As Integer
'    Dim oAlodine_Lote As New clsAlodine_lotes
'    For i = 1 To lista.ListItems.Count
'        If lista.ListItems(i).Checked = True Then
'            ' Updateamos el lote con el numero del documento de pago
'            oAlodine_Lote.setDOC_ID = 1
'            oAlodine_Lote.Informar_Documento_Pago (lista.ListItems(i).SubItems(6))
'        End If
'    Next
'End Sub
Private Sub PushButton2_Click()
    Dim strcadena As String
   On Error GoTo PushButton2_Click_Error

    If contar_marcados = 0 Then
        MsgBox "Debe seleccionar algún lote de alodine.", vbCritical, App.Title
        Exit Sub
    End If
    If txtnumero = "" Then
        MsgBox "Debe informar el número de factura.", vbCritical, App.Title
        Exit Sub
    End If
    If txtanno = "" Then
        MsgBox "Debe informar el año de la factura.", vbCritical, App.Title
        Exit Sub
    End If
    Dim oDocPago As New clsDocs_pago
    Dim ID As Long
    ID = oDocPago.recuperarIdPorNumero(txtnumero, txtanno)
    
    If ID = 0 Then
        MsgBox "El número de factura indicado NO EXISTE.", vbCritical, App.Title
        Exit Sub
    Else
        Dim OAP As New clsAlodine_planificacion
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                OAP.doc_pago lista.ListItems(i).SubItems(6), lista.ListItems(i).SubItems(7), lista.ListItems(i).SubItems(8), ID
            End If
        Next
        MsgBox "Lotes actualizados correctamente.", vbInformation, App.Title
        frmFacturaManual.visible = False
        cargar_lista
    End If

   On Error GoTo 0
   Exit Sub

PushButton2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PushButton2_Click of Formulario frmAlodine_Facturacion"
End Sub
