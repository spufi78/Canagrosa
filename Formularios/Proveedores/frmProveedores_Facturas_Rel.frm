VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmProveedores_Facturas_Rel 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Listado de Relaciones de las facturas"
   ClientHeight    =   9390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14595
   Icon            =   "frmProveedores_Facturas_Rel.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   14595
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   510
      Left            =   7875
      TabIndex        =   7
      Top             =   540
      Width           =   6675
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ped. Generales"
         Height          =   240
         Index           =   3
         Left            =   4680
         TabIndex        =   11
         Top             =   180
         Width           =   1590
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ped. Reactivos"
         Height          =   240
         Index           =   2
         Left            =   2925
         TabIndex        =   10
         Top             =   180
         Width           =   1680
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcontrataciones"
         Height          =   240
         Index           =   1
         Left            =   1080
         TabIndex        =   9
         Top             =   180
         Width           =   1635
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   180
         Value           =   -1  'True
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   13455
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8415
      Width           =   1095
   End
   Begin MSComctlLib.ListView listaSi 
      Height          =   7785
      Left            =   45
      TabIndex        =   0
      Top             =   570
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   13732
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
   Begin XtremeSuiteControls.PushButton cmdAnadir 
      Height          =   480
      Index           =   0
      Left            =   6750
      TabIndex        =   4
      Top             =   4050
      Width           =   1110
      _Version        =   851970
      _ExtentX        =   1958
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Añadir"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas_Rel.frx":030A
   End
   Begin XtremeSuiteControls.PushButton cmdborrar 
      Height          =   480
      Left            =   6750
      TabIndex        =   5
      Top             =   4545
      Width           =   1110
      _Version        =   851970
      _ExtentX        =   1958
      _ExtentY        =   847
      _StockProps     =   79
      Caption         =   "Eliminar"
      Appearance      =   5
      Picture         =   "frmProveedores_Facturas_Rel.frx":6B6C
   End
   Begin MSComctlLib.ListView listaNo 
      Height          =   7290
      Left            =   7875
      TabIndex        =   6
      Top             =   1080
      Width           =   6675
      _ExtentX        =   11774
      _ExtentY        =   12859
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Relaciones de las Facturas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   3
      Top             =   270
      Width           =   2460
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   14040
      Picture         =   "frmProveedores_Facturas_Rel.frx":D3CE
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturas del proveedor : "
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
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   0
      Width           =   2625
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   17640
   End
End
Attribute VB_Name = "frmProveedores_Facturas_Rel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_PROVEEDOR_ID As Long
Public PK_FACTURA_ID As Long
       
Private Sub cabecera()
    With listaSi.ColumnHeaders
        .Add , , "", 300, lvwColumnLeft
        .Add , , "TOBJETO", 1, lvwColumnLeft
        .Add , , "COBJETO", 1, lvwColumnLeft
        .Add , , "Tipo", 4000, lvwColumnLeft
        .Add , , "Número", 1100, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
    End With
    With listaNo.ColumnHeaders
        .Add , , "", 300, lvwColumnLeft
        .Add , , "COBJETO", 1, lvwColumnLeft
        .Add , , "Tipo", 3900, lvwColumnLeft
        .Add , , "Número", 1100, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
    End With
End Sub

Private Sub cmdAnadir_Click(Index As Integer)
    Dim oPFR As New clsProveedores_facturas_rel
    For i = 1 To listaNo.ListItems.Count
        If listaNo.ListItems(i).Checked = True Then
            With oPFR
                .setFACTURA_ID = PK_FACTURA_ID
                .setTOBJETO = listaNo.ListItems(i).Text
                .setCOBJETO = listaNo.ListItems(i).SubItems(1)
                .Insertar
            End With
        End If
    Next
    cargarListaSi
    cargarListaNo
End Sub

Private Sub cmdBorrar_Click()
    Dim existe As Boolean
    existe = False
    For i = 1 To listaSi.ListItems.Count
        If listaSi.ListItems(i).Checked = True Then
            existe = True
        End If
    Next
    If Not existe Then
        MsgBox "Marque las relaciones a eliminar.", vbCritical, App.Title
        Exit Sub
    End If
    Dim oPFR As New clsProveedores_facturas_rel
    For i = 1 To listaSi.ListItems.Count
        If listaSi.ListItems(i).Checked = True Then
            oPFR.Eliminar listaSi.ListItems(i).Text
        End If
    Next
    cargarListaSi
    cargarListaNo
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_proveedor
    cargarListaSi
    cargarListaNo
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cargar_proveedor()
    Dim oProveedor As New clsProveedor
    With oProveedor
        .Carga (PK_PROVEEDOR_ID)
        lbltitulo(0) = "Documentos Relacionados : " & .getNOMBRE
        Me.Caption = lbltitulo(0)
    End With
    Set oProveedor = Nothing
End Sub
Private Sub cargarListaSi()
    Dim rs As New ADODB.Recordset
    Dim oPFR As New clsProveedores_facturas_rel
    Set rs = oPFR.Listado(PK_FACTURA_ID)
    listaSi.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With listaSi.ListItems.Add(, , Format(rs("ID"), "000")) ' TOBJETO
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
            .SubItems(4) = rs(4)
            .SubItems(5) = Format(rs(5), "dd-mm-yyyy")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub
Private Sub cargarListaNo()
    listaNo.ListItems.Clear
    Dim tipo As Integer
    If opTipo(0).Value = True Then
        cargarSC PK_PROVEEDOR_ID
        cargarREX PK_PROVEEDOR_ID
        cargarPP PK_PROVEEDOR_ID
    ElseIf opTipo(1).Value = True Then
        cargarSC PK_PROVEEDOR_ID
    ElseIf opTipo(2).Value = True Then
        cargarREX PK_PROVEEDOR_ID
    ElseIf opTipo(3).Value = True Then
        cargarPP PK_PROVEEDOR_ID
    End If
End Sub
Private Sub cargarSC(PK_PROVEEDOR_ID As Long)
    Dim rs As New ADODB.Recordset
    Dim oSC As New clsSC_Paquetes
    Set rs = oSC.ListadoProveedor(PK_PROVEEDOR_ID)
    Dim encontrado As Boolean
    If rs.RecordCount <> 0 Then
        Do
           encontrado = False
           For i = 1 To listaSi.ListItems.Count
            If CInt(rs(0)) = CInt(listaSi.ListItems(i).SubItems(1)) And rs(1) = listaSi.ListItems(i).SubItems(2) Then
                encontrado = True
            End If
           Next
           If Not encontrado Then
            With listaNo.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
            End With
           End If
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oSC = Nothing
End Sub
Private Sub cargarREX(PROVEEDOR_ID As Long)
    Dim rs As New ADODB.Recordset
    Dim oPB As New clsPedidos_bote_ex
    Set rs = oPB.ListadoRelaciones(PROVEEDOR_ID)
    Dim encontrado As Boolean
    If rs.RecordCount <> 0 Then
        Do
           encontrado = False
           For i = 1 To listaSi.ListItems.Count
            If TOBJETO_PEDIDO_BOTE_EX = CInt(listaSi.ListItems(i).SubItems(1)) And rs(0) = listaSi.ListItems(i).SubItems(2) Then
                encontrado = True
            End If
           Next
           If Not encontrado Then
            With listaNo.ListItems.Add(, , TOBJETO_PEDIDO_BOTE_EX)
             .SubItems(1) = rs(0)
             .SubItems(2) = "Pedido de Reactivo"
             .SubItems(3) = rs(0)
             .SubItems(4) = Format(rs(1), "dd-mm-yyyy")
            End With
           End If
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oPB = Nothing
End Sub
Private Sub cargarPP(PROVEEDOR_ID As Long)
    Dim rs As New ADODB.Recordset
    Dim oPB As New clsPP
    Set rs = oPB.ListadoRelaciones(PROVEEDOR_ID)
    Dim encontrado As Boolean
    If rs.RecordCount <> 0 Then
        Do
           encontrado = False
           For i = 1 To listaSi.ListItems.Count
            If TOBJETO_PEDIDO_PROVEEDOR = CInt(listaSi.ListItems(i).SubItems(1)) And rs(0) = listaSi.ListItems(i).SubItems(2) Then
                encontrado = True
            End If
           Next
           If Not encontrado Then
            With listaNo.ListItems.Add(, , TOBJETO_PEDIDO_PROVEEDOR)
             .SubItems(1) = rs(0)
             .SubItems(2) = "Pedido General"
             .SubItems(3) = rs(1)
             .SubItems(4) = Format(rs(2), "dd-mm-yyyy")
            End With
           End If
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oPB = Nothing
End Sub
Private Sub opTipo_Click(Index As Integer)
    cargarListaNo
End Sub
