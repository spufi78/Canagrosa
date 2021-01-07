VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmTesoreria 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Listado de Facturas de Proveedores"
   ClientHeight    =   12540
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   15465
   Icon            =   "frmTesoreria.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   12540
   ScaleWidth      =   15465
   Begin VB.Frame frmProveedores 
      BackColor       =   &H00C0C0C0&
      Height          =   9285
      Left            =   45
      TabIndex        =   3
      Top             =   1170
      Visible         =   0   'False
      Width           =   15360
      Begin VB.TextBox txtListaTipo 
         Height          =   375
         Left            =   5175
         TabIndex        =   15
         Top             =   8325
         Width           =   420
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cerrar"
         Height          =   870
         Left            =   14175
         Picture         =   "frmTesoreria.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   8190
         Width           =   1050
      End
      Begin MSComctlLib.ListView lista 
         Height          =   7485
         Left            =   90
         TabIndex        =   4
         Top             =   540
         Width           =   15165
         _ExtentX        =   26749
         _ExtentY        =   13203
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
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
      Begin VB.Label lblProveedor 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "INFORME DE TESORERIA"
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
         Left            =   180
         TabIndex        =   14
         Top             =   180
         Width           =   3300
      End
      Begin VB.Label lblRetencion 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   11835
         TabIndex        =   12
         Top             =   8640
         Width           =   2085
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Retención"
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
         Height          =   285
         Index           =   3
         Left            =   10620
         TabIndex        =   11
         Top             =   8640
         Width           =   1275
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Base"
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
         Height          =   285
         Index           =   1
         Left            =   10620
         TabIndex        =   10
         Top             =   8100
         Width           =   870
      End
      Begin VB.Label lblBase 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   11835
         TabIndex        =   9
         Top             =   8100
         Width           =   2085
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "IVA"
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
         Height          =   285
         Index           =   0
         Left            =   10620
         TabIndex        =   8
         Top             =   8370
         Width           =   645
      End
      Begin VB.Label lblIVA 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   11835
         TabIndex        =   7
         Top             =   8370
         Width           =   2085
      End
      Begin VB.Label lbltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   11835
         TabIndex        =   6
         Top             =   8910
         Width           =   2085
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
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
         Height          =   285
         Index           =   2
         Left            =   10620
         TabIndex        =   5
         Top             =   8910
         Width           =   825
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00FFFFFF&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   375
         Left            =   90
         Top             =   180
         Width           =   15165
      End
   End
   Begin MSFlexGridLib.MSFlexGrid glista 
      Height          =   10980
      Left            =   45
      TabIndex        =   2
      Top             =   540
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   19368
      _Version        =   393216
      FixedRows       =   0
      FixedCols       =   0
      BackColor       =   16777215
      BackColorFixed  =   14737632
      BackColorSel    =   8553090
      BackColorBkg    =   12632256
      HighLight       =   2
      MergeCells      =   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   14400
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   11610
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "INFORME DE TESORERIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   180
      TabIndex        =   1
      Top             =   0
      Width           =   4725
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   15705
   End
End
Attribute VB_Name = "frmTesoreria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const color1 As Long = &HC0E0FF
Const color2 As Long = &HE0E0E0
Const color3 As Long = &HC0FFFF

Const colFP_ID As Integer = 14
Const colTIPO As Integer = 15
Const colCABECERA As Integer = 16

Const numCols As Integer = 14
Private Enum COLS
    C_ID = 0
    C_PROVEEDOR = 1
    C_fecha = 2
    C_concepto = 3
    C_NUMERO = 4
    C_FAMILIA = 5
    C_SUBCUENTA = 6
    C_BASE = 7
    C_IVA_PORCENTAJE = 8
    C_IVA = 9
    C_RETENCION = 10
    C_total = 11
    C_FP = 12
    C_vencimiento = 13
    C_PAGO = 14
    C_TOBJETO = 15
    C_cOBJETO = 16
    C_IDPROVEEDOR = 17
'M1335-I
    C_CUENTA = 18
'M1335-F
    C_ENVIADA = 19
End Enum
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    frmProveedores.Visible = False
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 10
    Me.Left = 10
    Me.Height = Screen.Height - 2500
'    fdesde = "01/01/" & Year(Date)
'    fhasta = "31/12/" & Year(Date)
    
'    fdesde = Date - 90
'    fhasta = Date
    
'    cargarCombos
    permisos
    cabecera
    Dim fila As Integer
    fila = 0
    fila = cargar_lista_pdtesPago(fila)
    fila = cargar_lista_pagoPrevisto(fila)
    fila = cargar_lista_cobrosAirbus(fila, False)
    fila = cargar_lista_cobrosAirbus(fila, True)
    fila = cargar_lista_cobros(fila)
'    cargar_lista
End Sub
Private Sub cargarCombos()
End Sub
Private Sub cabecera()
    With glista
       .Clear
       .FixedCols = 0
       .Rows = 1
       .COLS = 14 + 3 ' columnas + ocultas
       .TextMatrix(0, 0) = ""
       .AllowUserResizing = flexResizeColumns
       .CellAlignment = 1
       .ColWidth(colFP_ID) = 0
       .ColWidth(colTIPO) = 0
       .ColWidth(colCABECERA) = 0
    End With
End Sub
Private Sub cabeceraProveedores()
    lista.ColumnHeaders.Clear
    With lista.ColumnHeaders
        .Add , , "Nº", 800, lvwColumnLeft
        .Add , , "Proveedor", 2400, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Concepto", 1900, lvwColumnCenter
        .Add , , "Numero", 1000, lvwColumnCenter
        .Add , , "Familia", 1, lvwColumnLeft
        .Add , , "Subcuenta", 1, lvwColumnLeft
        .Add , , "Base", 1050, lvwColumnRight
        .Add , , "Iva %", 1, lvwColumnCenter
        .Add , , "Iva", 1000, lvwColumnRight
        .Add , , "Retención", 1000, lvwColumnRight
        .Add , , "Total", 1050, lvwColumnRight
        .Add , , "Forma Pago", 0, lvwColumnCenter
        .Add , , "Fecha Vencimiento", 1050, lvwColumnCenter
        .Add , , "Fecha Pago", 0, lvwColumnCenter
        .Add , , "TOBJETO", 1, lvwColumnLeft
        .Add , , "COBJETO", 1, lvwColumnLeft
        .Add , , "ID_PROVEEDOR", 1, lvwColumnLeft
        'M1335-I
        .Add , , "CUENTA_BANCARIA", 1, lvwColumnLeft
        'M1335-F
        .Add , , "Env", 350, lvwColumnLeft
    End With
End Sub
Private Sub cabeceraFacturas()
    lista.ColumnHeaders.Clear
    With lista.ColumnHeaders
        .Add , , "NºDoc", 900, lvwColumnLeft
        .Add , , "Cliente", 2500, lvwColumnLeft
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Importe", 1100, lvwColumnRight
        .Add , , "Dto%", 500, lvwColumnCenter
        .Add , , "Base", 1100, lvwColumnRight
        .Add , , "IVA%", 500, lvwColumnRight
        .Add , , "Cuota I.V.A.", 1100, lvwColumnRight
        .Add , , "Total", 1200, lvwColumnRight
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "FACTURA_CONCEPTOS", 1, lvwColumnCenter
        .Add , , "PAGADO", 1, lvwColumnCenter
        .Add , , "TIPO_DOCUMENTO", 1, lvwColumnCenter
        .Add , , "FACTORIZADA", 1, lvwColumnCenter
        .Add , , "Pedido", 1500, lvwColumnLeft
        .Add , , "FP", 1, lvwColumnCenter
        .Add , , "Asiento", 700, lvwColumnCenter
        .Add , , "F.Vencim.", 1050, lvwColumnCenter
        .Add , , "Comentario", 1, lvwColumnLeft
        .Add , , "ClienteFactura", 1, lvwColumnLeft
        .Add , , "F.Cobro", 1050, lvwColumnCenter
    End With
End Sub

Private Sub permisos()
    If usuario.getPER_TESORERIA_FP = False Then
    End If
End Sub
Private Sub cargar_lista(FP_ID As Integer, tipo As Integer, PERIODO As String)
    Dim rs As New ADODB.Recordset
    Dim oPF As New clsProveedores_Facturas
    Dim ID As Long
   On Error GoTo cargar_lista_Error
    Me.MousePointer = 11
    ' TIPO
    ' 1: PENDIENTES PAGO
    ' 2: PAGO PREVISTO
    Set rs = oPF.ListadoTesoreria(FP_ID, tipo, PERIODO)
    Dim BASE As Currency
    Dim IVA As Currency
    Dim retencion As Currency
    Dim total As Currency
    BASE = 0
    IVA = 0
    retencion = 0
    total = 0
    lista.ListItems.Clear
'    lblsubtitulo = "Se han detectado " & rs.RecordCount & " registros."
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "000000")) ' ID
           .SubItems(COLS.C_PROVEEDOR) = rs(17)
           .SubItems(COLS.C_IDPROVEEDOR) = rs(18)
            .SubItems(COLS.C_fecha) = Format(rs(1), "dd/mm/yyyy")  ' Fecha
            If Not IsNull(rs(2)) Then
                .SubItems(COLS.C_concepto) = rs(2)  ' Concepto
            End If
            If Not IsNull(rs(3)) Then
                .SubItems(COLS.C_NUMERO) = rs(3)  ' Numero
            End If
            If Not IsNull(rs(4)) Then
                .SubItems(COLS.C_FAMILIA) = rs(4)  ' Familia
            End If
            If Not IsNull(rs(5)) Then
                .SubItems(COLS.C_SUBCUENTA) = rs(5)  ' Subcuenta
            End If
            .SubItems(COLS.C_BASE) = Format(rs(6), "currency")  ' BI
            .SubItems(COLS.C_IVA_PORCENTAJE) = rs(7)  ' IVA PORCENTAJE
            .SubItems(COLS.C_IVA) = Format(rs(8), "currency")  ' IVA
            .SubItems(COLS.C_total) = Format(rs(9), "currency")  ' TOTAL
            BASE = BASE + rs(6)
            IVA = IVA + rs(8)
            retencion = retencion + rs(16)
            total = total + rs(9)
            If Not IsNull(rs(10)) Then
                .SubItems(COLS.C_FP) = rs(10)  ' FP
            End If
            If Not IsNull(rs(11)) Then
                .SubItems(COLS.C_vencimiento) = rs(11)  ' F.Vencimiento
            End If
            If Not IsNull(rs(13)) Then
                .SubItems(COLS.C_TOBJETO) = rs(13)  ' Tobjeto
            End If
            If Not IsNull(rs(14)) Then
                .SubItems(COLS.C_cOBJETO) = rs(14)  ' Cobjeto
            End If
            If Not IsNull(rs(15)) Then
                .Checked = True
            Else
                .Checked = False
            End If
            If Not IsNull(rs(16)) Then
                .SubItems(COLS.C_RETENCION) = Format(rs(16), "currency") ' RETENCION
            End If
            If Not IsNull(rs(12)) Then
                .SubItems(COLS.C_PAGO) = rs(12)
            End If
            'M1335-I
            If Not IsNull(rs(19)) Then
                .SubItems(COLS.C_CUENTA) = rs(19)
            End If
            'M1335-F
            If rs(20) = 0 Then
                .SubItems(COLS.C_ENVIADA) = "N"
            Else
                .SubItems(COLS.C_ENVIADA) = "S"
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    lblBase = Format(BASE, "currency")
    lblIVA = Format(IVA, "currency")
    lblRetencion = Format(retencion, "currency")
    lbltotal = Format(total, "currency")
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmTesoreria"

End Sub

Private Sub Form_Resize()
    glista.Width = Me.ScaleWidth
    glista.Left = 10
    
    cmdcancel.top = Me.ScaleHeight - 60 - cmdcancel.Height
    cmdcancel.Left = Me.ScaleWidth - 60 - cmdcancel.Width
    
    glista.Height = Me.ScaleHeight - glista.top - cmdcancel.Height - 120
    
'    fraDatos.Width = Me.ScaleWidth
'    cmdBuscar.Left = fraDatos.Width - cmdBuscar.Width - 100
'    frmTipo.Left = cmdBuscar.Left - frmTipo.Width - 100
    
    fondo.Width = Me.ScaleWidth
'    imagen.Left = fondo.Width - imagen.Width - 100
'    frmLeyenda.top = Me.Height - frmLeyenda.Height - 600
    
'    glista.ColWidth(1) = glista.Width * 0.4

'    frmUsuarios.Left = (Screen.Width / 2) - (frmUsuarios.Width / 2)
'    frmUsuarios.top = (Screen.Height / 2) - (frmUsuarios.Height / 2)

End Sub

Private Sub glista_DblClick()
'    MsgBox "TIPO : " & glista.TextMatrix(glista.Row, colTIPO) & " FP_ID : " & glista.TextMatrix(glista.Row, colFP_ID) & " CABECERA : " & glista.TextMatrix(glista.TextMatrix(glista.Row, colCABECERA), glista.col)
    Dim texto As String
    Dim tipo As Integer
    Dim FP_ID As Integer
    Dim cabecera As String
    texto = glista.TextMatrix(glista.Row, glista.Col)
    tipo = glista.TextMatrix(glista.Row, colTIPO)
    FP_ID = glista.TextMatrix(glista.Row, colFP_ID)
    cabecera = glista.TextMatrix(glista.TextMatrix(glista.Row, colCABECERA), glista.Col)
    If texto <> "" Then
        If (tipo = 1 Or tipo = 2) Then
            If tipo = 1 Then
                lblProveedor.Caption = "FACTURAS PENDIENTES DE PAGO, FORMA_PAGO : " & glista.TextMatrix(glista.Row, 0) & ", PERIODO : " & glista.TextMatrix(glista.TextMatrix(glista.Row, colCABECERA), glista.Col)
            Else
                lblProveedor.Caption = "FACTURAS PAGO PREVISTO, FORMA_PAGO : " & glista.TextMatrix(glista.Row, 0) & ", PERIODO : " & glista.TextMatrix(glista.TextMatrix(glista.Row, colCABECERA), glista.Col)
            End If
            frmProveedores.Left = (Me.Width / 2) - (frmProveedores.Width / 2)
            frmProveedores.top = (Me.Height / 2) - (frmProveedores.Height / 2)
            cabeceraProveedores
            txtListaTipo = "1"
            frmProveedores.Visible = True
            cargar_lista FP_ID, tipo, cabecera
        Else
            Select Case tipo
            Case 3
                lblProveedor.Caption = "PENDIENTES DE COBRO AIRBUS (SIN PEDIDO), FORMA_PAGO : " & glista.TextMatrix(glista.Row, 0) & ", PERIODO : " & glista.TextMatrix(glista.TextMatrix(glista.Row, colCABECERA), glista.Col)
            Case 4
                lblProveedor.Caption = "PENDIENTES DE COBRO AIRBUS (CON PEDIDO), FORMA_PAGO : " & glista.TextMatrix(glista.Row, 0) & ", PERIODO : " & glista.TextMatrix(glista.TextMatrix(glista.Row, colCABECERA), glista.Col)
            Case 5
                lblProveedor.Caption = "PENDIENTES DE COBRO OTROS, FORMA_PAGO : " & glista.TextMatrix(glista.Row, 0) & ", PERIODO : " & glista.TextMatrix(glista.TextMatrix(glista.Row, colCABECERA), glista.Col)
            End Select
            frmProveedores.Left = (Me.Width / 2) - (frmProveedores.Width / 2)
            frmProveedores.top = (Me.Height / 2) - (frmProveedores.Height / 2)
            cabeceraFacturas
            txtListaTipo = "2"
            frmProveedores.Visible = True
            cargarListaFacturas FP_ID, tipo, cabecera
        End If
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

Private Sub lista_DblClick()
    If lista.ListItems.Count = 0 Then Exit Sub
    If txtListaTipo = "1" Then
        With frmProveedores_Facturas
            .PK = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_IDPROVEEDOR)
            .PK_FACTURA_ID = lista.ListItems(lista.selectedItem.Index).Text
            .TOBJETO = 0
            .COBJETO = 0
            .Show 1
        End With
    Else
        gdoc = lista.ListItems(lista.selectedItem.Index).SubItems(9)
        frmListadoDocPago.Show
    End If
End Sub
Private Function colorearLinea(fila As Integer, color As Long, size As Integer, bold As Boolean, Col As Integer, alineacion As Integer)
    Dim i As Integer
    glista.Row = fila
    If Col < 0 Then
        For i = 0 To glista.COLS - 1
            glista.Col = i
            glista.CellBackColor = color
            glista.CellFontBold = bold
            glista.CellFontSize = size
            glista.CellAlignment = alineacion
        Next
    Else
            glista.Col = Col
            glista.CellBackColor = color
            glista.CellFontBold = bold
            glista.CellFontSize = size
            glista.CellAlignment = alineacion
    End If
End Function
Private Function cargar_lista_pdtesPago(fila As Integer) As Integer
    Dim i As Integer
    Dim fecha As Date
    Dim filaInicio As Integer
    filaInicio = fila
    With glista
'       .Rows = .Rows + 1
       For i = 0 To 13
        .TextMatrix(fila, i) = "PENDIENTE DE PAGO"
        .MergeCol(i) = True
       Next
       colorearLinea fila, color1, 12, True, -1, 1
       .MergeRow(fila) = True
       fila = fila + 1
       .Rows = .Rows + 1
       .TextMatrix(fila, 1) = "VENCIDAS"
       .ColWidth(0) = 2500
       .ColWidth(1) = 1700
       fecha = Format(Date, "mm-yyyy")
       For i = 2 To 13
        .TextMatrix(fila, i) = Format(fecha, "mm-yyyy")
        .ColWidth(i) = 1700
        fecha = DateAdd("m", 1, fecha)
       Next
       colorearLinea fila, color2, 8, True, -1, 3
    End With
    Dim rs As ADODB.Recordset
    Dim oT As New clsTesoreria_prevision
    Set rs = oT.ListadoPdtesPago()
    Dim fp As Integer
    Dim j As Integer
    Dim Col As Integer
    fila = fila + 1
    glista.Rows = glista.Rows + 1
    If rs.RecordCount > 0 Then
        Do
            fp = rs(2)
            With glista
                .TextMatrix(fila, 0) = rs(3)
                .TextMatrix(fila, colFP_ID) = rs(2)
                .TextMatrix(fila, colTIPO) = 1
                .TextMatrix(fila, colCABECERA) = filaInicio + 1
                colorearLinea fila, color2, 8, True, 0, 1
                If Trim(rs(1)) = "" Then
                    .TextMatrix(fila, 1) = moneda(rs(4))
                Else
                    Col = 0
                    For j = 0 To .COLS - 1
                       If .TextMatrix(1, j) = rs(1) Then
                            Col = j
                            Exit For
                       End If
                    Next
                    If Col <> 0 Then
                        .TextMatrix(fila, Col) = moneda(rs(4))
                    End If
                End If
            End With
            rs.MoveNext
            If Not rs.EOF Then
                If fp <> rs(2) Then
                    fila = fila + 1
                    glista.Rows = glista.Rows + 1
                End If
            End If
        Loop Until rs.EOF
        ' TOTALIZADOR
        fila = fila + 1
        glista.Rows = glista.Rows + 1
        glista.TextMatrix(fila, 0) = "Total"
        Dim t As Currency
        For i = 1 To numCols - 1
            t = 0
            For j = (filaInicio + 2) To fila - 1
                If Trim(glista.TextMatrix(j, i)) <> "" Then
                    t = t + glista.TextMatrix(j, i)
                End If
            Next
            If t <> 0 Then
                glista.TextMatrix(fila, i) = moneda(CStr(t))
            End If
        Next
        colorearLinea fila, color3, 10, True, -1, flexAlignRightCenter
        
        fila = fila + 1
    End If
    cargar_lista_pdtesPago = fila
'    Dim rs As New ADODB.Recordset
'    Dim oPF As New clsProveedores_Facturas
'    Dim ID As Long
'    Me.MousePointer = 11
'    Set rs = oPF.ListadoPdtesPago(fdesde, fhasta)
'    Dim total As Currency
'    total = 0
'    listaPdtePago.ListItems.Clear
'    If rs.RecordCount <> 0 Then
'        Do
'           With listaPdtePago.ListItems.Add(, , rs(0)) ' ID
'           .SubItems(1) = rs(1) ' FP
'           .SubItems(2) = moneda(rs(2)) ' Total
'           End With
'           total = total + rs(2)
'           rs.MoveNext
'        Loop Until rs.EOF
'    End If
'    lbltotalPdte = Format(total, "currency")
'    Me.MousePointer = 0


End Function

Private Function cargar_lista_pagoPrevisto(fila As Integer) As Integer
    Dim i As Integer
    Dim fecha As Date
    Dim filaInicio As Integer
    filaInicio = fila
    With glista
       .Rows = .Rows + 1
       For i = 0 To 13
        .TextMatrix(fila, i) = "PAGO PREVISTO"
        .MergeCol(i) = True
       Next
       colorearLinea fila, color1, 12, True, -1, 1
       .MergeRow(fila) = True
       fila = fila + 1
       .Rows = .Rows + 1
       .TextMatrix(fila, 1) = "VENCIDAS"
       .ColWidth(0) = 2500
       .ColWidth(1) = 1700
       fecha = Format(Date, "mm-yyyy")
       For i = 2 To 13
        .TextMatrix(fila, i) = Format(fecha, "mm-yyyy")
        .ColWidth(i) = 1700
        fecha = DateAdd("m", 1, fecha)
       Next
       colorearLinea fila, color2, 8, True, -1, 3
    End With
    Dim rs As ADODB.Recordset
    Dim oT As New clsTesoreria_prevision
    Set rs = oT.ListadoPagoPrevisto()
    Dim fp As Integer
    Dim j As Integer
    Dim Col As Integer
    fila = fila + 1
    glista.Rows = glista.Rows + 1
    If rs.RecordCount > 0 Then
        Do
            fp = rs(2)
            With glista
                .TextMatrix(fila, 0) = rs(3)
                .TextMatrix(fila, colFP_ID) = rs(2)
                .TextMatrix(fila, colTIPO) = 2
                .TextMatrix(fila, colCABECERA) = filaInicio + 1
                colorearLinea fila, color2, 8, True, 0, 1
                If Trim(rs(1)) = "" Then
                    .TextMatrix(fila, 1) = moneda(rs(4))
                Else
                    Col = 0
                    For j = 0 To .COLS - 1
                       If .TextMatrix(1, j) = rs(1) Then
                            Col = j
                            Exit For
                       End If
                    Next
                    If Col <> 0 Then
                        .TextMatrix(fila, Col) = moneda(rs(4))
                    End If
                End If
            End With
            rs.MoveNext
            If Not rs.EOF Then
                If fp <> rs(2) Then
                    fila = fila + 1
                    glista.Rows = glista.Rows + 1
                End If
            End If
        Loop Until rs.EOF
        ' TOTALIZADOR
        fila = fila + 1
        glista.Rows = glista.Rows + 1
        glista.TextMatrix(fila, 0) = "Total"
        Dim t As Currency
        For i = 1 To numCols - 1
            t = 0
            For j = (filaInicio + 2) To fila - 1
                If Trim(glista.TextMatrix(j, i)) <> "" Then
                    t = t + glista.TextMatrix(j, i)
                End If
            Next
            If t <> 0 Then
                glista.TextMatrix(fila, i) = moneda(CStr(t))
            End If
        Next
        colorearLinea fila, color3, 10, True, -1, flexAlignRightCenter
        fila = fila + 1
    End If
    cargar_lista_pagoPrevisto = fila
'    Dim rs As New ADODB.Recordset
'    Dim oPF As New clsProveedores_Facturas
'    Dim ID As Long
'   On Error GoTo cargar_lista_pagoPrevisto_Error
'
'    Me.MousePointer = 11
'    Set rs = oPF.ListadoPagoPrevisto(fdesde, fhasta)
'    Dim total As Currency
'    total = 0
'    listaPagoPrevisto.ListItems.Clear
''    lblsubtitulo = "Se han detectado " & rs.RecordCount & " registros."
'    If rs.RecordCount <> 0 Then
'        Do
'           With listaPagoPrevisto.ListItems.Add(, , rs(0)) ' ID
'           .SubItems(1) = rs(1) ' FP
'           .SubItems(2) = moneda(rs(2)) ' Total
'           End With
'           total = total + rs(2)
'           rs.MoveNext
'        Loop Until rs.EOF
'    End If
'    lbltotalPagoPrevisto = Format(total, "currency")
'    Me.MousePointer = 0
'
'   On Error GoTo 0
'   Exit Sub
'
'cargar_lista_pagoPrevisto_Error:
'
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista_pagoPrevisto of Formulario frmTesoreria"
End Function

Private Function cargar_lista_cobros(fila As Integer) As Integer
    Dim i As Integer
    Dim fecha As Date
    Dim filaInicio As Integer
    filaInicio = fila
    With glista
       .Rows = .Rows + 1
       For i = 0 To 13
        .TextMatrix(fila, i) = "PENDIENTES DE COBRO (OTROS)"
        .MergeCol(i) = True
       Next
       colorearLinea fila, color1, 12, True, -1, 1
       .MergeRow(fila) = True
       fila = fila + 1
       .Rows = .Rows + 1
       .TextMatrix(fila, 1) = "VENCIDAS"
       .ColWidth(0) = 2500
       .ColWidth(1) = 1700
       fecha = Format(Date, "mm-yyyy")
       For i = 2 To 13
        .TextMatrix(fila, i) = Format(fecha, "mm-yyyy")
        .ColWidth(i) = 1700
        fecha = DateAdd("m", 1, fecha)
       Next
       colorearLinea fila, color2, 8, True, -1, 3
    End With
    Dim rs As ADODB.Recordset
    Dim oT As New clsTesoreria_prevision
    Set rs = oT.ListadoCobro()
    Dim fp As Integer
    Dim j As Integer
    Dim Col As Integer
    fila = fila + 1
    glista.Rows = glista.Rows + 1
    If rs.RecordCount > 0 Then
        Do
            fp = rs(2)
            With glista
                .TextMatrix(fila, 0) = rs(3)
                .TextMatrix(fila, colFP_ID) = rs(2)
                .TextMatrix(fila, colTIPO) = 5
                .TextMatrix(fila, colCABECERA) = filaInicio + 1
                colorearLinea fila, color2, 8, True, 0, 1
                If Trim(rs(1)) = "" Then
                    .TextMatrix(fila, 1) = moneda(rs(4))
                Else
                    Col = 0
                    For j = 0 To .COLS - 1
                       If .TextMatrix(1, j) = rs(1) Then
                            Col = j
                            Exit For
                       End If
                    Next
                    If Col <> 0 Then
                        .TextMatrix(fila, Col) = moneda(rs(4))
                    End If
                End If
            End With
            rs.MoveNext
            If Not rs.EOF Then
                If fp <> rs(2) Then
                    fila = fila + 1
                    glista.Rows = glista.Rows + 1
                End If
            End If
        Loop Until rs.EOF
        ' TOTALIZADOR
        fila = fila + 1
        glista.Rows = glista.Rows + 1
        glista.TextMatrix(fila, 0) = "Total"
        Dim t As Currency
        For i = 1 To numCols - 1
            t = 0
            For j = (filaInicio + 2) To fila - 1
                If Trim(glista.TextMatrix(j, i)) <> "" Then
                    t = t + glista.TextMatrix(j, i)
                End If
            Next
            If t <> 0 Then
                glista.TextMatrix(fila, i) = moneda(CStr(t))
            End If
        Next
        colorearLinea fila, color3, 10, True, -1, flexAlignRightCenter
        fila = fila + 1
    End If
    cargar_lista_cobros = fila

'    Dim rs As New ADODB.Recordset
'    Dim oT As New clsTesoreria_prevision
'    Dim ID As Long
'
'    Me.MousePointer = 11
'    Set rs = oT.ListadoCobro(fdesde, fhasta)
'    Dim total As Currency
'    total = 0
'    listaCobros.ListItems.Clear
'    If rs.RecordCount <> 0 Then
'        Do
'           With listaCobros.ListItems.Add(, , rs(0)) ' ID
'           .SubItems(1) = rs(1) ' FP
'           .SubItems(2) = moneda(rs(2)) ' Total
'           End With
'           total = total + rs(2)
'           rs.MoveNext
'        Loop Until rs.EOF
'    End If
''    lbltotalPagoPrevisto = Format(total, "currency")
'    Me.MousePointer = 0
End Function
Private Function cargar_lista_cobrosAirbus(fila As Integer, pedido As Boolean) As Integer
    Dim i As Integer
    Dim fecha As Date
    Dim filaInicio As Integer
    filaInicio = fila
    With glista
       .Rows = .Rows + 1
       For i = 0 To 13
        If pedido = False Then
            .TextMatrix(fila, i) = "PENDIENTES DE COBRO AIRBUS (SIN PEDIDO)"
        Else
            .TextMatrix(fila, i) = "PENDIENTES DE COBRO AIRBUS (CON PEDIDO)"
        End If
        .MergeCol(i) = True
       Next
       colorearLinea fila, color1, 12, True, -1, 1
       .MergeRow(fila) = True
       fila = fila + 1
       .Rows = .Rows + 1
       .TextMatrix(fila, 1) = "VENCIDAS"
       .ColWidth(0) = 2500
       .ColWidth(1) = 1700
       fecha = Format(Date, "mm-yyyy")
       For i = 2 To 13
        .TextMatrix(fila, i) = Format(fecha, "mm-yyyy")
        .ColWidth(i) = 1700
        fecha = DateAdd("m", 1, fecha)
       Next
       colorearLinea fila, color2, 8, True, -1, 3
    End With
    Dim rs As ADODB.Recordset
    Dim oT As New clsTesoreria_prevision
    Set rs = oT.ListadoCobroAirbus(pedido)
    Dim fp As Integer
    Dim j As Integer
    Dim Col As Integer
    fila = fila + 1
    glista.Rows = glista.Rows + 1
    If rs.RecordCount > 0 Then
        Do
            fp = rs(2)
            With glista
                .TextMatrix(fila, 0) = rs(3)
                .TextMatrix(fila, colFP_ID) = rs(2)
                If pedido = False Then
                    .TextMatrix(fila, colTIPO) = 3
                Else
                    .TextMatrix(fila, colTIPO) = 4
                End If
                .TextMatrix(fila, colCABECERA) = filaInicio + 1
                colorearLinea fila, color2, 8, True, 0, 1
                If Trim(rs(1)) = "" Then
                    .TextMatrix(fila, 1) = moneda(rs(4))
                Else
                    Col = 0
                    For j = 0 To .COLS - 1
                       If .TextMatrix(1, j) = rs(1) Then
                            Col = j
                            Exit For
                       End If
                    Next
                    If Col <> 0 Then
                        .TextMatrix(fila, Col) = moneda(rs(4))
                    End If
                End If
            End With
            rs.MoveNext
            If Not rs.EOF Then
                If fp <> rs(2) Then
                    fila = fila + 1
                    glista.Rows = glista.Rows + 1
                End If
            End If
        Loop Until rs.EOF
        ' TOTALIZADOR
        fila = fila + 1
        glista.Rows = glista.Rows + 1
        glista.TextMatrix(fila, 0) = "Total"
        Dim t As Currency
        For i = 1 To numCols - 1
            t = 0
            For j = (filaInicio + 2) To fila - 1
                If Trim(glista.TextMatrix(j, i)) <> "" Then
                    t = t + glista.TextMatrix(j, i)
                End If
            Next
            If t <> 0 Then
                glista.TextMatrix(fila, i) = moneda(CStr(t))
            End If
        Next
        colorearLinea fila, color3, 10, True, -1, flexAlignRightCenter
        fila = fila + 1
    End If
    cargar_lista_cobrosAirbus = fila
End Function

Private Sub cargarListaFacturas(FP_ID As Integer, tipo As Integer, PERIODO As String)
    Dim IMPORTE As Currency
    Dim BASE As Currency
    Dim IVA As Currency
    
    Dim T_BASE As Currency
    Dim T_IVA As Currency
    Dim T_TOTAL As Currency
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    Dim oDoc As New clsDocs_pago
    Me.MousePointer = 11
    Set rs = oDoc.DocumentosTesoreria(FP_ID, tipo, PERIODO)
    If rs.RecordCount <> 0 Then
        Dim NUMERO As String
        Do
                Select Case rs(6)
                Case 1
                    NUMERO = "A-" & Format(rs(1), "0000")
                Case 2
                    NUMERO = "F-" & Format(rs(1), "0000")
                Case 3
                    NUMERO = "B-" & Format(rs(1), "0000")
                Case Else
                    NUMERO = Format(rs(1), "0000")
                End Select
            
                    With lista.ListItems.Add(, , NUMERO)
                        .SubItems(1) = rs.Fields(2)
                        .SubItems(2) = rs.Fields(3)
                        .SubItems(9) = rs.Fields(0)
                        IMPORTE = Format(rs(8), "currency")
                        If IsNull(rs.Fields("descuento")) Or rs.Fields("descuento") = "0" Then
                            BASE = Format(IMPORTE, "0.00")
                        Else
                            BASE = Format(IMPORTE - ((IMPORTE * rs.Fields("descuento")) / 100), "0.00")
                        End If
                        IVA = Format((BASE * rs.Fields("iva")) / 100, "0.00")
                        .SubItems(3) = Format(IMPORTE, "currency")
                        .SubItems(4) = Format(rs.Fields("descuento"), "Standard")
                        .SubItems(5) = Format(BASE, "currency")
                        .SubItems(6) = rs.Fields("iva")
                        .SubItems(7) = Format(IVA, "currency")
                        .SubItems(8) = Format(BASE + IVA, "currency")
                        .SubItems(10) = rs(9)
                        .SubItems(11) = rs(7)
                        .SubItems(12) = rs(6)
                        .SubItems(13) = rs(11) ' Factorizada
                        .SubItems(14) = rs(12) ' PEdido
                        .SubItems(15) = rs(13) ' Forma de pago
                        .SubItems(16) = rs(16) ' Asiento
                        If rs(13) <> 0 Then
                            .SubItems(17) = Format(CDate(rs(3)) + rs(17), "dd/mm/yyyy") ' F.Vencimiento
                        Else
                            .SubItems(17) = ""
                        End If
                        .SubItems(18) = rs(18)
                        .SubItems(19) = rs(19)
                        If IsNull(rs(20)) Then
                            .SubItems(20) = ""
                        Else
                            .SubItems(20) = rs(20)
                        End If
                        T_BASE = T_BASE + CCur(.SubItems(5))
                        T_IVA = T_IVA + CCur(.SubItems(7))
                        T_TOTAL = T_TOTAL + CCur(.SubItems(8))
                        ' 14 CLIENTE_ID, 15 CLIENTE_ID_FACTURA
'                        If rs(14) <> rs(15) Then
'                            colorear lista.ListItems.Count, vbRed
'                        End If
                    End With
            rs.MoveNext
        Loop Until rs.EOF
        lblBase = moneda(CStr(T_BASE))
        lblIVA = moneda(CStr(T_IVA))
        lbltotal = moneda(CStr(T_TOTAL))
    End If
    Me.MousePointer = 0
    Set oDoc = Nothing
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al buscar los documentos del cliente.", vbCritical, Err.Description
End Sub

