VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmRemesas_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Detalle de Remesas de Pago"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14910
   Icon            =   "frmRemesas_Detalle.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8565
   ScaleWidth      =   14910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Exportar Excel"
      Height          =   960
      Left            =   10980
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Crear nuevo paquete"
      Top             =   7560
      Width           =   1350
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   960
      Left            =   12375
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7560
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la Remesa"
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
      Height          =   1185
      Left            =   0
      TabIndex        =   7
      Top             =   630
      Width           =   14865
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   4815
         TabIndex        =   1
         Top             =   720
         Width           =   6045
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   12240
         TabIndex        =   15
         Top             =   315
         Width           =   2355
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   8955
         TabIndex        =   13
         Top             =   315
         Width           =   1905
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   4815
         TabIndex        =   11
         Top             =   315
         Width           =   2670
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1035
         TabIndex        =   2
         Top             =   315
         Width           =   2535
      End
      Begin MSDataListLib.DataCombo cmbBanco 
         Bindings        =   "frmRemesas_Detalle.frx":0442
         Height          =   315
         Left            =   1035
         TabIndex        =   0
         Top             =   720
         Width           =   2550
         _ExtentX        =   4498
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
      Begin MSDataListLib.DataCombo cmbEstado 
         Bindings        =   "frmRemesas_Detalle.frx":0488
         Height          =   315
         Left            =   12240
         TabIndex        =   21
         Top             =   720
         Width           =   2370
         _ExtentX        =   4180
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
         Caption         =   "Estado"
         Height          =   195
         Index           =   0
         Left            =   11385
         TabIndex        =   22
         Top             =   765
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Banco"
         Height          =   195
         Index           =   22
         Left            =   90
         TabIndex        =   19
         Top             =   765
         Width           =   465
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Index           =   4
         Left            =   3870
         TabIndex        =   17
         Top             =   765
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe"
         Height          =   240
         Index           =   3
         Left            =   11385
         TabIndex        =   16
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Documentos"
         Height          =   240
         Index           =   2
         Left            =   7740
         TabIndex        =   14
         Top             =   360
         Width           =   1365
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Creación"
         Height          =   240
         Index           =   0
         Left            =   3870
         TabIndex        =   12
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Remesa"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   8
         Top             =   360
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC - Salir"
      Height          =   960
      Left            =   13635
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Salir"
      Top             =   7560
      Width           =   1215
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir Facturas"
      Height          =   960
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Crear nuevo paquete"
      Top             =   7560
      Width           =   1350
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar Factura"
      Height          =   960
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Eliminar paquete seleccionado"
      Top             =   7560
      Width           =   1350
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5670
      Left            =   0
      TabIndex        =   3
      Top             =   1845
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   10001
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
      Caption         =   "Detalle de Remesas de Pago"
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
      TabIndex        =   10
      Top             =   45
      Width           =   3090
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Remesas de Pago"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   360
      Width           =   2070
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   14940
   End
End
Attribute VB_Name = "frmRemesas_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Public MODO As String

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


Private Sub cabecera()
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
        .Add , , "Forma Pago", 1200, lvwColumnCenter
        .Add , , "Fecha Vencimiento", 1050, lvwColumnCenter
        .Add , , "Fecha Pago", 1050, lvwColumnCenter
        .Add , , "TOBJETO", 1, lvwColumnLeft
        .Add , , "COBJETO", 1, lvwColumnLeft
        .Add , , "ID_PROVEEDOR", 1, lvwColumnLeft
        .Add , , "CUENTA_BANCARIA", 1, lvwColumnLeft
        .Add , , "Env", 1, lvwColumnLeft
    End With
End Sub

Private Sub cmdExcel_Click()
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    
   On Error GoTo Command1_Click_Error

    Me.MousePointer = 11
    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Add
    Set XLS = XLW.Worksheets(1)
    XLW.Worksheets(3).Delete
    XLW.Worksheets(2).Delete
    XLW.Worksheets(1).Name = "Listado de Remesa"
    XLS.Range("1:1").HorizontalAlignment = xlCenter
    XLS.Range("1:1").VerticalAlignment = xlCenter
    XLS.Range("1:1").RowHeight = 20
    XLS.Range("1:1").WrapText = True
    'Cabecera
    Dim i As Integer
    Dim fila As Integer
    For i = 1 To 15
        XLS.Cells(1, i) = lista.ColumnHeaders(i).Text
    Next
    fila = 2
    ' Datos
    For i = 1 To lista.ListItems.Count
        XLS.Range(XLS.Cells(fila, 8), XLS.Cells(fila, 8)).NumberFormat = "0.00"
        XLS.Range(XLS.Cells(fila, 9), XLS.Cells(fila, 9)).NumberFormat = "0.00"
        XLS.Range(XLS.Cells(fila, 10), XLS.Cells(fila, 10)).NumberFormat = "0.00"
        XLS.Range(XLS.Cells(fila, 11), XLS.Cells(fila, 11)).NumberFormat = "0.00"
        XLS.Range(XLS.Cells(fila, 12), XLS.Cells(fila, 12)).NumberFormat = "0.00"
        
        XLS.Cells(fila, 1) = lista.ListItems(i).Text ' Num
        XLS.Cells(fila, 2) = lista.ListItems(i).SubItems(1) ' Proveedor
        XLS.Cells(fila, 3) = Format(lista.ListItems(i).SubItems(2), "yyyy-mm-dd") ' Fecha
        XLS.Cells(fila, 4) = lista.ListItems(i).SubItems(3) ' Concepto
        XLS.Cells(fila, 5) = lista.ListItems(i).SubItems(4) ' Numero
        XLS.Cells(fila, 6) = lista.ListItems(i).SubItems(5) ' Familia
        XLS.Cells(fila, 7) = lista.ListItems(i).SubItems(6) ' Subuenta
        XLS.Cells(fila, 8) = moneda_bd(lista.ListItems(i).SubItems(7)) ' Base
        XLS.Cells(fila, 9) = moneda_bd(lista.ListItems(i).SubItems(8)) ' Iva  %
        XLS.Cells(fila, 10) = moneda_bd(lista.ListItems(i).SubItems(9)) ' Iva
        XLS.Cells(fila, 11) = moneda_bd(lista.ListItems(i).SubItems(10)) ' Retencion
        XLS.Cells(fila, 12) = moneda_bd(lista.ListItems(i).SubItems(11)) ' Total
        XLS.Cells(fila, 13) = lista.ListItems(i).SubItems(12) ' FP
        XLS.Cells(fila, 14) = Format(lista.ListItems(i).SubItems(13), "yyyy-mm-dd") ' F.V
        XLS.Cells(fila, 15) = Format(lista.ListItems(i).SubItems(14), "yyyy-mm-dd") ' FP
        fila = fila + 1
    Next
    Me.MousePointer = 0
    XLA.Visible = True

   On Error GoTo 0
   Exit Sub

Command1_Click_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command1_Click of Formulario frmRemesas_Detalle"

End Sub

Private Sub Command1_Click()
   

End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If cmbBanco.Text = "" Then
        MsgBox "Debe indicar el banco.", vbCritical, App.Title
        cmbBanco.SetFocus
        Exit Sub
    End If
    Dim oBanco As New clsBancos
    oBanco.Carga cmbBanco.BoundText
    If oBanco.getSUBCUENTA = "" Then
        MsgBox "El banco no tiene informada la subcuenta.", vbCritical, App.Title
        Exit Sub
    End If
    Dim oRemesa As New clsRemesas
    Dim ID_REMESA As Long
    With oRemesa
        .setDESCRIPCION = txtDatos(4)
        .setTIPO_ID = REMESAS.REMESA_PAGO
        .setBANCO_ID = cmbBanco.BoundText
        If IsNumeric(txtDatos(2)) Then
            .setDOCUMENTOS = txtDatos(2)
        Else
            .setDOCUMENTOS = 0
        End If
        .setIMPORTE = moneda_bd(Replace(txtDatos(3), "€", ""))
        .setESTADO_ID = cmbEstado.BoundText
        If PK = 0 Then
            ID_REMESA = .Insertar
        Else
            .Modificar PK
            ID_REMESA = PK
        End If
    End With
    If ID_REMESA = 0 Then
        MsgBox "Error al insertar la remesa.", vbCritical, App.Title
        Exit Sub
    End If
    ' Documentos
    Dim oPF As New clsProveedores_Facturas
    Dim oRD As New clsRemesas_documentos
    oRD.Eliminar ID_REMESA
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        With oRD
            .setREMESA_ID = ID_REMESA
            .setFACTURA_ID = lista.ListItems(i).Text
            .Insertar
        End With
        oPF.marcarPagada CLng(lista.ListItems(i).Text), ID_REMESA, oBanco.getSUBCUENTA
    Next
    Set oRemesa = Nothing
    MsgBox "Remesa generada correctamente.", vbInformation, App.Title
    Unload Me

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmRemesas_Detalle"
End Sub


Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combo cmbBanco, New clsBancos
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbEstado, DECODIFICADORA.DECODIFICADORA_REMESAS_ESTADOS
    If PK = 0 Then
        lbltitulo = "Creación de Remesas de Pago"
        cmbEstado.BoundText = 0
    Else
        lbltitulo = "Modificación de Remesa de Pago"
        Dim oRemesa As New clsRemesas
        With oRemesa
            If .Carga(PK) Then
                cmbBanco.BoundText = .getBANCO_ID
                txtDatos(1) = .getNUMERO & "/" & oRemesa.getANNO
                txtDatos(0) = .getFECHA
                txtDatos(2) = .getDOCUMENTOS
                txtDatos(4) = .getDESCRIPCION
                cmbEstado.BoundText = .getESTADO_ID
                If .getESTADO_ID = 1 Then
                    MODO = "C"
                End If
            End If
        End With
        Dim lista As String
        lista = oRemesa.ListadoIds(PK)
        If Trim(lista) <> "" Then
            cargar_lista_ids lista
        End If
    End If
    If MODO = "C" Then
        cmdAnadir.Visible = False
        cmdEliminar.Visible = False
        cmdok.Visible = False
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
Private Sub cmdAnadir_Click()
    Dim oFP As New frmProveedores_Facturas_Listado_Modal
    oFP.Show 1
    If Trim(oFP.LISTA_SEL) <> "" Then
        cargar_lista_ids oFP.LISTA_SEL
    End If
    Set oFP = Nothing
    Unload oFP
End Sub
Private Sub cargar_lista_ids(Documentos As String)
    Dim rs As New ADODB.Recordset
    Dim oPF As New clsProveedores_Facturas
   On Error GoTo cargar_lista_ids_Error

    Me.MousePointer = 11
    Set rs = oPF.ListadoCompletoId(Documentos)
    Dim existe As Boolean
    Dim i As Integer
    If rs.RecordCount <> 0 Then
        Do
          existe = False
          For i = 1 To lista.ListItems.Count
           If CLng(lista.ListItems(i).Text) = CLng(rs(0)) Then
            existe = True
           End If
          Next
          ' Validar cuenta informada
          If rs(19) = "" Or rs(19) = "____-____-____-____-____-____" Then
            MsgBox "El proveedor " & rs(17) & " no tiene la cuenta bancaria informada.", vbCritical, App.Title
            existe = True
          End If
          If Not existe Then
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
                If Not IsNull(rs(16)) Then
                    .SubItems(COLS.C_RETENCION) = Format(rs(16), "currency") ' RETENCION
                End If
                If Not IsNull(rs(12)) Then
                    .SubItems(COLS.C_PAGO) = rs(12)
                End If
                If Not IsNull(rs(19)) Then
                    .SubItems(COLS.C_CUENTA) = rs(19)
                End If
                If rs(20) = 0 Then
                    .SubItems(COLS.C_ENVIADA) = "N"
                Else
                    .SubItems(COLS.C_ENVIADA) = "S"
                End If
               End With
            End If
            rs.MoveNext
         Loop Until rs.EOF
    End If
    calcular_totales
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cargar_lista_ids_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista_ids of Formulario frmRemesas_Detalle"
End Sub
Private Sub cmdEliminar_Click()
    If Not (lista.selectedItem Is Nothing) Then
        lista.ListItems.Remove lista.selectedItem.Index
        calcular_totales
    Else
        MsgBox "Debe seleccionar la factura que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub
Private Sub calcular_totales()
    Dim i As Integer
    Dim total As Single
    txtDatos(2) = lista.ListItems.Count
    For i = 1 To lista.ListItems.Count
        total = total + (Replace(Replace(lista.ListItems(i).SubItems(COLS.C_total), "€", ""), ".", ""))
    Next
    txtDatos(3) = moneda(CStr(total))
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count = 0 Then Exit Sub
    With frmProveedores_Facturas
        .PK = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_IDPROVEEDOR)
        .PK_FACTURA_ID = lista.ListItems(lista.selectedItem.Index).Text
        .TOBJETO = 0
        .COBJETO = 0
        .Show 1
    End With
End Sub
