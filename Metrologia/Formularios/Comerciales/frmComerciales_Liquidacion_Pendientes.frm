VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmComerciales_Liquidacion_Pendientes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas cobradas pendientes de Liquidar"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11190
   Icon            =   "frmComerciales_Liquidacion_Pendientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   11190
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCobrar2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar Factura Seleccionada"
      Height          =   885
      Left            =   7080
      Picture         =   "frmComerciales_Liquidacion_Pendientes.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8280
      Width           =   2835
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   885
      Index           =   0
      Left            =   1290
      Picture         =   "frmComerciales_Liquidacion_Pendientes.frx":1F14
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8280
      Width           =   1155
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   885
      Index           =   1
      Left            =   90
      Picture         =   "frmComerciales_Liquidacion_Pendientes.frx":27DE
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8280
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la Liquidación"
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
      Height          =   1125
      Left            =   60
      TabIndex        =   3
      Top             =   390
      Width           =   11055
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   4890
         MaxLength       =   75
         TabIndex        =   15
         Top             =   240
         Width           =   6060
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   2340
         TabIndex        =   10
         Top             =   660
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
         Left            =   4890
         TabIndex        =   11
         Top             =   660
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
      Begin MSComCtl2.DTPicker fliquidacion 
         Height          =   330
         Left            =   2340
         TabIndex        =   13
         Top             =   240
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   0
         Left            =   3930
         TabIndex        =   14
         Top             =   300
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   3930
         TabIndex        =   12
         Top             =   750
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Liquidación"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Top             =   270
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Período Liquidación. Desde "
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   4
         Top             =   750
         Width           =   2025
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   9960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3210
      Left            =   60
      TabIndex        =   0
      Top             =   5010
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5662
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
   Begin MSComctlLib.ListView listaLiq 
      Height          =   3090
      Left            =   60
      TabIndex        =   16
      Top             =   1500
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   5450
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
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturas Cobradas pendientes de Liquidar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   -30
      TabIndex        =   2
      Top             =   4650
      Width           =   11235
   End
   Begin VB.Label lbltit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturas Cobradas pendientes de Liquidar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11235
   End
End
Attribute VB_Name = "frmComerciales_Liquidacion_Pendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK_AGENTE As Long
Public PK_LIQUIDACION As Long
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

Private Sub cmdCobrar2_Click()
    If lista.ListItems.Count > 0 Then
        frmDocumento.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).Text
        frmDocumento.Show 1
        actualizar_lista
    End If

End Sub

Private Sub cmdMarcar_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = Index
    Next
End Sub

Private Sub fliquidacion_Change()
    txtdatos(1) = "LIQUIDACION FECHA : " & Format(fliquidacion, "dd-mm-yyyy")
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
    Dim ocomercial As New clsComercial
    ocomercial.Cargar PK_AGENTE
    lbltit = "Liquidación del Agente : " & ocomercial.getNOMBRE
    Me.Caption = lbltit
    cargar_lista_pendientes
    If PK_LIQUIDACION <> 0 Then
        cargar_liquidacion
    Else
        fliquidacion = Date
        fdesde = Date
        fhasta = Date
        txtdatos(1) = "LIQUIDACION FECHA : " & Format(Date, "dd-mm-yyyy")
    End If
End Sub
Private Sub cargar_lista_pendientes()
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim numero As String
    Dim anno As String
    tipo = " AND A.TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.factura
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT A.ID_DOCUMENTO, A.NUMERO,A.FECHA,D.NOMBRE,C.NOMBRE,A.TOTAL " & _
               "  FROM DOCUMENTOS A " & _
               "  LEFT JOIN OBRAS C ON A.OBRA_ID = C.ID_OBRA " & _
               "  LEFT JOIN CLIENTES D ON C.CLIENTE_ID = D.ID_CLIENTE " & _
               "  LEFT JOIN LIQUIDACION_DOCUMENTOS B ON A.ID_DOCUMENTO = B.DOCUMENTO_ID " & _
               " WHERE A.ESTADO_ID = " & ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_COBRADA & _
               "   AND A.ANULADO = 0 " & _
               "   AND B.DOCUMENTO_ID Is Null " & _
               tipo & _
               "   AND C.COMERCIAL_ID = " & PK_AGENTE
    
    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    Dim ocomercial As New clsComercial
    ocomercial.Cargar PK
    If rs.RecordCount >= 1 Then
        Dim total As Currency
        total = 0
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = Format(rs.Fields(1), "0000")
                .SubItems(2) = Format(rs.Fields(2), "dd-mm-yyyy")
                .SubItems(3) = rs.Fields(3)
                .SubItems(4) = rs(4)
                .SubItems(5) = moneda(rs(5))
                .SubItems(6) = moneda((rs(5) * ocomercial.getCOMISION) / 100)
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
    With lista.ColumnHeaders
        .Add , , "ID", 300, lvwColumnLeft
        .Add , , "NºFactura", 800, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Cliente", 3000, lvwColumnLeft
        .Add , , "Obra", 3000, lvwColumnLeft
        .Add , , "Importe", 1200, lvwColumnRight
        .Add , , "Comision", 1200, lvwColumnCenter
    End With
    With listaLiq.ColumnHeaders
        .Add , , "ID", 300, lvwColumnLeft
        .Add , , "NºFactura", 800, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Cliente", 3000, lvwColumnLeft
        .Add , , "Obra", 3000, lvwColumnLeft
        .Add , , "Importe", 1200, lvwColumnRight
        .Add , , "Comision", 1200, lvwColumnCenter
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
    cmdCobrar2_Click
End Sub

Private Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
   On Error GoTo actualizar_lista_Error

    consulta = "SELECT A.ID_DOCUMENTO, A.NUMERO,A.FECHA,D.NOMBRE,C.NOMBRE,A.TOTAL " & _
               "  FROM DOCUMENTOS A " & _
               "  LEFT JOIN OBRAS C ON A.OBRA_ID = C.ID_OBRA " & _
               "  LEFT JOIN CLIENTES D ON C.CLIENTE_ID = D.ID_CLIENTE " & _
               "  LEFT JOIN LIQUIDACION_DOCUMENTOS B ON A.ID_DOCUMENTO = B.DOCUMENTO_ID " & _
               " WHERE A.ID_DOCUMENTO = " & lista.ListItems(lista.SelectedItem.Index).Text

    Set rs = datos_bd(consulta)
    Dim ocomercial As New clsComercial
    ocomercial.Cargar PK
    If rs.RecordCount >= 1 Then
            With lista.ListItems(lista.SelectedItem.Index)
                .SubItems(1) = Format(rs.Fields(1), "0000")
                .SubItems(2) = Format(rs.Fields(2), "dd-mm-yyyy")
                .SubItems(3) = rs.Fields(3)
                .SubItems(4) = rs(4)
                .SubItems(5) = moneda(rs(5))
                .SubItems(6) = moneda((rs(5) * ocomercial.getCOMISION) / 100)
            End With
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

actualizar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_lista of Formulario frmComerciales_Liquidacion_Pendientes"
End Sub
Private Sub cargar_combos()
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    llenar_combo cmbObra, New clsObras, 0, frmObras, ""
End Sub

