VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmFacturas_Impresion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresión de Facturas"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmFacturas_Impresion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   13635
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   885
      Index           =   1
      Left            =   90
      Picture         =   "frmFacturas_Impresion.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8280
      Width           =   1155
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   885
      Index           =   0
      Left            =   1290
      Picture         =   "frmFacturas_Impresion.frx":1F14
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8280
      Width           =   1155
   End
   Begin VB.TextBox txtdatos 
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
      Index           =   1
      Left            =   6930
      TabIndex        =   15
      Text            =   "2"
      Top             =   8730
      Width           =   945
   End
   Begin VB.CheckBox chkCopia 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incluir ES COPIA en la Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5400
      TabIndex        =   14
      Top             =   8325
      Value           =   1  'Checked
      Width           =   3075
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Facturas Seleccionadas"
      Height          =   885
      Index           =   0
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8280
      Width           =   2880
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro de Selección de Albaranes"
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
      Height          =   1410
      Left            =   45
      TabIndex        =   3
      Top             =   360
      Width           =   13485
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   1050
         Left            =   12150
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   1185
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbCliente 
         Height          =   345
         Left            =   1380
         TabIndex        =   7
         Top             =   240
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   609
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbObra 
         Height          =   345
         Left            =   1380
         TabIndex        =   8
         Top             =   600
         Width           =   10365
         _ExtentX        =   18283
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1395
         TabIndex        =   10
         Top             =   945
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
         Format          =   51118081
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3585
         TabIndex        =   11
         Top             =   945
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
         Format          =   51118081
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   13
         Top             =   1005
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2955
         TabIndex        =   12
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   6
         Top             =   270
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Obra"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   630
         Width           =   345
      End
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
      Height          =   6090
      Left            =   60
      TabIndex        =   0
      Top             =   2130
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   10742
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"frmFacturas_Impresion.frx":27DE
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
      Height          =   285
      Left            =   90
      TabIndex        =   19
      Top             =   1800
      Width           =   13380
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Numero Copias"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   5400
      TabIndex        =   16
      Top             =   8820
      Width           =   1410
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Impresión de Facturas"
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
      Left            =   30
      TabIndex        =   2
      Top             =   0
      Width           =   13545
   End
End
Attribute VB_Name = "frmFacturas_Impresion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub cmdMarcar_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
            lista.ListItems(i).Checked = Index
    Next
End Sub


Private Sub cmdImprimir_Click(Index As Integer)
    Dim i As Integer
    Dim algo As Boolean

   On Error GoTo cmdImprimir_Click_Error

    If lista.ListItems.Count > 0 Then
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                algo = True
            End If
        Next
        If Not algo Then
            MsgBox "Marque las facturas que desea imprimir.", vbExclamation, App.Title
            Exit Sub
        End If
    Else
        MsgBox "No existen facturas para imprimir.", vbExclamation, App.Title
        Exit Sub
    End If
    If MsgBox("Se imprimiran todas las facturas marcadas. ¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    Me.MousePointer = 11
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            Dim oDOCUMENTO As New clsDocumentos
            oDOCUMENTO.imprimir lista.ListItems(i).Text, True, False, txtdatos(1), , chkCopia.Value
            Set oDOCUMENTO = Nothing
        End If
    Next
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cmdImprimir_Click_Error:

    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmFacturas_Impresion"
    
End Sub


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
    Me.Left = 100
    Me.Top = 100
    fdesde = Date - 31
    fhasta = Date
    cabecera_lista
    cargar_combos
    cargar_lista
End Sub
Public Sub cargar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim cliente As String
    Dim obra As String
    Dim numero As String
    Dim anno As String
    Dim ESTADO As String
    Dim agente As String
    tipo = " AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.factura
    
    If cmbCliente.getTEXTO <> "" Then
        cliente = " AND O.CLIENTE_ID = " & cmbCliente.getPK_SALIDA
    End If
    If cmbObra.getTEXTO <> "" Then
        obra = " AND D.OBRA_ID = " & cmbObra.getPK_SALIDA
    End If
    
    ESTADO = " AND D.ESTADO_ID <> " & ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_ANULADO
    
    Dim rs As New ADODB.Recordset
    consulta = "SELECT D.ID_DOCUMENTO,D.NUMERO,D.FECHA,C.ID_CLIENTE,O.NOMBRE,C.NOMBRE,D.TOTAL,D.IVA " & _
               "  FROM DOCUMENTOS D " & _
               "  LEFT JOIN DOCUMENTOS_TIPOS TD ON TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "  LEFT JOIN OBRAS O ON D.OBRA_ID = O.ID_OBRA " & _
               "  LEFT JOIN CLIENTES C ON C.ID_CLIENTE = O.CLIENTE_ID " & _
               "  LEFT JOIN COMERCIALES COMER ON O.COMERCIAL_ID = COMER.ID_COMERCIAL " & _
               " WHERE 1 = 1 " & _
               "   AND D.FECHA >= '" & Format(fdesde, "YYYY-MM-DD") & "'" & _
               "   AND D.FECHA <= '" & Format(fhasta, "YYYY-MM-DD") & "'" & _
               tipo & cliente & obra & numero & anno & ESTADO & agente & _
               " ORDER BY D.NUMERO ASC "
    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim total As Currency
        total = 0
        While Not rs.EOF
        
                With lista.ListItems.Add(, , rs(0))
                    .SubItems(1) = Format(rs.Fields(1), "0000") ' nUMERO
                    .SubItems(2) = Format(rs.Fields(2), "dd-mm-yyyy") ' fECHA
                    .SubItems(3) = rs.Fields(3) ' ID_CLIENTE
                    .SubItems(4) = rs(4) ' CLIENTE
                    .SubItems(5) = rs(5) ' OBRA
                    .SubItems(6) = moneda(rs(6) + (rs(6) * rs(7) / 100)) ' IMPORTE
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
    ' Pendientes
    With lista.ColumnHeaders
        .Add , , "ID", 310, lvwColumnLeft
        .Add , , "NºFactura", 800, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Cod.Cliente", 1300, lvwColumnCenter
        .Add , , "Cliente", 4100, lvwColumnLeft
        .Add , , "Obra", 4100, lvwColumnLeft
        .Add , , "Importe", 1200, lvwColumnRight
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
Private Sub actualizar_lista()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
   On Error GoTo actualizar_lista_Error

    consulta = "SELECT D.ID_DOCUMENTO,D.NUMERO,D.FECHA,C.ID_CLIENTE,O.NOMBRE,C.NOMBRE,D.TOTAL,D.IVA " & _
               "  FROM DOCUMENTOS D " & _
               "  LEFT JOIN DOCUMENTOS_TIPOS TD ON TD.ID_TIPO_DOCUMENTO = D.TIPO_DOCUMENTO_ID " & _
               "  LEFT JOIN OBRAS O ON D.OBRA_ID = O.ID_OBRA " & _
               "  LEFT JOIN CLIENTES C ON C.ID_CLIENTE = O.CLIENTE_ID " & _
               "  LEFT JOIN COMERCIALES COMER ON O.COMERCIAL_ID = COMER.ID_COMERCIAL " & _
               " WHERE 1 = 1 " & _
               "   AND D.ID_DOCUMENTO = " & lista.ListItems(lista.SelectedItem.Index).Text

    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
            With lista.ListItems(lista.SelectedItem.Index)
                    .SubItems(1) = Format(rs.Fields(1), "0000") ' nUMERO
                    .SubItems(2) = Format(rs.Fields(2), "dd-mm-yyyy") ' fECHA
                    .SubItems(3) = rs.Fields(3) ' ID_CLIENTE
                    .SubItems(4) = rs(4) ' CLIENTE
                    .SubItems(5) = rs(5) ' OBRA
                    .SubItems(6) = moneda(rs(6) + (rs(6) * rs(7) / 100)) ' IMPORTE
            End With
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

actualizar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure actualizar_lista of Formulario frmFacturas_Impresion"
End Sub
Private Sub cargar_combos()
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    llenar_combo cmbObra, New clsObras, 0, frmObras, ""
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
            frmDocumento.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).Text
            frmDocumento.Show 1
            actualizar_lista
    End If

End Sub
