VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmAlbaranes_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Albaranes"
   ClientHeight    =   9210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmAlbaranes_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9210
   ScaleWidth      =   13635
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Listado Desglosado"
      Height          =   885
      Index           =   1
      Left            =   2310
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8280
      Width           =   2205
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Listado"
      Height          =   885
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8280
      Width           =   2205
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
      Height          =   1395
      Left            =   60
      TabIndex        =   3
      Top             =   390
      Width           =   13485
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Situación"
         Height          =   1215
         Left            =   9840
         TabIndex        =   9
         Top             =   120
         Width           =   1815
         Begin VB.OptionButton opEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Facturados"
            Height          =   225
            Index           =   2
            Left            =   300
            TabIndex        =   12
            Top             =   810
            Width           =   1245
         End
         Begin VB.OptionButton opEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pendientes"
            Height          =   225
            Index           =   1
            Left            =   300
            TabIndex        =   11
            Top             =   540
            Width           =   1305
         End
         Begin VB.OptionButton opEstado 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            Height          =   225
            Index           =   0
            Left            =   300
            TabIndex        =   10
            Top             =   270
            Value           =   -1  'True
            Width           =   975
         End
      End
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
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   609
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbObra 
         Height          =   345
         Left            =   1380
         TabIndex        =   8
         Top             =   600
         Width           =   7785
         _ExtentX        =   13732
         _ExtentY        =   609
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1380
         TabIndex        =   14
         Top             =   990
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
         Format          =   50855937
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3570
         TabIndex        =   15
         Top             =   990
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
         Format          =   50855937
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   2940
         TabIndex        =   16
         Top             =   1080
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   1050
         Width           =   465
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
      Height          =   6390
      Left            =   60
      TabIndex        =   0
      Top             =   1830
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   11271
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
      Left            =   9720
      TabIndex        =   18
      Top             =   8400
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
      Left            =   9720
      TabIndex        =   17
      Top             =   8670
      Width           =   2550
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Albaranes"
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
Attribute VB_Name = "frmAlbaranes_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmbCliente_change()
    cargar_lista
End Sub

Private Sub cmbObra_change()
    cargar_lista
End Sub

Private Sub cmbTipoFacturacion_Change()
    cargar_lista
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Command1_Click()

End Sub

Private Sub cmdImprimir_Click(Index As Integer)
    Dim FILTRO As String
   On Error GoTo cmdImprimir_Click_Error

    FILTRO = " {documentos.ANULADO} = 0 AND {documentos.TIPO_DOCUMENTO_ID}=" & ENUM_TIPOS_DOCUMENTOS.ALBARAN
    If cmbCliente.getTEXTO <> "" Then
        FILTRO = FILTRO & " AND {clientes.ID_CLIENTE} = " & cmbCliente.getPK_SALIDA
    End If
    If cmbObra.getTEXTO <> "" Then
        FILTRO = FILTRO & " AND {obras.ID_OBRA} = " & cmbObra.getPK_SALIDA
    End If
    FILTRO = FILTRO & " AND {documentos.FECHA} in Date (" & Year(fdesde) & "," & Month(fdesde) & "," & Day(fdesde) & ") to Date (" & Year(fhasta) & "," & Month(fhasta) & "," & Day(fhasta) & ")"
    
    If opEstado(1).Value = True Then
        FILTRO = FILTRO & " AND {documentos.FACTURADO} = 0 "
    ElseIf opEstado(2).Value = True Then
        FILTRO = FILTRO & " AND {documentos.FACTURADO} = 1 "
    End If
    Me.MousePointer = 11
    Dim p1() As String
    Dim p2() As String
    ReDim p1(2) As String
    ReDim p2(2) As String
    p1(1) = "FECHA_DESDE"
    p1(2) = "FECHA_HASTA"
    
    p2(1) = fdesde
    p2(2) = fhasta
    With frmReport
        .iniciar
        .CRITERIO = FILTRO
        If Index = 0 Then
            .informe = "rptalbaranes_listado"
        Else
            .informe = "rptalbaranes_listado_desglosado"
        End If
        .ParametrosNombre = p1
        .ParametrosValores = p2
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cmdImprimir_Click_Error:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmAlbaranes_Listado"

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
    cabecera_lista
    cargar_combos
    fdesde = Date
    fhasta = Date
    cargar_lista
End Sub
Private Sub cargar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim tipo As String
    Dim cliente As String
    Dim OBRA As String
    Dim numero As String
    Dim anno As String
    Dim ESTADO As String

    tipo = " AND TIPO_DOCUMENTO_ID = " & ENUM_TIPOS_DOCUMENTOS.ALBARAN
        
    If opEstado(0).Value = True Then
        ESTADO = ""
    ElseIf opEstado(1).Value = True Then
        ESTADO = " AND FACTURADO = 0 "
    Else
        ESTADO = " AND FACTURADO = 1 "
    End If
    
    If cmbCliente.getTEXTO <> "" Then
        cliente = " AND O.CLIENTE_ID = " & cmbCliente.getPK_SALIDA
    End If
    If cmbObra.getTEXTO <> "" Then
        OBRA = " AND D.OBRA_ID = " & cmbObra.getPK_SALIDA
    End If
    Dim total As Currency
    total = 0
    Dim rs As New ADODB.Recordset
    consulta = "SELECT D.ID_DOCUMENTO,D.FECHA,D.NUMERO,C.NOMBRE,O.NOMBRE,D.TOTAL,D.FACTURADO, D.DESCUENTO " & _
               "  FROM DOCUMENTOS D " & _
               " INNER JOIN OBRAS O ON D.OBRA_ID = O.ID_OBRA " & _
               " INNER JOIN CLIENTES C ON O.CLIENTE_ID = C.ID_CLIENTE " & _
               " WHERE 1 = 1 AND D.ANULADO = 0 " & _
               tipo & cliente & OBRA & numero & anno & ESTADO & _
               " AND FECHA >= '" & Format(fdesde, "YYYY-MM-DD") & "'" & _
               " AND FECHA <= '" & Format(fhasta, "YYYY-MM-DD") & "'" & _
               " ORDER BY D.NUMERO ASC"
    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    lblalbaranes = "Total (" & rs.RecordCount & " albaranes)"
    If rs.RecordCount >= 1 Then
        total = 0
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = Format(rs.Fields(1), "yyyy-mm-dd")
                .SubItems(2) = rs.Fields(2) ' Numero de factura
                .SubItems(3) = rs.Fields(3) ' Cliente
                .SubItems(4) = rs.Fields(4) ' Obra
                .SubItems(5) = Format(Replace(rs.Fields(5) - rs.Fields(7), ".", ","), "currency")
                total = total + .SubItems(5)
                If rs(6) = 0 Then
                    .SubItems(6) = "P"
                Else
                    .SubItems(6) = "F"
                End If
            End With
            rs.MoveNext
        Wend
'        lista.SetFocus
'    Else
'        MsgBox "No existen albaranes con esos criterios.", vbInformation, App.Title
    End If
    lbltotal = Format(total, "currency")
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
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Fecha", 1300, lvwColumnLeft
        .Add , , "Numero", 1400, lvwColumnCenter
        .Add , , "Cliente", 3700, lvwColumnLeft
        .Add , , "Obra", 3700, lvwColumnLeft
        .Add , , "Base", 1600, lvwColumnRight
        .Add , , "Situación", 1000, lvwColumnCenter
    End With
End Sub
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
            frmDocumento.PK_DOCUMENTO = lista.ListItems(lista.SelectedItem.Index).Text
            frmDocumento.Show 1
    End If
End Sub
Private Sub cargar_combos()
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    llenar_combo cmbObra, New clsObras, 0, frmObras, ""
End Sub

Private Sub opEstado_Click(Index As Integer)
    cargar_lista
End Sub
