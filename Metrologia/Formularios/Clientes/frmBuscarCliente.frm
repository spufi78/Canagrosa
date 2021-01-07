VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBuscarCliente 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localizar Cliente"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9885
   Icon            =   "frmBuscarCliente.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   9885
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   7590
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   855
      Left            =   8580
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7590
      Width           =   1245
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Default         =   -1  'True
      Height          =   855
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7590
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Nuevo"
      Height          =   885
      Left            =   90
      MaskColor       =   &H00E0E0E0&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7590
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criterios de búsqueda de Cliente"
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
      Height          =   1500
      Left            =   45
      TabIndex        =   2
      Top             =   45
      Width           =   9765
      Begin VB.CheckBox chkAnulados 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Buscar clientes anulados"
         Height          =   225
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   2805
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   5310
         TabIndex        =   10
         Top             =   720
         Width           =   2595
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1230
         TabIndex        =   9
         Top             =   750
         Width           =   2955
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1230
         TabIndex        =   0
         Top             =   300
         Width           =   2955
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   1005
         Left            =   8280
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   300
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   5310
         TabIndex        =   1
         Top             =   300
         Width           =   2595
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Móvil"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   4440
         TabIndex        =   5
         Top             =   750
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Teléfono"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   4
         Top             =   780
         Width           =   855
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.I.F."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   4440
         TabIndex        =   3
         Top             =   360
         Width           =   555
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5985
      Left            =   60
      TabIndex        =   7
      Top             =   1560
      Width           =   9765
      _ExtentX        =   17224
      _ExtentY        =   10557
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
End
Attribute VB_Name = "frmBuscarCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    gcliente = 0
    Unload Me
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        gcliente = lista.ListItems(lista.SelectedItem.Index)
        frmClientes.Show 1
        gcliente = 0
    End If
End Sub

Private Sub cmdNuevo_Click()
    gcliente = -1
    frmClientes.Show 1
    If gcliente > 0 Then
'        cargar_clientes
'        Dim ocli As New clsCliente
'        ocli.CargaCliente (gcliente)
'        cmbclientes.Text = ocli.getNOMBRE
'        gcliente = 0
        Unload Me
    End If
'    txtDatos(0).SetFocus
End Sub

Private Sub cmdok_Click()
    If lista.ListItems.Count > 0 Then
        gcliente = lista.ListItems(lista.SelectedItem.Index)
        Unload Me
    Else
        gcliente = 0
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 27
        cmdcancel_Click
     Case 121 ' F10
        cmdok_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    cargar_lista
End Sub
Private Sub lista_DblClick()
    cmdok_Click
End Sub

Private Sub txtdatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub
Private Sub txtDatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = vbWhite
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Codigo", 600, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 3400, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Dirección", 3400, lvwColumnLeft)
        .Tag = "Dirección"
    End With
    With lista.ColumnHeaders.Add(, , "Teléfono", 1000, lvwColumnCenter)
        .Tag = "Teléfono"
    End With
    With lista.ColumnHeaders.Add(, , "Móvil", 1000, lvwColumnCenter)
        .Tag = "Móvil"
    End With
End Sub

Public Sub cargar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim NOMBRE As String
    Dim nif As String
    Dim telefono As String
    Dim movil As String
    NOMBRE = ""
    nif = ""
    telefono = ""
    movil = ""
    NOMBRE = " NOMBRE like '%" & txtdatos(1) & "%'"
    If txtdatos(0).Text <> "" Then
        nif = " AND CIF like '%" & txtdatos(0) & "%'"
    End If
    If txtdatos(2).Text <> "" Then
        telefono = " AND telefono like '%" & txtdatos(2) & "%'"
    End If
    If txtdatos(3).Text <> "" Then
        movil = " AND movil like '%" & txtdatos(3) & "%'"
    End If
    Dim ANULADOS As String
    If chkanulados.Value = Unchecked Then
        ANULADOS = " AND ANULADO = 0 "
    End If
    Dim rs As New ADODB.Recordset
    consulta = "SELECT id_Cliente, " & _
               "       nombre, " & _
               "       direccion, " & _
               "       telefono, " & _
               "       movil " & _
               " FROM Clientes " & _
               " WHERE " & _
               NOMBRE & _
               nif & _
               telefono & _
               movil & _
               ANULADOS & _
               " ORDER BY nombre"
    lista.ListItems.Clear
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , rs.Fields(0))
                .SubItems(1) = rs.Fields(1)
                If Not IsNull(rs.Fields(2)) Then
                    .SubItems(2) = rs.Fields(2)
                Else
                    .SubItems(2) = ""
                End If
                If Not IsNull(rs.Fields(3)) Then
                    .SubItems(3) = rs.Fields(3)
                Else
                    .SubItems(3) = ""
                End If
                If Not IsNull(rs.Fields(4)) Then
                    .SubItems(4) = rs.Fields(4)
                Else
                    .SubItems(4) = ""
                End If
            End With
            rs.MoveNext
        Wend
'        lista.SetFocus
    Else
        MsgBox "No existen clientes con esos criterios.", vbInformation, App.Title
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los clientes.", vbCritical, Err.Description
End Sub
