VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{76C4A5C3-6A01-4523-911A-8FA5928ECD6B}#1.0#0"; "miComboBCA.ocx"
Begin VB.Form frmObras_Buscar 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Localizar Obra"
   ClientHeight    =   8835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12105
   Icon            =   "frmObras_Buscar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   12105
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1350
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7950
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   855
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7950
      Width           =   1245
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Default         =   -1  'True
      Height          =   855
      Left            =   9540
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7950
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
      TabIndex        =   8
      Top             =   7950
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criterios de búsqueda de Obra"
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
      Height          =   1140
      Left            =   45
      TabIndex        =   5
      Top             =   405
      Width           =   11985
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
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
         Left            =   1020
         TabIndex        =   0
         Top             =   300
         Width           =   1755
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   915
         Left            =   10680
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   150
         Width           =   1245
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
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
         Left            =   7320
         TabIndex        =   3
         Top             =   300
         Width           =   1905
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
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
         Left            =   3840
         TabIndex        =   2
         Top             =   300
         Width           =   2085
      End
      Begin vb6projectpryComboBCA.miComboBCA cmbCliente 
         Height          =   375
         Left            =   1020
         TabIndex        =   14
         Top             =   690
         Width           =   9285
         _ExtentX        =   16378
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cod.Cliente"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   13
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   12
         Top             =   780
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   3135
         TabIndex        =   7
         Top             =   360
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Teléfono"
         Height          =   195
         Index           =   0
         Left            =   6300
         TabIndex        =   6
         Top             =   360
         Width           =   630
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6315
      Left            =   60
      TabIndex        =   1
      Top             =   1590
      Width           =   11985
      _ExtentX        =   21140
      _ExtentY        =   11139
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Búsqueda de Obra"
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
      Index           =   0
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   12195
   End
End
Attribute VB_Name = "frmObras_Buscar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_CLIENTE As Long
Private Sub cmbCliente_change()
cargar_lista
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    gobra = 0
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
        gobra = lista.ListItems(lista.SelectedItem.Index)
        Unload Me
    Else
        gobra = 0
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
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, " ANULADO = 0 "
    If PK_CLIENTE <> 0 Then
        cmbCliente.MostrarElemento PK_CLIENTE
    End If
    cargar_lista
End Sub

Private Sub lista_DblClick()
    cmdok_Click
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Cod.Obra", 0, lvwColumnLeft)
        .Tag = "Cod.Obra"
    End With
    With lista.ColumnHeaders.Add(, , "Cod.Cliente", 1100, lvwColumnCenter)
        .Tag = "Cod.Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre Obra", 3400, lvwColumnLeft)
        .Tag = "Nombre Obra"
    End With
    With lista.ColumnHeaders.Add(, , "Dirección", 3000, lvwColumnLeft)
        .Tag = "Dirección"
    End With
    With lista.ColumnHeaders.Add(, , "Teléfono", 1400, lvwColumnCenter)
        .Tag = "Teléfono"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 2800, lvwColumnCenter)
        .Tag = "Cliente"
    End With
End Sub

Public Sub cargar_lista()
    On Error GoTo fallo
    Dim consulta As String
    Dim NOMBRE As String
    Dim TELEFONO As String
    Dim cliente As String
    NOMBRE = ""
    TELEFONO = ""
    cliente = ""
    If txtDatos(0) <> "" Then
        numero = " and o.CLIENTE_ID like '%" & Trim(txtDatos(0)) & "%'"
    End If
    If txtDatos(1) <> "" Then
        NOMBRE = " and o.NOMBRE like '%" & txtDatos(1) & "%'"
    End If
    If txtDatos(2).Text <> "" Then
        TELEFONO = " AND o.telefono like '%" & txtDatos(2) & "%'"
    End If
    If cmbCliente.getTEXTO <> "" Then
        cliente = " AND o.cliente_id = " & cmbCliente.getPK_SALIDA
    End If
    Dim rs As New ADODB.Recordset
    consulta = "SELECT o.id_obra, " & _
               "       o.cliente_id, " & _
               "       o.nombre, " & _
               "       o.direccion, " & _
               "       o.telefono, " & _
               "       cli.nombre " & _
               " FROM obras o, clientes cli" & _
               " WHERE o.cliente_id = cli.id_cliente " & _
               numero & _
               NOMBRE & _
               TELEFONO & _
               cliente & _
               " ORDER BY o.nombre"
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
                If Not IsNull(rs.Fields(5)) Then
                    .SubItems(5) = rs.Fields(5)
                Else
                    .SubItems(5) = ""
                End If
            End With
            rs.MoveNext
        Wend
'        lista.SetFocus
    Else
        MsgBox "No existen obras con esos criterios.", vbInformation, App.Title
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar los clientes.", vbCritical, Err.Description
End Sub
