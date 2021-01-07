VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmIndicador_Detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Definición de Indicadores"
   ClientHeight    =   10260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13290
   Icon            =   "frmIndicador_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10260
   ScaleWidth      =   13290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   735
      Left            =   90
      TabIndex        =   10
      Top             =   8550
      Width           =   13200
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   990
         TabIndex        =   13
         Top             =   225
         Width           =   9465
      End
      Begin XtremeSuiteControls.PushButton cmdAdd 
         Height          =   435
         Left            =   10530
         TabIndex        =   11
         Top             =   180
         Width           =   1290
         _Version        =   851970
         _ExtentX        =   2275
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmIndicador_Detalle.frx":08CA
      End
      Begin XtremeSuiteControls.PushButton cmdEliminar 
         Height          =   435
         Left            =   11880
         TabIndex        =   12
         Top             =   180
         Width           =   1215
         _Version        =   851970
         _ExtentX        =   2143
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmIndicador_Detalle.frx":712C
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   14
         Top             =   270
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12225
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9360
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Pedido"
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
      Height          =   1035
      Left            =   30
      TabIndex        =   3
      Top             =   600
      Width           =   13230
      Begin pryCombo.miCombo cmbDepartamento 
         Height          =   330
         Left            =   1575
         TabIndex        =   0
         Top             =   270
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbApartado 
         Height          =   330
         Left            =   1620
         TabIndex        =   9
         Top             =   675
         Width           =   11535
         _ExtentX        =   20346
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Departamento"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   5
         Top             =   315
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apartado"
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   4
         Top             =   660
         Width           =   1005
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6600
      Left            =   60
      TabIndex        =   1
      Top             =   1920
      Width           =   13185
      _ExtentX        =   23257
      _ExtentY        =   11642
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
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Indicadores"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   90
      TabIndex        =   8
      Top             =   1650
      Width           =   13155
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Definición de Indicadores"
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
      TabIndex        =   7
      Top             =   30
      Width           =   2670
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12720
      Picture         =   "frmIndicador_Detalle.frx":D98E
      Top             =   30
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Especifique el detalle del indicador"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   285
      Width           =   2445
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   585
      Left            =   0
      Top             =   0
      Width           =   13500
   End
End
Attribute VB_Name = "frmIndicador_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmbApartado_change()
    cargarLista
End Sub

Private Sub cmbDepartamento_change()
    cargarLista
End Sub

Private Sub cmdAdd_Click()
    If validar Then
        Dim oIndicador As New clsIndicador
        With oIndicador
            .setDEPARTAMENTO_ID = cmbDepartamento.getPK_SALIDA
            .setAPARTADO_ID = cmbApartado.getPK_SALIDA
            .setDESCRIPCION = txtDatos(0)
            .Insertar
        End With
        cargarLista
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cargarLista()
    lista.ListItems.Clear
    If cmbDepartamento.getTEXTO = "" Or cmbApartado.getTEXTO = "" Then Exit Sub
    Dim oIndicador As New clsIndicador
    Dim rs As ADODB.Recordset
    Set rs = oIndicador.Listado(cmbDepartamento.getPK_SALIDA, cmbApartado.getPK_SALIDA)
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(3)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oIndicador = Nothing
    Set rs = Nothing
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        Dim oIndicador As New clsIndicador
        oIndicador.Eliminar lista.ListItems(lista.selectedItem.Index).Text
        Set oIndicador = Nothing
        cargarLista
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID_INDICADOR", 1, lvwColumnLeft
        .Add , , "Descripcion", lista.Width - 300, lvwColumnLeft
    End With
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
'    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Function validar() As Boolean
    validar = True
    If cmbDepartamento.getTEXTO = "" Then
        MsgBox "Debe seleccionar el Departamento.", vbInformation, App.Title
        cmbDepartamento.SetFocus
        validar = False
        Exit Function
    End If
    If cmbApartado.getTEXTO = "" Then
        MsgBox "Debe seleccionar el Apartado.", vbInformation, App.Title
        cmbApartado.SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(0) = "" Then
        MsgBox "Debe especificar la Descripción del indicador.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
End Function

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbDepartamento, DECODIFICADORA.INDICADOR_DEPARTAMENTOS
    oDeco.cargar_mi_combo cmbApartado, DECODIFICADORA.INDICADOR_APARTADOS
    Set oDeco = Nothing
End Sub
