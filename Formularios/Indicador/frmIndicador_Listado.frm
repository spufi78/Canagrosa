VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmIndicador_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Indicadores"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13680
   Icon            =   "frmIndicador_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   13680
   Begin VB.CommandButton cmdApartados 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Apartados"
      Height          =   870
      Left            =   11385
      Picture         =   "frmIndicador_Listado.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7590
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   45
      TabIndex        =   4
      Top             =   675
      Width           =   13545
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   795
         Left            =   12540
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   915
      End
      Begin pryCombo.miCombo cmbDepartamento 
         Height          =   330
         Left            =   1320
         TabIndex        =   10
         Top             =   240
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbApartado 
         Height          =   330
         Left            =   1320
         TabIndex        =   12
         Top             =   630
         Width           =   11115
         _ExtentX        =   19606
         _ExtentY        =   582
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apartado"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   690
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Departamento"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdDepartamentos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Departamentos"
      Height          =   870
      Left            =   10170
      Picture         =   "frmIndicador_Listado.frx":1B3C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7590
      Width           =   1185
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12615
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7590
      Width           =   1020
   End
   Begin VB.CommandButton cmdGestion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Gestión"
      Height          =   870
      Left            =   90
      Picture         =   "frmIndicador_Listado.frx":2406
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7605
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5745
      Left            =   60
      TabIndex        =   0
      Top             =   1770
      Width           =   13545
      _ExtentX        =   23892
      _ExtentY        =   10134
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
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13095
      Picture         =   "frmIndicador_Listado.frx":2CD0
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
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
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   6
      Top             =   60
      Width           =   2400
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Especifique los datos necesarios para localizar los Indicadores"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   360
      Width           =   4395
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   13690
   End
End
Attribute VB_Name = "frmIndicador_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbApartado_change()
    cargar_lista
End Sub

Private Sub cmbDepartamento_change()
    cargar_lista
End Sub
Private Sub cmdApartados_Click()
    frmDecodificadora.CODIGO = DECODIFICADORA.INDICADOR_APARTADOS
    frmDecodificadora.Show
End Sub

Private Sub cmdDepartamentos_Click()
    frmDecodificadora.CODIGO = DECODIFICADORA.INDICADOR_DEPARTAMENTOS
    frmDecodificadora.Show
End Sub

Private Sub cmdGestion_Click()
    frmIndicador_Detalle.PK = 0
    frmIndicador_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdLimpiar_Click()
    cmbDepartamento.Limpiar
    cmbApartado.Limpiar
    cargar_lista
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_combos
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID_INDICADOR", 1, lvwColumnLeft
        .Add , , "Departamento", 2000, lvwColumnLeft
        .Add , , "Apartado", 3500, lvwColumnLeft
        .Add , , "Descripción", 7700, lvwColumnLeft
    End With
End Sub
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbDepartamento, DECODIFICADORA.INDICADOR_DEPARTAMENTOS
    oDeco.cargar_mi_combo cmbApartado, DECODIFICADORA.INDICADOR_APARTADOS
    Set oDeco = Nothing
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oIndicador As New clsIndicador
    lista.ListItems.Clear
    Dim departamento As Long
    Dim apartado As Long
    If cmbDepartamento.getTEXTO <> "" Then
        departamento = cmbDepartamento.getPK_SALIDA
    End If
    If cmbApartado.getTEXTO <> "" Then
        apartado = cmbApartado.getPK_SALIDA
    End If
    Set rs = oIndicador.Listado(departamento, apartado)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0)) ' id
             .SubItems(1) = rs(1) ' depar
             .SubItems(2) = rs(2) ' apar
             .SubItems(3) = rs(3) ' descripcion
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oIndicador = Nothing
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
