VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsuarios_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Usuarios"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   Icon            =   "frmUsuarios_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8325
   ScaleWidth      =   10830
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por"
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
      Height          =   690
      Left            =   45
      TabIndex        =   7
      Top             =   765
      Width           =   10770
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   4680
         TabIndex        =   11
         Top             =   270
         Width           =   2265
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1035
         TabIndex        =   9
         Top             =   270
         Width           =   2265
      End
      Begin VB.CheckBox chkAnulados 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Anulados"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   8415
         TabIndex        =   8
         Top             =   315
         Width           =   1980
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   3735
         TabIndex        =   12
         Top             =   315
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apellidos"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   315
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7425
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7410
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7410
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1110
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7410
      Width           =   1080
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5865
      Left            =   45
      TabIndex        =   0
      Top             =   1485
      Width           =   10740
      _ExtentX        =   18944
      _ExtentY        =   10345
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Usuarios no anulados del geslab"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   420
      Width           =   3075
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10260
      Picture         =   "frmUsuarios_Listado.frx":08CA
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Usuarios"
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
      TabIndex        =   5
      Top             =   120
      Width           =   2100
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   10815
   End
End
Attribute VB_Name = "frmUsuarios_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkAnulados_Click()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    gempleado = 0
    frmUsuarios.Show 1
    cargar_lista
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR al usuario " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oUsu As New clsUsuarios
        If oUsu.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
            cargar_lista
        End If
        Set oUsu = Nothing
    End If
End Sub

Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_Error

    gempleado = lista.ListItems(lista.selectedItem.Index)
    frmUsuarios.Show 1
    actualizar_lista
    gempleado = 0

   On Error GoTo 0
   Exit Sub

cmdModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmUsuarios_Listado"
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdCancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    With lista.ColumnHeaders
        .Add , , "Codigo", 600, lvwColumnLeft
        .Add , , "Apellidos", 3300, lvwColumnLeft
        .Add , , "Nombre", 1500, lvwColumnLeft
        .Add , , "Usuario", 1300, lvwColumnCenter
        .Add , , "Email", 3000, lvwColumnLeft
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ocli As New clsUsuarios
    Set rs = ocli.Listado_Filtro(txtfiltro(0), txtfiltro(1), chkAnulados.value)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_empleado"), "000"))
            .SubItems(1) = rs("apellidos")
            .SubItems(2) = rs("nombre")
            .SubItems(3) = rs("usuario")
            .SubItems(4) = rs("email")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set ocli = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
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
Private Sub lista_Click()
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim oUsu As New clsUsuarios
    If oUsu.CARGAR(gempleado) = True Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = oUsu.getAPELLIDOS
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = oUsu.getNOMBRE
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = oUsu.getUSUARIO
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = oUsu.getEMAIL
    End If
    Set oUsu = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
