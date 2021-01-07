VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoUsuarios 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Usuarios"
   ClientHeight    =   7260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   Icon            =   "frmListadoUsuarios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7260
   ScaleWidth      =   10440
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6330
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   9210
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6330
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6330
      Width           =   1155
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6330
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5910
      Left            =   45
      TabIndex        =   0
      Top             =   360
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   10425
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
      BackColor       =   &H00C0FFFF&
      Caption         =   "Usuarios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   4
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   10320
   End
End
Attribute VB_Name = "frmListadoUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAnadir_Click()
    gusuario = 0
    frmUsuarios.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        Dim pos As Integer
        Dim usuario As Integer
        If MsgBox("Va a ELIMINAR al usuario " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
            Dim ousu As New ClsUsuario
            If ousu.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
                cargar_lista
            End If
            Set ousu = Nothing
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        gusuario = lista.ListItems(lista.SelectedItem.Index)
        frmUsuarios.Show 1
        actualizar_lista
        gusuario = 0
        frmMenu.permisos
    End If
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
    Me.Top = 100
    Me.Left = 100
    cargar_cabecera
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ocli As New ClsUsuario
    Set rs = ocli.Listado
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("id_empleado"))
            .SubItems(1) = rs("usuario")
            .SubItems(2) = rs("nombre")
            .SubItems(3) = rs("per_1")
            .SubItems(4) = rs("per_2")
            .SubItems(5) = rs("per_3")
            .SubItems(6) = rs("per_4")
            .SubItems(7) = rs("per_5")
            .SubItems(8) = rs("per_6")
            .SubItems(9) = rs("per_7")
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
    If lista.ListItems(lista.SelectedItem.Index) <> "" Then
      cmdmodificar.Enabled = True
      cmdeliminar.Enabled = True
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim ousu As New ClsUsuario
    If ousu.Cargar(gusuario) = True Then
        lista.ListItems(lista.SelectedItem.Index).Text = gusuario
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = ousu.getUSUARIO
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = ousu.getNOMBRE
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = ousu.getPER_1
        lista.ListItems(lista.SelectedItem.Index).SubItems(4) = ousu.getPER_2
        lista.ListItems(lista.SelectedItem.Index).SubItems(5) = ousu.getPER_3
        lista.ListItems(lista.SelectedItem.Index).SubItems(6) = ousu.getPER_4
        lista.ListItems(lista.SelectedItem.Index).SubItems(7) = ousu.getPER_5
        lista.ListItems(lista.SelectedItem.Index).SubItems(8) = ousu.getPER_6
        lista.ListItems(lista.SelectedItem.Index).SubItems(9) = ousu.getPER_7
    End If
    Set ousu = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Public Sub cargar_cabecera()
    With lista.ColumnHeaders.Add(, , "Codigo", 600, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Usuario", 1300, lvwColumnLeft)
        .Tag = "Usuario"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 4500, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Impresión", 500, lvwColumnCenter)
        .Tag = "Impresión"
    End With
    With lista.ColumnHeaders.Add(, , "Alta", 500, lvwColumnCenter)
        .Tag = "Alta"
    End With
    With lista.ColumnHeaders.Add(, , "Modificación", 500, lvwColumnCenter)
        .Tag = "Modificación"
    End With
    With lista.ColumnHeaders.Add(, , "Baja", 500, lvwColumnCenter)
        .Tag = "Baja"
    End With
    With lista.ColumnHeaders.Add(, , "Usuarios", 500, lvwColumnCenter)
        .Tag = "Usuarios"
    End With
    With lista.ColumnHeaders.Add(, , "Recalculo", 500, lvwColumnCenter)
        .Tag = "Recalculo"
    End With
    With lista.ColumnHeaders.Add(, , "Expedientes", 500, lvwColumnCenter)
        .Tag = "Expedientes"
    End With
End Sub
