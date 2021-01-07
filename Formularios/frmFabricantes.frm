VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFabricantes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Fabricantes"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   Icon            =   "frmFabricantes.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   4470
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2250
      Picture         =   "frmFabricantes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4950
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   60
      Picture         =   "frmFabricantes.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4950
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   3360
      Picture         =   "frmFabricantes.frx":1C16
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4950
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1140
      Picture         =   "frmFabricantes.frx":1F20
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4950
      Width           =   1080
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
      Height          =   360
      Index           =   0
      Left            =   1020
      TabIndex        =   0
      Top             =   4500
      Width           =   3405
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4005
      Left            =   45
      TabIndex        =   1
      Top             =   450
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   7064
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
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   60
      TabIndex        =   7
      Top             =   4560
      Width           =   885
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mantenimiento Fabricantes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   4350
   End
End
Attribute VB_Name = "frmFabricantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "El nombre no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a insertar el fabricante. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim ofab As New clsFabricantes
            ofab.setNOMBRE = txtDatos(0)
            ofab.Insertar
            cargar_lista
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR el fabricante. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim ofab As New clsFabricantes
            ofab.Eliminar (lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "El nombre no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a modificar el fabricante. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim ofab As New clsFabricantes
            ofab.setNOMBRE = txtDatos(0)
            ofab.Modificar (lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            cargar_lista
            txtDatos(0) = ""
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 300
    Me.Left = 300
    With lista.ColumnHeaders.Add(, , "Nombre", 3500, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 400, lvwColumnCenter)
        .Tag = "ID"
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oFabricantes As New clsFabricantes
    Set rs = oFabricantes.Listado
    txtDatos(0) = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("nombre"))
            .SubItems(1) = rs("id_fabricante")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oFabricantes = Nothing
    Set rs = Nothing
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtDatos(0).Text = lista.ListItems(lista.SelectedItem.Index).Text
        txtDatos(0).SetFocus
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
