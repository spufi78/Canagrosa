VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTD_Familias 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento Familias de determinaciones"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTD_Familias.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   6090
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdFamilia 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Relación"
      Height          =   870
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6375
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Index           =   0
      Left            =   915
      TabIndex        =   1
      Top             =   5925
      Width           =   5100
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   4995
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6375
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6375
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2290
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6375
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1175
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6375
      Width           =   1080
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5340
      Left            =   60
      TabIndex        =   0
      Top             =   555
      Width           =   6000
      _ExtentX        =   10583
      _ExtentY        =   9419
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
      Left            =   5580
      Picture         =   "frmTD_Familias.frx":08CA
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Familias de determinaciones"
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
      TabIndex        =   8
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nombre"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   105
      TabIndex        =   7
      Top             =   5985
      Width           =   660
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   10305
   End
End
Attribute VB_Name = "frmTD_Familias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a insertar la familia. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim obj As New clsTipos_determinacion_familias
            With obj
                .setNOMBRE = txtDatos(0)
                .Insertar
            End With
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
        If MsgBox("Va a ELIMINAR la familia. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim obj As New clsTipos_determinacion_familias
            obj.Eliminar (lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdFamilia_Click()
    If lista.ListItems.Count > 0 Then
        frmTD_Familias_Relacion.FAMILIA = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        frmTD_Familias_Relacion.Show 1
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a modificar la familia. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim obj As New clsTipos_determinacion_familias
            With obj
                .setNOMBRE = txtDatos(0)
                .Modificar (lista.ListItems(lista.SelectedItem.Index).SubItems(1))
            End With
            cargar_lista
            txtDatos(0) = ""
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 200
    Me.Left = 200
    With lista.ColumnHeaders.Add(, , "Nombre", 5300, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 400, lvwColumnCenter)
        .Tag = "ID"
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim obj As New clsTipos_determinacion_familias
    Set rs = obj.Listado
    txtDatos(0) = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(1))
            .SubItems(1) = rs(0)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set obj = Nothing
    Set rs = Nothing
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
    If lista.ListItems.Count > 0 Then
        txtDatos(0).Text = lista.ListItems(lista.SelectedItem.Index).Text
    End If
End Sub
