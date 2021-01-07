VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTipos_caducidad 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tipos de Caducidad"
   ClientHeight    =   6660
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "frmTipos_caducidad.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7035
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
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
      Left            =   5910
      Picture         =   "frmTipos_caducidad.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5730
      Width           =   1050
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
      Index           =   1
      Left            =   1080
      TabIndex        =   1
      Top             =   5535
      Width           =   2745
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
      Left            =   1080
      TabIndex        =   0
      Top             =   5130
      Width           =   2760
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
      Height          =   465
      Left            =   3900
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5130
      Width           =   1515
   End
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
      Height          =   465
      Left            =   5460
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5130
      Width           =   1545
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4725
      Left            =   60
      TabIndex        =   4
      Top             =   360
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   8334
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
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dias"
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
      Left            =   60
      TabIndex        =   7
      Top             =   5580
      Width           =   390
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Caducidad"
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
      Left            =   60
      TabIndex        =   6
      Top             =   5190
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mantenimiento de Tipos de Caducidad"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   3
      Left            =   45
      TabIndex        =   5
      Top             =   15
      Width           =   6900
   End
End
Attribute VB_Name = "frmTipos_caducidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "El nombre de la caducidad no puede estar en blanco.", vbCritical, App.Title
        txtDatos(0).SetFocus
    ElseIf txtDatos(1).Text = "" Then
        MsgBox "Los dias no pueden estar en blanco.", vbCritical, App.Title
        txtDatos(1).SetFocus
    Else
        If MsgBox("Va a insertar la caducidad. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim ode As New clsTipos_caducidad
            ode.setCADUCIDAD = txtDatos(0)
            ode.setDIAS = txtDatos(1)
            ode.Insertar
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR la caducidad. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim ode As New clsTipos_caducidad
            ode.Eliminar (lista.ListItems(lista.SelectedItem.Index).SubItems(2))
            cargar_lista
        End If
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 200
    Me.Left = 200
    With lista.ColumnHeaders.Add(, , "Caducidad", 4450, lvwColumnLeft)
        .Tag = "Caducidad"
    End With
    With lista.ColumnHeaders.Add(, , "Dias", 1450, lvwColumnLeft)
        .Tag = "Dias"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 500, lvwColumnCenter)
        .Tag = "ID"
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oTipos_caducidad As New clsTipos_caducidad
    Set rs = oTipos_caducidad.Listado
    txtDatos(0) = ""
    txtDatos(1) = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("caducidad"))
            .SubItems(1) = rs("dias")
            .SubItems(2) = rs("id_tipo_caducidad")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oTipos_caducidad = Nothing
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
