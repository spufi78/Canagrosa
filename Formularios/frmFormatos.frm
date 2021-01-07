VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmformatos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Datos Envases"
   ClientHeight    =   6495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "frmFormatos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6495
   ScaleWidth      =   7035
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   5580
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5580
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5580
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   5940
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5580
      Width           =   1050
   End
   Begin VB.ComboBox cmbPrecintado 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmFormatos.frx":08CA
      Left            =   5985
      List            =   "frmFormatos.frx":08D4
      TabIndex        =   5
      Top             =   5175
      Width           =   945
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
      Left            =   870
      TabIndex        =   0
      Top             =   5130
      Width           =   3945
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4635
      Left            =   60
      TabIndex        =   1
      Top             =   450
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   8176
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
      Caption         =   "Precintado"
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
      Left            =   4905
      TabIndex        =   4
      Top             =   5220
      Width           =   930
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
      Left            =   60
      TabIndex        =   3
      Top             =   5190
      Width           =   660
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mantenimiento Envases"
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
      Width           =   6900
   End
End
Attribute VB_Name = "frmformatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    ElseIf cmbPrecintado.Text = "" Then
        MsgBox "El campo de precinto no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a insertar el envase. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim ode As New clsformatos
            ode.setDESCRIPCION = txtDatos(0)
            If UCase(cmbPrecintado.Text) = "SI" Then
                ode.setPRECINTADO = 1
            Else
                ode.setPRECINTADO = 0
            End If
            ode.Insertar
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
        If MsgBox("Va a ELIMINAR el envase. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim ode As New clsformatos
            ode.Eliminar (lista.ListItems(lista.SelectedItem.Index).SubItems(2))
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbCritical, App.Title
    Else
        If MsgBox("Va a modificar el envase. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim ofor As New clsformatos
            ofor.setDESCRIPCION = txtDatos(0)
            If UCase(cmbPrecintado.Text) = "SI" Then
                ofor.setPRECINTADO = 1
            Else
                ofor.setPRECINTADO = 0
            End If
            ofor.Modificar (lista.ListItems(lista.SelectedItem.Index).SubItems(2))
            cargar_lista
            txtDatos(0) = ""
            cmbPrecintado.Text = ""
        End If
    End If
    txtDatos(0).SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 200
    Me.Left = 200
    With lista.ColumnHeaders.Add(, , "Nombre", 5000, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Precintado", 1000, lvwColumnCenter)
        .Tag = "Unidad"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 500, lvwColumnCenter)
        .Tag = "ID"
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oformatos As New clsformatos
    Set rs = oformatos.Listado
    txtDatos(0) = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("descripcion"))
            If rs("precintado") = 0 Then
                .SubItems(1) = "No"
            Else
                .SubItems(1) = "Si"
            End If
            .SubItems(2) = rs("id_formato")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oformatos = Nothing
    Set rs = Nothing
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtDatos(0).Text = lista.ListItems(lista.SelectedItem.Index).Text
        cmbPrecintado.Text = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
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

