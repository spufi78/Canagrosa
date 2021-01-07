VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAlodine_Capacidades 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Capacidades"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   Icon            =   "frmAlodine_Capacidades.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   5160
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6840
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6840
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6840
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6840
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
      Index           =   0
      Left            =   1455
      TabIndex        =   0
      Top             =   6045
      Width           =   3660
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5655
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   9975
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
   Begin MSDataListLib.DataCombo cmbetiqueta 
      Height          =   360
      Left            =   1455
      TabIndex        =   8
      Top             =   6420
      Width           =   3630
      _ExtentX        =   6403
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tam.Etiqueta"
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
      Left            =   60
      TabIndex        =   9
      Top             =   6480
      Width           =   1320
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Capacidad"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   60
      TabIndex        =   7
      Top             =   6090
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Mantenimiento Capacidades"
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
      Height          =   315
      Index           =   3
      Left            =   60
      TabIndex        =   2
      Top             =   30
      Width           =   5070
   End
End
Attribute VB_Name = "frmAlodine_Capacidades"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAnadir_Click()
    On Error GoTo fallo
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbInformation, App.Title
        txtDatos(0).SetFocus
        Exit Sub
    End If
    If cmbEtiqueta.BoundText = "" Then
        MsgBox "El tamaño de etiqueta no puede estar en blanco.", vbInformation, App.Title
        cmbEtiqueta.SetFocus
        Exit Sub
    End If
    If MsgBox("Va a insertar la Capacidad. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oAlodine_capacidad As New clsAlodine_capacidad
        oAlodine_capacidad.setDESCRIPCION = txtDatos(0)
        oAlodine_capacidad.setTAMANO_ETIQUETA_ID = cmbEtiqueta.BoundText
        oAlodine_capacidad.Insertar
        cargar_lista
    End If
    txtDatos(0).SetFocus
    Exit Sub
fallo:
    error_grave (Err.Description)
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo fallo
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR la capacidad. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oAlodine_capacidad As New clsAlodine_capacidad
            oAlodine_capacidad.Eliminar (lista.ListItems(lista.selectedItem.Index).SubItems(2))
            cargar_lista
        End If
    End If
    Exit Sub
fallo:
    error_grave (Err.Description)
End Sub

Private Sub cmdModificar_Click()
    On Error GoTo fallo
    If txtDatos(0).Text = "" Then
        MsgBox "La descripción no puede estar en blanco.", vbInformation, App.Title
        txtDatos(0).SetFocus
        Exit Sub
    End If
    If cmbEtiqueta.BoundText = "" Then
        MsgBox "El tamaño de etiqueta no puede estar en blanco.", vbInformation, App.Title
        cmbEtiqueta.SetFocus
        Exit Sub
    End If
    If MsgBox("Va a modificar la capacidad. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oAlodine_capacidad As New clsAlodine_capacidad
        oAlodine_capacidad.setDESCRIPCION = txtDatos(0)
        oAlodine_capacidad.setTAMANO_ETIQUETA_ID = cmbEtiqueta.BoundText
        oAlodine_capacidad.Modificar (lista.ListItems(lista.selectedItem.Index).SubItems(2))
        cargar_lista
    End If
    txtDatos(0).SetFocus
    Exit Sub
fallo:
    error_grave (Err.Description)
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 200
    Me.Left = 200
    With lista.ColumnHeaders.Add(, , "Descripción", 3225, lvwColumnLeft)
        .Tag = "Descripción"
    End With
    With lista.ColumnHeaders.Add(, , "Etiqueta", 1100, lvwColumnCenter)
        .Tag = "Etiqueta"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 500, lvwColumnCenter)
        .Tag = "ID"
    End With
    cargar_lista
    cargar_combos
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.RecordSet
    Dim oAlodine_capacidad As New clsAlodine_capacidad
    Set rs = oAlodine_capacidad.Listado
    txtDatos(0) = ""
    cmbEtiqueta = ""
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(0))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oCapacidades = Nothing
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
        txtDatos(0).Text = lista.ListItems(lista.selectedItem.Index).Text
        cmbEtiqueta.Text = lista.ListItems(lista.selectedItem.Index).SubItems(1)
    End If
End Sub

Public Sub cargar_combos()
    cargar_combo cmbEtiqueta, New clsTamanos_etiqueta
End Sub
