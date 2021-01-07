VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFormacion_PF_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cursos de formación"
   ClientHeight    =   9285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9285
   ScaleWidth      =   11130
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   1230
      Left            =   45
      TabIndex        =   7
      Top             =   810
      Width           =   11040
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   870
         Index           =   1
         Left            =   9765
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   180
         Width           =   1140
      End
      Begin VB.TextBox txtCod 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   1755
         TabIndex        =   10
         Top             =   315
         Width           =   1905
      End
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1755
         TabIndex        =   9
         Top             =   720
         Width           =   7890
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   960
         Index           =   0
         Left            =   11205
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1140
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ID Plan:"
         Height          =   285
         Index           =   0
         Left            =   585
         TabIndex        =   12
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción:"
         Height          =   285
         Index           =   1
         Left            =   270
         TabIndex        =   11
         Top             =   765
         Width           =   1830
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   10035
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8370
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8355
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8355
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8355
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6120
      Left            =   45
      TabIndex        =   3
      Top             =   2070
      Width           =   11025
      _ExtentX        =   19447
      _ExtentY        =   10795
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cursos de formación"
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
      Left            =   180
      TabIndex        =   6
      Top             =   135
      Width           =   2145
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   225
      TabIndex        =   5
      Top             =   450
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   10440
      Picture         =   "frmFormacion_PF_Listado.frx":0000
      Top             =   135
      Width           =   480
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11295
      Picture         =   "frmFormacion_PF_Listado.frx":08CA
      Top             =   135
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmFormacion_PF_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private seleccionado As Integer

Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "ID Plan", 1200, lvwColumnLeft)
        .Tag = "Código"
    End With
    With lista.ColumnHeaders.Add(, , "Descripción", 6000, lvwColumnCenter)
        .Tag = "Descripción"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha de creación", 3600, lvwColumnCenter)
        .Tag = "Fecha de creación"
    End With
End Sub

Private Sub cmdAnadir_Click()
    frmFormacion_PF_Detalle.PK = 0
    frmFormacion_PF_Detalle.Show 1
    cargar_lista
    marcar_lista
End Sub

Private Sub cmdBuscar_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a ELIMINAR el plan de formación Nº " & lista.ListItems(lista.selectedItem.Index) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oPlan As New clsFormacion_pf
        Dim oPlanDocs As New clsFormacion_pf_docs
        
        oPlan.Carga CLng(lista.selectedItem.Text)
                
        oPlanDocs.Eliminar CLng(lista.selectedItem.Text)
        oPlan.Eliminar CLng(lista.selectedItem.Text)
       
        Set oPlanDocs = Nothing
        Set oPlan = Nothing
    End If
    cargar_lista
    lista.SetFocus
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Eliminar of Formulario frmFormacion_PlanAnual_Listado"
End Sub


Private Sub cmdModificar_Click()
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then Exit Sub
    frmFormacion_PF_Detalle.PK = CLng(lista.ListItems(lista.selectedItem.Index).Text)
    frmFormacion_PF_Detalle.Show 1
    
    seleccionado = lista.selectedItem.Index
    cargar_lista
    marcar_lista
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Modificar of Formulario frmFormacion_PlanAnual_Listado"
End Sub


Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    cabecera
    cargar_lista
End Sub
 

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oPlanes As New clsFormacion_pf
    
    Set rs = oPlanes.ListadoFiltro(txtCod.Text, txtDescripcion.Text)
 
    lista.ListItems.Clear
     
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("ID_PLAN_FORMACION"))
                 .SubItems(1) = rs("DESCRIPCION")
                 .SubItems(2) = rs("FTIMESTP")
                 
            End With
      
            rs.MoveNext
        Loop Until rs.EOF
    End If
     
    Set oPlanes = Nothing
    Set rs = Nothing
End Sub

Private Sub marcar_lista()
    If seleccionado = 0 Then
        seleccionado = 1
    End If
 
    If lista.ListItems.Count > 0 And lista.ListItems.Count >= seleccionado Then
      lista.ListItems(seleccionado).Selected = True
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub optModalidad_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub optNivel_Click(Index As Integer)
    cargar_lista
End Sub
