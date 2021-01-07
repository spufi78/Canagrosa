VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmTareas_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Tareas"
   ClientHeight    =   8565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8910
   Icon            =   "frmTareas_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8565
   ScaleWidth      =   8910
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
      Height          =   1140
      Left            =   45
      TabIndex        =   7
      Top             =   855
      Width           =   8835
      Begin VB.CheckBox chkActiva 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo las tareas activas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   135
         TabIndex        =   13
         Top             =   720
         Value           =   1  'Checked
         Width           =   2895
      End
      Begin VB.TextBox txtdato 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   4770
         TabIndex        =   11
         Top             =   225
         Width           =   2835
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   870
         Left            =   7830
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   180
         Width           =   870
      End
      Begin MSDataListLib.DataCombo cmbtipos 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   225
         Width           =   2835
         _ExtentX        =   5001
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
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
         Caption         =   "Tarea"
         Height          =   195
         Index           =   1
         Left            =   4185
         TabIndex        =   12
         Top             =   270
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Módulo"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   8
         Top             =   285
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2220
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7650
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7650
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7650
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7650
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5580
      Left            =   45
      TabIndex        =   2
      Top             =   2010
      Width           =   8835
      _ExtentX        =   15584
      _ExtentY        =   9843
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
      Caption         =   "Listado de Tareas"
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
      TabIndex        =   10
      Top             =   90
      Width           =   1920
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   8370
      Picture         =   "frmTareas_Listado.frx":08CA
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado con las tareas existentes en el sistema."
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   9
      Top             =   405
      Width           =   3300
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   8955
   End
End
Attribute VB_Name = "frmTareas_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub cmborigen_Change()
'    cargar_lista
'End Sub
Private Sub cmbtipos_Change()
    cmdBuscar_Click
End Sub
Private Sub cmdAnadir_Click()
    frmTareas_Detalle.PK = 0
    frmTareas_Detalle.Show 1
    cargar_lista
End Sub
Private Sub cmdBuscar_Click()
    cargar_lista
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar la Tarea : " & lista.ListItems(lista.SelectedItem.Index).SubItems(2), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oTarea As New clsTareas
            If oTarea.Eliminar(lista.ListItems(lista.SelectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmTareas_Detalle.PK = lista.ListItems(lista.SelectedItem.Index).Text
        frmTareas_Detalle.Show 1
        actualizar_lista
    End If
End Sub

Private Sub chkActiva_Click()
    cmdBuscar_Click
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = (Screen.Width - frmMenu.ButtonBar.Width - Me.Width) / 2
    Me.Top = (Screen.Height - (frmMenu.SmartMenuXP1.Height * 2) - Me.Height - 1000) / 2
    cabecera
    cargar_botones Me
    cargar_combos
    cargar_lista
'    permisos
'    If USUARIO.getUSUARIO = "julio" Then
'        cmdCargar.Visible = True
'    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Modulo", 2000, lvwColumnCenter
        .Add , , "Descripcion", 3500, lvwColumnLeft
        .Add , , "F.Alta", 1000, lvwColumnCenter
        .Add , , "F.Baja", 1000, lvwColumnCenter
        .Add , , "Activa", 1000, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As ADOdb.RecordSet
    Dim oTareas As New clsTareas
    lista.ListItems.Clear
    Dim tipo As Long
    If cmbtipos.Text = "" Then
        tipo = 0
    Else
        tipo = cmbtipos.BoundText
    End If
    Set rs = oTareas.Listado(tipo, txtdato(0), chkActiva.value)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             If IsDate(rs(3)) Then
                .SubItems(3) = Format(rs(3), "dd-mm-yyyy")
             Else
                .SubItems(3) = ""
             End If
             If IsDate(rs(4)) Then
                 .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
             Else
                .SubItems(4) = ""
             End If
             If rs(5) = 1 Then
                .SubItems(5) = "Si"
             Else
                .SubItems(5) = "No"
             End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oTareas = Nothing
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
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim rs As ADOdb.RecordSet
    Dim oTarea As New clsTareas
    Set rs = oTarea.Listado_por_Codigo(lista.ListItems(lista.SelectedItem.Index).Text)
    If rs.RecordCount <> 0 Then
            With lista.ListItems(lista.SelectedItem.Index)
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = Format(rs(3), "dd-mm-yyyy")
             .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
             If rs(5) = 1 Then
                .SubItems(5) = "Si"
             Else
                .SubItems(5) = "No"
             End If
            End With
            rs.MoveNext
    End If
    Set rs = Nothing
    Set oTarea = Nothing
End Sub

Public Sub cargar_combos()
    Dim oDECODIFICADORA As New clsDecodificadora
    oDECODIFICADORA.cargar_combo cmbtipos, decodificadora.TAREAS_MODULOS
End Sub
Private Sub txtDato_Change(Index As Integer)
    cmdBuscar_Click
End Sub
