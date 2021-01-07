VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEquipos_Listado_Avisos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Equipos fuera del periodo de calibración, verificación y mantenimiento"
   ClientHeight    =   8610
   ClientLeft      =   405
   ClientTop       =   1005
   ClientWidth     =   14025
   Icon            =   "frmEquipos_Listado_Avisos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   14025
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   7740
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   1320
      Left            =   45
      TabIndex        =   0
      Top             =   720
      Width           =   13965
      Begin VB.CheckBox chkbaja 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar Equipos en Baja"
         Height          =   195
         Left            =   9855
         TabIndex        =   14
         Top             =   1035
         Visible         =   0   'False
         Width           =   2085
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   7380
         TabIndex        =   13
         Top             =   180
         Visible         =   0   'False
         Width           =   2670
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1125
         TabIndex        =   12
         Top             =   540
         Width           =   1545
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1125
         TabIndex        =   11
         Top             =   180
         Width           =   1545
      End
      Begin VB.OptionButton opListado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Equipos con Calibración"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   1035
         Width           =   2310
      End
      Begin VB.OptionButton opListado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Equipos con Verificación"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   2565
         TabIndex        =   9
         Top             =   1035
         Width           =   2400
      End
      Begin VB.OptionButton opListado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Equipos con Mantenimiento"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   5175
         TabIndex        =   8
         Top             =   1035
         Width           =   2670
      End
      Begin VB.OptionButton opListado 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   7920
         TabIndex        =   7
         Top             =   1035
         Value           =   -1  'True
         Width           =   1005
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   3645
         TabIndex        =   6
         Top             =   180
         Width           =   2670
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   6
         Left            =   10935
         TabIndex        =   5
         Top             =   135
         Visible         =   0   'False
         Width           =   105
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Estado revisión"
         Height          =   1095
         Left            =   11970
         TabIndex        =   1
         Top             =   135
         Width           =   1905
         Begin VB.OptionButton opRevision 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mostrar sin revisar"
            Height          =   240
            Index           =   2
            Left            =   180
            TabIndex        =   4
            Top             =   810
            Width           =   1590
         End
         Begin VB.OptionButton opRevision 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mostrar revisados"
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   3
            Top             =   495
            Width           =   1590
         End
         Begin VB.OptionButton opRevision 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Mostrar todos"
            Height          =   240
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   225
            Value           =   -1  'True
            Width           =   1590
         End
      End
      Begin MSDataListLib.DataCombo cmbFiltro 
         Height          =   315
         Index           =   0
         Left            =   3645
         TabIndex        =   15
         Top             =   540
         Width           =   2680
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbFiltro 
         Height          =   315
         Index           =   1
         Left            =   7380
         TabIndex        =   16
         Top             =   540
         Width           =   2680
         _ExtentX        =   4736
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label Label2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descripción"
         Height          =   240
         Left            =   6480
         TabIndex        =   23
         Top             =   225
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Num. Serie"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   22
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nº Equipo"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   21
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nombre"
         Height          =   240
         Index           =   2
         Left            =   2790
         TabIndex        =   20
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Familia"
         Height          =   240
         Index           =   3
         Left            =   6480
         TabIndex        =   19
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Situación"
         Height          =   240
         Left            =   2790
         TabIndex        =   18
         Top             =   585
         Width           =   915
      End
      Begin VB.Label Label4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Proveedor"
         Height          =   240
         Left            =   10080
         TabIndex        =   17
         Top             =   180
         Visible         =   0   'False
         Width           =   915
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5640
      Left            =   45
      TabIndex        =   25
      Top             =   2025
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   9948
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
      AutoSize        =   -1  'True
      Caption         =   "Doble click para ver el detalle del equipo"
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
      Index           =   7
      Left            =   5850
      TabIndex        =   28
      Top             =   7785
      Width           =   3525
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Avisos - Gestión de Equipos de Medición y Ensayo"
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
      TabIndex        =   27
      Top             =   120
      Width           =   5310
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13230
      Picture         =   "frmEquipos_Listado_Avisos.frx":164A
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Ventana de gestión de Equipos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   26
      Top             =   420
      Width           =   2220
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   14070
   End
End
Attribute VB_Name = "frmEquipos_Listado_Avisos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public strAvisos_a_mostrar As String ' todos, calibración, verificación o mantenimiento

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 100
    Me.Left = 100
    
    cargar_botones Me
    Call cargar_combos
    Call cabecera
    strAvisos_a_mostrar = "TODOS"
    Call cargar_lista
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    
    oDeco.cargar_combo cmbFiltro(1), decodificadora.EQ_FAMILIAS
    oDeco.cargar_combo cmbFiltro(0), decodificadora.EQ_SITUACIONES
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "NºEquipo", 900, lvwColumnLeft
        .Add , , "Nombre Equipo", 4000, lvwColumnLeft
        .Add , , "NºSerie", 1200, lvwColumnCenter
        .Add , , "Proveedor", 2000, lvwColumnLeft
        .Add , , "Situación", 2600, lvwColumnLeft
        .Add , , "Familia", 2500, lvwColumnLeft
        .Add , , "C.Amb.", 700, lvwColumnLeft
    End With
End Sub

Private Sub opListado_Click(Index As Integer)
    Select Case Index
        Case 0: ' todos
            strAvisos_a_mostrar = "TODOS"
        Case 1: ' calibración
            strAvisos_a_mostrar = "CALIBRACION"
        Case 2: ' verificación
            strAvisos_a_mostrar = "VERIFICACION"
        Case 3: ' mantenimiento
            strAvisos_a_mostrar = "MANTENIMIENTO"
    End Select
    cargar_lista
End Sub

Private Sub opRevision_Click(Index As Integer)
    cargar_lista
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub chkSinRevisar_Click()
    cargar_lista
End Sub

Private Sub cmbFiltro_Change(Index As Integer)
    cargar_lista
End Sub

Public Sub cargar_lista()
    lista.ListItems.Clear
    
    Select Case strAvisos_a_mostrar
        Case "TODOS"
            Call cargar_avisos_calibracion
            Call cargar_avisos_verificacion
            Call cargar_avisos_mantenimiento
            
        Case "CALIBRACION"
            Call cargar_avisos_calibracion
            
        Case "VERIFICACION"
            Call cargar_avisos_verificacion
            
        Case "MANTENIMIENTO"
            Call cargar_avisos_mantenimiento
            
    End Select
    lblsubtitulo = "Ventana de gestión de Equipos. Número de equipos mostrados : " & lista.ListItems.Count
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        'frmEquipos_Detalle.ES_AVISO = True
        'frmEquipos_Detalle.PK = lista.ListItems(lista.SelectedItem.Index).Text
        'frmEquipos_Detalle.Show 1
        cargar_lista
    Else
        MsgBox "Debe seleccionar el equipo que desea modificar.", vbOKOnly + vbInformation, App.Title
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

Private Sub cmdSalir_Click()
    Unload Me
End Sub

' procedimiento que carga los equipos que están fuera de plazo de calibración
Private Sub cargar_avisos_calibracion()
    Dim rs As ADODB.RecordSet
    Dim oEQ As New clsEquipos
    
    Set rs = oEQ.Listado_calibracion_avisos(txtFiltro(0), txtFiltro(1), txtFiltro(2), txtFiltro(3), cmbFiltro(1), cmbFiltro(0), txtFiltro(6), chkbaja.value, opRevision(0).value, opRevision(1).value, opRevision(2).value)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "00000")) ' id_equipo
                .SubItems(1) = rs(1) ' nombre
                .SubItems(2) = rs(2) ' serie
                .SubItems(3) = rs(3) ' proveedor
                .SubItems(4) = rs(4) ' situación
                .SubItems(5) = rs(5) ' familia
                .SubItems(6) = rs(6) ' cond_amb
            End With
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oEQ = Nothing
End Sub

' procedimiento que carga los equipos que están fuera de plazo de verificación
Private Sub cargar_avisos_verificacion()
    Dim rs As ADODB.RecordSet
    Dim oEQ As New clsEquipos
    
    Set rs = oEQ.Listado_verificacion_avisos(txtFiltro(0), txtFiltro(1), txtFiltro(2), txtFiltro(3), cmbFiltro(1), cmbFiltro(0), txtFiltro(6), chkbaja.value, opRevision(0).value, opRevision(1).value, opRevision(2).value)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "00000")) ' id_equipo
                .SubItems(1) = rs(1) ' nombre
                .SubItems(2) = rs(2) ' serie
                .SubItems(3) = rs(3) ' proveedor
                .SubItems(4) = rs(4) ' situación
                .SubItems(5) = rs(5) ' familia
                .SubItems(6) = rs(6) ' cond_amb
            End With
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oEQ = Nothing
End Sub

' procedimiento que carga los equipos que están fuera de plazo de mantenimiento
Private Sub cargar_avisos_mantenimiento()
    Dim rs As ADODB.RecordSet
    Dim oEQ As New clsEquipos
    Dim insertar_aviso As Boolean
    
    Set rs = oEQ.Listado_mantenimiento_avisos(txtFiltro(0), txtFiltro(1), txtFiltro(2), txtFiltro(3), cmbFiltro(1), cmbFiltro(0), txtFiltro(6), chkbaja.value, opRevision(0).value, opRevision(1).value, opRevision(2).value)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "00000")) ' id_equipo
                .SubItems(1) = rs(1) ' nombre
                .SubItems(2) = rs(2) ' serie
                .SubItems(3) = rs(3) ' proveedor
                .SubItems(4) = rs(4) ' situación
                .SubItems(5) = rs(5) ' familia
                .SubItems(6) = rs(6) ' cond_amb
            End With
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oEQ = Nothing
End Sub
