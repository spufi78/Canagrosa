VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEquipos_Planes_Acciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento Acciones"
   ClientHeight    =   8130
   ClientLeft      =   4140
   ClientTop       =   2940
   ClientWidth     =   9165
   Icon            =   "frmEquipos_Planes_Acciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8130
   ScaleWidth      =   9165
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3420
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7245
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1140
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Modificar acción"
      Top             =   7245
      Width           =   1080
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2295
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Eliimnar acción"
      Top             =   7245
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Añadir acción"
      Top             =   7245
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado de acciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5820
      Left            =   0
      TabIndex        =   17
      Top             =   1350
      Width           =   9150
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1260
         MaxLength       =   255
         TabIndex        =   1
         Top             =   4935
         Width           =   7790
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   1260
         MaxLength       =   3
         TabIndex        =   2
         Top             =   5355
         Width           =   1230
      End
      Begin MSComctlLib.ListView lista 
         Height          =   4245
         Left            =   90
         TabIndex        =   18
         Top             =   225
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   7488
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
      Begin MSDataListLib.DataCombo cmbFamiliaAcc 
         Height          =   360
         Left            =   1245
         TabIndex        =   0
         Top             =   4545
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   635
         _Version        =   393216
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Acción"
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
         Height          =   255
         Left            =   90
         TabIndex        =   22
         Top             =   5010
         Width           =   915
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "T. Previsto"
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
         Height          =   255
         Left            =   90
         TabIndex        =   21
         Top             =   5430
         Width           =   1140
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Minutos"
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
         Left            =   2565
         TabIndex        =   20
         Top             =   5430
         Width           =   810
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Familia"
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
         Height          =   255
         Left            =   90
         TabIndex        =   19
         Top             =   4590
         Width           =   915
      End
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
      Height          =   1005
      Left            =   0
      TabIndex        =   12
      Top             =   315
      Width           =   9150
      Begin VB.TextBox txtFiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   7020
         MaxLength       =   3
         TabIndex        =   9
         Top             =   180
         Width           =   915
      End
      Begin VB.TextBox txtFiltro 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   810
         MaxLength       =   255
         TabIndex        =   8
         Top             =   540
         Width           =   4920
      End
      Begin MSDataListLib.DataCombo cmbFiltro 
         Height          =   315
         Left            =   810
         TabIndex        =   7
         Top             =   180
         Width           =   4935
         _ExtentX        =   8705
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
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
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Familia"
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   16
         Top             =   225
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Minutos"
         Height          =   195
         Index           =   1
         Left            =   8010
         TabIndex        =   15
         Top             =   225
         Width           =   555
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "T. Previsto"
         Height          =   240
         Index           =   0
         Left            =   6165
         TabIndex        =   14
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nombre"
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   13
         Top             =   585
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7245
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mantenimiento de Acciones"
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
      Height          =   315
      Index           =   3
      Left            =   15
      TabIndex        =   10
      Top             =   0
      Width           =   9120
   End
End
Attribute VB_Name = "frmEquipos_Planes_Acciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbFiltro_Click(AREA As Integer)
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cargar_combos
    
    Me.Top = 200
    Me.Left = 200
    With lista.ColumnHeaders.Add(, , "Familia", 1500, lvwColumnLeft)
        .Tag = "Familia"
    End With
    With lista.ColumnHeaders.Add(, , "Acción", 6000, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "T. Previsto (Min)", 1450, lvwColumnLeft)
        .Tag = "TPrevisto"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 0, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "FAMILIA_ACC_ID", 0, lvwColumnLeft)
        .Tag = "FAMILIA_ACC_ID"
    End With
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    If datos_correctos() Then
        If MsgBox("Va a insertar la acción. ¿Está seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oAccion As New clsEquipos_planes_Acciones
            
            oAccion.setNOMBRE = txtDatos(0)
            oAccion.setT_PREVISTO = txtDatos(1)
            oAccion.setFAMILIA_ACC_ID = cmbFamiliaAcc.BoundText
            oAccion.Insertar
            cargar_lista
            txtDatos(0).SetFocus
        End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        If datos_correctos() Then
            If MsgBox("Va a modificar la acción. ¿Está seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                Dim oAccion As New clsEquipos_planes_Acciones
                oAccion.setNOMBRE = txtDatos(0)
                oAccion.setT_PREVISTO = txtDatos(1)
                oAccion.setFAMILIA_ACC_ID = cmbFamiliaAcc.BoundText
                oAccion.Modificar (lista.ListItems(lista.SelectedItem.Index).SubItems(4))
                cargar_lista
                Set oAccion = Nothing
            End If
            txtDatos(0).SetFocus
        End If
    Else
        MsgBox "Debe seleccionar la acción que desea modificar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR la acción. ¿Está seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oAccion As New clsEquipos_planes_Acciones
            
            oAccion.Eliminar (lista.ListItems(lista.SelectedItem.Index).SubItems(4))
            cargar_lista
        End If
    Else
        MsgBox "Debe seleccionar la acción que desea eliminar.", vbOKOnly + vbInformation, App.Title
    End If
End Sub

Private Sub cmdImprimir_Click()
'    If lista.ListItems.Count > 0 Then
'        With frmReport
'            .iniciar
'            .informe = "rptEquipos_Planes_Acciones"
'            .criterio = ""
'            .imprimir = False
'            .generar
'            .Visible = True
'        End With
'    End If
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
        txtDatos(0).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        txtDatos(1).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
        cmbFamiliaAcc.BoundText = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
    End If
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 1 ' sólo aquellos controles que requieran un numérico
            If Not ((KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or KeyAscii = 8) Then ' Si no es un número o el "." no se permite
                KeyAscii = 0
            End If
    End Select
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

' Funciones auxiliares del módulo
' -------------------------------
' Procedimiento que carga la lista con las acciones
Public Sub cargar_lista()
    Dim rs As New ADOdb.RecordSet
    Dim oAcciones As New clsEquipos_planes_Acciones
    
    txtDatos(0) = ""
    txtDatos(1) = ""
    cmbFamiliaAcc.BoundText = ""
    lista.ListItems.Clear
    Set rs = oAcciones.Listado(txtFiltro(0), txtFiltro(1), cmbFiltro)
    If rs.RecordCount <> 0 Then
        Do
'            With lista.ListItems.Add(, , rs("NOMBRE"))
'                .SubItems(1) = rs("T_PREVISTO")
'                .SubItems(2) = rs("ID_ACCION")
'                .SubItems(3) = rs("FAMILIA_ACC_ID")
'            End With
            With lista.ListItems.Add(, , rs("FAMILIA"))
                .SubItems(1) = rs("NOMBRE")
                .SubItems(2) = rs("T_PREVISTO")
                .SubItems(3) = rs("FAMILIA_ACC_ID")
                .SubItems(4) = rs("ID_ACCION")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    Set oAcciones = Nothing
    Set rs = Nothing
End Sub

' Función que comprueba si los datos introducidos son correctos
Private Function datos_correctos() As Boolean
    datos_correctos = True
    
    If cmbFamiliaAcc.BoundText = "" Or cmbFamiliaAcc.BoundText = "0" Then ' Familia de acciones
        MsgBox "Debe seleccionar una familia para la acción.", vbCritical, App.Title
        datos_correctos = False
        Exit Function
    End If
    If Len(Trim(txtDatos(0))) = 0 Then
        MsgBox "Debe introducir un nombre para la acción.", vbCritical, App.Title
        datos_correctos = False
        txtDatos(0).SetFocus
        Exit Function
    End If
    If Len(Trim(txtDatos(1))) = 0 Then
        MsgBox "Debe introducir un tiempo previsto para la acción.", vbCritical, App.Title
        datos_correctos = False
        txtDatos(1).SetFocus
        Exit Function
    ElseIf Not IsNumeric(txtDatos(1)) Then
        MsgBox "Debe introducir un tiempo previsto que sea numérico (Minutos).", vbCritical, App.Title
        datos_correctos = False
        txtDatos(1).SetFocus
        Exit Function
    End If
End Function

Private Sub txtFiltro_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    
    oDeco.cargar_combo cmbFiltro, decodificadora.EQ_FAMILIAS_ACCIONES_PLANES_MTO
    oDeco.cargar_combo cmbFamiliaAcc, decodificadora.EQ_FAMILIAS_ACCIONES_PLANES_MTO
    
End Sub
