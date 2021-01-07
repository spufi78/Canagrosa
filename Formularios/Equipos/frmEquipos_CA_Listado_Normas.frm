VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEquipos_CA_Listado_Normas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "EQUIPOS - Listado de NORMAS CONTROLADAS"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   13875
   Begin VB.CommandButton cmdVincular_a_equipo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vincular"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   7200
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12735
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7170
      Width           =   1050
   End
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
      Height          =   1320
      Left            =   0
      TabIndex        =   0
      Top             =   285
      Width           =   13785
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1140
         MaxLength       =   255
         TabIndex        =   7
         Top             =   960
         Width           =   4020
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   870
         Left            =   11520
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   270
         Width           =   1050
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   870
         Left            =   12645
         Picture         =   "frmEquipos_CA_Listado_Normas.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   270
         Width           =   1050
      End
      Begin VB.CheckBox chkNADCAP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "NADCAP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10395
         TabIndex        =   4
         Top             =   630
         Width           =   960
      End
      Begin VB.CheckBox chkENAC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENAC"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10395
         TabIndex        =   3
         Top             =   270
         Width           =   810
      End
      Begin VB.CheckBox chkEQA 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "EQA"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10395
         TabIndex        =   2
         Top             =   990
         Width           =   750
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   6255
         MaxLength       =   255
         TabIndex        =   1
         Top             =   960
         Width           =   3750
      End
      Begin MSDataListLib.DataCombo cmbtipos 
         Height          =   315
         Left            =   1140
         TabIndex        =   8
         Top             =   225
         Width           =   4050
         _ExtentX        =   7144
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
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   315
         Left            =   1140
         TabIndex        =   9
         Top             =   600
         Width           =   4035
         _ExtentX        =   7117
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
      Begin MSDataListLib.DataCombo cmbSectores 
         Height          =   315
         Left            =   6255
         TabIndex        =   10
         Top             =   585
         Width           =   3780
         _ExtentX        =   6668
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
      Begin MSDataListLib.DataCombo cmbSubtipo 
         Height          =   315
         Left            =   6255
         TabIndex        =   11
         Top             =   225
         Width           =   3780
         _ExtentX        =   6668
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
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   17
         Top             =   960
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pdte.Estado"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   16
         Top             =   645
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   285
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sector"
         Height          =   195
         Index           =   3
         Left            =   5535
         TabIndex        =   14
         Top             =   645
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   4
         Left            =   5535
         TabIndex        =   13
         Top             =   990
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subtipo"
         Height          =   195
         Index           =   5
         Left            =   5535
         TabIndex        =   12
         Top             =   285
         Width           =   540
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5490
      Left            =   0
      TabIndex        =   19
      Top             =   1620
      Width           =   13785
      _ExtentX        =   24315
      _ExtentY        =   9684
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Listado de NORMAS CONTROLADAS"
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
      Height          =   270
      Index           =   3
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   13770
   End
End
Attribute VB_Name = "frmEquipos_CA_Listado_Normas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'E0035-I
Public FK_EQUIPO As Long
'E0035-F
'-----------------------------------------

Private Sub cmbestados_Change()
    cmdBuscar_Click
End Sub
Private Sub cmbSubtipo_Change()
    cmdBuscar_Click
End Sub
Private Sub cmbtipos_Change()
    cmdBuscar_Click
End Sub
Private Sub cmbsectores_Change()
    cmdBuscar_Click
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdLimpiar_Click()
    txtDatos(1) = ""
    txtDatos(0) = ""
    cmbtipos.Text = ""
    cmbSectores.Text = ""
    cmbestados.Text = ""
    cmbSubtipo.Text = ""
    cmdBuscar_Click
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 100
    Me.Left = 100
    cabecera
    cargar_botones Me
    cargar_combos
    cargar_lista
    'permisos
    If USUARIO.getUSUARIO = "julio" Then
        'E0034-I
        ' Se ha eliminado el botón, este listado será sólo para seleccionar y vincular al equipo
        'cmdCargar.Visible = True
        'E0034-F
    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Norma", 4000, lvwColumnLeft
        .Add , , "Tipo", 2000, lvwColumnLeft
        .Add , , "SubTipo", 2000, lvwColumnLeft
        .Add , , "Código", 2100, lvwColumnCenter
        .Add , , "Edición", 1000, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Estado", 1200, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.RecordSet
    Dim oca_normas As New clsCa_normas
    lista.ListItems.Clear
    Dim tipo As String
    Dim SECTOR As String
    Dim estado As String
    Dim nombre As String
    Dim CODIGO As String
    Dim SUBTIPO As String
    If cmbtipos.Text = "" Then
        tipo = 0
    Else
        tipo = cmbtipos.BoundText
    End If
    If cmbSectores.Text = "" Then
        SECTOR = 0
    Else
        SECTOR = cmbSectores.BoundText
    End If
    If cmbestados.Text = "" Then
        estado = 0
    Else
        estado = cmbestados.BoundText
    End If
    If cmbSubtipo.Text = "" Then
        SUBTIPO = 0
    Else
        SUBTIPO = cmbSubtipo.BoundText
    End If
    nombre = txtDatos(1)
    CODIGO = txtDatos(0)
    Set rs = oca_normas.Listado(tipo, SECTOR, estado, nombre, CODIGO, SUBTIPO, chkENAC.value, chkNADCAP.value, chkEQA.value)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = Format(rs(6), "dd-mm-yyyy")
             .SubItems(7) = rs(7)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    'Set oca_documentos = Nothing 'ERROR
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
Public Sub cargar_combos()
    Dim oDECODIFICADORA As New clsDecodificadora
    oDECODIFICADORA.Cargar_Combo cmbtipos, decodificadora.CA_NORMAS_TIPOS
    oDECODIFICADORA.Cargar_Combo cmbSectores, decodificadora.CA_NORMAS_SECTORES
    oDECODIFICADORA.Cargar_Combo cmbestados, decodificadora.CA_NORMAS_ESTADOS
    oDECODIFICADORA.Cargar_Combo cmbSubtipo, decodificadora.CA_NORMAS_SUBTIPOS
End Sub
'Public Sub permisos()
'    If Not USUARIO.getPER_DOCUMENTACION_CALIDAD Then
'        'cmdAnadir.Enabled = False
'        'cmdModificar.Enabled = False
'        'cmdEliminar.Enabled = False
'    End If
'End Sub
Private Sub txtdatos_Change(Index As Integer)
    cmdBuscar_Click
End Sub
Private Sub lista_DblClick()
    Dim oEquipos_Normas As New clsEquipos_Normas

    oEquipos_Normas.setNORMA_ID = lista.ListItems(lista.SelectedItem.Index).Text ' Este es el ID_NORMA
    oEquipos_Normas.setEQUIPO_ID = FK_EQUIPO
    Call oEquipos_Normas.Insertar    ' Se vincula la norma al equipo
    Unload Me
    frmEquipos_Detalle.Show
End Sub
