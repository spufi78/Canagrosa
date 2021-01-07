VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmDocumentator 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ficheros Adjuntos"
   ClientHeight    =   9345
   ClientLeft      =   3045
   ClientTop       =   3495
   ClientWidth     =   13575
   Icon            =   "frmDocumentator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9345
   ScaleWidth      =   13575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
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
      Height          =   1320
      Index           =   1
      Left            =   30
      TabIndex        =   7
      Top             =   555
      Width           =   13500
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   6870
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   960
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   855
         TabIndex        =   0
         Top             =   315
         Width           =   1650
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Height          =   675
         Left            =   12465
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   135
         Width           =   915
      End
      Begin MSDataListLib.DataCombo cmbTipoFiltro 
         Height          =   330
         Left            =   3060
         TabIndex        =   1
         Top             =   330
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Texto"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   405
         Width           =   405
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   2595
         TabIndex        =   8
         Top             =   390
         Width           =   315
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8400
      Width           =   1005
   End
   Begin MSComctlLib.ListView listaCarpetas 
      Height          =   6405
      Left            =   45
      TabIndex        =   5
      Top             =   1920
      Width           =   6480
      _ExtentX        =   11430
      _ExtentY        =   11298
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14603217
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
   Begin MSComctlLib.ListView listaFicheros 
      Height          =   6405
      Left            =   6570
      TabIndex        =   10
      Top             =   1935
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   11298
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
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
      Caption         =   "Ficheros adjuntos: "
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
      TabIndex        =   6
      Top             =   150
      Width           =   1980
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   540
      Left            =   0
      Top             =   0
      Width           =   13560
   End
End
Attribute VB_Name = "frmDocumentator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkTodos_Click()
    If chkTodos.value = Checked Then
        cmbTipoFiltro.Text = ""
        cmbTipoFiltro.Enabled = False
    Else
        cmbTipoFiltro.Enabled = True
    End If
    PresentarDatos
End Sub
Private Sub cmbTipoFiltro_Change()
    PresentarDatos
End Sub

Private Sub cmdLimpiar_Click()
    txtFiltro(0) = ""
    cmbTipoFiltro.Text = ""
    cmbTipoFiltro.Enabled = False
    chkTodos.value = Checked
    PresentarDatos
End Sub
Private Sub PresentarDatos()
    listaFicheros.ListItems.Clear
    listaCarpetas.ListItems.Clear
    Dim rs As ADODB.Recordset
    Dim consulta As String
    Dim filtro As String
    ' nombre
    If Len(txtFiltro(0)) > 4 Then
        filtro = filtro & " AND (CARPETA LIKE '%" & txtFiltro(0) & "%' OR NOMBRE LIKE '%" & txtFiltro(0) & "%')"
    End If
    ' TIPO
    If cmbTipoFiltro.Text <> "" Then
        filtro = filtro & " AND TYPE = '" & cmbTipoFiltro.Text & "'"
    End If
    consulta = "select distinct departamento, anno, ruta, carpeta from aaa where 1 = 1 "
    consulta = consulta & filtro
    consulta = consulta & " LIMIT 250"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Do
            With listaCarpetas.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(1)
                 .SubItems(2) = rs(2)
                 .SubItems(3) = rs(3)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oAdjunto = Nothing
    If listaCarpetas.ListItems.Count > 0 Then
        listaCarpetas_Click
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
'    PresentarDatos
End Sub

Public Sub cabecera()
    With listaCarpetas.ColumnHeaders
        .Add , , "DEPARTAMENTO", 1000, lvwColumnLeft
        .Add , , "AÑO", 550, lvwColumnLeft
        .Add , , "RUTA", 1, lvwColumnLeft
        .Add , , "CARPETA", 4800, lvwColumnLeft
    End With
    With listaFicheros.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "NOMBRE", 4800, lvwColumnLeft
        .Add , , "SIZE", 1000, lvwColumnRight
        .Add , , "TYPE", 900, lvwColumnLeft
    End With
End Sub
Private Sub cargar_combos()
'    Dim oDeco As New clsDecodificadora
'    oDeco.cargar_combo cmbTipoFiltro, decodificadora.ADJUNTOS_TIPOS_DOCUMENTOS
    Dim consulta As String
    consulta = "select distinct type from aaa"
    Set rs = datos_bd(consulta)
    Set cmbTipoFiltro.RowSource = rs
    cmbTipoFiltro.ListField = rs(0).Name
    cmbTipoFiltro.BoundColumn = rs(0).Name
    Set rs = Nothing
End Sub

Private Sub listaCarpetas_Click()
    listaFicheros.ListItems.Clear
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "select id, nombre, size, type from aaa " & _
               " where departamento = " & listaCarpetas.ListItems(listaCarpetas.selectedItem.Index).Text & _
               "   and anno = " & listaCarpetas.ListItems(listaCarpetas.selectedItem.Index).SubItems(1) & _
               "   and ruta = '" & listaCarpetas.ListItems(listaCarpetas.selectedItem.Index).SubItems(2) & "'" & _
               "   and carpeta = '" & listaCarpetas.ListItems(listaCarpetas.selectedItem.Index).SubItems(3) & "'"
'    consulta = consulta & " LIMIT 100"
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Do
            With listaFicheros.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(1)
                 .SubItems(2) = rs(2)
                 .SubItems(3) = rs(3)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub
Private Sub listaFicheros_DblClick()
    Dim iret As Long
    Dim ruta As String
    ruta = listaCarpetas.ListItems(listaCarpetas.selectedItem.Index).SubItems(2)
    ruta = ruta & "\" & listaCarpetas.ListItems(listaCarpetas.selectedItem.Index).SubItems(3)
    ruta = ruta & "\" & listaFicheros.ListItems(listaFicheros.selectedItem.Index).SubItems(1)
    ruta = Replace(ruta, "/", "\")
    iret = ShellExecute(0, vbNullString, ruta, vbNullString, App.Path, 1)
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    PresentarDatos
End Sub
Private Sub mostrar_pdf(DOC As String)
    If DOC <> "" Then
        If UCase(Right(DOC, 3)) = "PDF" Then
            If Dir(DOC) <> "" Then
                pdf1.Visible = True
                cmdMostrar.Visible = True
                pdf1.LoadFile DOC
                pdf1.setShowToolbar False
            End If
        Else
            pdf1.Visible = False
            cmdMostrar.Visible = False
        End If
    End If
End Sub

