VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBusquedaGeneral 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formulario de búsqueda"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "frmBusquedaGeneral.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   8130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Nuevo"
      Height          =   855
      Left            =   90
      MaskColor       =   &H00E0E0E0&
      Picture         =   "frmBusquedaGeneral.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   7245
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consultar"
      Height          =   855
      Left            =   4410
      Picture         =   "frmBusquedaGeneral.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7245
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Default         =   -1  'True
      Height          =   855
      Left            =   5640
      Picture         =   "frmBusquedaGeneral.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7230
      Width           =   1185
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Cancel          =   -1  'True
      Caption         =   "ESC-Salir"
      Height          =   855
      Left            =   6870
      Picture         =   "frmBusquedaGeneral.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7230
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criterios de búsqueda"
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
      Height          =   780
      Left            =   45
      TabIndex        =   2
      Top             =   765
      Width           =   8055
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
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
         Height          =   330
         Left            =   960
         TabIndex        =   3
         Top             =   270
         Width           =   6915
      End
      Begin VB.Label lblCampo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Filtro"
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
         Left            =   180
         TabIndex        =   4
         Top             =   330
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5580
      Left            =   30
      TabIndex        =   5
      Top             =   1605
      Width           =   8085
      _ExtentX        =   14261
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
      Caption         =   "lbltitulo"
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
      TabIndex        =   1
      Top             =   90
      Width           =   765
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7560
      Picture         =   "frmBusquedaGeneral.frx":2BF2
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "lblsubtitulo"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   390
      Width           =   735
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   -30
      Width           =   8190
   End
End
Attribute VB_Name = "frmBusquedaGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TABLA As String
Public DESCRIPCION As String
Public PK As String
Public CAMPO As String
Public PK_SALIDA As Long
Public FK_CAMPO As String
Public FK_VALOR As Long
Public MUESTRA_DETALLE As Boolean
Public FILTRO As String
Public QUERY As String
Public ORDENACION As String
Public FORMULARIO As Form
'Public conn As ADODB.Connection

Private Sub cmdcancel_Click()
    PK_SALIDA = 0
    Unload Me
End Sub

Private Sub cmdmodificar_Click()
    If lista.ListItems.Count > 0 Then
        FORMULARIO.PK = CLng(lista.ListItems(lista.SelectedItem.Index))
        FORMULARIO.Show 1
    End If
End Sub

Private Sub cmdNuevo_Click()
    FORMULARIO.PK = 0
    FORMULARIO.Show 1
End Sub

Private Sub cmdok_Click()
    If lista.ListItems.Count > 0 Then
        PK_SALIDA = CLng(lista.ListItems(lista.SelectedItem.Index).Text)
    Else
        PK_SALIDA = 0
    End If
    Unload Me
End Sub
Private Sub Form_Load()
    If TABLA <> "" Then
        lbltitulo = "Búsqueda de " & DESCRIPCION
        lblsubtitulo = "Específique un críterio para la búsqueda de " & DESCRIPCION
        cabecera
        cargar_lista
    End If
    If MUESTRA_DETALLE = False Then
        cmdNuevo.Visible = False
        cmdmodificar.Visible = False
    End If
End Sub

Public Sub cargar_lista()
    On Error GoTo fallo
    Dim CONSULTA As String
    Dim rs As ADODB.Recordset
    lista.ListItems.Clear
    Dim s As String
    If QUERY <> "" Then
        CONSULTA = QUERY & " AND " & CAMPO & " LIKE '%" & txtDatos & "%' " & " ORDER BY " & CAMPO & " " & ORDENACION
    Else
        If FILTRO <> "" Then
            s = " AND " & FILTRO
        End If
        If FK_CAMPO <> "" And FK_VALOR <> 0 Then
            CONSULTA = "SELECT " & PK & "," & CAMPO & _
                       "  FROM " & TABLA & _
                       " WHERE " & FK_CAMPO & " = " & FK_VALOR & _
                       "   AND " & CAMPO & " LIKE '%" & txtDatos & "%' " & _
                       s & _
                       " ORDER BY " & CAMPO & " " & ORDENACION
        Else
            CONSULTA = "SELECT " & PK & "," & CAMPO & _
                       "  FROM " & TABLA & _
                       " WHERE " & CAMPO & " LIKE '%" & txtDatos & "%' " & _
                       s & _
                       " ORDER BY " & CAMPO & " " & ORDENACION
        End If
    End If
    Me.MousePointer = 11
    Set rs = datos_bd(CONSULTA)
    If rs.RecordCount > 0 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , rs.Fields(0))
                .SubItems(1) = rs.Fields(1)
            End With
            rs.MoveNext
        Wend
    End If
    Set rs = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar en la tabla : " & TABLA, vbCritical, App.Title
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdok_Click
    End If
End Sub

Private Sub txtDatos_Change()
    cargar_lista
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , PK, 1, lvwColumnLeft
        .Add , , DESCRIPCION, lista.Width - 350, lvwColumnLeft
    End With
End Sub
