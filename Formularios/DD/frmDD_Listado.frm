VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDD_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tipos de Dependencias"
   ClientHeight    =   7770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11925
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDD_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   11925
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10830
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6870
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6870
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6870
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6870
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5985
      Left            =   60
      TabIndex        =   0
      Top             =   855
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   10557
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Establezca la relación de dependencias que se tomarán en las determinaciones de los análisis"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   420
      Width           =   6645
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11250
      Picture         =   "frmDD_Listado.frx":000C
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dependencias de determinaciones"
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
      TabIndex        =   5
      Top             =   120
      Width           =   3645
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   870
      Left            =   0
      Top             =   0
      Width           =   12015
   End
End
Attribute VB_Name = "frmDD_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAnadir_Click()
    frmDD_Detalle.PK_DETERMINACION = 0
    frmDD_Detalle.PK_CAMPO = 0
    frmDD_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a eliminar el tipo de dependencia.", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oTD As New clsTipos_determinacion_dep
        oTD.Eliminar lista.ListItems(lista.SelectedItem.Index).SubItems(3), lista.ListItems(lista.SelectedItem.Index).SubItems(4)
        Set oTD = Nothing
        cargar_lista
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmDD_Detalle.PK_DETERMINACION = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
        frmDD_Detalle.PK_CAMPO = lista.ListItems(lista.SelectedItem.Index).SubItems(4)
        frmDD_Detalle.Show 1
        cargar_lista
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    With lista.ColumnHeaders.Add(, , "Determinacion", 3650, lvwColumnLeft)
        .Tag = "Determinacion"
    End With
    With lista.ColumnHeaders.Add(, , "Campo", 2500, lvwColumnLeft)
        .Tag = "Formula"
    End With
    With lista.ColumnHeaders.Add(, , "Depende", 3650, lvwColumnLeft)
        .Tag = "Depende"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 500, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "ID2", 500, lvwColumnCenter)
        .Tag = "ID2"
    End With
    With lista.ColumnHeaders.Add(, , "ID3", 500, lvwColumnCenter)
        .Tag = "ID3"
    End With
    cargar_lista
End Sub
Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim odd As New clsTipos_determinacion_dep
    Dim oTD As New clsTipos_determinacion
    Dim ocf As New clsFormulas_campos
    Set rs = odd.Listado
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(0))
           .SubItems(1) = rs(1)
           .SubItems(2) = rs(2)
           .SubItems(3) = rs(3)
           .SubItems(4) = rs(4)
           .SubItems(5) = rs(5)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set odd = Nothing
    Set oTD = Nothing
    Set ocf = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_Click()
    If lista.ListItems(lista.SelectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
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
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

