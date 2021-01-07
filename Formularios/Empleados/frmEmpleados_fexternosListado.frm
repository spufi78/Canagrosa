VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEmpleados_fexternosListado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Formadores externos"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11685
   Icon            =   "frmEmpleados_fexternosListado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8205
   ScaleWidth      =   11685
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
      Height          =   645
      Left            =   45
      TabIndex        =   6
      Top             =   630
      Width           =   11580
      Begin VB.TextBox txtb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   0
         Left            =   870
         TabIndex        =   9
         Top             =   225
         Width           =   2580
      End
      Begin VB.TextBox txtb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   1
         Left            =   4320
         TabIndex        =   8
         Top             =   225
         Width           =   3075
      End
      Begin VB.TextBox txtb 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   2
         Left            =   8595
         TabIndex        =   7
         Top             =   225
         Width           =   2625
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CIF"
         Height          =   195
         Index           =   1
         Left            =   3735
         TabIndex        =   11
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   195
         Index           =   2
         Left            =   7695
         TabIndex        =   10
         Top             =   270
         Width           =   630
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7290
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7290
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7290
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7290
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5955
      Left            =   45
      TabIndex        =   0
      Top             =   1320
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   10504
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
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
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11070
      Picture         =   "frmEmpleados_fexternosListado.frx":030A
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Formadores Externos"
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
      TabIndex        =   13
      Top             =   180
      Width           =   3390
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre el formador para ver el detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   2
      Left            =   3960
      TabIndex        =   1
      Top             =   7605
      Width           =   4275
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   11685
   End
End
Attribute VB_Name = "frmEmpleados_fexternosListado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAnadir_Click()
    frmEmpleados_fexternos.PK = 0
    frmEmpleados_fexternos.Show 1
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdeliminar_Click()
    If MsgBox("Va a ELIMINAR al formador " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim oF As New clsEmpleados_fexternos
        oF.setID_FEXTERNO = lista.ListItems(lista.SelectedItem.Index)
        If oF.Eliminar = True Then
            cargar_lista
        End If
        Set oF = Nothing
    End If
End Sub
Private Sub cmdModificar_Click()
    frmEmpleados_fexternos.PK = lista.ListItems(lista.SelectedItem.Index)
    frmEmpleados_fexternos.Show 1
    cargar_lista
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 100
    Me.Left = 100
    With lista.ColumnHeaders
        .Add , , "Codigo", 600, lvwColumnLeft
        .Add , , "Nombre", 4400, lvwColumnLeft
        .Add , , "Direccion", 4400, lvwColumnLeft
        .Add , , "Telefono", 1500, lvwColumnCenter
    End With
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.RecordSet
    Dim oF As New clsEmpleados_fexternos
    Set rs = oF.Listado(txtb(0), txtb(1), txtb(2))
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_fexterno"), "000"))
            .SubItems(1) = rs("nombre")
            If IsNull(rs("direccion")) = False Then
                .SubItems(2) = rs("direccion")
            End If
            If IsNull(rs("telefono")) = False Then
                .SubItems(3) = rs("telefono")
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oF = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
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
Private Sub lista_Click()
    If lista.ListItems(lista.SelectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim oF As New clsEmpleados_fexternos
    If oF.Carga(CLng(lista.ListItems(lista.SelectedItem.Index).Text)) = True Then
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = oF.getNOMBRE
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = oF.getDIRECCION
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = oF.getTELEFONO
    End If
    Set oF = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub txtb_Change(Index As Integer)
    cargar_lista
End Sub
