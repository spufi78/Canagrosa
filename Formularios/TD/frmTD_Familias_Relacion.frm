VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmTD_Familias_Relacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relación de determinaciones de una familia"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13485
   Icon            =   "frmTD_Familias_Relacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   13485
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   12300
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   30
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7680
      Width           =   1155
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7290
      Left            =   30
      TabIndex        =   0
      Top             =   330
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   12859
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
      BackColor       =   14609914
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
   Begin MSComctlLib.ListView lista2 
      Height          =   7290
      Left            =   6750
      TabIndex        =   4
      Top             =   330
      Width           =   6690
      _ExtentX        =   11800
      _ExtentY        =   12859
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Determinaciones asignadas"
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
      Index           =   0
      Left            =   6750
      TabIndex        =   6
      Top             =   0
      Width           =   6690
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Determinaciones sin asignar"
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
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   6690
   End
End
Attribute VB_Name = "frmTD_Familias_Relacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public familia As Integer

Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        Dim oTD As New clsTipos_determinacion
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                oTD.Modificar_Familia lista.ListItems(i), familia
            End If
        Next
        cargar_lista
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdanadir_Click of Formulario frmTD_Familias"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo cmdEliminar_Click_Error
    If lista2.ListItems.Count > 0 Then
        Dim i As Integer
        Dim oTD As New clsTipos_determinacion
        For i = 1 To lista2.ListItems.Count
            If lista2.ListItems(i).Checked = True Then
                oTD.Modificar_Familia lista2.ListItems(i), 0
            End If
        Next
        cargar_lista
    End If
   On Error GoTo 0
   Exit Sub

cmdEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdeliminar_Click of Formulario frmTD_Familias"
End Sub

Private Sub Form_Activate()
    log (Me.Name)
    cargar_botones Me
    If familia <> 0 Then
        Dim oFamilia As New clsTipos_determinacion_familias
        oFamilia.Carga (familia)
        Me.Caption = "Listado de determinaciones de la FAMILIA : " & oFamilia.getNOMBRE
        Label1(0) = Me.Caption
        cargar_lista
        Me.SetFocus
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    Me.Left = 100
    Me.top = 100
    With lista.ColumnHeaders.Add(, , "ID", 900, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 4000, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "PNT", 1500, lvwColumnCenter)
        .Tag = "PNT"
    End With
    With lista2.ColumnHeaders.Add(, , "ID", 900, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista2.ColumnHeaders.Add(, , "Nombre", 4000, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista2.ColumnHeaders.Add(, , "PNT", 1500, lvwColumnCenter)
        .Tag = "PNT"
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oFamilia As New clsTipos_determinacion
    Set rs = oFamilia.lista("", "", "", "", "", 0)
    lista.ListItems.Clear
    lista2.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            If rs(7) = familia Then
                    With lista2.ListItems.Add(, , Format(rs(3), "0000"))
                     .SubItems(1) = rs(0)
                     .SubItems(2) = rs(2)
                    End With
            ElseIf rs(7) = 0 Then
                    With lista.ListItems.Add(, , Format(rs(3), "0000"))
                     .SubItems(1) = rs(0)
                     .SubItems(2) = rs(2)
                    End With
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oFamilia = Nothing
End Sub

Private Sub lista2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista2.ListItems.Count > 0 Then
     lista2.SortKey = ColumnHeader.Index - 1
     If lista2.SortOrder = 0 Then
        lista2.SortOrder = 1
     Else
        lista2.SortOrder = 0
     End If
     lista2.Sorted = True
   End If

End Sub

Private Sub lista2_DblClick()
    If lista2.ListItems.Count > 0 Then
        frmTD_Detalle.PK = lista2.ListItems(lista2.selectedItem.Index)
        frmTD_Detalle.Show 1
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
    If lista.ListItems.Count > 0 Then
        frmTD_Detalle.PK = lista.ListItems(lista.selectedItem.Index)
        frmTD_Detalle.Show 1
    End If
End Sub
