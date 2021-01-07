VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmComerciales_ListadoClientes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relación de obras de un comercial"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13485
   Icon            =   "frmComerciales_ListadoClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   13485
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7680
      Width           =   2235
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   12330
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7650
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
      Width           =   2715
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
   Begin MSComctlLib.ListView lcomercial 
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Obras"
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
      Caption         =   "Obras sin comercial asignado"
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
Attribute VB_Name = "frmComerciales_ListadoClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public COMERCIAL As Integer

Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        Dim oObra As New clsObras
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                oObra.modificar_comercial lista.ListItems(i), COMERCIAL
            End If
        Next
        cargar_lista
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdanadir_Click of Formulario frmComerciales_ListadoClientes"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo cmdEliminar_Click_Error
    If lcomercial.ListItems.Count > 0 Then
        Dim i As Integer
        Dim oObra As New clsObras
        Dim algo As Boolean
        algo = False
        For i = 1 To lcomercial.ListItems.Count
            If lcomercial.ListItems(i).Checked = True Then
                algo = True
                oObra.modificar_comercial lcomercial.ListItems(i), 0
            End If
        Next
        If Not algo Then
            MsgBox "Marque alguna obra para eliminar.", vbExclamation, App.Title
        Else
            cargar_lista
        End If
    End If
   On Error GoTo 0
   Exit Sub

cmdEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdeliminar_Click of Formulario frmComerciales_ListadoClientes"
End Sub

Private Sub Form_Activate()
    log (Me.Name)
    cargar_botones Me
    If COMERCIAL <> 0 Then
        Dim oComercial As New clsComercial
        oComercial.Cargar (COMERCIAL)
        Me.Caption = "Listado de clientes del comercial : " & oComercial.getNOMBRE
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
    Me.Top = 100
    With lista.ColumnHeaders.Add(, , "Codigo", 1200, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 3700, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Provincia", 1500, lvwColumnCenter)
        .Tag = "Provincia"
    End With
    With lcomercial.ColumnHeaders.Add(, , "Codigo", 1200, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lcomercial.ColumnHeaders.Add(, , "Nombre", 3700, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lcomercial.ColumnHeaders.Add(, , "Provincia", 1500, lvwColumnCenter)
        .Tag = "Provincia"
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oObra As New clsObras
    Set rs = oObra.Listado_Comercial
    lista.ListItems.Clear
    lcomercial.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            If rs(1) = COMERCIAL Then
                    With lcomercial.ListItems.Add(, , Format(rs(0), "0000"))
                     .SubItems(1) = rs(2)
                     If Not IsNull(rs(3)) Then
                         .SubItems(2) = rs(3)
                     End If
                    End With
            ElseIf rs(1) = 0 Then
                    With lista.ListItems.Add(, , Format(rs(0), "0000"))
                     .SubItems(1) = rs(2)
                     If Not IsNull(rs(3)) Then
                     .SubItems(2) = rs(3)
                     End If
                    End With
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oObra = Nothing
End Sub

Private Sub lcomercial_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lcomercial.ListItems.Count > 0 Then
     lcomercial.SortKey = ColumnHeader.Index - 1
     If lcomercial.SortOrder = 0 Then
        lcomercial.SortOrder = 1
     Else
        lcomercial.SortOrder = 0
     End If
     lcomercial.Sorted = True
   End If

End Sub

Private Sub lcomercial_DblClick()
    If lcomercial.ListItems.Count > 0 Then
'        gcliente = lcomercial.ListItems(lcomercial.SelectedItem.Index)
        frmClientes.pk = lcomercial.ListItems(lcomercial.SelectedItem.Index)
        frmClientes.Show 1
        gcliente = 0
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
        frmClientes.pk = lista.ListItems(lista.SelectedItem.Index)
        frmClientes.Show 1
        gcliente = 0
    End If
End Sub
