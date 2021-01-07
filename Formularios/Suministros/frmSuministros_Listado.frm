VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmSuministros_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tipos de Suministros"
   ClientHeight    =   8355
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12195
   Icon            =   "frmSuministros_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8355
   ScaleWidth      =   12195
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   4410
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7425
      Width           =   1050
   End
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
      Height          =   825
      Left            =   45
      TabIndex        =   8
      Top             =   630
      Width           =   12120
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   4815
         TabIndex        =   13
         Top             =   315
         Width           =   2310
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   900
         TabIndex        =   10
         Top             =   315
         Width           =   2310
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   645
         Left            =   11115
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   915
      End
      Begin VB.CheckBox chkAnulados 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Anulados"
         Height          =   240
         Left            =   8010
         TabIndex        =   14
         Top             =   360
         Width           =   1860
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Producto"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   12
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   0
         Left            =   3555
         TabIndex        =   11
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdReactivar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Reactivar"
      Height          =   870
      Left            =   3330
      Picture         =   "frmSuministros_Listado.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7425
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11070
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7425
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7425
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1170
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7425
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   870
      Left            =   2250
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7425
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5895
      Left            =   45
      TabIndex        =   0
      Top             =   1485
      Width           =   12120
      _ExtentX        =   21378
      _ExtentY        =   10398
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
      Caption         =   "Listado de Tipos de Suministros"
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
      TabIndex        =   6
      Top             =   30
      Width           =   3360
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado completo de tipos de Suministros"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   5
      Top             =   330
      Width           =   2865
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   12150
   End
End
Attribute VB_Name = "frmSuministros_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdImprimir_Click()
   On Error GoTo cmdImprimir_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    Dim cadena As String
    Me.MousePointer = vbHourglass
    Dim rs As New ADODB.Recordset
    rs.Fields.Append "c1", adChar, 250, adFldUpdatable
    rs.Fields.Append "c2", adChar, 250, adFldUpdatable
    rs.Fields.Append "c3", adChar, 250, adFldUpdatable
    rs.Fields.Append "c4", adChar, 250, adFldUpdatable
    rs.Fields.Append "c5", adChar, 250, adFldUpdatable
    rs.Open
    
    Dim i As Integer

    For i = 1 To lista.ListItems.Count
         rs.AddNew
         rs("c1") = lista.ListItems(i).Text
         rs("c2") = lista.ListItems(i).SubItems(1)
         rs("c3") = lista.ListItems(i).SubItems(2)
         rs("c4") = lista.ListItems(i).SubItems(3)
         rs("c5") = lista.ListItems(i).SubItems(4)
         rs.Update
     Next i
     
     Dim XLA As excel.Application
     Dim XLW As excel.Workbook
     Dim XLS As excel.Worksheet
     
     Set XLA = New excel.Application
     Set XLW = XLA.Workbooks.Add
     Set XLS = XLW.Worksheets(1)
     XLW.Worksheets(3).Delete
     XLW.Worksheets(2).Delete
     XLW.Worksheets(1).Name = "Listado de Suministros"

     'Cabecera
     With XLS.Range("A1:E1")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlCenter
         .WrapText = True
         .Orientation = 0
         .AddIndent = False
         .IndentLevel = 0
         .ShrinkToFit = False
         .ReadingOrder = xlContext
         .MergeCells = False
     End With
     With XLS.Range("A1:E1").Interior
         .Pattern = xlSolid
         .PatternColorIndex = xlAutomatic
         .color = &HC0C0FF
     End With
     With XLS.Range("A1:E1").Borders
         .LineStyle = vbSolid
     End With
     
     XLS.Range("A1:A1").ColumnWidth = 15
     XLS.Range("B1:B1").ColumnWidth = 40
     XLS.Range("C1:C1").ColumnWidth = 40
     XLS.Range("D1:D1").ColumnWidth = 40
     XLS.Range("E1:E1").ColumnWidth = 15
     XLS.Cells(1, 1) = "ID"
     XLS.Cells(1, 2) = "Producto"
     XLS.Cells(1, 3) = "Procedimiento"
     XLS.Cells(1, 4) = "Reactivo"
     XLS.Cells(1, 5) = "Anulado"

     i = 2
     If rs.RecordCount > 0 Then
       rs.MoveFirst
       Do
         XLS.Cells(i, 1) = ClrStr(rs("c1"), False, True, True)
         XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True)
         XLS.Cells(i, 3) = ClrStr(rs("c3"), False, True, True)
         XLS.Cells(i, 4) = ClrStr(rs("c4"), False, True, True)
         XLS.Cells(i, 5) = ClrStr(rs("c5"), False, True, True)
         i = i + 1
         XLS.Range("A" & i).EntireRow.Insert
         rs.MoveNext
       Loop Until rs.EOF
     End If
     Me.MousePointer = vbNormal
     XLA.Visible = True
     Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cmdImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmFormacion_Listado"

End Sub
Private Sub chkAnulados_Click()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmSuministros_Tipos.PK = 0
    frmSuministros_Tipos.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a anular el suministro : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oSuministro As New clsSuministros_tipos
            If oSuministro.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmSuministros_Tipos.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmSuministros_Tipos.Show 1
        actualizar_lista
    End If
End Sub

Private Sub cmdReactivar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a REACTIVAR el suministro : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oSuministro As New clsSuministros_tipos
            If oSuministro.Reactivar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "ID", 700, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Producto", 5000, lvwColumnLeft)
        .Tag = "Producto"
    End With
    With lista.ColumnHeaders.Add(, , "Procedimiento", 2600, lvwColumnCenter)
        .Tag = "Procedimiento"
    End With
    With lista.ColumnHeaders.Add(, , "Reactivo", 2700, lvwColumnCenter)
        .Tag = "Reactivo"
    End With
    With lista.ColumnHeaders.Add(, , "Anulado", 700, lvwColumnCenter)
        .Tag = "Anulado"
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oSuministro As New clsSuministros_tipos
    lista.ListItems.Clear
    Set rs = oSuministro.Listado(txtfiltro(0), txtfiltro(1), chkAnulados.Value)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             If rs(4) = 0 Then
                .SubItems(4) = ""
             Else
                .SubItems(4) = "X"
             End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oSuministro = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
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
Public Sub actualizar_lista()
    Dim oSuministro As New clsSuministros_tipos
    Dim rs As ADODB.Recordset
    Set rs = oSuministro.Listado_ID(CLng(lista.ListItems(lista.selectedItem.Index).Text))
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs(3)
    End If
    Set oSuministro = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
