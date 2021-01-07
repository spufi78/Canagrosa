VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEmpleados_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Empleados"
   ClientHeight    =   8595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11640
   Icon            =   "frmEmpleados_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8595
   ScaleWidth      =   11640
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   10410
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2460
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      Height          =   885
      Left            =   3660
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7680
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   885
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7680
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CheckBox chkfiltro 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por estado"
      Height          =   195
      Left            =   7335
      TabIndex        =   2
      Top             =   8235
      Width           =   1770
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7590
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   13388
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
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
   Begin MSDataListLib.DataCombo cmbestados 
      Height          =   360
      Left            =   7335
      TabIndex        =   1
      Top             =   7785
      Width           =   2070
      _ExtentX        =   3651
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      Object.DataMember      =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmEmpleados_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkfiltro_Click()
    cargar_lista
End Sub

Private Sub cmbestados_Change()
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    gOperario = 0
    frmEmpleados_Gestion.Show 1
    cargar_lista
    lista.SetFocus
End Sub

Private Sub cmdBuscar_Click()
    gOperario = 0
    frmEmpleados_Buscar.Show 1
    If gOperario <> 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i) = gOperario Then
                lista.ListItems(i).Selected = True
                lista.ListItems(i).EnsureVisible
                lista.SetFocus
                Exit Sub
            End If
        Next
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        Dim pos As Integer
        Dim operario As Integer
        If MsgBox("Va a ELIMINAR al Operario " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
            Dim oOperario As New clsEmpleados
            If oOperario.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
                cargar_lista
            End If
            Set oOperario = Nothing
        End If
        lista.SetFocus
    End If
End Sub

'Private Sub cmdImprimir_Click()
'    On Error GoTo fallo
'    Dim i As Integer
'    ' Generamos los datos del listado
'    Dim rs As New ADODB.Recordset
'    rs.Fields.Append "c1", adChar, 5, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 30, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 30, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 15, adFldUpdatable
'    rs.Fields.Append "c5", adChar, 15, adFldUpdatable
'    rs.Open
'    For i = 1 To lista.ListItems.Count
'        rs.AddNew
'        rs("c1") = Left(lista.ListItems(i), 5)
'        If lista.ListItems(i).SubItems(1) <> "" Then
'            rs("c2") = Left(lista.ListItems(i).SubItems(1), 30)
'        Else
'            rs("c2") = ""
'        End If
'        If lista.ListItems(i).SubItems(2) <> "" Then
'            rs("c3") = Left(lista.ListItems(i).SubItems(2), 30)
'        Else
'            rs("c3") = ""
'        End If
'        If lista.ListItems(i).SubItems(3) <> "" Then
'            rs("c4") = Left(lista.ListItems(i).SubItems(3), 15)
'        Else
'            rs("c4") = ""
'        End If
'        If lista.ListItems(i).SubItems(4) <> "" Then
'            rs("c5") = Left(lista.ListItems(i).SubItems(4), 15)
'        Else
'            rs("c5") = ""
'        End If
'        rs.Update
'    Next
'
'    ' Generar Listado
'    Dim Listado As New rptListado
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Operarios"
'    End With
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("d1").DataField = rs.Fields("c1").Name
'        .Controls("d2").DataField = rs.Fields("c2").Name
'        .Controls("d3").DataField = rs.Fields("c3").Name
'        .Controls("d4").DataField = rs.Fields("c4").Name
'        .Controls("d5").DataField = rs.Fields("c5").Name
'    End With
'    ' Pie de Pagina
'    With Listado.Sections("pie")
''MD1        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd/mm/yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & usuario.getNOMBRE
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Operarios"
''    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado de documentos.", vbCritical, Err.Description
'End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
'        If usuario.getPER_MODIFICACION = 0 Then
'            Exit Sub
'        End If
        gOperario = lista.ListItems(lista.SelectedItem.Index)
        frmEmpleados_Gestion.Show 1
        actualizar_lista
        gOperario = 0
        lista.SetFocus
    End If
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me

    Me.Left = 100
    Me.Top = 100
    With lista.ColumnHeaders.Add(, , "Codigo", 600, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 3900, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Direccion", 3900, lvwColumnLeft)
        .Tag = "Direccion"
    End With
    With lista.ColumnHeaders.Add(, , "Telefono", 1400, lvwColumnCenter)
        .Tag = "Telefono"
    End With
    With lista.ColumnHeaders.Add(, , "Movil", 1400, lvwColumnCenter)
        .Tag = "Movil"
    End With
    cargar_estados
    cargar_lista
    permisos
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ocli As New clsEmpleados
    If chkfiltro.Value = Unchecked Or chkfiltro.Value = Checked And cmbestados.BoundText = "" Then
        Set rs = ocli.Listado
    Else
        Set rs = ocli.Listado_Estado(cmbestados.BoundText)
    End If
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_empleado"), "000"))
            .SubItems(1) = rs("nombre")
            If IsNull(rs("direccion")) = False Then
                .SubItems(2) = rs("direccion")
            End If
            If IsNull(rs("telefono")) = False Then
                .SubItems(3) = rs("telefono")
            End If
            If IsNull(rs("movil")) = False Then
                .SubItems(4) = rs("movil")
            Else
                .SubItems(4) = ""
            End If
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set ocli = Nothing
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
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index) <> "" Then
          cmdmodificar.Enabled = True
          cmdEliminar.Enabled = True
        End If
        permisos
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim ocli As New clsEmpleados
    If ocli.cargar(CLng(gOperario)) = True Then
'        lista.ListItems(lista.SelectedItem.Index).Text = gOperario
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = ocli.getNOMBRE
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = ocli.getDIRECCION
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = ocli.getTELEFONO
        lista.ListItems(lista.SelectedItem.Index).SubItems(4) = ocli.getMOVIL
    End If
    Set ocli = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       cmdModificar_Click
    End If
End Sub
Public Sub permisos()
End Sub

Public Sub cargar_estados()
    Dim ooe As New clsEmpleados_Estados
    Set cmbestados.RowSource = ooe.Listado
    cmbestados.ListField = "nombre"
    cmbestados.BoundColumn = "id_estado"
    Set ooe = Nothing
End Sub

