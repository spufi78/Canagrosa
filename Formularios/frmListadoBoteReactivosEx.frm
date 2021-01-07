VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListadoBoteReactivosEx 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Botes de Reactivos Externos"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9855
   Icon            =   "frmListadoBoteReactivosEx.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   9855
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   3285
      Picture         =   "frmListadoBoteReactivosEx.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6435
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   2210
      Picture         =   "frmListadoBoteReactivosEx.frx":2F6C
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6435
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   1135
      Picture         =   "frmListadoBoteReactivosEx.frx":3276
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6435
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   60
      Picture         =   "frmListadoBoteReactivosEx.frx":4140
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6435
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   8730
      Picture         =   "frmListadoBoteReactivosEx.frx":5182
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6420
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5970
      Left            =   45
      TabIndex        =   0
      Top             =   390
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   10530
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
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Listado de Botes de Reactivos Externos"
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
      Height          =   330
      Index           =   4
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   9735
   End
End
Attribute VB_Name = "frmListadoBoteReactivosEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdAnadir_Click()
    gbotereactivoex = 0
    frmBoteReactivoEx.Show 1
    cargar_lista
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a ELIMINAR el bote " & lista.ListItems(lista.SelectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
        Dim obe As New clsTipos_bote_ex
        If obe.Eliminar(lista.ListItems(lista.SelectedItem.Index)) = True Then
            cargar_lista
        End If
        Set obe = Nothing
    End If
End Sub

Private Sub cmdImprimir_Click()
    On Error GoTo fallo
    Dim i As Integer
    ' Generamos los datos del listado
    Dim rs As New ADODB.Recordset
    rs.Fields.Append "c1", adChar, 5, adFldUpdatable
    rs.Fields.Append "c2", adChar, 150, adFldUpdatable
    rs.Fields.Append "c3", adChar, 150, adFldUpdatable
    rs.Fields.Append "c4", adChar, 150, adFldUpdatable
    rs.Open
    For i = 1 To lista.ListItems.Count
        rs.AddNew
        rs("c1") = lista.ListItems(i)
        If Trim(lista.ListItems(i).SubItems(1)) <> "" Then
            rs("c2") = lista.ListItems(i).SubItems(1)
        End If
        If Trim(lista.ListItems(i).SubItems(2)) <> "" Then
            rs("c3") = lista.ListItems(i).SubItems(2)
        End If
        If Trim(lista.ListItems(i).SubItems(3)) <> "" Then
            rs("c4") = lista.ListItems(i).SubItems(3)
        End If
        rs.Update
    Next
    ' Generar Listado
    Dim Listado As New rptListado
    ' Cabecera
    With Listado.Sections("cabecera")
        .Controls("titulo").Caption = "Listado de Botes Externos"
        .Controls("etiqueta4").Caption = "ID"
        .Controls("etiqueta5").Caption = "Código"
        .Controls("etiqueta10").Caption = "Reactivo"
        .Controls("etiqueta11").Caption = "Proveedor"
    End With
    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
    'Detalle
    With Listado.Sections("detalle")
        .Controls("d1").DataField = rs.Fields("c1").Name
        .Controls("d2").DataField = rs.Fields("c2").Name
        .Controls("d3").DataField = rs.Fields("c3").Name
        .Controls("d4").DataField = rs.Fields("c4").Name
    End With
    
    ' Pie de Pagina
    With Listado.Sections("pie")
        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
        .Controls("pie2").Caption = "Impreso por : " & EMPLEADO.getNOMBRE
    End With
    Set Listado.DataSource = rs
    Listado.Caption = "Listado de Botes Externos"
'    Listado.WindowState = vbMaximized
    Listado.Show
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox "Error al generar el listado.", vbCritical, Err.Description
End Sub

Private Sub cmdModificar_Click()
    gbotereactivoex = lista.ListItems(lista.SelectedItem.Index)
    frmBoteReactivoEx.Show 1
    actualizar_lista
    gbotereactivoex = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 100
    Me.Top = 100
    With lista.ColumnHeaders.Add(, , "Id", 600, lvwColumnLeft)
        .Tag = "Id"
    End With
    With lista.ColumnHeaders.Add(, , "Codigo", 1300, lvwColumnLeft)
        .Tag = "Codigo"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo de Reactivo", 3800, lvwColumnLeft)
        .Tag = "Tipo de Reactivo"
    End With
    With lista.ColumnHeaders.Add(, , "Proveedor", 3800, lvwColumnLeft)
        .Tag = "Proveedor"
    End With
    cargar_lista
End Sub

Public Sub cargar_lista()
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    Dim otbe As New clsTipos_bote_ex
    Set rs = otbe.Listado
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs(0), "0000"))
            .SubItems(1) = rs(1)
            .SubItems(2) = rs(2)
            .SubItems(3) = rs(3)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set otbe = Nothing
    lista_Click
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
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
    Else
      cmdModificar.Enabled = False
      cmdEliminar.Enabled = False
    End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdModificar_Click
    End If
End Sub

Public Sub actualizar_lista()
    Dim ocli As New clsTipos_bote_ex
    Dim otr As New clsTipos_reactivo_ex
    Dim oprov As New clsProveedor
    If ocli.Cargar(CLng(gbotereactivoex)) = True Then
        otr.Cargar (ocli.getTIPO_REACTIVO_EX_ID)
        oprov.Carga (ocli.getPROVEEDOR_ID)
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = ocli.getCODIGO
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = otr.getNOMBRE
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = oprov.getNOMBRE
    End If
    Set ocli = Nothing
    lista_Click
End Sub
