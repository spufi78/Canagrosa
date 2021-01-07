VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmIndicadores_Lista_Campos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento Campos de Indicadores"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   Icon            =   "frmIndicadores_Lista_Campos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   10380
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6030
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6030
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6030
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6030
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6030
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5535
      Left            =   60
      TabIndex        =   0
      Top             =   450
      Width           =   10275
      _ExtentX        =   18124
      _ExtentY        =   9763
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
      Caption         =   "Mantenimiento Campos de Indicadores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   60
      TabIndex        =   1
      Top             =   60
      Width           =   10260
   End
End
Attribute VB_Name = "frmIndicadores_Lista_Campos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdAnadir_Click()
    gindicadores_campos = 0
    frmIndicadores_Campos.Show 1
    cargar_lista
    gindicadores_campos = 0
End Sub

Private Sub cmdeliminar_Click()
    If MsgBox("Va a eliminar el campo : " & lista.ListItems(lista.SelectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oIndicadores_campos As New clsIndicadores_campos
        oIndicadores_campos.Eliminar (lista.ListItems(lista.SelectedItem.Index).SubItems(3))
        Set oIndicadores_campos = Nothing
        cargar_lista
    End If
End Sub
'Private Sub cmdImprimir_Click_old()
'    On Error GoTo fallo
'    Dim i As Integer
'    ' Generamos los datos del listado
'    Dim rs As New ADODB.RecordSet
'    rs.Fields.Append "c1", adChar, 5, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 25, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 25, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 25, adFldUpdatable
'    rs.Open
'    For i = 1 To lista.ListItems.Count
'        rs.AddNew
'        rs("c1") = Left(lista.ListItems(i).SubItems(3), 5)
'        If Trim(lista.ListItems(i)) <> "" Then
'            rs("c2") = Left(lista.ListItems(i), 25)
'        End If
'        If Trim(lista.ListItems(i).SubItems(1)) <> "" Then
'            rs("c3") = Left(lista.ListItems(i).SubItems(1), 25)
'        End If
'        If Trim(lista.ListItems(i).SubItems(2)) <> "" Then
'            rs("c4") = Left(lista.ListItems(i).SubItems(2), 25)
'        End If
'        rs.Update
'    Next
'    ' Generar Listado
'    Dim Listado As New rptListado
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Campos"
'        .Controls("etiqueta4").Caption = "ID"
'        .Controls("etiqueta5").Caption = "Nombre"
'        .Controls("etiqueta10").Caption = "Valor"
'        .Controls("etiqueta11").Caption = "Función"
'    End With
'    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("d1").DataField = rs.Fields("c1").Name
'        .Controls("d2").DataField = rs.Fields("c2").Name
'        .Controls("d3").DataField = rs.Fields("c4").Name
'        .Controls("d4").DataField = rs.Fields("c3").Name
'    End With
'    Listado.Sections("detalle").Controls("d4").Alignment = 0
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & Usuario.getNOMBRE
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Campos"
''    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado.", vbCritical, Err.Description
'End Sub

Private Sub cmdImprimir_Click()
    
    Dim oIndicadores_campos As New clsIndicadores_campos
        
        oIndicadores_campos.Imprimir_Listado
    
    Set oIndicadores_campos = Nothing
    
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        gindicadores_campos = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
        frmIndicadores_Campos.Show 1
        modificar_lista
        gindicadores_campos = 0
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
    cabecera
    cargar_lista
End Sub
Public Sub cargar_lista()
    Dim rs As New ADODB.RecordSet
    Dim oIndicadores_campos As New clsIndicadores_campos
    Set rs = oIndicadores_campos.lista
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = Format(rs(3), "0000")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oIndicadores_campos = Nothing
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

Public Sub modificar_lista()
    Dim rs As New ADODB.RecordSet
    Dim oIndicadores_campos As New clsIndicadores_campos
    Set rs = oIndicadores_campos.Listado_por_Codigo(gindicadores_campos)
    If rs.RecordCount <> 0 Then
        lista.ListItems(lista.SelectedItem.Index).Text = rs(0)
        lista.ListItems(lista.SelectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.SelectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.SelectedItem.Index).SubItems(3) = Format(rs(3), "000")
    End If
    Set oIndicadores_campos = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Nombre", 3000, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Función", 2900, lvwColumnLeft)
        .Tag = "Función"
    End With
    With lista.ColumnHeaders.Add(, , "Valor", 3500, lvwColumnLeft)
        .Tag = "Valor"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 550, lvwColumnCenter)
        .Tag = "ID"
    End With
End Sub
