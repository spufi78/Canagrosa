VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCa_Filtro_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de documentos para Auditoria"
   ClientHeight    =   9600
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   10200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkMTL 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "MTL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   8910
      Width           =   6180
   End
   Begin VB.CheckBox chkNADCAP 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "NADCAP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   11
      Top             =   8640
      Width           =   6180
   End
   Begin VB.CheckBox chkENAC 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "ENAC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   10
      Top             =   8370
      Width           =   6120
   End
   Begin VB.CheckBox chkEQA 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "EQA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   135
      TabIndex        =   9
      Top             =   9180
      Width           =   6150
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todos"
      Height          =   330
      Left            =   5130
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7965
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todos"
      Height          =   330
      Left            =   6525
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7965
      Width           =   1590
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todos"
      Height          =   330
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7965
      Width           =   1365
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todos"
      Height          =   330
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7965
      Width           =   1590
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8640
      Width           =   1830
   End
   Begin MSComctlLib.ListView lista1 
      Height          =   7260
      Left            =   45
      TabIndex        =   1
      Top             =   630
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   12806
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9090
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8640
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista2 
      Height          =   7260
      Left            =   5130
      TabIndex        =   2
      Top             =   630
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   12806
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Marque los Tipos y Subtipos para los que desea generar el listado."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   45
      TabIndex        =   3
      Top             =   90
      Width           =   10095
   End
End
Attribute VB_Name = "frmCa_Filtro_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TIPO_LLAMADA As Integer
' 1 para DOCUMENTACION
' 2 para NORMAS

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista1.ListItems.Count
        lista1.ListItems(i).Checked = False
    Next
End Sub

'Private Sub Normas_completo()
'    On Error GoTo fallo
'    Dim i As Integer
'    Dim onormas As New clsCa_normas
'    Dim rs As ADOdb.RecordSet
'    Dim filtro1 As String
'    For i = 1 To lista1.ListItems.Count
'        If lista1.ListItems(i).Checked = True Then
'            filtro1 = filtro1 & lista1.ListItems(i).SubItems(1) & ","
'        End If
'    Next
'    If filtro1 <> "" Then
'        filtro1 = Left(filtro1, Len(filtro1) - 1)
'    End If
'    Dim filtro2 As String
'    For i = 1 To lista2.ListItems.Count
'        If lista2.ListItems(i).Checked = True Then
'            filtro2 = filtro2 & lista2.ListItems(i).SubItems(1) & ","
'        End If
'    Next
'    If filtro2 <> "" Then
'        filtro2 = Left(filtro2, Len(filtro2) - 1)
'    End If
'    Set rs = onormas.Listado_Auditoria(filtro1, filtro2, chkNADCAP.value, chkEQA.value, chkENAC.value)
'    If rs.RecordCount = 0 Then
'        MsgBox "No existen documentos con esos criterios.", vbExclamation, App.Title
'        Exit Sub
'    End If
'    ' Generar Listado
'    Dim Listado As New rptListadoNormasAuditoria
'    Listado.Orientation = rptOrientLandscape
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Normas Controladas (LI-01)"
'    End With
'    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
' '       .Controls("linea1").Visible = True
'        .Controls("d1").DataField = rs.Fields(0).Name
'        .Controls("d2").DataField = rs.Fields(1).Name
'        .Controls("d3").DataField = rs.Fields(2).Name
'        .Controls("d4").DataField = rs.Fields(3).Name
'        .Controls("d5").DataField = rs.Fields(4).Name
'        .Controls("d6").DataField = rs.Fields(5).Name
'        .Controls("d7").DataField = rs.Fields(6).Name
'        .Controls("d8").DataField = rs.Fields(7).Name
'    End With
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie3").Caption = "Firmado, Margarita Halcón"
'        .Controls("pie3").Visible = True
'        .Controls("firma").Visible = True
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Documentos de calidad."
''    Listado.WindowState = vbMaximized
'    Listado.Show 1
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado.", vbCritical, Err.Description
'
'End Sub

Private Sub cmdImprimir_Click()
    Select Case TIPO_LLAMADA
    Case 1
        Documentos
    Case 2
        Normas
    End Select
End Sub

Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista1.ListItems.Count
        lista1.ListItems(i).Checked = True
    Next
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    Dim i As Integer
    For i = 1 To lista2.ListItems.Count
        lista2.ListItems(i).Checked = False
    Next

End Sub

Private Sub Command2_Click()
    Dim i As Integer
    For i = 1 To lista2.ListItems.Count
        lista2.ListItems(i).Checked = True
    Next

End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    Select Case TIPO_LLAMADA
    Case 1
        cargar_lista_documentos
    
    Case 2
        cargar_lista_normas
    End Select
End Sub
Private Sub cargar_lista_normas()
    Dim rs As ADODB.Recordset
    Dim oDecodificadora As New clsDecodificadora
    Set rs = oDecodificadora.Listado(DECODIFICADORA.CA_NORMAS_TIPOS)
    If rs.RecordCount <> 0 Then
        Do
           With lista1.ListItems.Add(, , rs("descripcion"))
            .SubItems(1) = rs("valor")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = oDecodificadora.Listado(DECODIFICADORA.CA_NORMAS_SUBTIPOS)
    If rs.RecordCount <> 0 Then
        Do
           With lista2.ListItems.Add(, , rs("descripcion"))
            .SubItems(1) = rs("valor")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oDecodificadora = Nothing
    Set rs = Nothing
End Sub

Private Sub cabecera()
    With lista1.ColumnHeaders
        .Add , , "Tipo", 4300, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnLeft
    End With
    With lista2.ColumnHeaders
        .Add , , "SubTipo", 4300, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnLeft
    End With
End Sub

Private Sub cargar_lista_documentos()
    Dim rs As ADODB.Recordset
    Dim oDecodificadora As New clsDecodificadora
    Set rs = oDecodificadora.Listado(DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS)
    If rs.RecordCount <> 0 Then
        Do
           With lista1.ListItems.Add(, , rs("descripcion"))
            .SubItems(1) = rs("valor")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = oDecodificadora.Listado(DECODIFICADORA.CA_DOCUMENTOS_SUBFAMILIAS)
    If rs.RecordCount <> 0 Then
        Do
           With lista2.ListItems.Add(, , rs("descripcion"))
            .SubItems(1) = rs("valor")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oDecodificadora = Nothing
    Set rs = Nothing
End Sub


Private Sub Documentos()
    On Error GoTo fallo
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim filtro1 As String
    For i = 1 To lista1.ListItems.Count
        If lista1.ListItems(i).Checked = True Then
            filtro1 = filtro1 & lista1.ListItems(i).SubItems(1) & ","
        End If
    Next
    If filtro1 <> "" Then
        filtro1 = Left(filtro1, Len(filtro1) - 1)
    End If
    Dim filtro2 As String
    For i = 1 To lista2.ListItems.Count
        If lista2.ListItems(i).Checked = True Then
            filtro2 = filtro2 & lista2.ListItems(i).SubItems(1) & ","
        End If
    Next
    If filtro2 <> "" Then
        filtro2 = Left(filtro2, Len(filtro2) - 1)
    End If
    Dim oDocumentos As New clsCa_documentos
    Set rs = oDocumentos.Listado_Auditoria(filtro1, filtro2, chkNADCAP.value, chkMTL.value, chkEQA.value, chkENAC.value)
    If rs.RecordCount = 0 Then
        MsgBox "No existen documentos con esos criterios.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim Listado As New rptListadoModal
'    Listado.Orientation = rptOrientLandscape
    ' Cabecera
    With Listado.Sections("cabecera")
        .Controls("titulo").Caption = "Lista de Documentos en Vigor (LI-01)"
        .Controls("etiqueta4").Left = 170
        .Controls("etiqueta4").Width = 5800
        .Controls("etiqueta4").Caption = "Documento"
        .Controls("etiqueta5").Left = 6000
        .Controls("etiqueta5").Width = 1500
        .Controls("etiqueta5").Caption = "Código"
        .Controls("etiqueta10").Left = 7800
        .Controls("etiqueta10").Width = 1500
        .Controls("etiqueta10").Caption = "Edición"
        .Controls("etiqueta11").Left = 9400
        .Controls("etiqueta11").Width = 1500
        .Controls("etiqueta11").Caption = "Fecha"
    End With
    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
    'Detalle
    With Listado.Sections("detalle")
 '       .Controls("linea1").Visible = True
        .Controls("d1").Left = 170
        .Controls("d1").Width = 5800
        .Controls("d1").CanGrow = True
        .Controls("d1").Alignment = 0
        .Controls("d1").DataField = rs.Fields(0).Name
        .Controls("d2").Left = 6000
        .Controls("d2").Width = 1500
        .Controls("d2").Alignment = 2
        .Controls("d2").DataField = rs.Fields(1).Name
        .Controls("d3").Left = 7800
        .Controls("d3").Width = 1500
        .Controls("d3").Alignment = 2
        .Controls("d3").DataField = rs.Fields(2).Name
        .Controls("d4").Left = 9400
        .Controls("d4").Width = 1500
        .Controls("d4").Alignment = 2
        .Controls("d4").DataField = rs.Fields(3).Name
    End With
    ' Pie de Pagina
    With Listado.Sections("pie")
        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & USUARIO.getNOMBRE
        .Controls("pie3").Caption = "Firmado, Margarita Halcón"
        .Controls("pie3").Visible = True
        .Controls("firma").Visible = True
    End With
    Set Listado.DataSource = rs
    Listado.Caption = "Listado de Documentos de calidad."
    Listado.Show 1
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox "Error al generar el listado.", vbCritical, Err.Description

End Sub
Private Sub Normas()
    On Error GoTo fallo
    Dim i As Integer
    Dim onormas As New clsCa_normas
    Dim rs As ADODB.Recordset
    Dim filtro1 As String
    For i = 1 To lista1.ListItems.Count
        If lista1.ListItems(i).Checked = True Then
            filtro1 = filtro1 & lista1.ListItems(i).SubItems(1) & ","
        End If
    Next
    If filtro1 <> "" Then
        filtro1 = Left(filtro1, Len(filtro1) - 1)
    End If
    Dim filtro2 As String
    For i = 1 To lista2.ListItems.Count
        If lista2.ListItems(i).Checked = True Then
            filtro2 = filtro2 & lista2.ListItems(i).SubItems(1) & ","
        End If
    Next
    If filtro2 <> "" Then
        filtro2 = Left(filtro2, Len(filtro2) - 1)
    End If
    Set rs = onormas.Listado_Auditoria(filtro1, filtro2, chkNADCAP.value, chkMTL.value, chkEQA.value, chkENAC.value)
    If rs.RecordCount = 0 Then
        MsgBox "No existen documentos con esos criterios.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim Listado As New rptListadoModal
'    Listado.Orientation = rptOrientLandscape
    ' Cabecera
    With Listado.Sections("cabecera")
        .Controls("titulo").Caption = "Listado de Normas Controladas (LI-02)"
        .Controls("etiqueta4").Left = 170
        .Controls("etiqueta4").Width = 5800
        .Controls("etiqueta4").Caption = "Norma"
        .Controls("etiqueta5").Left = 6000
        .Controls("etiqueta5").Width = 1500
        .Controls("etiqueta5").Caption = "Código"
        .Controls("etiqueta10").Left = 7800
        .Controls("etiqueta10").Width = 1500
        .Controls("etiqueta10").Caption = "Edición"
        .Controls("etiqueta11").Left = 9400
        .Controls("etiqueta11").Width = 1500
        .Controls("etiqueta11").Caption = "Fecha"
    End With
    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
    'Detalle
    With Listado.Sections("detalle")
 '       .Controls("linea1").Visible = True
        .Controls("d1").Left = 170
        .Controls("d1").Width = 5800
        .Controls("d1").CanGrow = True
        .Controls("d1").Alignment = 0
        .Controls("d1").DataField = rs.Fields(0).Name
        .Controls("d2").Left = 6000
        .Controls("d2").Width = 1500
        .Controls("d2").Alignment = 2
        .Controls("d2").DataField = rs.Fields(1).Name
        .Controls("d3").Left = 7800
        .Controls("d3").Width = 1500
        .Controls("d3").Alignment = 2
        .Controls("d3").DataField = rs.Fields(2).Name
        .Controls("d4").Left = 9400
        .Controls("d4").Width = 1500
        .Controls("d4").Alignment = 2
        .Controls("d4").DataField = rs.Fields(3).Name
    End With
    ' Pie de Pagina
    With Listado.Sections("pie")
        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & USUARIO.getNOMBRE
        .Controls("pie3").Caption = "Firmado, Margarita Halcón"
        .Controls("pie3").Visible = True
        .Controls("firma").Visible = True
    End With
    Set Listado.DataSource = rs
    Listado.Caption = "Listado de Documentos de calidad."
    Listado.Show 1
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox "Error al generar el listado.", vbCritical, Err.Description

End Sub

