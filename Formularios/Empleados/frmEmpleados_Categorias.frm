VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmEmpleados_Categorias 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categorias"
   ClientHeight    =   8985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13380
   Icon            =   "frmEmpleados_Categorias.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   13380
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   885
      Left            =   9360
      Picture         =   "frmEmpleados_Categorias.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8010
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la Categoría"
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
      Height          =   2400
      Left            =   45
      TabIndex        =   7
      Top             =   6525
      Width           =   9240
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver Ficha"
         Height          =   840
         Left            =   2250
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1485
         Width           =   1095
      End
      Begin VB.TextBox txtcategoria 
         Height          =   420
         Left            =   5085
         TabIndex        =   17
         Top             =   135
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton cmdmodificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   840
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1485
         Width           =   1095
      End
      Begin VB.CommandButton cmdeliminar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   840
         Left            =   5625
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1485
         Width           =   1095
      End
      Begin VB.CommandButton cmdanadir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   840
         Left            =   3375
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1485
         Width           =   1095
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   990
         MaxLength       =   75
         TabIndex        =   11
         Top             =   315
         Width           =   2535
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   990
         MaxLength       =   75
         TabIndex        =   8
         Top             =   720
         Width           =   7080
      End
      Begin pryCombo.miCombo cmbFicha 
         Height          =   330
         Left            =   990
         TabIndex        =   19
         Top             =   1125
         Width           =   8130
         _ExtentX        =   14340
         _ExtentY        =   582
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ficha"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   20
         Top             =   1170
         Width           =   690
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Categoría"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   9
         Top             =   765
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Listado de Empleados"
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
      Height          =   5655
      Left            =   6930
      TabIndex        =   5
      Top             =   855
      Width           =   6405
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   330
         Index           =   2
         Left            =   1215
         MaxLength       =   75
         TabIndex        =   16
         Top             =   225
         Width           =   5145
      End
      Begin MSComctlLib.ListView lista 
         Height          =   4995
         Left            =   90
         TabIndex        =   6
         Top             =   585
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   8811
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
      Begin VB.Label lblcodigo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Departamento"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   270
         Width           =   1365
      End
   End
   Begin VB.CommandButton cmdDepartamentos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Departamentos"
      Height          =   885
      Left            =   10755
      Picture         =   "frmEmpleados_Categorias.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8010
      Width           =   1335
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   12150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8010
      Width           =   1155
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   5595
      Left            =   45
      TabIndex        =   2
      Top             =   855
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   9869
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Categorias.frx":1A5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Categorias.frx":2338
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Categorias.frx":2C12
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Relación de departamentos y distinas categorías dentro de la empresa"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   4980
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12780
      Picture         =   "frmEmpleados_Categorias.frx":34EC
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Departamentos y Categorías"
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
      TabIndex        =   0
      Top             =   120
      Width           =   2985
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   13365
   End
End
Attribute VB_Name = "frmEmpleados_Categorias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'M1002-I
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Codigo", 1200, lvwColumnLeft
        .Add , , "Nombre", 4750, lvwColumnCenter
    End With
End Sub
'M1002-F
Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If validar Then
        Dim oCat As New clsEmpleados_categorias
        Dim departamento As Integer
        With oCat
            .setCODIGO = txtDatos(1)
            .setDESCRIPCION = txtDatos(0)
            .setDOCUMENTO_ID = cmbFicha.getPK_SALIDA
            departamento = recuperar_id_departamento
            If departamento > 0 Then
                .setDEPARTAMENTO_ID = recuperar_id_departamento
                .Insertar
            Else
                MsgBox "Selecccione el departamento al que corresponde.", vbCritical, App.Title
            End If
        End With
        cargar_categorias
        Set oCat = Nothing
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdanadir_Click of Formulario frmEmpleados_Categorias"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDepartamentos_Click()
    Dim oform As New frmDecodificadora
    oform.CODIGO = DECODIFICADORA.EMPLEADOS_DEPARTAMENTOS
    oform.Show
    Set oform = Nothing
End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo cmdEliminar_Click_Error

    If validar Then
        If txtcategoria <> "" Then
            If CInt(txtcategoria) > 0 Then
                Dim oCat As New clsEmpleados_categorias
                oCat.Eliminar txtcategoria
                Set oCat = Nothing
                cargar_categorias
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdeliminar_Click of Formulario frmEmpleados_Categorias"

End Sub

Private Sub cmdImprimir_Click()
    frmReport.iniciar
    frmReport.informe = "\Empleados\rptEmpleados_Categorias"
    frmReport.criterio = "{decodificadora.CODIGO} = " & DECODIFICADORA.EMPLEADOS_DEPARTAMENTOS
    frmReport.imprimir = False
    frmReport.generar
    frmReport.Show 1
    Unload frmReport
End Sub

Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_Error

    If validar Then
        If txtcategoria <> "" Then
            If CInt(txtcategoria) > 0 Then
                Dim oCat As New clsEmpleados_categorias
                Dim departamento As Integer
                With oCat
                    .setCODIGO = txtDatos(1)
                    .setDESCRIPCION = txtDatos(0)
                    .setDOCUMENTO_ID = cmbFicha.getPK_SALIDA
                    departamento = recuperar_id_departamento
                    If departamento > 0 Then
                        .setDEPARTAMENTO_ID = recuperar_id_departamento
                        .Modificar txtcategoria
                    Else
                        MsgBox "Selecccione el departamento al que corresponde.", vbCritical, App.Title
                    End If
                End With
                cargar_categorias
                Set oCat = Nothing
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdmodificar_Click of Formulario frmEmpleados_Categorias"

End Sub

Private Sub cmdMostrar_Click()
    If cmbFicha.getTEXTO <> "" Then
        Dim oDoc As New clsCa_documentos
        oDoc.mostrar cmbFicha.getPK_SALIDA, False
        Set oDoc = Nothing
    End If
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    'M1002-I
    cabecera
    'M1002-F
    cargar_categorias
    llenar_combo cmbFicha, New clsCa_documentos, 0, frmCA_Documento, " FAMILIA_ID = 29 "
End Sub

Public Sub cargar_categorias()
    Dim nodX As Node
    Tree.Nodes.Clear
     Dim departamento As Integer
     Dim categoria As Integer
    Dim rs As ADODB.Recordset
    Dim oCat As New clsEmpleados_categorias
    Set rs = oCat.Listado
    
    If rs.RecordCount > 0 Then
        Do
'            Tree.Nodes(nodX.Index).Bold = True
            If departamento <> rs(0) Then
                departamento = rs(0)
                Set nodX = Tree.Nodes.Add(, , "ID-" & departamento, rs(1), 1)
                If Not IsNull(rs(2)) Then
                    categoria = rs(2)
                    Set nodX = Tree.Nodes.Add("ID-" & departamento, tvwChild, "ID-" & departamento & "-" & categoria & "-" & Format(rs(4), "0000"), rs(3), 2)
                End If
            End If
            If categoria <> rs(2) Then
                categoria = rs(2)
                Set nodX = Tree.Nodes.Add("ID-" & departamento, tvwChild, "ID-" & departamento & "-" & categoria & "-" & Format(rs(4), "0000"), rs(3), 2)
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oCat = Nothing
    Dim i As Integer
    For i = 1 To Tree.Nodes.Count - 1
        Tree.Nodes(i).Expanded = True
    Next
    
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        If USUARIO.getPER_MODIFICACION = 0 Then
            Exit Sub
        End If
        frmEmpleados_Gestion.PK = lista.ListItems(lista.selectedItem.Index)
        frmEmpleados_Gestion.Show 1
        lista.SetFocus
    End If
End Sub

Private Sub Tree_Click()
    txtcategoria = ""
    If Tree.Nodes.Count = 0 Then Exit Sub
    Dim d() As String
    d = Split(Tree.Nodes(Tree.selectedItem.Index).Key, "-")
    If UBound(d) = 1 Then
        lblCodigo.Caption = "Departamento"
    Else
        lblCodigo.Caption = "Categoria"
    End If
    txtDatos(2) = Tree.Nodes(Tree.selectedItem.Index).Text
    Dim categoria As Integer
    Dim documento As Long
    categoria = recuperar_id_categoria
    documento = recuperar_id_documento
    If categoria > 0 Then
        Dim oCat As New clsEmpleados_categorias
        oCat.Carga CLng(categoria)
        txtDatos(1) = oCat.getCODIGO
        txtDatos(0) = oCat.getDESCRIPCION
        txtcategoria = categoria
        'M1002-I
        lista.ListItems.Clear
        Dim oEmpleados As New clsEmpleados_categorias_historia
        Dim oDesc As New clsEmpleados
        Dim rsEmpleados As New ADODB.Recordset
        Set rsEmpleados = oEmpleados.Listado_Categoria(CLng(categoria))
        If rsEmpleados.RecordCount > 0 Then
            Do
                With lista.ListItems.Add(, , Format(rsEmpleados("empleado_id"), "000"))
                     oDesc.CARGAR rsEmpleados("empleado_id")
                     .SubItems(1) = oDesc.getNOMBRE
                End With
                rsEmpleados.MoveNext
            Loop Until rsEmpleados.EOF
        End If
        Set rsEmpleados = Nothing
        'M1002-F
        cmbFicha.limpiar
        If documento <> 0 Then
            cmbFicha.MostrarElemento documento
        End If
        Set oCat = Nothing
    End If
End Sub

Private Function recuperar_id_departamento() As Integer
    If Tree.Nodes.Count = 0 Then Exit Function
    Dim d() As String
    On Error GoTo fallo
    d = Split(Tree.Nodes(Tree.selectedItem.Index).Key, "-")
    If UBound(d) > 0 Then
        recuperar_id_departamento = d(1)
    Else
        recuperar_id_departamento = 0
    End If
    Exit Function
fallo:
    recuperar_id_departamento = 0
End Function


Private Function validar() As Boolean
    validar = True
    If txtDatos(0) = "" Then
        MsgBox "Debe indicar la categoria.", vbCritical, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If txtDatos(1) = "" Then
        MsgBox "Debe indicar el código de la categoria.", vbCritical, App.Title
        txtDatos(1).SetFocus
        validar = False
        Exit Function
    End If
    If cmbFicha.getTEXTO = "" Then
        MsgBox "Debe indicar la ficha de la categoria.", vbCritical, App.Title
        cmbFicha.SetFocus
        validar = False
        Exit Function
    End If

End Function
Private Function recuperar_id_categoria() As Integer
    If Tree.Nodes.Count = 0 Then Exit Function
    Dim d() As String
    On Error GoTo fallo
    d = Split(Tree.Nodes(Tree.selectedItem.Index).Key, "-")
    If UBound(d) > 1 Then
        recuperar_id_categoria = d(2)
    Else
        recuperar_id_categoria = 0
    End If
    Exit Function
fallo:
    recuperar_id_categoria = 0
End Function
Private Function recuperar_id_documento() As Long
    If Tree.Nodes.Count = 0 Then Exit Function
    Dim d() As String
    On Error GoTo fallo
    d = Split(Tree.Nodes(Tree.selectedItem.Index).Key, "-")
    If UBound(d) > 1 Then
        recuperar_id_documento = CLng(d(3))
    Else
        recuperar_id_documento = 0
    End If
    Exit Function
fallo:
    recuperar_id_documento = 0
End Function

