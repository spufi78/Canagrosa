VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEquipos_PlanesMantenimiento_listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Equipos - Planes de Mantenimiento - Listado"
   ClientHeight    =   8010
   ClientLeft      =   1185
   ClientTop       =   1860
   ClientWidth     =   12855
   ClipControls    =   0   'False
   Icon            =   "frmEquipos_PlanesMantenimiento_listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   12855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdVincular 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vincular"
      Enabled         =   0   'False
      Height          =   870
      Left            =   4455
      Picture         =   "frmEquipos_PlanesMantenimiento_listado.frx":058A
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Vincular plan de mantenimiento al equipo"
      Top             =   7110
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   1095
      Left            =   0
      TabIndex        =   7
      Top             =   555
      Width           =   12840
      Begin VB.TextBox txtNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1110
         MaxLength       =   255
         TabIndex        =   11
         Top             =   315
         Width           =   3930
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1110
         MaxLength       =   255
         TabIndex        =   8
         Top             =   630
         Width           =   7575
      End
      Begin MSDataListLib.DataCombo cmbFrecuencia 
         Height          =   315
         Left            =   6255
         TabIndex        =   12
         Top             =   300
         Width           =   2460
         _ExtentX        =   4339
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Frecuencia"
         Height          =   330
         Index           =   42
         Left            =   5310
         TabIndex        =   13
         Top             =   315
         Width           =   795
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   675
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11790
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7110
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Añadir plan de mantenimiento"
      Top             =   7110
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Modificar plan de mantenimiento"
      Top             =   7110
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2265
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Eliminar plan de mantenimiento"
      Top             =   7110
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7110
      Visible         =   0   'False
      Width           =   1020
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5400
      Left            =   0
      TabIndex        =   5
      Top             =   1665
      Width           =   12840
      _ExtentX        =   22648
      _ExtentY        =   9525
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
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12240
      Picture         =   "frmEquipos_PlanesMantenimiento_listado.frx":0BC8
      Top             =   45
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Planes de Mantenimiento"
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
      Height          =   570
      Index           =   3
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   12840
   End
End
Attribute VB_Name = "frmEquipos_PlanesMantenimiento_listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public FK_EQUIPO As Long ' Clave foránea del equipo al que se le va a asignar el plan de mto


Private Sub cmbFrecuencia_Click(Area As Integer)
    cmdBuscar_Click
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el Plan de Mantenimiento: " & lista.ListItems(lista.SelectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oEQ_Plan As New clsPlanMantenimiento
            If oEQ_Plan.Eliminar(lista.ListItems(lista.SelectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    Else
        MsgBox "Debe seleccionar un plan de mantenimiento.", vbInformation, App.Title
    End If
End Sub

Private Sub cmdImprimir_Click()


' JONATHAN.17-08-2010
' --------> ESTE BOTON NO ESTÁ VISIBLE Y NO SE ACTIVA EN NINGUNA PARTE DEL FORMULARIO, POR LO QUE SE COMENTA EL CÓDIGO, PARA QUE NO HAGA REFERENCIA A rptListado
' JONATHAN.17-08-2010

'    On Error GoTo fallo
'    If lista.ListItems.Count = 0 Then
'        Exit Sub
'    End If
'    Dim i As Integer
'    ' Generamos los datos del listado
'    Dim rs As New ADODB.RecordSet
'    rs.Fields.Append "c1", adChar, 250, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 50, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 20, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 10, adFldUpdatable
'    rs.Open
''    Dim existe As Boolean
''    existe = False
'    For i = 1 To lista.ListItems.Count
''        If Trim(lista.ListItems(i).SubItems(6)) = "En vigor" Then
'            rs.AddNew
''            existe = True
'            rs("c1") = Trim(Left(lista.ListItems(i).SubItems(1), 250))
'            rs("c2") = Left(lista.ListItems(i).SubItems(4), 50)
'            rs("c3") = Left(lista.ListItems(i).SubItems(5), 20)
'            rs("c4") = Left(lista.ListItems(i).SubItems(6), 10)
'            rs.Update
''        End If
'    Next
''    If Not existe Then
''        MsgBox "No existen en la lista documentos en vigor.", vbInformation, App.Title
''        Exit Sub
''    End If
'
'    ' Generar Listado
'    Dim Listado As New rptListado
''    Listado.Orientation = rptOrientLandscape
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Normas Controladas (LI-02)"
'        .Controls("etiqueta4").Left = 170
'        .Controls("etiqueta4").Width = 5800
'        .Controls("etiqueta4").Caption = "Norma"
'        .Controls("etiqueta5").Left = 6000
'        .Controls("etiqueta5").Width = 1500
'        .Controls("etiqueta5").Caption = "Código"
'        .Controls("etiqueta10").Left = 7800
'        .Controls("etiqueta10").Width = 1500
'        .Controls("etiqueta10").Caption = "Edición"
'        .Controls("etiqueta11").Left = 9400
'        .Controls("etiqueta11").Width = 1500
'        .Controls("etiqueta11").Caption = "Fecha"
'    End With
'    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
' '       .Controls("linea1").Visible = True
'        .Controls("d1").Left = 170
'        .Controls("d1").Width = 5800
'        .Controls("d1").CanGrow = True
'        .Controls("d1").Alignment = 0
'        .Controls("d1").DataField = rs.Fields("c1").Name
'        .Controls("d2").Left = 6000
'        .Controls("d2").Width = 1500
'        .Controls("d2").Alignment = 2
'        .Controls("d2").DataField = rs.Fields("c2").Name
'        .Controls("d3").Left = 7800
'        .Controls("d3").Width = 1500
'        .Controls("d3").Alignment = 2
'        .Controls("d3").DataField = rs.Fields("c3").Name
'        .Controls("d4").Left = 9400
'        .Controls("d4").Width = 1500
'        .Controls("d4").Alignment = 2
'        .Controls("d4").DataField = rs.Fields("c4").Name
'    End With
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
''        .Controls("pie2").Caption = "Impreso por : " & USUARIO.getNOMBRE
'        .Controls("pie3").Caption = "Firmado, Margarita Halcón"
'        .Controls("pie3").Visible = True
'        .Controls("firma").Visible = True
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Documentos de calidad."
''    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado.", vbCritical, Err.Description
End Sub

Private Sub cmdAnadir_Click()
    Dim objfrm As New frmEquipoPlanMtoEdicion
    
    objfrm.TipoEdicion = ALTA
    objfrm.Show vbModal
    
    If objfrm.Resultado Then
        cargar_lista
    End If
    
    Unload objfrm
    Set objfrm = Nothing
    
    
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    'frmEquipos_Detalle_Mto.FK_PLAN_MTO = 0
    Unload Me
End Sub

Private Sub cmdEliminar_Click_old()
'    If lista.ListItems.Count > 0 Then
'        If MsgBox("Va a eliminar el Plan de Mantenimiento: " & lista.ListItems(lista.SelectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
'            Dim oEQ_Plan As New clsEquipos_planes_Mantenimiento
'            If oEQ_Plan.Eliminar(lista.ListItems(lista.SelectedItem.Index).Text) = True Then
'                cargar_lista
'            End If
'        End If
'    Else
'        MsgBox "Debe seleccionar un plan de mantenimiento.", vbInformation, App.Title
'    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmEquipoPlanMtoEdicion.PK = lista.ListItems(lista.SelectedItem.Index).Text
        frmEquipoPlanMtoEdicion.Show 1
        'frmEquipos_PlanesMantenimiento_detalle.PK = lista.ListItems(lista.SelectedItem.Index).Text
        'frmEquipos_PlanesMantenimiento_detalle.Show 1
        If frmEquipoPlanMtoEdicion.Resultado Then
            cargar_lista
        End If
        Unload frmEquipoPlanMtoEdicion
    Else
        MsgBox "Debe seleccionar un plan de mantenimiento.", vbInformation, App.Title
    End If
End Sub

Private Sub cmdVincular_Click()
'    If lista.ListItems.Count > 0 Then
'        frmEquipos_Detalle_Mto.FK_PLAN_MTO = lista.ListItems(lista.SelectedItem.Index).Text ' Este es el ID_PLAN_MTO
'        Unload Me
'    Else
'        MsgBox "Debe seleccionar un plan de mantenimiento.", vbInformation, App.Title
'    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Top = 100
    Me.Left = 100
    cabecera
    cargar_botones Me
    cargar_combos
    cargar_lista
    permisos
    
    'E0053-I
    'El botón vincular sólo aparecerá si se ha abierto el formulario desde el mto de planes de un equipo
    If Me.FK_EQUIPO = 0 Then
        cmdVincular.Visible = False
    Else
        cmdVincular.Visible = True
    End If
    'E0053-F
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Se restablece FK_EQUIPO a 0
    Me.FK_EQUIPO = 0
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Plan", 4000, lvwColumnLeft
        .Add , , "Descripción", 7000, lvwColumnLeft
        .Add , , "Frecuencia", 1500, lvwColumnLeft
        .Add , , "Orden", 0, lvwColumnLeft
    End With
End Sub
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    
    oDeco.cargar_combo cmbFrecuencia, decodificadora.EQ_periodicidad

End Sub
Public Sub cargar_lista()
    Dim rs As ADODB.RecordSet
    Dim oEQ_Plan_Mto As New clsPlanMantenimiento
    
   On Error GoTo cargar_lista_Error

    Set rs = oEQ_Plan_Mto.Listado(txtNombre.Text, txtDescripcion.Text, getDataComboSel(cmbFrecuencia))

    lista.ListItems.Clear
    
    If rs.RecordCount <> 0 Then
        While Not rs.EOF
            With lista.ListItems.Add(, , rs("ID_PLAN_MTO"))
             .SubItems(1) = rs("NOMBRE")
             .SubItems(2) = rs("DESCRIPCION")
             .SubItems(3) = rs("FRECUENCIA")
            End With
            rs.MoveNext
        Wend
    End If
    Set oEQ_Plan_Mto = Nothing

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmEquipos_PlanesMantenimiento_listado"
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

Public Sub permisos()
'    If Not USUARIO.getPER_DOCUMENTACION_CALIDAD Then
'        cmdAnadir.Enabled = False
'        cmdModificar.Enabled = False
'        cmdEliminar.Enabled = False
'    End If
End Sub
