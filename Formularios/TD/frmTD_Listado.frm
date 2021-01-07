VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmTD_Listado 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Tipos de Determinaciones"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12465
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTD_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9375
   ScaleWidth      =   12465
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   915
      Left            =   9135
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8415
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por"
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
      Height          =   1365
      Left            =   45
      TabIndex        =   16
      Top             =   675
      Width           =   12390
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   4680
         TabIndex        =   23
         Top             =   630
         Width           =   2265
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1035
         TabIndex        =   21
         Top             =   630
         Width           =   2265
      End
      Begin VB.CheckBox chkAnuladas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Solo las Anuladas"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   3240
         TabIndex        =   20
         Top             =   1035
         Width           =   2475
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   780
         Left            =   11250
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         Width           =   1050
      End
      Begin VB.CheckBox chkSubcontratables 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcontratables"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   1665
         TabIndex        =   3
         Top             =   1035
         Width           =   1530
      End
      Begin VB.CheckBox chkusadas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "No utilizadas"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   90
         TabIndex        =   2
         Top             =   1035
         Width           =   2475
      End
      Begin VB.TextBox txtfiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1035
         TabIndex        =   0
         Top             =   270
         Width           =   2265
      End
      Begin pryCombo.miCombo cmbFamilia 
         Height          =   375
         Left            =   4680
         TabIndex        =   1
         Top             =   270
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Norma"
         Height          =   195
         Index           =   1
         Left            =   3780
         TabIndex        =   24
         Top             =   675
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "PNT"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   675
         Width           =   330
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   3
         Left            =   3780
         TabIndex        =   18
         Top             =   315
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   17
         Top             =   315
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdFamilia 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Familias"
      Height          =   915
      Left            =   6525
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8430
      Width           =   1050
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   915
      Left            =   4365
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8430
      Width           =   1050
   End
   Begin VB.CommandButton cmdQuien 
      BackColor       =   &H00E0E0E0&
      Caption         =   "¿Donde?"
      Height          =   915
      Left            =   5445
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8430
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   915
      Left            =   1125
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8430
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   915
      Left            =   3285
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8430
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   915
      Left            =   2205
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8430
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   915
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8430
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   11430
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8430
      Width           =   1005
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6300
      Left            =   60
      TabIndex        =   13
      Top             =   2070
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   11113
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de los tipos de determinaciones"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   15
      Top             =   360
      Width           =   2760
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11880
      Picture         =   "frmTD_Listado.frx":000C
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Tipos de determinación"
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
      TabIndex        =   14
      Top             =   45
      Width           =   3630
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   12450
   End
End
Attribute VB_Name = "frmTD_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkAnuladas_Click()
    cargar_lista
End Sub

Private Sub cmbfamilia_Change()
    cargar_lista
End Sub

Private Sub cmdLimpiar_Click()
    txtfiltro(0) = ""
    txtfiltro(1) = ""
    txtfiltro(2) = ""
    cmbFamilia.limpiar
    chkSubcontratables.Value = Unchecked
    chkusadas.Value = Unchecked
    chkDia.Value = Unchecked
    cargar_lista
End Sub
Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a duplicar el tipo de determinación. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
      Dim DETERMINACION As Long
      Dim oDeter As New clsTipos_determinacion
      If oDeter.duplicar(lista.ListItems(lista.selectedItem.Index).SubItems(4)) <> 0 Then
          MsgBox "El tipo de determinación se ha duplicado correctamente.", vbOKOnly + vbInformation, App.Title
          cargar_lista
      End If
      Set oDeter = Nothing
    End If
    Exit Sub
fallo:
    MsgBox "Error en el proceso de duplicación.", vbCritical, App.Title

End Sub

Private Sub cmdFamilia_Click()
    frmTD_Familias.Show 1
    cargar_lista
End Sub

Private Sub cmdQuien_Click()
    If lista.ListItems.Count > 0 Then
        frmTD_Donde.PK = lista.ListItems(lista.selectedItem.Index).SubItems(4)
        frmTD_Donde.Show 1
    End If
End Sub
Private Sub chkDia_Click()
    cargar_lista
End Sub

Private Sub chkusadas_Click()
    cargar_lista
End Sub

'E0126-I
Private Sub chkSubcontratables_Click()
    cargar_lista
End Sub
'E0126-F

Private Sub cmdAnadir_Click()
    frmTD_Detalle.PK = 0
    frmTD_Detalle.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Va a eliminar el tipo de determinación : " & lista.ListItems(lista.selectedItem.Index), vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oTD As New clsTipos_determinacion
        If oTD.Eliminar(lista.ListItems(lista.selectedItem.Index).SubItems(4)) = True Then
            cargar_lista
        End If
        Set oTD = Nothing
    End If
End Sub

'Private Sub cmdImprimir_Click()
'    On Error GoTo fallo
'    Dim i As Integer
'    ' Generamos los datos del listado
'    Dim rs As New ADODB.RecordSet
'    rs.Fields.Append "c1", adChar, 50, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 50, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 15, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 5, adFldUpdatable
'    rs.Open
'    For i = 1 To lista.ListItems.Count
'        rs.AddNew
'        rs("c1") = Left(lista.ListItems(i), 50)
'        If Trim(lista.ListItems(i).SubItems(2)) <> "" Then
'            rs("c2") = Left(lista.ListItems(i).SubItems(2), 50)
'        End If
'        If Trim(lista.ListItems(i).SubItems(3)) <> "" Then
'            rs("c3") = Left(lista.ListItems(i).SubItems(3), 15)
'        End If
'        If Trim(lista.ListItems(i).SubItems(4)) <> "" Then
'            rs("c4") = Left(lista.ListItems(i).SubItems(4), 5)
'        End If
'        rs.Update
'    Next
'    ' Generar Listado
'    Dim Listado As New rptListado
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("titulo").Caption = "Listado de Determinaciones"
'        .Controls("etiqueta4").Caption = "ID"
'        .Controls("etiqueta5").Caption = "Determinacion"
'        .Controls("etiqueta10").Caption = "Formula"
'        .Controls("etiqueta11").Caption = "PNT"
'    End With
'    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("d1").DataField = rs.Fields("c4").Name
'        .Controls("d2").DataField = rs.Fields("c1").Name
'        .Controls("d3").DataField = rs.Fields("c2").Name
'        .Controls("d4").DataField = rs.Fields("c3").Name
'    End With
'    ' Pie de Pagina
'    With Listado.Sections("pie")
'        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Impreso por : " & Usuario.getNOMBRE
'    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Determinaciones"
''    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado.", vbCritical, Err.Description
'End Sub

Private Sub cmdImprimir_Click()

    Dim objTD As New clsTipos_determinacion

    objTD.Imprimir_Listado Trim(txtfiltro(0).Text), cmbFamilia.getPK_SALIDA, (chkDia.Value = vbChecked), (chkusadas.Value = vbChecked), (chkSubcontratables.Value = vbChecked)
    
    Set objTD = Nothing


End Sub
Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmTD_Detalle.PK = lista.ListItems(lista.selectedItem.Index).SubItems(4)
        frmTD_Detalle.Show 1
        modificar_determinacion
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
    If MsgBox("Va a eliminar todas las determinaciones listadas.", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim oTD As New clsTipos_determinacion
        Dim i As Integer
        Me.MousePointer = 11
        For i = 1 To lista.ListItems.Count
            oTD.Eliminar lista.ListItems(i).SubItems(4)
        Next
        Set oTD = Nothing
        Me.MousePointer = 0
        MsgBox "OK"
        cargar_lista
    End If

End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_lista
    llenar_combo cmbFamilia, New clsTipos_determinacion_familias, 0, Me, ""
    If UCase(USUARIO.getUSUARIO) = "JULIO" Then
        Command1.visible = True
    Else
        Command1.visible = False
    End If
End Sub
Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oDET As New clsTipos_determinacion
    
    If chkusadas.Value = Unchecked Then
        Set rs = oDET.lista(txtfiltro(0), txtfiltro(1), txtfiltro(2), IIf(chkSubcontratables.Value = Checked, 1, ""), cmbFamilia.getPK_SALIDA, chkAnuladas.Value)
    Else
        Set rs = oDET.lista_no_utilizadas(txtfiltro(0), txtfiltro(1), txtfiltro(2), IIf(chkSubcontratables.Value = Checked, 1, ""), cmbFamilia.getPK_SALIDA, chkAnuladas.Value)
    End If
    lista.ListItems.Clear
    Dim total As Integer
    Dim nombre As String
    total = rs.RecordCount
    If rs.RecordCount <> 0 Then
        Do
           nombre = rs("NOMBRE")
           If rs("DESCRIPCION") <> "" Then
               nombre = nombre & " -> " & rs("DESCRIPCION")
           End If
           With lista.ListItems.Add(, , nombre)
           .SubItems(1) = rs("FAMILIA")
           If Not IsNull(rs("FORMULA")) Then
               .SubItems(2) = rs("FORMULA")
           End If
           .SubItems(3) = rs("PNT")
           .SubItems(4) = Format(rs("ID_TIPO_DETERMINACION"), "0000")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    lblsubtitulo = "Número de tipos de determinación mostrados : " & total
    Set oDET = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub


Private Sub lista_Click()
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

Public Sub modificar_determinacion()
    Dim oDET As New clsTipos_determinacion
    Dim ofor As New clsFormulas
    Dim ofam As New clsTipos_determinacion_familias
    If oDET.CargarTipoDeterminacion(lista.ListItems(lista.selectedItem.Index).SubItems(4)) = True Then
        lista.ListItems(lista.selectedItem.Index).Text = oDET.getNOMBRE & " -> " & oDET.getDESCRIPCION
        ofam.Carga oDET.getFAMILIA_ID
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = ofam.getNOMBRE
        If oDET.getFORMULA_ID = 0 Then
           lista.ListItems(lista.selectedItem.Index).SubItems(2) = ""
        Else
           ofor.CARGAR (oDET.getFORMULA_ID)
           lista.ListItems(lista.selectedItem.Index).SubItems(2) = ofor.getNOMBRE
        End If
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = oDET.getPNT
    End If
    Set oDET = Nothing
    Set ofor = Nothing
    Set ofam = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Nombre", 4000, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Familia", 3000, lvwColumnLeft)
        .Tag = "Familia"
    End With
    With lista.ColumnHeaders.Add(, , "Formula", 3000, lvwColumnLeft)
        .Tag = "Formula"
    End With
    With lista.ColumnHeaders.Add(, , "PNT", 1400, lvwColumnCenter)
        .Tag = "PNT"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 550, lvwColumnCenter)
        .Tag = "ID"
    End With
End Sub

Private Sub txtfiltro_Change(Index As Integer)
    cargar_lista
End Sub
