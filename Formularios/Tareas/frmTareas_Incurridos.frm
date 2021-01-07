VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form frmTareas_Incurridos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Incurridos"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11835
   Icon            =   "frmTareas_Incurridos.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   11835
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7155
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
      TabIndex        =   8
      Top             =   855
      Width           =   11760
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo las facturables"
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   1035
         Width           =   2760
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   870
         Left            =   10845
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   180
         Width           =   870
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   870
         Left            =   9945
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   870
      End
      Begin MSDataListLib.DataCombo cmbtipos 
         Height          =   315
         Left            =   900
         TabIndex        =   0
         Top             =   225
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo cmbusuario 
         Height          =   315
         Left            =   900
         TabIndex        =   2
         Top             =   630
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo cmbTarea 
         Height          =   315
         Left            =   5220
         TabIndex        =   1
         Top             =   225
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSComCtl2.DTPicker fechadesde 
         Height          =   330
         Left            =   5220
         TabIndex        =   3
         Top             =   630
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   60096513
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechahasta 
         Height          =   330
         Left            =   7470
         TabIndex        =   4
         Top             =   630
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   60096513
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   1
         Left            =   6885
         TabIndex        =   15
         Top             =   720
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   3
         Left            =   4680
         TabIndex        =   14
         Top             =   720
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarea"
         Height          =   195
         Index           =   13
         Left            =   4680
         TabIndex        =   13
         Top             =   270
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   12
         Top             =   675
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Módulo"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   285
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10755
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   7155
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4815
      Left            =   45
      TabIndex        =   6
      Top             =   2250
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   8493
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
      Caption         =   "Informe de Incurridos"
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
      Index           =   0
      Left            =   90
      TabIndex        =   11
      Top             =   90
      Width           =   2190
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11250
      Picture         =   "frmTareas_Incurridos.frx":08CA
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione los filtros necesarios para generar los datos de incurridos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   405
      Width           =   4815
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   11880
   End
End
Attribute VB_Name = "frmTareas_Incurridos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Private Sub cmborigen_Change()
'    cargar_lista
'End Sub

Private Sub Check1_Click()
    cmdBuscar_Click
End Sub

Private Sub cmbTarea_Change()
    cmdBuscar_Click
End Sub

Private Sub cmbtipos_Change()
   On Error GoTo cmbtipos_Change_Error

    If cmbtipos.Text <> "" Then
     If IsNumeric(cmbtipos.BoundText) Then
        cmbTarea.Text = ""
        Dim oTarea As New clsTareas
        Set cmbTarea.RowSource = oTarea.Listado_Combo_Filtro(CLng(cmbtipos.BoundText), fechadesde)
        cmbTarea.ListField = "descripcion"
        cmbTarea.DataField = "descripcion" 'campo asociado
        cmbTarea.BoundColumn = "id_tarea" 'lo que realmente envia
        Set oTarea = Nothing
     End If
    End If
    cmdBuscar_Click

   On Error GoTo 0
   Exit Sub

cmbtipos_Change_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmbtipos_Change of Formulario frmTareas_Incurridos"
End Sub
Private Sub cmbUsuario_Change()
    cmdBuscar_Click
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    If lista.ListItems.Count > 0 Then
        Dim consulta As String
        consulta = "{decodificadora.CODIGO}=" & DECODIFICADORA.TAREAS_MODULOS
        consulta = consulta & " AND {decodificadora_tipo_hora.CODIGO} = " & DECODIFICADORA.TAREAS_TIPOS_HORAS
        If Check1.value = Checked Then
            consulta = consulta & " AND {tareas_incurridos.FACTURABLE} = 1"
        End If
        Dim tipo As String
        Dim USUARIO As String
        Dim tarea As String
        If cmbtipos.Text <> "" Then
            consulta = consulta & " AND {tareas.MODULO_ID} = " & cmbtipos.BoundText
        End If
        If cmbUsuario.BoundText <> "0" Then
            consulta = consulta & " AND {usuarios.ID_EMPLEADO}=" & cmbUsuario.BoundText
        End If
        consulta = consulta & " AND {tareas_incurridos.FECHA} in Date " & _
                   "(" & Format(fechadesde, "yyyy") & "," & Format(fechadesde, "mm") & "," & Format(fechadesde, "dd") & ") to Date " & _
                   "(" & Format(fechahasta, "yyyy") & "," & Format(fechahasta, "mm") & "," & Format(fechahasta, "dd") & ")"
        With frmReport
            .iniciar
            .informe = "\Tareas\rptTareas_Listado"
            .criterio = consulta
            .imprimir = False
            .generar
            .Visible = True
        End With
    Else
        MsgBox "No existen datos en la lista.", vbExclamation, App.Title
    End If
End Sub

Private Sub cmdLimpiar_Click()
    cmbtipos.Text = ""
    cmbTarea.Text = ""
    cmbUsuario.Text = ""
    fechadesde = "01-" & Format(Date, "mm-yyyy")
    fechahasta = Date
End Sub

Private Sub fechadesde_Change()
    cmdBuscar_Click
End Sub

Private Sub fechahasta_Change()
    cmdBuscar_Click
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = (Screen.Width - frmMenu.ButtonBar.Width - Me.Width) / 2
    Me.Top = (Screen.Height - (frmMenu.SmartMenuXP1.Height * 2) - Me.Height - 1000) / 2
    cabecera
    cargar_botones Me
    cargar_combos
    fechadesde = "01-" & Format(Date, "mm-yyyy")
    fechahasta = Date
    cmbUsuario.BoundText = USUARIO.getID_EMPLEADO
    If USUARIO.getPER_INCURRIDOS = False Then
        cmbUsuario.Enabled = False
    End If
'    cargar_lista
'    permisos
'    If USUARIO.getUSUARIO = "julio" Then
'        cmdCargar.Visible = True
'    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Usuario", 1000, lvwColumnCenter
        .Add , , "Modulo", 1300, lvwColumnCenter
        .Add , , "Tarea", 3100, lvwColumnLeft
        .Add , , "Fecha", 1000, lvwColumnCenter
        .Add , , "Horas", 800, lvwColumnCenter
        .Add , , "Referencia", 2000, lvwColumnCenter
        .Add , , "Observación", 2000, lvwColumnLeft
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oTareas As New clsTareas_incurridos
   On Error GoTo cargar_lista_Error

    lista.ListItems.Clear
    Dim tipo As Long
    Dim USUARIO As Long
    Dim tarea As Long
    If cmbtipos.Text = "" Then
        tipo = 0
    Else
        tipo = cmbtipos.BoundText
    End If
    If cmbUsuario.BoundText = "" Then
        USUARIO = 0
    Else
        USUARIO = cmbUsuario.BoundText
    End If
    If cmbTarea.BoundText = "" Then
        tarea = 0
    Else
        tarea = cmbTarea.BoundText
    End If
    Set rs = oTareas.Listado_Filtro(USUARIO, tipo, tarea, fechadesde, fechahasta, Check1.value)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
             .SubItems(5) = rs(5)
             .SubItems(6) = rs(6)
             .SubItems(7) = rs(7)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oTareas = Nothing

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmTareas_Incurridos"
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
'    cmdModificar_Click
End Sub

Public Sub actualizar_lista()
    Dim rs As ADODB.Recordset
    Dim oTarea As New clsTareas
    Set rs = oTarea.Listado_por_Codigo(lista.ListItems(lista.selectedItem.Index).Text)
    If rs.RecordCount <> 0 Then
            With lista.ListItems(lista.selectedItem.Index)
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = Format(rs(3), "dd-mm-yyyy")
             .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
             If rs(5) = 1 Then
                .SubItems(5) = "Si"
             Else
                .SubItems(5) = "No"
             End If
            End With
            rs.MoveNext
    End If
    Set rs = Nothing
    Set oTarea = Nothing
End Sub

Public Sub cargar_combos()
    Dim oDecodificadora As New clsDecodificadora
    oDecodificadora.cargar_combo cmbtipos, DECODIFICADORA.TAREAS_MODULOS
    cargar_combo cmbUsuario, New clsUsuarios
End Sub
'Private Sub txtDato_Change(Index As Integer)
'    cmdBuscar_Click
'End Sub
