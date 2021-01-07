VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEquipoListadoMtoPte 
   Caption         =   "Listado de Equipos Sin Mantenimientos Previstos"
   ClientHeight    =   8220
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10245
   Icon            =   "frmEquipoListadoMtoPte.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8220
   ScaleWidth      =   10245
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9150
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7290
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6480
      Left            =   0
      TabIndex        =   0
      Top             =   750
      Width           =   10245
      _ExtentX        =   18071
      _ExtentY        =   11430
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
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
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Equipos sin Registros de Mantenimiento Previstos"
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
      TabIndex        =   2
      Top             =   120
      Width           =   5220
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9630
      Picture         =   "frmEquipoListadoMtoPte.frx":1272
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   $"frmEquipoListadoMtoPte.frx":157C
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   420
      Width           =   9420
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   10230
   End
End
Attribute VB_Name = "frmEquipoListadoMtoPte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Enum COLS
    COL_NOMBRE_EQUIPO = 1
    COL_PLAN_MTO = 2
    COL_PERIODICIDAD = 3
    COL_ANNO = 4
    COL_ULT_MTO = 5
    COL_ID_PLAN_MTO = 6
    COL_ID_PERIODICIDAD = 7
End Enum

Private Sub carga_lista()

    Dim oM As New clsEquipoMantenimiento
    Dim rs As ADODB.RecordSet
        
    Set rs = oM.devolver_listado_equipos_sin_mto_previstos()
    
    Set oM = Nothing
    lista.ListItems.Clear
    
    If rs.RecordCount = 0 Then Exit Sub
    
    rs.MoveFirst
    
    While Not rs.EOF
        With lista.ListItems.Add(, , rs!ID_EQUIPO)
            .SubItems(COLS.COL_NOMBRE_EQUIPO) = rs!nombre_equipo
            .SubItems(COL_PLAN_MTO) = rs!nombre_plan
            .SubItems(COL_PERIODICIDAD) = rs!nombre_periodicicad
            .SubItems(COL_ANNO) = CStr(rs!ANNO)
            .SubItems(COL_ULT_MTO) = CStr(rs!FECHA_ULT_MTO)
            .SubItems(COL_ID_PLAN_MTO) = rs!PLAN_MANTENIMIENTO_ID
            .SubItems(COL_ID_PERIODICIDAD) = rs!PERIODICIDAD_ID
        End With
        rs.MoveNext
    Wend
        
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    
    log (Me.Name)

    cabecera
    
    carga_lista
    
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Item(1).Text = "NºEquipo"
        .Item(1).Width = lista.Width * 0.09
        .Item(1).Alignment = lvwColumnLeft
        
        .Add , , "Nombre Equipo", lista.Width * 0.426, lvwColumnLeft
        .Add , , "Plan Mto.", lista.Width * 0.1, lvwColumnCenter
        .Add , , "Periodicidad", lista.Width * 0.18, lvwColumnLeft
        .Add , , "Año", lista.Width * 0.07, lvwColumnCenter
        .Add , , "Ult. Mto.", lista.Width * 0.11, lvwColumnCenter
        .Add , , "plan_mto_id", 0, lvwColumnLeft
        .Add , , "periodicidad_id", 0, lvwColumnLeft
    End With




End Sub

Private Sub lista_DblClick()
Dim objfrm As New frmEquipoCrearFechasMtoPrevisto
Dim oM As New clsEquipoMantenimiento
Dim lng_id_responsable As Long
Dim str_fecha_ult As String

With lista.SelectedItem
    If InStr(.SubItems(COLS.COL_PERIODICIDAD), "*") <> 0 Then
        MsgBox "No se pueden crear Registros de Mantenimiento Previstos para este tipo de Periodicidad desde esta opción.", vbInformation, "Crear Registros Mantenimiento Presvistos"
        Exit Sub
    End If

    If .SubItems(COLS.COL_ULT_MTO) <> "--" Then
        oM.Carga_ultimo_mto_realizado CLng(.Text), CLng(.SubItems(COLS.COL_ID_PLAN_MTO))
        lng_id_responsable = oM.getMANTENEDOR_ID
        str_fecha_ult = oM.getFECHA_ACTUAL
    Else
        lng_id_responsable = 0
        str_fecha_ult = "--"
    End If

    objfrm.EQUIPO_ID = CLng(.Text)
    objfrm.EQUIPO = CLng(.SubItems(COLS.COL_NOMBRE_EQUIPO))
    objfrm.PLAN_ID = CLng(.SubItems(COLS.COL_ID_PLAN_MTO))
    objfrm.RESPONSABLE_ID = lng_id_responsable
    objfrm.FECHA_ULT_MTO = str_fecha_ult
    objfrm.ANNO = CInt(.SubItems(COLS.COL_ANNO))
    
    objfrm.Show 1
        
    Set objfrm = Nothing
End With

carga_lista
    
End Sub


