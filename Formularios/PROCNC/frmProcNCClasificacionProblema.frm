VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProcNCClasificacionProblema 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Clasificacion del Problema"
   ClientHeight    =   9465
   ClientLeft      =   1800
   ClientTop       =   1395
   ClientWidth     =   12210
   Icon            =   "frmProcNCClasificacionProblema.frx":0000
   LinkTopic       =   "frmProcNCClasificacionProblema"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9465
   ScaleWidth      =   12210
   Begin VB.CommandButton cmdAdjuntar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   990
      Left            =   11310
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7095
      Width           =   855
   End
   Begin VB.ListBox lstCausas 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   930
      Index           =   5
      ItemData        =   "frmProcNCClasificacionProblema.frx":6852
      Left            =   1290
      List            =   "frmProcNCClasificacionProblema.frx":6854
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   5610
      Width           =   10875
   End
   Begin VB.ListBox lstCausas 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   930
      Index           =   4
      ItemData        =   "frmProcNCClasificacionProblema.frx":6856
      Left            =   1290
      List            =   "frmProcNCClasificacionProblema.frx":6858
      Style           =   1  'Checkbox
      TabIndex        =   6
      Top             =   4650
      Width           =   10875
   End
   Begin VB.ListBox lstCausas 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   930
      Index           =   3
      ItemData        =   "frmProcNCClasificacionProblema.frx":685A
      Left            =   1290
      List            =   "frmProcNCClasificacionProblema.frx":685C
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   3690
      Width           =   10875
   End
   Begin VB.ListBox lstCausas 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   930
      Index           =   2
      ItemData        =   "frmProcNCClasificacionProblema.frx":685E
      Left            =   1290
      List            =   "frmProcNCClasificacionProblema.frx":6860
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   2730
      Width           =   10875
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   11130
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   8520
      Width           =   1020
   End
   Begin VB.TextBox txtTotalProblemas 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2130
      MaxLength       =   65000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   10005
   End
   Begin VB.TextBox txtAfectados 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2130
      MaxLength       =   65000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1200
      Width           =   10005
   End
   Begin VB.TextBox txtAlcance 
      Appearance      =   0  'Flat
      Height          =   495
      Left            =   2130
      MaxLength       =   65000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   660
      Width           =   10005
   End
   Begin VB.ListBox lstCausas 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   930
      Index           =   1
      ItemData        =   "frmProcNCClasificacionProblema.frx":6862
      Left            =   1290
      List            =   "frmProcNCClasificacionProblema.frx":6864
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   1770
      Width           =   10875
   End
   Begin VB.Frame fraProblemasHumanos 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   1320
      TabIndex        =   8
      Top             =   6600
      Width           =   4665
      Begin VB.Frame fraProblemasHumanos_pregunta 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   5
         Left            =   90
         TabIndex        =   24
         Top             =   1500
         Width           =   4515
         Begin VB.OptionButton OptProblemasHumanos_proceso_inusual_complejo_si 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   3780
            TabIndex        =   25
            Top             =   0
            Width           =   225
         End
         Begin VB.OptionButton OptProblemasHumanos_proceso_inusual_complejo_no 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   4170
            TabIndex        =   26
            Top             =   0
            Width           =   225
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Es un proceso inusual o complejo?"
            Height          =   225
            Index           =   3
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   2550
         End
      End
      Begin VB.Frame fraProblemasHumanos_pregunta 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   21
         Top             =   1260
         Width           =   4515
         Begin VB.OptionButton OptProblemasHumanos_objetivos_marcados_claramente_no 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   4170
            TabIndex        =   23
            Top             =   0
            Width           =   225
         End
         Begin VB.OptionButton OptProblemasHumanos_objetivos_marcados_claramente_si 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   3780
            TabIndex        =   22
            Top             =   0
            Width           =   225
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Han quedado los objetivos claramente marcados?"
            Height          =   225
            Index           =   2
            Left            =   0
            TabIndex        =   34
            Top             =   0
            Width           =   3630
         End
      End
      Begin VB.Frame fraProblemasHumanos_pregunta 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   18
         Top             =   1020
         Width           =   4515
         Begin VB.OptionButton OptProblemasHumanos_formacion_suficiente_no 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   4170
            TabIndex        =   19
            Top             =   0
            Width           =   225
         End
         Begin VB.OptionButton OptProblemasHumanos_formacion_suficiente_si 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   3780
            TabIndex        =   20
            Top             =   0
            Width           =   225
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Ha sido suficiente la formación?"
            Height          =   225
            Index           =   5
            Left            =   0
            TabIndex        =   33
            Top             =   0
            Width           =   2340
         End
      End
      Begin VB.Frame fraProblemasHumanos_pregunta 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   15
         Top             =   780
         Width           =   4515
         Begin VB.OptionButton OptProblemasHumanos_herramientas_adecuadas_si 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   3780
            TabIndex        =   16
            Top             =   0
            Width           =   225
         End
         Begin VB.OptionButton OptProblemasHumanos_herramientas_adecuadas_no 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   4170
            TabIndex        =   17
            Top             =   0
            Width           =   225
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Son adecuadas las Herramientas de Trabajo?"
            Height          =   225
            Index           =   1
            Left            =   0
            TabIndex        =   32
            Top             =   0
            Width           =   3330
         End
      End
      Begin VB.Frame fraProblemasHumanos_pregunta 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   1
         Left            =   90
         TabIndex        =   12
         Top             =   540
         Width           =   4515
         Begin VB.OptionButton OptProblemasHumanos_instrucciones_incompletas_si 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   3780
            TabIndex        =   13
            Top             =   0
            Width           =   225
         End
         Begin VB.OptionButton OptProblemasHumanos_instrucciones_incompletas_no 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   4170
            TabIndex        =   14
            Top             =   0
            Width           =   225
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Están incompletas las Instrucciones de Tabajo?"
            Height          =   225
            Index           =   4
            Left            =   0
            TabIndex        =   31
            Top             =   0
            Width           =   3465
         End
      End
      Begin VB.Frame fraProblemasHumanos_pregunta 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   300
         Width           =   4515
         Begin VB.OptionButton OptProblemasHumanos_operador_sustituido_no 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   4170
            TabIndex        =   11
            Top             =   0
            Width           =   225
         End
         Begin VB.OptionButton OptProblemasHumanos_operador_sustituido_si 
            BackColor       =   &H00C0C0C0&
            Height          =   225
            Left            =   3780
            TabIndex        =   10
            Top             =   0
            Width           =   225
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Ha sido sustituido el Operador?"
            Height          =   225
            Index           =   0
            Left            =   0
            TabIndex        =   30
            Top             =   0
            Width           =   2295
         End
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sí  -  No"
         Height          =   195
         Index           =   9
         Left            =   3870
         TabIndex        =   36
         Top             =   90
         Width           =   600
      End
   End
   Begin MSComctlLib.ListView lstDocumentacion 
      Height          =   1695
      Left            =   6090
      TabIndex        =   27
      Top             =   6780
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   2990
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
      NumItems        =   0
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "(Doble Click para abrir archivo)"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   9060
      TabIndex        =   46
      Top             =   6540
      Width           =   2175
   End
   Begin VB.Label lblCapCausas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Problemas Humanos"
      Height          =   465
      Index           =   6
      Left            =   30
      TabIndex        =   45
      Top             =   6660
      Width           =   1215
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "¿Hay más de un problema?"
      Height          =   195
      Index           =   8
      Left            =   60
      TabIndex        =   44
      Top             =   180
      Width           =   1935
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "¿A qué/quién afecta?"
      Height          =   195
      Index           =   7
      Left            =   60
      TabIndex        =   43
      Top             =   1230
      Width           =   1560
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "¿Cual es el alcance?"
      Height          =   195
      Index           =   6
      Left            =   60
      TabIndex        =   42
      Top             =   690
      Width           =   1485
   End
   Begin VB.Label lblCapCausas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Problemas de Requerimientos"
      Height          =   435
      Index           =   0
      Left            =   30
      TabIndex        =   37
      Top             =   1830
      Width           =   1215
   End
   Begin VB.Label lblCapCausas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Problemas de Equipamiento / Material"
      Height          =   585
      Index           =   1
      Left            =   30
      TabIndex        =   38
      Top             =   2760
      Width           =   1215
   End
   Begin VB.Label lblCapCausas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Problemas de Producción"
      Height          =   585
      Index           =   2
      Left            =   30
      TabIndex        =   39
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label lblCapCausas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Problemas de Aseguramiento de la Calidad"
      Height          =   585
      Index           =   4
      Left            =   30
      TabIndex        =   40
      Top             =   4680
      Width           =   1185
   End
   Begin VB.Label lblCapCausas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Problemas de Planificación"
      Height          =   495
      Index           =   5
      Left            =   30
      TabIndex        =   41
      Top             =   5610
      Width           =   1185
   End
End
Attribute VB_Name = "frmProcNCClasificacionProblema"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long


Private mvarblnBloqueoClick As Boolean

Private mvarobjProcNC As New clsProcNc
Private RS As ADODB.RecordSet
Private strSql As String
Private mvarblnEditable As Boolean



Private Function guardar_datos() As Boolean
On Error GoTo guardar_datos_Error
    
Dim int_operador_sustituido As Integer
Dim int_instrucciones_incompletas As Integer
Dim int_objetivos_marcados_claramente As Integer
Dim int_formacion_suficiente As Integer
Dim int_herramientas_adecuadas As Integer
Dim int_proceso_inusual_complejo As Integer
    
    
int_operador_sustituido = 0
int_instrucciones_incompletas = 0
int_objetivos_marcados_claramente = 0
int_formacion_suficiente = 0
int_herramientas_adecuadas = 0
int_proceso_inusual_complejo = 0

If OptProblemasHumanos_operador_sustituido_si.value Then int_operador_sustituido = 1
    
If OptProblemasHumanos_instrucciones_incompletas_si.value Then int_instrucciones_incompletas = 1
    
If OptProblemasHumanos_objetivos_marcados_claramente_si.value Then int_objetivos_marcados_claramente = 1
    
If OptProblemasHumanos_herramientas_adecuadas_si.value Then int_herramientas_adecuadas = 1
    
If OptProblemasHumanos_formacion_suficiente_si.value Then int_formacion_suficiente = 1

If OptProblemasHumanos_proceso_inusual_complejo_si.value Then int_proceso_inusual_complejo = 1
    
guardar_datos = False
    
mvarobjProcNC.guardar_datos_clasificacion_problemas txtTotalProblemas.Text, txtAfectados.Text, txtAlcance.Text, _
int_operador_sustituido, int_instrucciones_incompletas, int_objetivos_marcados_claramente, _
int_formacion_suficiente, int_herramientas_adecuadas, int_proceso_inusual_complejo
    
    
   guardar_datos = True
On Error GoTo 0
    Exit Function
    
guardar_datos_Error:
    guardar_datos = False
End Function

Private Sub opciones_edicion()

    txtTotalProblemas.Enabled = mvarblnEditable
    txtAfectados.Enabled = mvarblnEditable
    txtAlcance.Enabled = mvarblnEditable
    'lstCausas(1).Enabled = mvarblnEditable
    'lstCausas(2).Enabled = mvarblnEditable
    'lstCausas(3).Enabled = mvarblnEditable
    'lstCausas(4).Enabled = mvarblnEditable
    'lstCausas(5).Enabled = mvarblnEditable
    'lstDocumentacion.Enabled = mvarblnEditable
    fraProblemasHumanos.Enabled = mvarblnEditable
    cmdAdjuntar.Enabled = mvarblnEditable
    

End Sub

Private Sub PresentarDatos_Otros()

    With mvarobjProcNC
        
        txtTotalProblemas.Text = .getPROBLEMAS_CLASIFICACION_MAS_DE_UN_PROBLEMA
        txtAfectados.Text = .getPROBLEMAS_CLASIFICACION_AFECTADOS
        txtAlcance.Text = .getPROBLEMAS_CLASIFICACION_ALCANCE
        
        If Not .getPROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS Then OptProblemasHumanos_herramientas_adecuadas_no.value = True
        If .getPROBLEMAS_HUMANOS_HERRAMIENTAS_ADECUADAS Then OptProblemasHumanos_herramientas_adecuadas_si.value = True
        
        If Not .getPROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS Then OptProblemasHumanos_instrucciones_incompletas_no.value = True
        If .getPROBLEMAS_HUMANOS_INSTRUCCIONES_INCOMPLETAS Then OptProblemasHumanos_instrucciones_incompletas_si.value = True
        
        If Not .getPROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE Then OptProblemasHumanos_objetivos_marcados_claramente_no.value = True
        If .getPROBLEMAS_HUMANOS_OBJETIVOS_MARCADOS_CLARAMENTE Then OptProblemasHumanos_objetivos_marcados_claramente_si.value = True
        
        If Not .getPROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO Then OptProblemasHumanos_operador_sustituido_no.value = True
        If .getPROBLEMAS_HUMANOS_OPERADOR_SUSTITUIDO Then OptProblemasHumanos_operador_sustituido_si.value = True
        
        If Not .getPROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO Then OptProblemasHumanos_proceso_inusual_complejo_no.value = True
        If .getPROBLEMAS_HUMANOS_PROCESO_INUSUAL_COMPLEJO Then OptProblemasHumanos_proceso_inusual_complejo_si.value = True
        
        If Not .getPROBLEMAS_HUMANOS_SUFICIENTE_FORMACION Then OptProblemasHumanos_formacion_suficiente_no.value = True
        If .getPROBLEMAS_HUMANOS_SUFICIENTE_FORMACION Then OptProblemasHumanos_formacion_suficiente_si.value = True
        
    End With
    
End Sub

Private Sub cmdAdjuntar_Click()
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_PROCNC_IDENTIFICACION_PROBLEMA
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
    Call PresentarDatos_DocumentosAdjuntos
End Sub

Private Sub cmdcancel_Click()

    If Not mvarblnEditable Then Unload Me

    If Not guardar_datos Then Exit Sub

    Unload Me
    
End Sub



Public Property Get Editable() As Boolean

    Editable = mvarblnEditable

End Property

Public Property Let Editable(ByVal blnEditable As Boolean)

    mvarblnEditable = blnEditable

End Property

Private Sub Form_Activate()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_Load()
    cabecera
    cargar_botones Me

    cargar_listados
    
    cargar_datos

    opciones_edicion
End Sub


Private Sub cargar_listados()


    Dim x As Integer
    
    For x = 1 To 5
        lstCausas(x).Clear
    Next x
        
        Set RS = mvarobjProcNC.devolver_listado_clasificacion_problemas(x)
        
        If RS.RecordCount <> 0 Then
            RS.MoveFirst
            While Not RS.EOF
                x = RS("id_tipocausa")
                lstCausas(x).AddItem RS!DESCRIPCION
                lstCausas(x).ItemData(lstCausas(x).ListCount - 1) = RS!id_causa
                RS.MoveNext
            Wend
        End If
        

    

End Sub
Private Sub cargar_datos()

mvarobjProcNC.Carga PK

PresentarDatos_Clasificacion
PresentarDatos_DocumentosAdjuntos

PresentarDatos_Otros


End Sub

Private Sub PresentarDatos_Clasificacion()

    Set RS = mvarobjProcNC.devolver_clasificacion_problemas_incidencia()
    Dim x As Integer, idx As Integer

    If RS.RecordCount <> 0 Then
        RS.MoveFirst
        While Not RS.EOF
            x = RS("id_tipocausa")
            For idx = 0 To lstCausas(x).ListCount - 1
                If CInt(RS("id_causa")) = lstCausas(x).ItemData(idx) Then
                    lstCausas(x).Selected(idx) = True
                    Exit For
                End If
            Next idx
            RS.MoveNext
        Wend
    End If

End Sub
Private Sub PresentarDatos_DocumentosAdjuntos()
    lstDocumentacion.ListItems.Clear
    Dim oAdjunto As New clsAdjuntos
    Dim RS As ADODB.RecordSet
    Set RS = oAdjunto.Listado(TOBJETO.TOBJETO_PROCNC_IDENTIFICACION_PROBLEMA, PK, "", "")
    If RS.RecordCount > 0 Then
        Do
            With lstDocumentacion.ListItems.Add(, , RS(0))
                 .SubItems(1) = RS(2)
            End With
            RS.MoveNext
        Loop Until RS.EOF
    End If
End Sub
Private Sub cabecera()
    With lstDocumentacion.ColumnHeaders
        .Add , , "id", 0, lvwColumnLeft
        .Add , , "Documento", 4995, lvwColumnLeft
    End With
End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then cmdcancel_Click
    
End Sub


Private Sub lstCausas_ItemCheck(Index As Integer, Item As Integer)

Static bloqueo_local As Boolean

    Dim blnMarcar As Boolean, ID As Long, NOMBRE As String
    
    
    If mvarblnBloqueoClick Then
        bloqueo_local = True
        
        lstCausas(Index).Selected(Item) = Not lstCausas(Index).Selected(Item)
        mvarblnBloqueoClick = False
        bloqueo_local = False
        Exit Sub
    End If
    
    blnMarcar = lstCausas(Index).Selected(Item)
    ID = lstCausas(Index).ItemData(Item)
    NOMBRE = lstCausas(Index).Text
    
    If blnMarcar Then
        mvarobjProcNC.anadir_clasificacion_problema_incidencia ID
    Else
        mvarobjProcNC.eliminar_clasificacion_problema_incidencia ID
    End If
End Sub


Private Sub lstCausas_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
If Not mvarblnEditable Then mvarblnBloqueoClick = True

End Sub
Private Sub lstDocumentacion_DblClick()
    If lstDocumentacion.ListItems.Count = 0 Then Exit Sub
    Dim oAdjunto As New clsAdjuntos
    oAdjunto.CargarDocumento TOBJETO.TOBJETO_PROCNC_IDENTIFICACION_PROBLEMA, PK, 0, lstDocumentacion.ListItems(lstDocumentacion.selectedItem.Index).Text, True
    Set oAdjunto = Nothing
End Sub

