VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProcNCPersonalInvestigacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Personal Investigacion"
   ClientHeight    =   4545
   ClientLeft      =   2655
   ClientTop       =   3210
   ClientWidth     =   7215
   Icon            =   "frmProcNCPersonalInvestigacion.frx":0000
   LinkTopic       =   "frmPersonalInvestigacion"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   7215
   Begin MSDataListLib.DataCombo cmbDepartamentos 
      Height          =   315
      Left            =   1680
      TabIndex        =   1
      Top             =   30
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   6150
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3570
      Width           =   1020
   End
   Begin VB.CommandButton cmdEliminarPersonal 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6900
      Picture         =   "frmProcNCPersonalInvestigacion.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Eliminar accesorio"
      Top             =   420
      Width           =   285
   End
   Begin VB.CommandButton cmdAnadirPersonal 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6570
      Picture         =   "frmProcNCPersonalInvestigacion.frx":0A5E
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Añadir accesorio"
      Top             =   420
      Visible         =   0   'False
      Width           =   285
   End
   Begin VB.CommandButton cmdAnadirDpto 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6570
      Picture         =   "frmProcNCPersonalInvestigacion.frx":0C83
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Añadir accesorio"
      Top             =   60
      Width           =   285
   End
   Begin VB.ListBox lstEquipo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   2730
      Left            =   60
      Style           =   1  'Checkbox
      TabIndex        =   7
      Top             =   780
      Width           =   7095
   End
   Begin VB.OptionButton optGrupo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Por Usuario"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   480
      Width           =   1365
   End
   Begin VB.OptionButton optDepartamento 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Por Departamento"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Value           =   -1  'True
      Width           =   1635
   End
   Begin MSDataListLib.DataCombo cmbPersonal 
      Height          =   315
      Left            =   1680
      TabIndex        =   4
      Top             =   390
      Visible         =   0   'False
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label lblCap 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nota: Debe Marcar el Jefe de Equipo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   3570
      Width           =   4245
   End
End
Attribute VB_Name = "frmProcNCPersonalInvestigacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private mvarobjProcNC As New clsProcNc
Private mvarblnEditable As Boolean
Private rs As ADODB.RecordSet
Private strSql As String

Private mvarlngid_jefe As Long

Private mvarblnModificandoCheckLista As Boolean

Private mvarblnBloqueoClick As Boolean
Private Sub cmdAnadirDpto_Click()

Dim lngid As Long
    
    lngid = getDataComboSel(cmbDepartamentos)
    
    If lngid <= 0 Then
        Exit Sub
    End If

    mvarobjProcNC.anadir_equipo_investigacion_departamento lngid
    
    
    mvarblnModificandoCheckLista = True
    
    PresentarDatos_Equipo

    mvarblnModificandoCheckLista = False
End Sub

Private Sub cmdAnadirPersonal_Click()
Dim lngid As Long
    
    lngid = getDataComboSel(cmbPersonal)
    
    If lngid <= 0 Then
        Exit Sub
    End If

    mvarobjProcNC.anadir_equipo_investigacion_usuario lngid
    
    mvarblnModificandoCheckLista = True
        PresentarDatos_Equipo
    mvarblnModificandoCheckLista = False
    
End Sub


Private Sub cmdEliminarPersonal_Click()

    If lstEquipo.ListIndex < 0 Then Exit Sub

    mvarobjProcNC.eliminar_equipo_investigacion_usuario lstEquipo.ItemData(lstEquipo.ListIndex)
    
    
    mvarblnModificandoCheckLista = True
        PresentarDatos_Equipo
    mvarblnModificandoCheckLista = False

End Sub


Private Sub Form_Activate()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_Load()

    mvarblnModificandoCheckLista = False

    cabecera
    cargar_botones Me

    cargar_listados
    
    cargar_datos

    opciones_edicion

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then cmdcancel_Click
    
End Sub

Public Property Get Editable() As Boolean

    Editable = mvarblnEditable

End Property

Public Property Let Editable(ByVal blnEditable As Boolean)

    mvarblnEditable = blnEditable

End Property

Private Sub cabecera()
    'With lstEquipo.ColumnHeaders
    '    .Add , , "id", 0, lvwColumnLeft
    '    .Add , , "Personal", 12000, lvwColumnLeft
    'End With
End Sub

Private Sub cargar_datos()

mvarobjProcNC.Carga PK

mvarlngid_jefe = -1


mvarblnModificandoCheckLista = True
    
PresentarDatos_Equipo

mvarblnModificandoCheckLista = False

End Sub

Private Sub cargar_listados()

    Dim oDeco As New clsDecodificadora
    ' Ahora carga la Combo
    oDeco.cargar_combo cmbDepartamentos, decodificadora.PROCNC_DEPARTAMENTOS
    cargar_combo cmbPersonal, New clsUsuarios
    
    Set oDeco = Nothing

End Sub

Private Sub opciones_edicion()

    'lstEquipo.Enabled = mvarblnEditable
    optDepartamento.Enabled = mvarblnEditable
    optGrupo.Enabled = mvarblnEditable
    cmbDepartamentos.Enabled = mvarblnEditable
    cmbPersonal.Enabled = mvarblnEditable
    cmdAnadirDpto.Enabled = mvarblnEditable
    cmdAnadirPersonal.Enabled = mvarblnEditable
    cmdEliminarPersonal.Enabled = mvarblnEditable

End Sub



Private Sub PresentarDatos_Equipo()

    Set rs = mvarobjProcNC.devolver_equipo_investigacion()

    lstEquipo.Clear

    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
            lstEquipo.AddItem rs("usuario")
            lstEquipo.ItemData(lstEquipo.ListCount - 1) = rs("id_usuario")
            
            If CInt(rs("jefe_equipo")) = 1 Then
                mvarlngid_jefe = rs("id_usuario")
                lstEquipo.Selected(lstEquipo.ListCount - 1) = True
            End If
            
            rs.MoveNext
        Wend
    End If


End Sub

Private Sub cmdcancel_Click()

    If Not mvarblnEditable Then Unload Me

    If lstEquipo.ListCount > 0 Then
    
        If Not mvarobjProcNC.comprobar_equipo_investigacion_tiene_jefe_equipo Then
            MsgBox "Debe señalar al menos un usuario como Jefe de Equipo de Investigacion"
            Exit Sub
        End If
    End If

Unload Me

End Sub

Private Sub lstEquipo_ItemCheck(Item As Integer)
Static bloqueo_local As Boolean

If bloqueo_local Then Exit Sub

    If Item < 0 Then Exit Sub
        
    If mvarblnBloqueoClick Then
        bloqueo_local = True
        
        lstEquipo.Selected(Item) = Not lstEquipo.Selected(Item)
        mvarblnBloqueoClick = False
        bloqueo_local = False
        Exit Sub
    End If
        
    If lstEquipo.ItemData(Item) = mvarlngid_jefe Then
        If Not lstEquipo.Selected(Item) Then
            lstEquipo.Selected(Item) = True
        End If
    Else
        mvarobjProcNC.establecer_jefe_equipo lstEquipo.ItemData(Item)
    End If


    If mvarblnModificandoCheckLista Then Exit Sub

    PresentarDatos_Equipo

End Sub


Private Sub lstEquipo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Not mvarblnEditable Then mvarblnBloqueoClick = True
End Sub


Private Sub optDepartamento_Click()

    If optDepartamento.value Then
        cmbDepartamentos.Visible = True
        cmdAnadirDpto.Visible = True
        
        cmbPersonal.Visible = False
        cmdAnadirPersonal.Visible = False
    End If
    
End Sub

Private Sub optGrupo_Click()
    If optGrupo.value Then
        cmbDepartamentos.Visible = False
        cmdAnadirDpto.Visible = False
        
        cmbPersonal.Visible = True
        cmdAnadirPersonal.Visible = True
        
    End If

End Sub
