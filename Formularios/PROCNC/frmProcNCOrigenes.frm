VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProcNCOrigenes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Origenes Incidencia"
   ClientHeight    =   6630
   ClientLeft      =   2760
   ClientTop       =   2535
   ClientWidth     =   10860
   Icon            =   "frmProcNCOrigenes.frx":0000
   LinkTopic       =   "frmProcNCOrigenes"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   10860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSDataListLib.DataCombo cmbDepartamentos 
      Height          =   315
      Left            =   2160
      TabIndex        =   6
      Top             =   4260
      Width           =   4605
      _ExtentX        =   8123
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
      Left            =   9750
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   5640
      Width           =   1020
   End
   Begin VB.CommandButton cmdAnadirDpto 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6810
      Picture         =   "frmProcNCOrigenes.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Añadir accesorio"
      Top             =   4290
      Width           =   285
   End
   Begin VB.CommandButton cmdEliminarDpto 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   7140
      Picture         =   "frmProcNCOrigenes.frx":052F
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar accesorio"
      Top             =   4290
      Width           =   285
   End
   Begin VB.TextBox txtOtros 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   3450
      Width           =   9975
   End
   Begin VB.ListBox lstOrigen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   1380
      Index           =   3
      Left            =   5430
      Style           =   1  'Checkbox
      TabIndex        =   2
      Top             =   1980
      Width           =   5385
   End
   Begin VB.ListBox lstOrigen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   1380
      Index           =   2
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1980
      Width           =   5385
   End
   Begin VB.ListBox lstOrigen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   1380
      Index           =   4
      Left            =   5430
      Style           =   1  'Checkbox
      TabIndex        =   3
      Top             =   240
      Width           =   5385
   End
   Begin VB.ListBox lstOrigen 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Height          =   1380
      Index           =   1
      Left            =   0
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   240
      Width           =   5355
   End
   Begin VB.CheckBox chkOtros 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Otros"
      Height          =   225
      Left            =   60
      TabIndex        =   4
      Tag             =   "15"
      Top             =   3480
      Width           =   705
   End
   Begin MSComctlLib.ListView lstDepartamentos 
      Height          =   1935
      Left            =   60
      TabIndex        =   9
      Top             =   4620
      Width           =   7365
      _ExtentX        =   12991
      _ExtentY        =   3413
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
   Begin VB.Label Label6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Departamentos Implicados"
      Height          =   225
      Left            =   90
      TabIndex        =   15
      Top             =   4320
      Width           =   2025
   End
   Begin VB.Label lblCapCausas 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Sistema de Calidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   0
      TabIndex        =   14
      Top             =   30
      Width           =   5385
   End
   Begin VB.Label lblCapCausas 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Metrologia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   2
      Left            =   5460
      TabIndex        =   13
      Top             =   30
      Width           =   5355
   End
   Begin VB.Label lblCapCausas 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fallo Técnico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   4
      Left            =   30
      TabIndex        =   12
      Top             =   1740
      Width           =   5355
   End
   Begin VB.Label lblCapCausas 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   5
      Left            =   5430
      TabIndex        =   11
      Top             =   1740
      Width           =   5385
   End
End
Attribute VB_Name = "frmProcNCOrigenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private mvarobjProcNC As New clsProcNc
Private mvarblnEditable As Boolean
Private rs As ADODB.RecordSet
Private strSql As String

Private mvarblnBloqueoClick As Boolean
Private Sub Form_Load()

    log Me.Name
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
    With lstDepartamentos.ColumnHeaders
        .Add , , "id", 0, lvwColumnLeft
        .Add , , "Departamento", lstDepartamentos.Width, lvwColumnLeft
    End With
End Sub

Private Sub cargar_datos()

mvarobjProcNC.Carga PK

PresentarDatos_Origenes
PresentarDatos_Departamentos


End Sub

Private Sub cargar_listados()


    Dim x As Integer
    Dim oDeco As New clsDecodificadora
    
    For x = 1 To 4
        lstOrigen(x).Clear
        
        Set rs = mvarobjProcNC.devolver_listado_origenes(x)
        
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            While Not rs.EOF
                lstOrigen(x).AddItem rs!descripcion
                lstOrigen(x).ItemData(lstOrigen(x).ListCount - 1) = rs!valor
                rs.MoveNext
            Wend
        End If
        
    Next x


    ' Ahora carga la Combo
    oDeco.cargar_combo cmbDepartamentos, decodificadora.PROCNC_DEPARTAMENTOS
    

End Sub

Private Sub opciones_edicion()

    'lstDepartamentos.FlatScrollBar = True

    'lstDepartamentos.Enabled = mvarblnEditable
    'lstOrigen(1).Enabled = mvarblnEditable
    'lstOrigen(2).Enabled = mvarblnEditable
    'lstOrigen(3).Enabled = mvarblnEditable
    'lstOrigen(4).Enabled = mvarblnEditable
    
    chkOtros.Enabled = mvarblnEditable
    txtOtros.Enabled = mvarblnEditable And txtOtros.Enabled
    cmdAnadirDpto.Enabled = mvarblnEditable
    cmdEliminarDpto.Enabled = mvarblnEditable
    cmbDepartamentos.Enabled = mvarblnEditable

End Sub

Private Sub guardar_datos()

    If chkOtros.value = vbUnchecked Then
        mvarobjProcNC.guardar_datos_origen_incidencia ""
    Else
        mvarobjProcNC.guardar_datos_origen_incidencia Trim(txtOtros.Text)
    End If

End Sub

Private Sub PresentarDatos_Departamentos()

    Set rs = mvarobjProcNC.devolver_departamentos_origen()

    lstDepartamentos.ListItems.Clear
    
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
            With lstDepartamentos.ListItems.Add(, , rs("id_departamento"))
                .SubItems(1) = rs("departamento")
            End With
            rs.MoveNext
        Wend
    End If

End Sub

Private Sub PresentarDatos_Origenes()

    Set rs = mvarobjProcNC.devolver_origenes_incidencia()
    Dim x As Integer, idx As Integer

    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
            If rs("id_origen_incidencia") = 15 Then
               chkOtros.value = vbChecked
               txtOtros.Text = mvarobjProcNC.getORIGEN_OTROS
               txtOtros.Enabled = True
            Else
                For x = 1 To 4
                    For idx = 0 To lstOrigen(x).ListCount - 1
                        If CInt(rs("id_origen_incidencia")) = lstOrigen(x).ItemData(idx) Then
                            lstOrigen(x).Selected(idx) = True
                            Exit For
                        End If
                    Next idx
                Next x
            End If
            rs.MoveNext
        Wend
    End If

    
    
    

End Sub

Private Sub cmdcancel_Click()

If Not mvarblnEditable Then Unload Me

guardar_datos

Unload Me

End Sub

Private Sub cmdAnadirDpto_Click()
    
    Dim lngid As Long
    
    lngid = getDataComboSel(cmbDepartamentos)
    If lngid <= 0 Then
        Exit Sub
    End If
    
    If mvarobjProcNC.anadir_departamento_origen(CStr(lngid)) Then
        Call PresentarDatos_Departamentos
    End If
    
End Sub

Private Sub cmdEliminarDpto_Click()

If lstDepartamentos.ListItems.Count = 0 Then Exit Sub

Dim lngid As Long

    lngid = lstDepartamentos.SelectedItem
        
    mvarobjProcNC.eliminar_departamento_origen lngid
    PresentarDatos_Departamentos
    
End Sub

Private Sub chkOtros_Click()
If chkOtros.value = vbChecked Then
    txtOtros.Enabled = True
    mvarobjProcNC.anadir_origen_incidencia chkOtros.Tag, chkOtros.Caption
Else
    txtOtros.Enabled = False
    mvarobjProcNC.eliminar_origen_incidencia CLng(chkOtros.Tag)
End If
End Sub

Private Sub lstOrigen_ItemCheck(Index As Integer, Item As Integer)

Static bloqueo_local As Boolean

If Item < 0 Then Exit Sub
If bloqueo_local Then Exit Sub

If mvarblnBloqueoClick Then
    bloqueo_local = True
    
    lstOrigen(Index).Selected(Item) = Not lstOrigen(Index).Selected(Item)
    mvarblnBloqueoClick = False
    bloqueo_local = False
    Exit Sub
End If

Dim blnMarcar As Boolean, ID As Long, nombre As String

blnMarcar = lstOrigen(Index).Selected(Item)
ID = lstOrigen(Index).ItemData(Item)
nombre = lstOrigen(Index).Text

If blnMarcar Then
    mvarobjProcNC.anadir_origen_incidencia ID, nombre
Else
    mvarobjProcNC.eliminar_origen_incidencia ID
End If




End Sub

Private Sub lstOrigen_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
If Not mvarblnEditable Then KeyCode = 0
End Sub


Private Sub lstOrigen_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

If Not mvarblnEditable Then mvarblnBloqueoClick = True

End Sub


