VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipoParametrosRes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parametros de Resultados"
   ClientHeight    =   7125
   ClientLeft      =   3210
   ClientTop       =   2880
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7125
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosParametro 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Parámetro"
      Height          =   1665
      Left            =   30
      TabIndex        =   11
      Top             =   30
      Width           =   7005
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1410
         MaxLength       =   100
         TabIndex        =   0
         Top             =   210
         Width           =   4560
      End
      Begin VB.TextBox txtIncertidumbre 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4470
         MaxLength       =   100
         TabIndex        =   5
         Text            =   "0"
         Top             =   870
         Width           =   1500
      End
      Begin VB.TextBox txtCorrecion 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   4470
         MaxLength       =   100
         TabIndex        =   3
         Text            =   "0"
         Top             =   540
         Width           =   1500
      End
      Begin VB.TextBox txtTolerancia 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1410
         MaxLength       =   100
         TabIndex        =   4
         Text            =   "0"
         Top             =   855
         Width           =   1500
      End
      Begin VB.TextBox txtRangoMax 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2220
         MaxLength       =   100
         TabIndex        =   2
         Text            =   "0"
         Top             =   540
         Width           =   690
      End
      Begin VB.TextBox txtRangoMin 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1410
         MaxLength       =   100
         TabIndex        =   1
         Text            =   "0"
         Top             =   540
         Width           =   690
      End
      Begin pryCombo.miCombo cmbUnidades 
         Height          =   330
         Left            =   1410
         TabIndex        =   6
         Top             =   1200
         Width           =   5625
         _ExtentX        =   9922
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   3
         Left            =   150
         TabIndex        =   18
         Top             =   240
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incertidumbre"
         Height          =   195
         Index           =   2
         Left            =   3210
         TabIndex        =   17
         Top             =   945
         Width           =   960
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Correción"
         Height          =   195
         Index           =   1
         Left            =   3210
         TabIndex        =   16
         Top             =   615
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tolerancia"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   900
         Width           =   750
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades"
         Height          =   195
         Index           =   66
         Left            =   690
         TabIndex        =   14
         Top             =   1290
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rango Min - Max"
         Height          =   195
         Index           =   62
         Left            =   150
         TabIndex        =   13
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "-"
         Height          =   195
         Index           =   61
         Left            =   2130
         TabIndex        =   12
         Top             =   600
         Width           =   45
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   7800
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   6210
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadirParametro 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir "
      Height          =   570
      Left            =   7110
      Picture         =   "frmEquipoParametrosRes.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Añadir accesorio"
      Top             =   1110
      Width           =   870
   End
   Begin VB.CommandButton cmdEliminarParametro 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   570
      Left            =   8010
      Picture         =   "frmEquipoParametrosRes.frx":0225
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Eliminar accesorio"
      Top             =   1110
      Width           =   870
   End
   Begin MSFlexGridLib.MSFlexGrid grdParam 
      Height          =   4095
      Left            =   60
      TabIndex        =   9
      Top             =   2070
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   7223
      _Version        =   393216
      FixedCols       =   0
      HighLight       =   2
      SelectionMode   =   1
      Appearance      =   0
   End
   Begin VB.Label lblCap 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Param"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   60
      TabIndex        =   19
      Top             =   1710
      Width           =   8775
   End
End
Attribute VB_Name = "frmEquipoParametrosRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarobjEquipo As clsEquipos
Private mvarintid_tipo_parametro As Integer

Private Sub GuardarDatos()


    Call mvarobjEquipo.AnadirParametroResultado(mvarobjEquipo.getID_EQUIPO, mvarintid_tipo_parametro, txtDescripcion.Text, _
    cmbUnidades.getPK_SALIDA, txtRangoMin.Text, _
    txtRangoMax.Text, txtTolerancia.Text, _
    txtCorrecion.Text, txtIncertidumbre.Text)


End Sub

Private Sub select_txt(ByRef txt As TextBox)
    txt.SelStart = 0
    txt.SelLength = Len(txt.Text)
End Sub

Private Sub cmdAnadirParametro_Click()



    If Not ComprobarDatos Then Exit Sub

    Call GuardarDatos
    
    Call PresentarDatos


End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdEliminarParametro_Click()


Dim lngFila As Long, id_param As String


    lngFila = grdParam.RowSel
    If lngFila <= 0 Then Exit Sub
    
    id_param = grdParam.TextMatrix(lngFila, 0)
    mvarobjEquipo.EliminarParametroResultado id_param
    
    Call PresentarDatos
    

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjEquipo = Nothing

End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combo
    Call configuraGrid
        
    Call PresentarDatos
    
End Sub


Private Sub configuraGrid()

    With grdParam
        .Cols = 8
        .ColWidth(0) = 0
        .ColWidth(1) = .Width * 0.3 ' Desc
        .ColWidth(2) = .Width * 0.2 ' Unidad
        .ColWidth(3) = .Width * 0.1 ' Rango Max
        .ColWidth(4) = .Width * 0.1 ' Rango Min
        .ColWidth(5) = .Width * 0.1 ' Tolerancia
        .ColWidth(6) = .Width * 0.1 ' Correcion
        .ColWidth(7) = .Width * 0.1 ' Incertidumbre
        
        
        .TextMatrix(0, 1) = "Descripción"
        .TextMatrix(0, 2) = "Unidad"
        .TextMatrix(0, 3) = "R.Min"
        .TextMatrix(0, 4) = "R.Max"
        .TextMatrix(0, 5) = "Toler."
        .TextMatrix(0, 6) = "Correc."
        .TextMatrix(0, 7) = "Incert."
        
        .Rows = 1
    End With

End Sub

Private Sub PresentarDatos()
Dim rs As ADOdb.RecordSet
    
    Set rs = mvarobjEquipo.DevolverParametrosResultados(mvarobjEquipo.getID_EQUIPO, mvarintid_tipo_parametro)
    
    If mvarintid_tipo_parametro = 1 Then
        lblCap.Caption = "Parámetros de Resultados para Calibraciones de este equipo"
    Else
        lblCap.Caption = "Parámetros de Resultados para Verificaciones de este equipo"
    End If
    Me.Caption = lblCap.Caption
    
    With grdParam
        .Rows = 1
        If rs.RecordCount <> 0 Then
            rs.MoveFirst
            While Not rs.EOF
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = rs("id_parametro")
                .TextMatrix(.Rows - 1, 1) = rs("descripcion")
                .TextMatrix(.Rows - 1, 2) = rs("unidad")
                .TextMatrix(.Rows - 1, 3) = rs("rango_min")
                .TextMatrix(.Rows - 1, 4) = rs("rango_max")
                .TextMatrix(.Rows - 1, 5) = rs("tolerancia_max")
                .TextMatrix(.Rows - 1, 6) = rs("correccion")
                .TextMatrix(.Rows - 1, 7) = rs("incertidumbre")
                rs.MoveNext
            Wend
        End If
    End With
            
End Sub


Public Property Get Equipo() As clsEquipos

    Set Equipo = mvarobjEquipo

End Property

Public Property Set Equipo(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Public Property Get id_tipo_parametro() As Integer

    id_tipo_parametro = mvarintid_tipo_parametro

End Property

Public Property Let id_tipo_parametro(ByVal intid_tipo_parametro As Integer)

    mvarintid_tipo_parametro = intid_tipo_parametro

End Property

Private Sub cargar_combo()
    llenar_combo cmbUnidades, New clsUnidades, 0, Me, ""
End Sub

Private Function ComprobarDatos() As Boolean

Dim strCad As String

    ComprobarDatos = False
    strCad = ""
    
    If Trim(txtDescripcion.Text) = "" Then _
        strCad = strCad & vbCrLf & " - Debe indicar la descripción del Parámetro"
        
    If cmbUnidades.getPK_SALIDA <= 0 Then _
        strCad = strCad & vbCrLf & " - Debe indicar las Inidades del Parámetro"
    
    If strCad <> "" Then
        ComprobarDatos = False
        MsgBox "Se han detectado los siguientes errores: " & strCad, vbInformation, "Añadir Parámetro de Resultado"
    Else
        ComprobarDatos = True
    End If

End Function


Private Sub txtCorrecion_GotFocus()
select_txt txtCorrecion
End Sub

Private Sub txtCorrecion_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtCorrecion, KeyAscii, True)

End Sub


Private Sub txtIncertidumbre_GotFocus()
select_txt txtIncertidumbre
End Sub

Private Sub txtIncertidumbre_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtIncertidumbre, KeyAscii, True)

End Sub

Private Sub txtRangoMax_GotFocus()
select_txt txtRangoMax
End Sub

Private Sub txtRangoMax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtRangoMax, KeyAscii, True)

End Sub

Private Sub txtRangoMin_GotFocus()
select_txt txtRangoMin

End Sub

Private Sub txtRangoMin_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtRangoMin, KeyAscii, True)
End Sub

Private Sub txtTolerancia_GotFocus()
select_txt txtTolerancia
End Sub

Private Sub txtTolerancia_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtTolerancia, KeyAscii, True)

End Sub


