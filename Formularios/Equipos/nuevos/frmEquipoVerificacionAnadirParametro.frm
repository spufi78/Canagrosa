VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEquipoVerificacionAnadirParametro 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Añadir Parámetro Verificación"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNMedidas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2805
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin MSDataListLib.DataCombo cmbTipo 
      Height          =   315
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.TextBox txtRangoMax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   7890
      TabIndex        =   3
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtRangoMin 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   6930
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txtDescripcion 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   120
      MaxLength       =   100
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "Cancelar"
      Height          =   1035
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   780
      Width           =   1245
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Aceptar"
      Height          =   1035
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   780
      Width           =   1245
   End
   Begin MSDataListLib.DataCombo cmbUnidades 
      Height          =   315
      Left            =   5040
      TabIndex        =   1
      Top             =   360
      Width           =   1905
      _ExtentX        =   3360
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Medidas"
      Height          =   195
      Index           =   5
      Left            =   2865
      TabIndex        =   13
      Top             =   750
      Width           =   885
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Unidad"
      Height          =   195
      Index           =   4
      Left            =   5070
      TabIndex        =   12
      Top             =   150
      Width           =   1845
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo"
      Height          =   195
      Index           =   3
      Left            =   150
      TabIndex        =   11
      Top             =   750
      Width           =   885
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rango Max."
      Height          =   195
      Index           =   2
      Left            =   7920
      TabIndex        =   10
      Top             =   150
      Width           =   885
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rango Min."
      Height          =   195
      Index           =   1
      Left            =   6990
      TabIndex        =   9
      Top             =   150
      Width           =   885
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Punto de Verificacion"
      Height          =   195
      Index           =   0
      Left            =   150
      TabIndex        =   8
      Top             =   120
      Width           =   1875
   End
End
Attribute VB_Name = "frmEquipoVerificacionAnadirParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarstrDescripcion As String
Private mvarlngTipo As Long
Private mvardblrmin As Double
Private mvardblrmax As Double
Private mvarintmedidas As Integer
Private mvarblnResultado As Boolean
Private mvarlngid_unidad As Long
Private mvarstrUnidad As String


Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    
    cargar_combo cmbUnidades, New clsUnidades
    oDeco.cargar_combo cmbTipo, DECODIFICADORA.EQ_TIPOS_PARAMETROS_RESULTADO

End Sub

Private Function comprobar_datos() As Boolean

Dim cad As String
Dim ra As String, ri As String

cad = ""
If getDataComboSel(cmbTipo) > 0 Then
    If Trim(txtNMedidas.Text) = "" Then
        cad = cad & vbCrLf & " - Debe indicar un Nº de Medidas (Mínimo 1)"
    ElseIf CInt(txtNMedidas.Text) = 0 Then
        cad = cad & vbCrLf & " - Debe indicar un Nº de Medidas (Mínimo 1)"
    End If
End If

If getDataComboSel(cmbUnidades) <= 0 And cmbUnidades.Enabled Then
    cad = cad & vbCrLf & " - Debe indicar una Unidad válida"
End If

ra = ""
ri = ""
If txtRangoMin.Enabled Then
    If Trim(txtRangoMin.Text) = "" Then
        cad = cad & vbCrLf & " - Debe indicar un Rango Mínimo válido"
    ElseIf Not IsNumeric(txtRangoMin.Text) Then
        cad = cad & vbCrLf & " - Debe indicar un Rango Mínimo válido"
    Else
        ri = txtRangoMin.Text
    End If
End If

If txtRangoMax.Enabled Then
    If Trim(txtRangoMax.Text) = "" Then
        cad = cad & vbCrLf & " - Debe indicar un Rango Máximo válido"
    ElseIf Not IsNumeric(txtRangoMax.Text) Then
        cad = cad & vbCrLf & " - Debe indicar un Rango Máximo válido"
    Else
        ra = txtRangoMax.Text
    End If
End If


If Trim(ra) <> "" And Trim(ri) <> "" Then
    If CDbl(ri) > CDbl(ra) Then
        cad = cad & vbCrLf & " - Rango Máximo no puede ser inferior al Rango Mínimo"
    End If
End If

If Trim(txtdescripcion.Text) = "" Then
    cad = cad & vbCrLf & " - Debe indicar un Punto de Verificación Correcto"
End If


If cad <> "" Then
    MsgBox "Se han detectado los siguientes errores: " & cad
    Exit Function
End If

comprobar_datos = True

End Function

Private Sub PresentarDatos()
    
    cmbUnidades.BoundText = mvarlngid_unidad
    cmbTipo.BoundText = mvarlngTipo
    cmbTipo_change
    
    txtdescripcion.Text = mvarstrDescripcion
    If txtRangoMin.Enabled Then txtRangoMin.Text = mvardblrmin
    If txtRangoMax.Enabled Then txtRangoMax.Text = mvardblrmax
    
    If txtNMedidas.Enabled Then txtNMedidas.Text = IIf(mvarintmedidas = 0, 1, mvarintmedidas)
    
    
    
    
End Sub

Private Sub recoger_datos()
    
    mvarstrDescripcion = txtdescripcion.Text
    mvarlngTipo = getDataComboSel(cmbTipo)
    If cmbUnidades.Enabled Then
        mvarlngid_unidad = getDataComboSel(cmbUnidades)
        mvarstrUnidad = cmbUnidades.Text
    Else
        mvarlngid_unidad = 0
        mvarstrUnidad = "N/A"
    End If
    If mvarlngTipo > 0 Then
        mvardblrmin = CDbl(txtRangoMin)
        mvardblrmax = CDbl(txtRangoMax)
        mvarintmedidas = CInt(txtNMedidas)
    Else
        mvardblrmin = 0
        mvardblrmax = 0
        mvarintmedidas = 1
    End If

End Sub

Private Sub cmbTipo_change()
If getDataComboSel(cmbTipo) = 0 Then
    txtNMedidas.Text = "N/A"
    txtNMedidas.Enabled = False
    txtRangoMin.Text = "N/A"
    txtRangoMin.Enabled = False
    txtRangoMax.Text = "N/A"
    txtRangoMax.Enabled = False
    cmbUnidades.BoundText = 0
    cmbUnidades.Enabled = False
Else
    txtNMedidas.Text = "1"
    txtNMedidas.Enabled = True
    txtRangoMin.Text = "0"
    txtRangoMin.Enabled = True
    txtRangoMax.Text = "0"
    txtRangoMax.Enabled = True
    
    cmbUnidades.Enabled = True
End If
End Sub

Private Sub cmdcancel_Click()
    mvarblnResultado = False
    Me.Hide

End Sub

Private Sub cmdok_Click()

    If Not comprobar_datos Then Exit Sub

    recoger_datos
    
    mvarblnResultado = True
    Me.Hide
    
End Sub

Private Sub Form_Load()
log Me.Name

cargar_botones Me

cargar_combos

PresentarDatos

End Sub


Public Property Get DESCRIPCION() As String

    DESCRIPCION = mvarstrDescripcion

End Property

Public Property Let DESCRIPCION(ByVal strDescripcion As String)

    mvarstrDescripcion = strDescripcion

End Property

Public Property Get tipo() As Long

    tipo = mvarlngTipo

End Property

Public Property Let tipo(ByVal lngTipo As Long)

    mvarlngTipo = lngTipo

End Property

Public Property Get rmin() As Double

    rmin = mvardblrmin

End Property

Public Property Let rmin(ByVal dblrmin As Double)

    mvardblrmin = dblrmin

End Property

Public Property Get rmax() As Double

    rmax = mvardblrmax

End Property

Public Property Let rmax(ByVal dblrmax As Double)

    mvardblrmax = dblrmax

End Property

Public Property Get medidas() As Integer

    medidas = mvarintmedidas

End Property

Public Property Let medidas(ByVal intmedidas As Integer)

    mvarintmedidas = intmedidas

End Property

Public Property Get RESULTADO() As Boolean

    RESULTADO = mvarblnResultado

End Property

Public Property Let RESULTADO(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Public Property Get id_unidad() As Long

    id_unidad = mvarlngid_unidad

End Property

Public Property Let id_unidad(ByVal lngid_unidad As Long)

    mvarlngid_unidad = lngid_unidad

End Property





Private Sub txtNMedidas_GotFocus()
txtNMedidas.SelStart = 0
txtNMedidas.SelLength = Len(txtNMedidas)
End Sub

Private Sub txtNMedidas_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloNumerico(txtNMedidas, KeyAscii, False)
End Sub


Private Sub txtRangoMax_GotFocus()
txtRangoMax.SelStart = 0
txtRangoMax.SelLength = Len(txtRangoMax)

End Sub

Private Sub txtRangoMin_GotFocus()
txtRangoMin.SelStart = 0
txtRangoMin.SelLength = Len(txtRangoMin)

End Sub

Private Sub txtRangoMin_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtRangoMin, KeyAscii, True)
End Sub
Private Sub txtRangoMax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtRangoMax, KeyAscii, True)
End Sub



Public Property Get Unidad() As String

    Unidad = mvarstrUnidad

End Property

Public Property Let Unidad(ByVal strUnidad As String)

    mvarstrUnidad = strUnidad

End Property

