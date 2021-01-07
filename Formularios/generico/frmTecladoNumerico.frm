VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmTecladoNumerico 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Teclado En Pantalla"
   ClientHeight    =   5280
   ClientLeft      =   10635
   ClientTop       =   2655
   ClientWidth     =   3240
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   3240
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox cmdguion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   3825
      Width           =   810
   End
   Begin VB.CommandButton cmdSiguiente 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Index           =   1
      Left            =   2430
      Picture         =   "frmTecladoNumerico.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   945
      Width           =   810
   End
   Begin MSComCtl2.DTPicker txtFecha 
      Height          =   495
      Left            =   450
      TabIndex        =   24
      Top             =   90
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   873
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckBox        =   -1  'True
      DateIsNull      =   -1  'True
      Format          =   16515073
      CurrentDate     =   40247
   End
   Begin VB.CommandButton cmdCancelar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   9
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   270
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4185
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.CheckBox chkNulo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "- -"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   3825
      Width           =   810
   End
   Begin VB.CheckBox chkMenorQue 
      BackColor       =   &H00E0E0E0&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   3825
      Width           =   810
   End
   Begin VB.CheckBox chkMayorque 
      BackColor       =   &H00E0E0E0&
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   810
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   3825
      Width           =   810
   End
   Begin VB.TextBox txtSubCabecera 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   270
      Width           =   3240
   End
   Begin VB.TextBox txtCabecera 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   0
      Width           =   3240
   End
   Begin VB.CommandButton cmdMasMenos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "+/-"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4545
      Width           =   810
   End
   Begin VB.CommandButton cmdSiguiente 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1440
      Index           =   0
      Left            =   2430
      Picture         =   "frmTecladoNumerico.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   2385
      Width           =   810
   End
   Begin VB.CommandButton cmdComa 
      BackColor       =   &H00E0E0E0&
      Caption         =   ","
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   810
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   3105
      Width           =   810
   End
   Begin VB.CommandButton cmdNumero 
      BackColor       =   &H00E0E0E0&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   0
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3105
      Width           =   810
   End
   Begin VB.CommandButton cmdNumero 
      BackColor       =   &H00E0E0E0&
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   9
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   945
      Width           =   810
   End
   Begin VB.CommandButton cmdNumero 
      BackColor       =   &H00E0E0E0&
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   8
      Left            =   810
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   945
      Width           =   810
   End
   Begin VB.CommandButton cmdNumero 
      BackColor       =   &H00E0E0E0&
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   7
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   945
      Width           =   810
   End
   Begin VB.CommandButton cmdNumero 
      BackColor       =   &H00E0E0E0&
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   6
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1665
      Width           =   810
   End
   Begin VB.CommandButton cmdNumero 
      BackColor       =   &H00E0E0E0&
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   5
      Left            =   810
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1665
      Width           =   810
   End
   Begin VB.CommandButton cmdNumero 
      BackColor       =   &H00E0E0E0&
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   4
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1665
      Width           =   810
   End
   Begin VB.CommandButton cmdNumero 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   3
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2385
      Width           =   810
   End
   Begin VB.CommandButton cmdNumero 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   2
      Left            =   810
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2385
      Width           =   810
   End
   Begin VB.CommandButton cmdNumero 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Index           =   1
      Left            =   0
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2385
      Width           =   810
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   2430
      Picture         =   "frmTecladoNumerico.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4545
      Width           =   810
   End
   Begin VB.TextBox txtResultado 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   11.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   0
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   540
      Width           =   3240
   End
   Begin VB.CommandButton cmdDel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Del"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   14.25
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1620
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3105
      Width           =   810
   End
   Begin VB.CheckBox chkConforme 
      BackColor       =   &H00E0E0E0&
      Caption         =   "CONFORME"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   810
      Picture         =   "frmTecladoNumerico.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   4545
      Width           =   810
   End
   Begin VB.CheckBox chkNoConforme 
      BackColor       =   &H00E0E0E0&
      Caption         =   "NO CONFORME"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   1620
      Picture         =   "frmTecladoNumerico.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   4545
      Width           =   810
   End
End
Attribute VB_Name = "frmTecladoNumerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event Change(ByVal res As String)
Public Event SiguienteElemento(ByRef cabecera As String, ByRef Subcabecera As String, ByRef RESULTADO As String, ByRef fecha As String, ByRef CONFORME As Integer, ByRef Cerrar As Boolean, ByRef desestimarEvento As Boolean)
Public Event AnteriorElemento(ByRef cabecera As String, ByRef Subcabecera As String, ByRef RESULTADO As String, ByRef fecha As String, ByRef CONFORME As Integer, ByRef Cerrar As Boolean, ByRef desestimarEvento As Boolean)
Public Event EstablecerConformidad(ByVal VALOR As Integer)
Public Event EstablecerFecha(ByVal fecha As String)
Public Event Salir()

Private mvarstrTextoInicial As String
Private mvarstrCabecera As String
Private mvarstrSubCabecera As String
Private mvarstrFecha As String
Private mvarblnCargando As Boolean
Private mvarintConforme As Integer
Private mvarblnOcultarConformidad As Boolean
Private mvarlngPosY As Long
Private mvarlngPosX As Long

Public Property Get OcultarConformidad() As Boolean

    OcultarConformidad = mvarblnOcultarConformidad

End Property

Public Property Let OcultarConformidad(ByVal blnOcultarConformidad As Boolean)

    mvarblnOcultarConformidad = blnOcultarConformidad

End Property

Public Property Get cabecera() As String

    cabecera = mvarstrCabecera

End Property

Public Property Let cabecera(ByVal strCabecera As String)

    mvarstrCabecera = strCabecera

    If Me.Visible Then
        txtCabecera.Text = strCabecera
    End If

End Property

Private Sub des_activar_botones(ByVal activo As Boolean)
cmdComa.Enabled = activo
cmdNumero(0).Enabled = activo
cmdNumero(1).Enabled = activo
cmdNumero(2).Enabled = activo
cmdNumero(3).Enabled = activo
cmdNumero(4).Enabled = activo
cmdNumero(5).Enabled = activo
cmdNumero(6).Enabled = activo
cmdNumero(7).Enabled = activo
cmdNumero(8).Enabled = activo
cmdNumero(9).Enabled = activo
cmdDel.Enabled = activo
cmdMasMenos.Enabled = activo
chkMayorque.Enabled = activo
chkMenorQue.Enabled = activo

End Sub

Private Sub evaluar_negativo_nulo_mayormenor_que(ByVal cad As String)
If InStr(1, cad, ">") > 0 Then
    chkMayorque.value = vbChecked
Else
    chkMayorque.value = vbUnchecked
End If

If InStr(1, cad, "<") > 0 Then
    chkMenorQue.value = vbChecked
Else
    chkMenorQue.value = vbUnchecked
End If

If Trim(cad) = "--" Or Trim(cad) = "- -" Then
    chkNulo.value = vbChecked
Else
    chkNulo.value = vbUnchecked
    txtResultado.Text = cad
    chkNulo_Click
End If
End Sub

Private Function MayMenQue() As String
If chkMayorque.value = vbChecked Then
    MayMenQue = "> "
ElseIf chkMenorQue.value = vbChecked Then
    MayMenQue = "< "
Else
    MayMenQue = ""
End If
End Function


Private Sub cmdCancelar_Click()
RaiseEvent EstablecerConformidad(mvarintConforme)
'RaiseEvent EstablecerFecha("")
RaiseEvent Change(mvarstrTextoInicial)
Me.Hide
RaiseEvent Salir

End Sub

Private Sub cmdcancel_Click()
    Me.Hide
    RaiseEvent Salir
End Sub

Private Sub cmdComa_Click()
If InStr(1, txtResultado, ".") = 0 Then
    ' SI NO ENCUENTRA LA COMA
    txtResultado.Text = txtResultado.Text & ","
    RaiseEvent Change(txtResultado.Text)
End If

End Sub

Private Sub cmdDel_Click()
    If Trim(txtResultado.Text) <> "" Then
        txtResultado.Text = Left(txtResultado.Text, Len(txtResultado.Text) - 1)
        RaiseEvent Change(txtResultado.Text)
    End If
End Sub

Private Sub cmdguion_Click()
    txtResultado.Text = txtResultado.Text & "-"
    RaiseEvent Change(txtResultado.Text)
End Sub

Private Sub cmdMasMenos_Click()

If (Left(txtResultado.Text, 1) = "-" Or Left(txtResultado.Text, 3) = "> -" Or Left(txtResultado.Text, 3) = "< -") Then
    txtResultado.Text = Replace(txtResultado.Text, "-", "")
Else
    txtResultado.Text = MayMenQue & "-" & Replace(Replace(txtResultado.Text, "< ", ""), "> ", "")
End If
RaiseEvent Change(txtResultado.Text)

End Sub

Private Sub cmdNumero_Click(Index As Integer)
txtResultado.Text = txtResultado.Text & CStr(Index)
RaiseEvent Change(txtResultado.Text)
End Sub


Public Property Get TextoInicial() As String

    TextoInicial = mvarstrTextoInicial

End Property

Public Property Let TextoInicial(ByVal strTextoInicial As String)

    mvarstrTextoInicial = strTextoInicial

    If Me.Visible Then
        txtResultado.Text = strTextoInicial
        evaluar_negativo_nulo_mayormenor_que strTextoInicial
    End If
    
End Property



Private Sub chkConforme_Click()
If chkConforme.value = vbChecked Then
    chkNoConforme.value = vbUnchecked
    RaiseEvent EstablecerConformidad(1)
Else
    RaiseEvent EstablecerConformidad(-1)
End If

End Sub


Private Sub chkMayorque_Click()

If chkMayorque.value = vbChecked Then _
    chkMenorQue.value = vbUnchecked

txtResultado.Text = MayMenQue & Replace(Replace(txtResultado.Text, "< ", ""), "> ", "")
RaiseEvent Change(txtResultado.Text)
End Sub

Private Sub chkMenorQue_Click()
If chkMenorQue.value = vbChecked Then _
    chkMayorque.value = vbUnchecked

txtResultado.Text = MayMenQue & Replace(Replace(txtResultado.Text, "< ", ""), "> ", "")
RaiseEvent Change(txtResultado.Text)

End Sub


Private Sub chkNoConforme_Click()
If chkNoConforme.value = vbChecked Then
    chkConforme.value = vbUnchecked
    RaiseEvent EstablecerConformidad(0)
Else
    RaiseEvent EstablecerConformidad(-1)
End If

End Sub

Private Sub chkNulo_Click()

If chkNulo.value = vbChecked Then
    des_activar_botones False
    txtResultado.Text = "--"
Else
    des_activar_botones True
    If mvarstrTextoInicial <> "--" Then
        txtResultado.Text = mvarstrTextoInicial
    Else
        txtResultado.Text = ""
    End If
    chkMayorque.value = vbUnchecked
    chkMenorQue.value = vbUnchecked
End If
    
If Not mvarblnCargando Then RaiseEvent Change(txtResultado.Text)
    
End Sub



Private Sub cmdSiguiente_Click(Index As Integer)
    Dim cab As String, subcab As String, res As String, Cerrar As Boolean, fecha As String, CONFORME As Integer, desestimarEvento As Boolean
    Cerrar = False
    CONFORME = -1
    fecha = ""
    desestimarEvento = False
    If Index = 0 Then
        RaiseEvent SiguienteElemento(cab, subcab, res, fecha, CONFORME, Cerrar, desestimarEvento)
    Else
        RaiseEvent AnteriorElemento(cab, subcab, res, fecha, CONFORME, Cerrar, desestimarEvento)
    End If
    If Cerrar Then
        Me.Visible = False
        Exit Sub
    ElseIf desestimarEvento Or Me.Visible = False Then
        Exit Sub
    End If
    evaluar_negativo_nulo_mayormenor_que res
' JGM-I
    txtCabecera.Text = cab
    txtSubCabecera.Text = subcab
    txtResultado.Text = res
    RaiseEvent Change(res)
' JGM-F
    ' Conformidad o no
    If CONFORME = 1 Then
        chkConforme.value = vbChecked
        chkNoConforme.value = vbUnchecked
    ElseIf CONFORME = 0 Then
        chkConforme.value = vbUnchecked
        chkNoConforme.value = vbChecked
    Else
        chkConforme.value = vbUnchecked
        chkNoConforme.value = vbUnchecked
    End If
    ' Fecha
    If Trim(fecha) = "" And IsDate(fecha) Then
        txtfecha.value = CDate(fecha)
    Else
        txtfecha.value = Null
    End If

End Sub

Private Sub Form_Activate()

    mvarblnCargando = True
    
    evaluar_negativo_nulo_mayormenor_que mvarstrTextoInicial
    
    mvarblnCargando = False
    
    txtResultado.Text = mvarstrTextoInicial
    txtCabecera.Text = mvarstrCabecera
    txtSubCabecera.Text = mvarstrSubCabecera
    If Trim(mvarstrFecha) <> "" And IsDate(mvarstrFecha) Then
        txtfecha.value = CDate(mvarstrFecha)
    Else
        txtfecha.value = Null
    End If
    
    
    Select Case mvarintConforme
        Case 1
            chkConforme.value = vbChecked
            chkNoConforme.value = vbUnchecked
        Case 0
            chkConforme.value = vbUnchecked
            chkNoConforme.value = vbChecked
        Case Else
            mvarintConforme = -1
            chkConforme.value = vbUnchecked
            chkNoConforme.value = vbUnchecked
    End Select
    
    'Me.Top = 0
    'Me.Left = Screen.Width - Me.Width
    Me.Top = mvarlngPosY
    Me.Left = mvarlngPosX
    
    

End Sub


Public Property Get Subcabecera() As String

    Subcabecera = mvarstrSubCabecera

End Property

Public Property Let Subcabecera(ByVal strSubCabecera As String)

    mvarstrSubCabecera = strSubCabecera

    If Me.Visible Then
        txtSubCabecera.Text = strSubCabecera
    End If

End Property

Private Sub Form_Initialize()
mvarlngPosX = 0
mvarlngPosY = 0
End Sub

Private Sub Form_Load()

' Height Original = 6315
' Height sin conformidades = 5490

If mvarblnOcultarConformidad Then
    chkConforme.Visible = False
    chkConforme.Enabled = False
    chkNoConforme.Visible = False
    chkNoConforme.Enabled = False
'    Me.Height = 5490
End If

End Sub

Private Sub txtFecha_Change()
Dim f As String

If IsNull(txtfecha.value) Then
    f = ""
Else
    f = Format(txtfecha.value, "dd/mm/yyyy")
End If

RaiseEvent EstablecerFecha(f)
End Sub
Public Property Get fecha() As String

    fecha = mvarstrFecha

End Property

Public Property Let fecha(ByVal strFecha As String)

    mvarstrFecha = strFecha

End Property

Public Property Get CONFORME() As Integer

    CONFORME = mvarintConforme

End Property

Public Property Let CONFORME(ByVal intConforme As Integer)

    mvarintConforme = intConforme

End Property

Public Property Get PosY() As Long

    PosY = mvarlngPosY

End Property

Public Property Let PosY(ByVal lngPosY As Long)

    mvarlngPosY = lngPosY

End Property

Public Property Get PosX() As Long

    PosX = mvarlngPosX

End Property

Public Property Let PosX(ByVal lngPosX As Long)

    mvarlngPosX = lngPosX

End Property

Private Sub txtResultado_GotFocus()
    txtResultado.SelStart = 0
    txtResultado.SelLength = Len(txtResultado)
End Sub

