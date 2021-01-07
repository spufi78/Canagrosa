VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEmpleados_Gestion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Empleados"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmEmpleados_Gestion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNominas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nominas"
      Enabled         =   0   'False
      Height          =   915
      Left            =   3990
      Picture         =   "frmEmpleados_Gestion.frx":09EA
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   5940
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   915
      Left            =   7950
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   5940
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   915
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   5940
      Width           =   1275
   End
   Begin VB.CommandButton cmdAnticipo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anticipos"
      Enabled         =   0   'False
      Height          =   915
      Left            =   2690
      Picture         =   "frmEmpleados_Gestion.frx":12B4
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   5940
      Width           =   1275
   End
   Begin VB.CommandButton cmdControl 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Control"
      Enabled         =   0   'False
      Height          =   915
      Left            =   1390
      Picture         =   "frmEmpleados_Gestion.frx":15BE
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   5940
      Width           =   1275
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   7110
      TabIndex        =   27
      Top             =   495
      Width           =   2085
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   360
         Left            =   90
         TabIndex        =   28
         Top             =   225
         Width           =   1890
         _ExtentX        =   3334
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdExpediente 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Expediente"
      Enabled         =   0   'False
      Height          =   915
      Left            =   90
      Picture         =   "frmEmpleados_Gestion.frx":1E88
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   5940
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5760
      Top             =   6075
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   2625
      Left            =   7110
      TabIndex        =   23
      Top             =   1260
      Width           =   2100
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   11
         Left            =   90
         MaxLength       =   30
         TabIndex        =   10
         Top             =   2160
         Width           =   1920
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1710
         Left            =   315
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apodo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   90
         TabIndex        =   25
         Top             =   1935
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Index           =   13
      Left            =   45
      TabIndex        =   21
      Top             =   4275
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   1230
         Index           =   9
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   240
         Width           =   8985
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos del Operario "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3750
      Left            =   15
      TabIndex        =   12
      Top             =   450
      Width           =   6975
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   13
         Left            =   1350
         MaxLength       =   25
         TabIndex        =   7
         Top             =   2790
         Width           =   4875
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   12
         Left            =   4500
         MaxLength       =   6
         TabIndex        =   6
         Top             =   2370
         Width           =   1725
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   10
         Left            =   1350
         TabIndex        =   8
         Top             =   3240
         Width           =   3975
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   390
         Left            =   5430
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3225
         Width           =   1095
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   4500
         MaxLength       =   30
         TabIndex        =   4
         Top             =   1980
         Width           =   2385
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1980
         Width           =   1725
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   8
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   5
         Top             =   2370
         Width           =   1725
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1350
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1140
         Width           =   960
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   1
         Top             =   735
         Width           =   5520
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   0
         Top             =   330
         Width           =   5505
      End
      Begin MSDataListLib.DataCombo cmbProvincia 
         Height          =   360
         Left            =   3420
         TabIndex        =   35
         Top             =   1140
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbMunicipio 
         Height          =   315
         Left            =   1350
         TabIndex        =   36
         Top             =   1575
         Width           =   5535
         _ExtentX        =   9763
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.C.C."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   9
         Left            =   225
         TabIndex        =   31
         Top             =   2850
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo Interno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   8
         Left            =   3150
         TabIndex        =   30
         Top             =   2430
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imagen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   210
         TabIndex        =   24
         Top             =   3285
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   7
         Left            =   210
         TabIndex        =   20
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Movil"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   15
         Left            =   3810
         TabIndex        =   19
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   11
         Left            =   210
         TabIndex        =   18
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.I.F."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   225
         TabIndex        =   17
         Top             =   2460
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2460
         TabIndex        =   16
         Top             =   1170
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   210
         TabIndex        =   15
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   225
         TabIndex        =   14
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   225
         TabIndex        =   13
         Top             =   375
         Width           =   735
      End
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nuevo Empleado"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   60
      TabIndex        =   22
      Top             =   30
      Width           =   9180
   End
End
Attribute VB_Name = "frmEmpleados_Gestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbProvincia_Change()
    If cmbProvincia.Text <> "" Then
        cargar_municipios (cmbProvincia.BoundText)
    End If
End Sub


Private Sub cmdAnticipo_Click()
    frmEmpleados_Anticipo.Show 1
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdControl_Click()
    frmEmpleados_Control.Show 1
End Sub

Private Sub cmdExpediente_Click()
    frmEmpleados_Expediente.Show 1
End Sub
Private Sub cmdEXplorar_Click()
    cd.DialogTitle = "Abrir fichero de imagen"
    cd.InitDir = App.Path & "\recursos\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        txtdatos(10).Text = cd.FileName  ' cd.FileTitle
    End If
End Sub

Private Sub cmdNominas_Click()
    frmEmpleados_Nominas.Show 1
End Sub

Private Sub cmdok_Click()
    If gOperario > 0 Then
        modificar_Operario
    Else
        insertar_Operario
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    
    cargar_estados
    ' LP005
    cargar_combo cmbProvincia, New clsProvincias
    permisos
    If gOperario > 0 Then
        consulta_Operario
        cmdExpediente.Enabled = True
        cmdControl.Enabled = True
        cmdAnticipo.Enabled = True
        cmdNominas.Enabled = True
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmEmpleados_Gestion = Nothing
End Sub

Private Sub txtdatos_Change(Index As Integer)
    On Error Resume Next
    If Index = 10 And txtdatos(10) <> "" Then
        If Dir(txtdatos(10)) <> "" Then
            Set img.Picture = LoadPicture(txtdatos(10))
        End If
    End If
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub

Private Sub txtdatos_Keyup(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40 ' Abajo
       If Index = 9 Then
        txtdatos(1).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 1 Then
        txtdatos(9).SetFocus
       Else
        SendKeys "+{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 27
        cmdcancel_Click
     Case 121 ' F10
        cmdok_Click
    End Select
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 9 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = &HFFFFFF
End Sub

Public Sub borrar_campos()
    Dim i As Integer
    For i = 1 To 13
        ' LP005
        If i <> 4 And i <> 5 Then
            txtdatos(i) = ""
        End If
    Next
    txtdatos(1).SetFocus
End Sub

Public Sub bloquear_campos()
    Dim i As Integer
    For i = 1 To 13
        ' LP005
        If i <> 4 And i <> 5 Then
        txtdatos(i).Locked = True
        End If
    Next
End Sub

Public Sub insertar_Operario()
    If valida_datos = False Then
        Exit Sub
    End If
    pregunta = "Va a dar de alta el empleado. ¿Esta seguro?"
    If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim aux As Long
        Set oOperario = mover_datos
        aux = gOperario
        gOperario = oOperario.Insertar
'        If aux = -1 Then
            Unload Me
'            Exit Sub
'        End If
'        If gOperario <> 0 Then
'            borrar_campos
'        End If
'        Set oOperario = Nothing
'        img.Picture = LoadPicture()
    End If
End Sub

Public Sub modificar_Operario()
    If valida_datos() = False Then
        Exit Sub
    End If
    Dim pos As Integer
    Dim operario As Integer
    pregunta = "Va a modificar los datos del empleado. ¿Esta seguro?"
    If MsgBox(pregunta, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Set oOperario = mover_datos
        If oOperario.Modificar(gOperario) = True Then
            Unload Me
        End If
        Set oOperario = Nothing
    End If

End Sub
Public Function valida_datos() As Boolean
    valida_datos = True
    If txtdatos(1) = "" Then
        MsgBox "El nombre del empleado no puede estar en blanco.", vbCritical, "Error"
        txtdatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtdatos(11) = "" Then
        MsgBox "El apodo del empleado no puede estar en blanco.", vbCritical, "Error"
        txtdatos(11).SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbestados.BoundText = "" Then
        MsgBox "Debe seleccionar un estado.", vbCritical, "Error"
        cmbestados.SetFocus
        valida_datos = False
        Exit Function
    End If
    
End Function

Public Sub consulta_Operario()
    On Error GoTo fallo
    Dim oOperario As New clsEmpleados
    lbltitulo.Caption = "Modificacion de empleado"
    oOperario.cargar (gOperario)
    With oOperario
        txtdatos(1) = .getNOMBRE
        txtdatos(2) = .getDIRECCION
        txtdatos(3) = .getCP
        ' LP005
'        txtdatos(4) = .getPROVINCIA
        cmbProvincia.BoundText = .getPROVINCIA_ID
        cargar_municipios .getPROVINCIA_ID
        cmbMunicipio.BoundText = .getMUNICIPIO_ID
        txtdatos(6) = .getTELEFONO
        txtdatos(7) = .getMOVIL
        txtdatos(8) = .getCIF
        txtdatos(9) = .getOBSERVACIONES
        txtdatos(10) = .getfoto
        txtdatos(11) = .getAPODO
        txtdatos(12) = .getCODIGO_INTERNO
        txtdatos(13) = .getCCC
        cmbestados.BoundText = .getESTADO_ID
    End With
    Set oOperario = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del empleado.", vbCritical, Err.Description
End Sub
Public Function mover_datos() As clsEmpleados
    On Error GoTo fallo
    Dim oOperario As New clsEmpleados
    With oOperario
        .setNOMBRE = txtdatos(1)
        .setDIRECCION = txtdatos(2)
        If Trim(txtdatos(3)) <> "" Then
            .setCP = CLng(txtdatos(3).Text)
        Else
            .setCP = 0
        End If
        ' LP005
'        .setPROVINCIA = txtdatos(4)
        If cmbProvincia.Text = "" Then
            .setPROVINCIA_ID = 0
        Else
            .setPROVINCIA_ID = cmbProvincia.BoundText
        End If
        If cmbMunicipio.Text = "" Then
            .setMUNICIPIO_ID = 0
        Else
            .setMUNICIPIO_ID = cmbMunicipio.BoundText
        End If
        .setCIF = txtdatos(8)
        .setTELEFONO = txtdatos(6)
        .setMOVIL = txtdatos(7)
        .setOBSERVACIONES = txtdatos(9)
        .setfoto = txtdatos(10)
        .setAPODO = txtdatos(11)
        .setESTADO_ID = cmbestados.BoundText
        .setCODIGO_INTERNO = txtdatos(12)
        .setCCC = txtdatos(13)
    End With
    Set mover_datos = oOperario
    Set oOperario = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del empleado.", vbCritical, Err.Description
End Function
Public Sub cargar_estados()
    Dim ooe As New clsEmpleados_Estados
    Set cmbestados.RowSource = ooe.Listado
    cmbestados.ListField = "nombre"
    cmbestados.BoundColumn = "id_estado"
    Set ooe = Nothing
End Sub

Public Sub permisos()
'    If usuario.getPER_EMPLEADOS = 0 Then
'        cmdControl.Enabled = False
'        cmdExpediente.Enabled = False
'        cmdAnticipo.Enabled = False
'    End If
End Sub
Public Sub cargar_municipios(PROVINCIA As Long)
     If IsNumeric(PROVINCIA) Then
        Dim omuni As New clsMunicipios
        Set cmbMunicipio.RowSource = omuni.Listado(PROVINCIA)
        cmbMunicipio.ListField = "nombre" 'campo que veo
        cmbMunicipio.DataField = "nombre" 'campo asociado
        cmbMunicipio.BoundColumn = "id_municipio" 'lo que realmente envia
        Set omuni = Nothing
     End If
End Sub
