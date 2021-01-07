VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmProveedores 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Proveedores"
   ClientHeight    =   9630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmProveedores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos del Proveedor"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4050
      Left            =   30
      TabIndex        =   22
      Top             =   420
      Width           =   9195
      Begin VB.CommandButton cmdaddprovincia 
         Caption         =   "+"
         Height          =   345
         Left            =   8730
         TabIndex        =   46
         Top             =   1710
         Width           =   315
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
         Index           =   13
         Left            =   1080
         MaxLength       =   75
         TabIndex        =   1
         Top             =   705
         Width           =   7980
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
         Index           =   4
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   10
         Top             =   3150
         Width           =   7665
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
         Index           =   10
         Left            =   4140
         MaxLength       =   30
         TabIndex        =   7
         Top             =   2340
         Width           =   1695
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
         Index           =   0
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   11
         Top             =   3600
         Width           =   7665
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
         Left            =   6720
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2340
         Width           =   1785
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
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   6
         Top             =   2340
         Width           =   1665
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Height          =   345
         Index           =   8
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   9
         Top             =   2730
         Width           =   1665
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
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   3
         Top             =   1500
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
         Left            =   1080
         MaxLength       =   75
         TabIndex        =   2
         Top             =   1095
         Width           =   7980
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
         Left            =   1080
         MaxLength       =   75
         TabIndex        =   0
         Top             =   330
         Width           =   7980
      End
      Begin MSDataListLib.DataCombo cmbProvincia 
         Height          =   360
         Left            =   3510
         TabIndex        =   4
         Top             =   1500
         Width           =   5115
         _ExtentX        =   9022
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
         Height          =   360
         Left            =   1080
         TabIndex        =   5
         Top             =   1935
         Width           =   7545
         _ExtentX        =   13309
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Actividad"
         Height          =   195
         Index           =   14
         Left            =   90
         TabIndex        =   45
         Top             =   720
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "e-Mail"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   39
         Top             =   3195
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Móvil"
         Height          =   195
         Index           =   6
         Left            =   3540
         TabIndex        =   38
         Top             =   2400
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   33
         Top             =   3645
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   30
         Top             =   1980
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fax"
         Height          =   195
         Index           =   15
         Left            =   6270
         TabIndex        =   29
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   28
         Top             =   2400
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.I.F."
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   27
         Top             =   2820
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
         Height          =   195
         Index           =   4
         Left            =   2580
         TabIndex        =   26
         Top             =   1560
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   25
         Top             =   1530
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   24
         Top             =   1140
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   23
         Top             =   375
         Width           =   555
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Envío de Notificaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1545
      Left            =   45
      TabIndex        =   40
      Top             =   4485
      Width           =   9165
      Begin VB.CommandButton cmdaddprovinciaN 
         Caption         =   "+"
         Height          =   345
         Left            =   8760
         TabIndex        =   47
         Top             =   810
         Width           =   315
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
         Index           =   12
         Left            =   1395
         MaxLength       =   75
         TabIndex        =   12
         Top             =   225
         Width           =   7710
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
         Index           =   5
         Left            =   1395
         MaxLength       =   5
         TabIndex        =   13
         Top             =   630
         Width           =   960
      End
      Begin MSDataListLib.DataCombo cmbProvinciaN 
         Height          =   360
         Left            =   3555
         TabIndex        =   14
         Top             =   630
         Width           =   5115
         _ExtentX        =   9022
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
      Begin MSDataListLib.DataCombo cmbMunicipioN 
         Height          =   360
         Left            =   1395
         TabIndex        =   15
         Top             =   1065
         Width           =   7275
         _ExtentX        =   12832
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   44
         Top             =   270
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   43
         Top             =   660
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
         Height          =   195
         Index           =   10
         Left            =   2625
         TabIndex        =   42
         Top             =   690
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   41
         Top             =   1110
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8670
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   915
      Left            =   6660
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8670
      Width           =   1275
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Condiciones Especiales "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1140
      Left            =   45
      TabIndex        =   34
      Top             =   6060
      Width           =   9165
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
         Height          =   315
         Index           =   11
         Left            =   1350
         TabIndex        =   17
         Top             =   690
         Width           =   3375
      End
      Begin MSMask.MaskEdBox txtcuenta 
         Height          =   315
         Left            =   5760
         TabIndex        =   18
         Top             =   705
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   23
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####-####-##-##########"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo cmbFP 
         Height          =   315
         Left            =   1350
         TabIndex        =   16
         Top             =   270
         Width           =   7680
         _ExtentX        =   13547
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
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Cuenta"
         Height          =   195
         Left            =   4905
         TabIndex        =   37
         Top             =   750
         Width           =   735
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Banco"
         Height          =   195
         Left            =   150
         TabIndex        =   36
         Top             =   690
         Width           =   465
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   150
         TabIndex        =   35
         Top             =   330
         Width           =   1080
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Index           =   13
      Left            =   60
      TabIndex        =   31
      Top             =   7230
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   1005
         Index           =   9
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   240
         Width           =   8985
      End
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nuevo Proveedor"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   60
      TabIndex        =   32
      Top             =   30
      Width           =   9180
   End
End
Attribute VB_Name = "frmProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pk As Long
Private Sub cmbProvincia_Change()
    If cmbProvincia.Text <> "" Then
        cargar_municipios (cmbProvincia.BoundText)
    End If
End Sub

Private Sub cmdaddprovincia_Click()
    frmProvincias.Show 1
    Dim aux As Long
    aux = 0
    If cmbProvincia.Text <> "" Then
        aux = cmbProvincia.BoundText
    End If
    Cargar_Combo cmbProvincia, New clsProvincias
    cmbProvincia.BoundText = aux
    cmbProvincia_Change
End Sub

Private Sub cmdaddprovinciaN_Click()
    frmProvincias.Show 1
    Cargar_Combo cmbProvinciaN, New clsProvincias
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    If pk <> 0 Then
        modificar_proveedor
    Else
        insertar_proveedor
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    If pk <> 0 Then
        consulta_proveedor
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmProveedores = Nothing
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_Keyup(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40 ' Abajo
       If Index = 9 Then
        txtDatos(1).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 1 Then
        txtDatos(9).SetFocus
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

Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = &HFFFFFF
End Sub

Public Sub borrar_campos()
    Dim i As Integer
    For i = 0 To 11
        If i <> 4 And i <> 5 Then
            On Error Resume Next
            txtDatos(i) = ""
        End If
    Next
    txtDatos(1).SetFocus
End Sub

Public Sub bloquear_campos()
    Dim i As Integer
    For i = 0 To 11
        If i <> 4 And i <> 5 Then
            On Error Resume Next
            txtDatos(i).Locked = True
        End If
    Next
End Sub

Public Sub insertar_proveedor()
    If valida_datos = False Then
        Exit Sub
    End If
    pregunta = "Va a dar de alta el proveedor. ¿Esta seguro?"
    If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Set oProveedor = mover_datos
        If oProveedor.Insertar > 0 Then
            MsgBox "El proveedor se ha insertado correctamente.", vbInformation, App.Title
        End If
        Set oProveedor = Nothing
        Unload Me
    End If
End Sub

Public Sub modificar_proveedor()
    If valida_datos() = False Then
        Exit Sub
    End If
    Dim pos As Integer
    Dim proveedor As Integer
    pregunta = "Va a modificar los datos del proveedor. ¿Esta seguro?"
    If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Set oProveedor = mover_datos
        oProveedor.setID_PROVEEDOR = pk
        If oProveedor.Modificar = True Then
            MsgBox "El proveedor se ha modificado correctamente.", vbInformation, App.Title
            Unload Me
        End If
        Set oProveedor = Nothing
    End If
End Sub
Public Function valida_datos() As Boolean
    valida_datos = True
    If txtDatos(1) = "" Then
        MsgBox "El nombre del proveedor no puede estar en blanco.", vbCritical, "Error"
        txtDatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtDatos(8) = "" Then
        MsgBox "El CIF no puede estar en blanco.", vbCritical, "Error"
        txtDatos(8).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtDatos(3) <> "" Then
        If Not IsNumeric(txtDatos(3)) Then
            MsgBox "El CP debe ser numérico.", vbCritical, "Error"
            txtDatos(3).SetFocus
            valida_datos = False
            Exit Function
        End If
    End If
    If txtDatos(5) <> "" Then
        If Not IsNumeric(txtDatos(5)) Then
            MsgBox "El CP de Notificaciones debe ser numérico.", vbCritical, "Error"
            txtDatos(5).SetFocus
            valida_datos = False
            Exit Function
        End If
    End If
    ' LP005
    If cmbProvincia.Text = "" Then
        MsgBox "Seleccione una provincia.", vbCritical, App.Title
        cmbProvincia.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbMunicipio.Text = "" Then
        MsgBox "Seleccione un municipio.", vbInformation, App.Title
        cmbMunicipio.SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Public Sub consulta_proveedor()
    On Error GoTo fallo
    Dim oProveedor As New clsProveedor
    lbltitulo.Caption = "Modificacion de proveedor"
    oProveedor.Carga (pk)
    With oProveedor
        txtDatos(1) = .getNOMBRE
        txtDatos(13) = .getACTIVIDAD
        txtDatos(2) = .getDIRECCION
        txtDatos(3) = .getCP
        cmbProvincia.BoundText = .getPROVINCIA_ID
        cargar_municipios (.getPROVINCIA_ID)
        cmbMunicipio.BoundText = .getMUNICIPIO_ID
        
        txtDatos(12) = .getDIRECCIONN
        txtDatos(5) = .getCPN
        cmbProvinciaN.BoundText = .getPROVINCIAN_ID
        cargar_municipiosN (.getPROVINCIAN_ID)
        cmbMunicipioN.BoundText = .getMUNICIPION_ID
        
        txtDatos(6) = .getTELEFONO
        txtDatos(4) = .getEMAIL
        txtDatos(7) = .getFAX
        txtDatos(8) = .getCIF
        txtDatos(9) = .getOBSERVACIONES
        txtDatos(0) = .getRESPONSABLE
        txtDatos(10) = .getMOVIL
        txtDatos(11) = .getBANCO
        cmbfp.BoundText = .getFORMA_PAGO
        txtcuenta = .getCCC
    End With
    Set oProveedor = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del proveedor.", vbCritical, Err.Description
End Sub

Public Function mover_datos() As clsProveedor
    On Error GoTo fallo
    Dim oProveedor As New clsProveedor
    With oProveedor
        .setNOMBRE = txtDatos(1)
        .setACTIVIDAD = txtDatos(13)
        .setDIRECCION = txtDatos(2)
        If Trim(txtDatos(3)) <> "" Then
            .setCP = CLng(txtDatos(3).Text)
        Else
            .setCP = 0
        End If
        .setPROVINCIA_ID = cmbProvincia.BoundText
        .setMUNICIPIO_ID = cmbMunicipio.BoundText
        
        .setDIRECCIONN = txtDatos(12)
        If Trim(txtDatos(5)) <> "" Then
            .setCPN = CLng(txtDatos(5).Text)
        Else
            .setCPN = 0
        End If
        If cmbProvinciaN.Text = "" Then
            .setPROVINCIAN_ID = 0
        Else
            .setPROVINCIAN_ID = cmbProvinciaN.BoundText
        End If
        If cmbMunicipioN.Text = "" Then
            .setMUNICIPION_ID = 0
        Else
            .setMUNICIPION_ID = cmbMunicipioN.BoundText
        End If
        .setCIF = txtDatos(8)
        .setTELEFONO = txtDatos(6)
        .setMOVIL = txtDatos(10)
        .setFAX = txtDatos(7)
        .setEMAIL = txtDatos(4)
        .setOBSERVACIONES = txtDatos(9)
        .setRESPONSABLE = txtDatos(0)
        If cmbfp.BoundText = "" Then
            .setFORMA_PAGO = 0
        Else
            .setFORMA_PAGO = cmbfp.BoundText
        End If
        .setBANCO = txtDatos(11)
        .setCCC = txtcuenta
    End With
    Set mover_datos = oProveedor
    Set oProveedor = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del proveedor.", vbCritical, Err.Description
End Function

Public Sub cargar_combos()
    Cargar_Combo cmbfp, New clsForma_pago
    Cargar_Combo cmbProvincia, New clsProvincias
    Cargar_Combo cmbProvinciaN, New clsProvincias
End Sub
Public Sub cargar_municipios(PROVINCIA As Long)
    cmbMunicipio.Text = ""
    cargar_combo_FK cmbMunicipio, New clsMunicipios, PROVINCIA
End Sub
Public Sub cargar_municipiosN(PROVINCIA As Long)
    cmbMunicipioN.Text = ""
    cargar_combo_FK cmbMunicipioN, New clsMunicipios, PROVINCIA
End Sub

