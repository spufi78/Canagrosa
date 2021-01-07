VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmClientes 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Clientes"
   ClientHeight    =   7980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9990
   Icon            =   "frmClientes.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   9990
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDirecciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Direcciones"
      Enabled         =   0   'False
      Height          =   915
      Left            =   2700
      Picture         =   "frmClientes.frx":1CFA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7035
      Visible         =   0   'False
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
      Height          =   1515
      Left            =   60
      TabIndex        =   35
      Top             =   4260
      Width           =   9885
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
         ForeColor       =   &H000000FF&
         Height          =   345
         Index           =   11
         Left            =   1350
         TabIndex        =   13
         Top             =   630
         Width           =   1035
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   1  'Right Justify
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
         Index           =   18
         Left            =   8460
         TabIndex        =   15
         Top             =   630
         Width           =   1305
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   1  'Right Justify
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
         Index           =   17
         Left            =   8460
         TabIndex        =   16
         Top             =   1035
         Width           =   1305
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
         ForeColor       =   &H000000FF&
         Height          =   345
         Index           =   5
         Left            =   4635
         TabIndex        =   12
         Top             =   1035
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
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
         Left            =   1350
         TabIndex        =   14
         Top             =   1035
         Width           =   1710
      End
      Begin MSDataListLib.DataCombo cmbFP 
         Height          =   315
         Left            =   1350
         TabIndex        =   11
         Top             =   270
         Width           =   8385
         _ExtentX        =   14790
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
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Copias Factura"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   43
         Top             =   675
         Width           =   1065
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Riesgo"
         Height          =   195
         Index           =   6
         Left            =   7320
         TabIndex        =   42
         Top             =   690
         Width           =   495
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Riesgo Real"
         Height          =   195
         Index           =   5
         Left            =   7320
         TabIndex        =   41
         Top             =   1095
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "I.V.A."
         Height          =   195
         Index           =   4
         Left            =   3465
         TabIndex        =   39
         Top             =   1125
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.contable"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   38
         Top             =   1095
         Width           =   765
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Forma de Pago"
         Height          =   195
         Left            =   180
         TabIndex        =   36
         Top             =   330
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   915
      Left            =   7350
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   7035
      Width           =   1275
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   8670
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   7035
      Width           =   1275
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
      Height          =   1185
      Index           =   13
      Left            =   60
      TabIndex        =   31
      Top             =   5820
      Width           =   9885
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   840
         Index           =   9
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   240
         Width           =   9615
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos del Cliente "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   60
      TabIndex        =   21
      Top             =   375
      Width           =   9885
      Begin VB.CommandButton cmdSiguiente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Siguiente Código Libre"
         Height          =   375
         Left            =   3870
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   225
         Width           =   1905
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   12
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   22
         Top             =   240
         Width           =   2430
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
         Index           =   16
         Left            =   1350
         MaxLength       =   50
         TabIndex        =   10
         Top             =   3450
         Width           =   8385
      End
      Begin VB.CommandButton cmdaddprovincia 
         Caption         =   "+"
         Height          =   345
         Left            =   9420
         TabIndex        =   37
         Top             =   2070
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
         Index           =   0
         Left            =   8055
         MaxLength       =   30
         TabIndex        =   8
         Top             =   2685
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
         Height          =   330
         Index           =   7
         Left            =   6045
         MaxLength       =   30
         TabIndex        =   7
         Top             =   2670
         Width           =   1440
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
         Height          =   345
         Index           =   10
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   1
         Top             =   1020
         Width           =   8430
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
         MaxLength       =   50
         TabIndex        =   6
         Top             =   2670
         Width           =   3915
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00FF0000&
         Height          =   345
         Index           =   8
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   9
         Top             =   3060
         Width           =   1755
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
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
         TabIndex        =   3
         Top             =   1830
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
         TabIndex        =   2
         Top             =   1425
         Width           =   8430
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
         Top             =   630
         Width           =   8430
      End
      Begin MSDataListLib.DataCombo cmbProvincia 
         Height          =   315
         Left            =   3690
         TabIndex        =   4
         Top             =   1830
         Width           =   5625
         _ExtentX        =   9922
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
      Begin MSDataListLib.DataCombo cmbMunicipio 
         Height          =   315
         Left            =   1350
         TabIndex        =   5
         Top             =   2250
         Width           =   7965
         _ExtentX        =   14049
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
         Caption         =   "Código"
         Height          =   195
         Index           =   6
         Left            =   240
         TabIndex        =   44
         Top             =   285
         Width           =   495
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "e-Mail"
         Height          =   195
         Index           =   9
         Left            =   225
         TabIndex        =   40
         Top             =   3510
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fax"
         Height          =   195
         Index           =   8
         Left            =   7695
         TabIndex        =   34
         Top             =   2745
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "P. Contacto"
         Height          =   195
         Index           =   3
         Left            =   225
         TabIndex        =   33
         Top             =   1065
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
         Height          =   195
         Index           =   7
         Left            =   225
         TabIndex        =   30
         Top             =   2310
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Movil"
         Height          =   195
         Index           =   15
         Left            =   5595
         TabIndex        =   29
         Top             =   2730
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   195
         Index           =   11
         Left            =   225
         TabIndex        =   28
         Top             =   2730
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.I.F."
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   27
         Top             =   3150
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
         Height          =   195
         Index           =   4
         Left            =   2580
         TabIndex        =   26
         Top             =   1890
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   25
         Top             =   1860
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   24
         Top             =   1470
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   23
         Top             =   675
         Width           =   555
      End
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Nuevo Cliente"
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
      Height          =   315
      Left            =   60
      TabIndex        =   32
      Top             =   30
      Width           =   9870
   End
End
Attribute VB_Name = "frmClientes"
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

Private Sub cmdcancel_Click()
    Unload Me
End Sub
'Private Sub cmdDirecciones_Click()
'    If valida_datos() = False Then
'        Exit Sub
'    End If
'    Set ocliente = mover_datos
'    ocliente.setID_CLIENTE = pk
'    If ocliente.modificar_cliente = True Then
'        Dim oagenda As New clsAgenda
'        With oagenda
'                .setNOMBRE = UCase(txtdatos(1))
'                .setTELEFONO = txtdatos(6)
'                .setFAX = txtdatos(0)
'                .setMOVIL = txtdatos(7)
'                .modificar_cliente
'        End With
'        Set ocliente = Nothing
'    End If
'    frmClientes_Direcciones.Show 1
'    consulta_Cliente
'End Sub

'Private Sub cmdFP_Click()
'    frmClientes_FP.Show 1
'End Sub

Private Sub cmdok_Click()
    If pk > 0 Then
        modificar_cliente
    Else
        insertar_cliente
    End If
End Sub


Private Sub cmdSiguiente_Click()
    Dim ocliente As New clsCliente
    txtdatos(12) = ocliente.BuscarId_ClienteLibre(txtdatos(12))
    Set ocliente = Nothing
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    If pk > 0 Then
        txtdatos(12).Enabled = False
        cmdSiguiente.Enabled = False
        cmdDirecciones.Enabled = True
        consulta_Cliente
    Else
        ' IVA
        txtdatos(5) = ReadINI(App.Path & "\config.ini", "parametros", "iva")
        ' Copias factura
        txtdatos(11) = ReadINI(App.Path & "\config.ini", "parametros", "Copias_facturas")
        txtdatos(4) = "7000001" ' Cuenta Cliente
        ' ID_CLIENTE
        txtdatos(12).Enabled = True
        Dim ocliente As New clsCliente
'        ocliente.CrearId_Cliente
'        txtDatos(12) = ocliente.getID_CLIENTE
        txtdatos(12) = ocliente.BuscarId_ClienteLibre
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmClientes = Nothing
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub

Private Sub txtdatos_Keyup(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40 ' Abajo
       If Index = 12 Then
        txtdatos(1).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 1 Then
        txtdatos(12).SetFocus
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
    If KeyAscii = 13 And Index <> 9 And Index <> 12 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = &HFFFFFF
    If Index = 17 Or Index = 18 Then
        txtdatos(Index) = moneda(txtdatos(Index))
    End If
    If Index = 11 Then ' cOPIAS FACTURA
        If txtdatos(Index) = "" Then
            txtdatos(Index) = ReadINI(App.Path & "\config.ini", "parametros", "Copias_facturas")
        Else
            If Not IsNumeric(txtdatos(11)) Then
                txtdatos(Index) = ReadINI(App.Path & "\config.ini", "parametros", "Copias_facturas")
            End If
        End If
    End If
    If Index = 12 Then ' ID
        Dim ocliente As New clsCliente
        If txtdatos(Index) = "" Then
            txtdatos(Index) = ocliente.BuscarId_ClienteLibre
'            ocliente.CrearId_Cliente
'            txtDatos(Index) = ocliente.getID_CLIENTE
        Else
            If Not IsNumeric(txtdatos(Index)) Then
                MsgBox "EL código de cliente debe ser numérico.", vbCritical, App.Title
                ocliente.CrearId_Cliente
                txtdatos(Index) = ocliente.getID_CLIENTE
                Exit Sub
            Else
                If ocliente.Verificar_Codigo(txtdatos(Index)) Then
                    MsgBox "El código del cliente ya existe. No se puede utilizar.", vbExclamation, App.Title
                    txtdatos(Index) = ocliente.BuscarId_ClienteLibre
'                    txtdatos(Index) = ocliente.getID_CLIENTE
                End If
            End If
        End If
    End If
End Sub

Public Sub borrar_campos()
    Dim i As Integer
    For i = 0 To 15
        If i <> 5 Then
            On Error Resume Next
            txtdatos(i) = ""
        End If
    Next
    txtdatos(1).SetFocus
End Sub

Public Sub bloquear_campos()
    Dim i As Integer
    For i = 1 To 15
        If i <> 4 And i <> 5 Then
            On Error Resume Next
            txtdatos(i).Locked = True
        End If
    Next
End Sub

Public Sub insertar_cliente()
    If valida_datos = False Then
        Exit Sub
    End If
    pregunta = "Va a dar de alta el Cliente. ¿Esta seguro?"
    If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim aux As Long
        Set ocliente = mover_datos
        aux = pk
        If ocliente.Verificar_Codigo(txtdatos(12)) Then
            MsgBox "El código del cliente ha sido utilizado por otro usuario.", vbExclamation, App.Title
            txtdatos(12).SetFocus
            Exit Sub
        End If
        ocliente.setID_CLIENTE = txtdatos(12)
        pk = ocliente.insertar_cliente
        If pk > 0 Then
            MsgBox "El cliente se ha insertado correctamente.", vbInformation, App.Title
        End If
'        If aux = -1 Then
            Unload Me
'            Exit Sub
'        End If
'        borrar_campos
'        Set oCliente = Nothing
    End If
End Sub

Public Sub modificar_cliente()
    If valida_datos() = False Then
        Exit Sub
    End If
    Dim pos As Integer
    Dim cliente As Integer
    pregunta = "Va a modificar los datos del Cliente. ¿Esta seguro?"
    If MsgBox(pregunta, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Set ocliente = mover_datos
        ocliente.setID_CLIENTE = pk
        If ocliente.modificar_cliente = True Then
            Dim oagenda As New clsAgenda
            With oagenda
                .setNOMBRE = UCase(txtdatos(1))
                .setTelefono = txtdatos(6)
                .setFAX = txtdatos(0)
                .setMOVIL = txtdatos(7)
                .modificar_cliente pk
            End With
        
            MsgBox "El cliente se ha modificado correctamente.", vbInformation, App.Title
            Unload Me
        End If
        Set ocliente = Nothing
    End If

End Sub
Public Function valida_datos() As Boolean
    valida_datos = True
    If txtdatos(1) = "" Then
        MsgBox "El nombre del cliente no puede estar en blanco.", vbCritical, "Error"
        txtdatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
    
    If frmMenu.StatusBar1.Panels(3) <> "Server: " & IP_RESPALDO Then
        If txtdatos(8) = "" Then
            MsgBox "El NIF no puede estar en blanco.", vbCritical, "Error"
            txtdatos(8).SetFocus
            valida_datos = False
            Exit Function
        End If
        If txtdatos(3) <> "" Then
            If IsNumeric(txtdatos(3)) = False Then
                MsgBox "El CP debe ser numérico.", vbCritical, "Error"
                txtdatos(3).SetFocus
                valida_datos = False
                Exit Function
            End If
        End If
        If txtdatos(5) <> "" Then
            If IsNumeric(txtdatos(5)) = False Then
                MsgBox "El IVA debe ser numérico.", vbCritical, "Error"
                txtdatos(5).SetFocus
                valida_datos = False
                Exit Function
            End If
        Else
            MsgBox "El IVA debe ser numérico.", vbCritical, "Error"
            txtdatos(5).SetFocus
            valida_datos = False
            Exit Function
        End If
        If cmbProvincia.Text = "" Then
            MsgBox "Seleccione una provincia.", vbInformation, App.Title
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
    End If
    
    If txtdatos(11) <> "" Then
        If IsNumeric(txtdatos(11)) = False Then
            MsgBox "El número de copias debe ser numérico.", vbCritical, "Error"
            txtdatos(11).SetFocus
            valida_datos = False
            Exit Function
        End If
    Else
        MsgBox "El número de copias debe ser numérico.", vbCritical, "Error"
        txtdatos(11).SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbFP.Text = "" Then
        MsgBox "Seleccione la Forma de PAGO.", vbInformation, App.Title
        cmbFP.SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Public Sub consulta_Cliente()
    On Error GoTo fallo
    Dim ocliente As New clsCliente
    lbltitulo.Caption = "Modificacion de Cliente"
    ocliente.CargaCliente (pk)
    With ocliente
        txtdatos(12) = .getID_CLIENTE
        txtdatos(1) = .getNOMBRE
        txtdatos(2) = .getDIRECCION
        txtdatos(3) = .getCP
        txtdatos(5) = .getIVA
        txtdatos(11) = .getCOPIAS_FACTURA
        ' LP005
'        txtdatos(4) = .getPROVINCIA
        cmbProvincia.BoundText = .getPROVINCIA_ID
'        txtdatos(5) = .getMUNICIPIO
        cargar_municipios (.getPROVINCIA_ID)
        cmbMunicipio.BoundText = .getMUNICIPIO_ID
        txtdatos(6) = .getTelefono
        txtdatos(7) = .getMOVIL
        txtdatos(8) = .getCIF
        txtdatos(9) = .getOBSERVACIONES
        txtdatos(10) = .getRAZON
        txtdatos(0) = .getFAX
        txtdatos(16) = .getEMAIL
        cmbFP.BoundText = .getFORMA_PAGO
        txtdatos(4) = .getCCONTABLE
        txtdatos(18) = moneda(.getRIESGO)
        txtdatos(17) = moneda(.getRIESGO_REAL)
    End With
    Set ocliente = Nothing
    Exit Sub
fallo:
    log ("Error al consultar los datos del cliente : " & Err.Description)
    MsgBox "Error al consultar los datos del cliente.", vbCritical, Err.Description
End Sub
Public Function mover_datos() As clsCliente
    On Error GoTo fallo
    Dim ocliente As New clsCliente
    With ocliente
        .setNOMBRE = txtdatos(1)
        .setDIRECCION = txtdatos(2)
        If txtdatos(3) <> "" Then
            .setCP = txtdatos(3)
        Else
            .setCP = 0
        End If
        ' LP005
'        .setPROVINCIA = txtdatos(4)
        .setPROVINCIA_ID = cmbProvincia.BoundText
'        .setMUNICIPIO = txtdatos(5)
        .setMUNICIPIO_ID = cmbMunicipio.BoundText
        .setCIF = txtdatos(8)
        .setTelefono = txtdatos(6)
        .setMOVIL = txtdatos(7)
        .setFAX = txtdatos(0)
        .setOBSERVACIONES = txtdatos(9)
        .setRAZON = txtdatos(10)
        .setEMAIL = txtdatos(16)
        .setIVA = txtdatos(5)
        .setCOPIAS_FACTURA = txtdatos(11)
        If cmbFP.BoundText = "" Then
            .setFORMA_PAGO = 0
        Else
            .setFORMA_PAGO = cmbFP.BoundText
        End If
'        .setCOMERCIAL_ID = 0
        .setCCONTABLE = txtdatos(4)
        If txtdatos(18) <> "" Then
            .setRIESGO = moneda_bd(txtdatos(18))
        Else
            .setRIESGO = moneda_bd("0")
        End If
        If txtdatos(17) <> "" Then
            .setRIESGO_REAL = moneda_bd(txtdatos(17))
        Else
            .setRIESGO_REAL = moneda_bd("0")
        End If
    End With
    Set mover_datos = ocliente
    Set ocliente = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del cliente.", vbCritical, Err.Description
End Function

Public Sub cargar_combos()
    Cargar_Combo cmbFP, New clsForma_pago
    Cargar_Combo cmbProvincia, New clsProvincias
End Sub
Public Sub cargar_municipios(PROVINCIA As Long)
    cmbMunicipio.Text = ""
    cargar_combo_FK cmbMunicipio, New clsMunicipios, PROVINCIA
End Sub
