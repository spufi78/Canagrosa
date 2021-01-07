VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmLoginTablet 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Geslab (Login)"
   ClientHeight    =   9360
   ClientLeft      =   2835
   ClientTop       =   3435
   ClientWidth     =   10320
   Icon            =   "frmLoginTablet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5530.2
   ScaleMode       =   0  'User
   ScaleWidth      =   9689.917
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2760
      Left            =   6705
      TabIndex        =   19
      Top             =   1350
      Width           =   3525
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   180
         PasswordChar    =   "*"
         TabIndex        =   23
         Top             =   1845
         Width           =   3210
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Index           =   0
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   765
         Width           =   3210
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   180
         TabIndex        =   21
         Top             =   1395
         Width           =   2985
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   135
         TabIndex        =   20
         Top             =   360
         Width           =   2985
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   1140
      Left            =   6750
      Picture         =   "frmLoginTablet.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8145
      Width           =   1725
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   1140
      Left            =   8550
      Picture         =   "frmLoginTablet.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8145
      Width           =   1635
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Height          =   3930
      Left            =   6705
      TabIndex        =   3
      Top             =   4140
      Width           =   3525
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
         Height          =   900
         Index           =   0
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2925
         Width           =   2160
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
         Height          =   900
         Index           =   9
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   225
         Width           =   1080
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
         Height          =   900
         Index           =   8
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   225
         Width           =   1080
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
         Height          =   900
         Index           =   7
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   225
         Width           =   1080
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
         Height          =   900
         Index           =   6
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1125
         Width           =   1080
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
         Height          =   900
         Index           =   5
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1125
         Width           =   1080
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
         Height          =   900
         Index           =   4
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1125
         Width           =   1080
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
         Height          =   900
         Index           =   3
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2025
         Width           =   1080
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
         Height          =   900
         Index           =   2
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2025
         Width           =   1080
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
         Height          =   900
         Index           =   1
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2025
         Width           =   1080
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
         Height          =   900
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2925
         Width           =   1080
      End
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   7650
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6705
      Left            =   45
      TabIndex        =   0
      Top             =   1350
      Width           =   6630
      Begin MSComctlLib.ListView lista 
         Height          =   6405
         Left            =   135
         TabIndex        =   18
         Top             =   225
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   11298
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
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label lblreg 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Left            =   210
         TabIndex        =   2
         Top             =   3750
         Width           =   3855
      End
   End
   Begin MSWinsockLib.Winsock wsck 
      Left            =   8055
      Top             =   6840
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Image imagen 
      Appearance      =   0  'Flat
      Height          =   1245
      Left            =   90
      Stretch         =   -1  'True
      Top             =   45
      Width           =   3795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "La fecha y hora del sistema no están correctas. "
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
      Left            =   45
      TabIndex        =   15
      Top             =   8910
      Visible         =   0   'False
      Width           =   4125
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
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
      Left            =   45
      TabIndex        =   1
      Top             =   9090
      Visible         =   0   'False
      Width           =   75
   End
End
Attribute VB_Name = "frmLoginTablet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdcancel_Click()
'Q    If glogin = 2 Then ' Inactividad
'Q       End
'Q    Else
        Unload Me
'Q    End If
End Sub

Private Sub cmdDel_Click()
    txtDatos(1) = ""
End Sub

Private Sub cmdNumero_Click(Index As Integer)
    txtDatos(1) = txtDatos(1) & cmdNumero(Index).Caption
End Sub

Private Sub cmdok_Click() 'comprobar si la contraseña es correcta
    Me.MousePointer = 11
    Dim oParametro As New clsParametros
'    If glogin = 2 Then
'        If USUARIO.Autentificacion(txtdatos(0), txtdatos(1)) Then
'            Unload Me
'            frmMenu.Show
'            oParametro.Carga parametros.TIEMPO_INACTIVIDAD, ""
'            Call frmMenu.Inactividad(oParametro.getVALOR)
'        Else
'            Me.MousePointer = 0
'            MsgBox "La contraseña o el usuario no es válido. Vuelva a intentarlo", vbOKOnly + vbInformation, "Inicio de sesión"
'            txtdatos(1) = ""
'            txtdatos(1).SetFocus
'        End If
'        Exit Sub
'    End If
    Set usuario = New clsUsuarios
    verificar_fecha_sistema
    Me.MousePointer = 0
    Dim CODIGO As Integer
    If usuario.Autentificacion(txtDatos(0), txtDatos(1)) Then
         ' Insertar el login
         Dim cadena As String
         Dim consulta As String
         NOMBRE_PC = Winsock1.LocalHostName
         cadena = NOMBRE_PC
         If cadena = "" Then
             cadena = "No identificado"
         End If
         consulta = "UPDATE usuarios set USO = '" & UCase(cadena) & "' where id_empleado=" & usuario.getID_EMPLEADO ' rs(0)
         usuario.setUSO = UCase(cadena)
         execute_bd consulta
         'USUARIO.CARGAR 5
         Unload Me
'Q        frmMenu.bCancel = False
         glogin = 1
         frmMenu.Show
'        oParametro.Carga parametros.TIEMPO_INACTIVIDAD, ""
'        Call frmMenu.Inactividad(oParametro.getVALOR)
    Else
         Me.MousePointer = 0
         MsgBox "La contraseña o el usuario no es válido. Vuelva a intentarlo", vbOKOnly + vbInformation, "Inicio de sesión"
         txtDatos(1) = ""
         txtDatos(1).SetFocus
    End If
End Sub

Private Sub Form_Activate()
'Q    If glogin = 2 Then
'Q        lblreg.Caption = "Sesión expirada. Reingrese."
'Q    End If
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
'    verificar_fecha_sistema
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then
'        MsgBox "El programa Geslab ya se encuentra en ejecución.", vbInformation, App.Title
'        End
    End If
    On Error Resume Next
    Set imagen.Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
    cabecera
'    If glogin = 1 Then
'        txtdatos(0) = USUARIO.getUSUARIO
'        lblLabels(2).Visible = True
'        txtdatos(2).Visible = True
'        Exit Sub
'    End If
'    If glogin = 2 Then
'        txtdatos(0) = USUARIO.getUSUARIO
'        txtdatos(0).Enabled = False
'        Exit Sub
'    End If
    ' Para Pruebas
'    If Not CrearConexionGlobal(txtDatos(0)) = True Then
'        MsgBox "No se pudo conectar a la base de datos.", vbCritical, App.Title
'        End
'    End If
    cargar_lista_usuarios
    If ReadINI(App.Path + "\config.ini", "Otros", "usuario") <> "" Then
        txtDatos(0) = ReadINI(App.Path + "\config.ini", "Otros", "usuario")
    End If
    If ReadINI(App.Path + "\config.ini", "Otros", "password") <> "" Then
        txtDatos(1) = ReadINI(App.Path + "\config.ini", "Otros", "password")
    End If
    If txtDatos(0) <> "" And txtDatos(1) <> "" Then
 '        cmdok_Click
    End If
End Sub


Private Sub lista_Click()
    txtDatos(0) = lista.ListItems(lista.selectedItem.Index).SubItems(1)
    txtDatos(1) = ""
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40
        SendKeys "{Tab}", True
     Case 38
        SendKeys "+{Tab}", True
     Case 27
        cmdcancel_Click
    End Select
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = &HFFFFFF
End Sub

Public Sub verificar_fecha_sistema()
    Dim cIT As cInternetTime
   On Error GoTo verificar_fecha_sistema_Error

    Set cIT = New cInternetTime
    DoEvents
    cIT.ObtenerFechayHora wsck
    'esperamos a obtener la respuesta
    Dim intentos As Integer
    intentos = 1
    Do While cIT.CodigoError = -1 And intentos < 2
        Espera (1)
        intentos = intentos + 1
        DoEvents
    Loop
    If cIT.CodigoError = 0 Then
        If Format(Date, "dd-mm-yyyy") <> Format(cIT.fecha, "dd-mm-yyyy") Then
            Label1.Visible = True
            Label3.Visible = True
            Label3.Caption = "Nuevos: " & cIT.FechayHora
            cIT.ActualizarFechaSistema
        End If
    End If
    Set cIT = Nothing

   On Error GoTo 0
   Exit Sub

verificar_fecha_sistema_Error:

    log ("Error " & Err.Number & " (" & Err.Description & ") in procedure verificar_fecha_sistema of Formulario frmLoginTablet")
End Sub
Public Sub cargar_lista_usuarios()
    Dim rs As ADODB.Recordset
    Dim oUsuarios As New clsUsuarios
    Set rs = oUsuarios.Listado_por_Usuario
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs("ID_EMPLEADO"))
                .SubItems(1) = rs("USUARIO")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oUsuarios = Nothing
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Usuario", lista.Width - 250, lvwColumnLeft
    End With
End Sub

