VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.ocx"
Begin VB.Form frmLogin 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E.R.P. Geslab (Login)"
   ClientHeight    =   4095
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   4275
   Icon            =   "frmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2419.462
   ScaleMode       =   0  'User
   ScaleWidth      =   4013.992
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3780
      Top             =   2745
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   4050
      Left            =   45
      TabIndex        =   5
      Top             =   15
      Width           =   4200
      Begin VB.CheckBox chkPrueba 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Acceder en Modo Pruebas"
         Height          =   195
         Left            =   810
         TabIndex        =   13
         Top             =   3780
         Width           =   2670
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   1845
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   2340
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ESC-Salir"
         Height          =   870
         Left            =   2250
         Picture         =   "frmLogin.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2790
         Width           =   1410
      End
      Begin VB.CommandButton cmdok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Height          =   870
         Left            =   540
         Picture         =   "frmLogin.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2790
         Width           =   1410
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1845
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1935
         Width           =   1965
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1845
         TabIndex        =   0
         Top             =   1530
         Width           =   1965
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
         TabIndex        =   12
         Top             =   3750
         Width           =   3855
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nueva"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   450
         TabIndex        =   8
         Top             =   2385
         Visible         =   0   'False
         Width           =   1305
      End
      Begin VB.Image imagen 
         Appearance      =   0  'Flat
         Height          =   930
         Left            =   450
         Stretch         =   -1  'True
         Top             =   315
         Width           =   3390
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   450
         TabIndex        =   7
         Top             =   1980
         Width           =   1305
      End
      Begin VB.Label lblLabels 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   450
         TabIndex        =   6
         Top             =   1590
         Width           =   1260
      End
   End
   Begin MSWinsockLib.Winsock wsck 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   30
      TabIndex        =   11
      Top             =   4530
      Width           =   4185
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Se actualizan correctamente..."
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
      Height          =   240
      Left            =   45
      TabIndex        =   10
      Top             =   4305
      Width           =   4155
   End
   Begin VB.Label Label1 
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   45
      TabIndex        =   9
      Top             =   4080
      Width           =   4155
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdcancel_Click()
'Q    If glogin = 2 Then ' Inactividad
'Q        End
'Q    Else
        Unload Me
'Q    End If
End Sub
Private Sub cmdok_Click() 'comprobar si la contraseña es correcta
    MODO_PRUEBA = chkPrueba.Value
    Me.MousePointer = 11
'Q    Dim oParametro As New clsParametros
'Q    If glogin = 2 Then
'Q        If USUARIO.Autentificacion(txtDatos(0), txtDatos(1)) Then
'Q            Unload Me
'Q            frmMenu.Show
'Q            oParametro.Carga parametros.TIEMPO_INACTIVIDAD, ""
'Q            Call frmMenu.Inactividad(oParametro.getVALOR)
'Q        Else
'Q            Me.MousePointer = 0
'Q            MsgBox "La contraseña o el usuario no es válido. Vuelva a intentarlo", vbOKOnly + vbInformation, "Inicio de sesión"
'Q            txtDatos(1) = ""
'Q            txtDatos(1).SetFocus
'Q        End If
'Q        Exit Sub
'Q    End If
    Set USUARIO = New clsUsuarios
'JGM    If CrearConexionGlobal(txtDatos(0)) = True Then
'        verificar_fecha_sistema
'        registrar_componentes_arranque (Me.hWnd)
        Me.MousePointer = 0
        Dim CODIGO As Integer
        If USUARIO.Autentificacion(txtDatos(0), txtDatos(1)) Then
            If txtDatos(2) <> "" Then
                If USUARIO.modificar_password(USUARIO.getID_EMPLEADO, Encripta(txtDatos(2), txtDatos(0))) Then
                    MsgBox "Se ha modificado el password correctamente.", vbOKOnly + vbInformation, App.Title
                End If
            Else
                ' Insertar el login
                Dim cadena As String
                Dim consulta As String
                NOMBRE_PC = Winsock1.LocalHostName
                cadena = NOMBRE_PC
                If cadena = "" Then
                 cadena = "No identificado"
                End If
                consulta = "UPDATE usuarios set USO = '" & UCase(cadena) & "' where id_empleado=" & USUARIO.getID_EMPLEADO ' rs(0)
                USUARIO.setUSO = UCase(cadena)
                execute_bd consulta
                
                ' Ejecutar la consulta de max_concat
                execute_bd "SET GLOBAL group_concat_max_len = 1000000"
            End If
            'USUARIO.CARGAR 5
'            On Error Resume Next
    
'            Call RegisterServer(Me.Hwnd, ReadINI(App.Path + "\config.ini", "otros", "ocx") & "\Codejock.CommandBars.v13.2.1.ocx", True)
'            FileCopy ReadINI(App.Path + "\config.ini", "version", "ruta") & "\version.exe", App.Path & "\version.exe"
            Unload Me
'Q            frmMenu.bCancel = False
            glogin = 1
            frmMenu.Show
'            oParametro.Carga parametros.TIEMPO_INACTIVIDAD, ""
'            Call frmMenu.Inactividad(oParametro.getVALOR)
        Else
            Me.MousePointer = 0
            MsgBox "La contraseña o el usuario no es válido. Vuelva a intentarlo", vbOKOnly + vbInformation, "Inicio de sesión"
            txtDatos(1) = ""
            txtDatos(1).SetFocus
'JGM            conn.Close
        End If
'JGM    Else
'JGM        Me.MousePointer = 0
'JGM        MsgBox "No se pudo conectar con la base de datos", vbCritical, Err.Description
'JGM    End If
End Sub

Private Sub Form_Activate()
'Q    If glogin = 2 Then
'Q        lblreg.Caption = "Sesión expirada. Reingrese."
'Q    End If
End Sub

Private Sub Form_Initialize()
    On Error Resume Next
    verificar_fecha_sistema
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error
    NOMBRE_PC = Winsock1.LocalHostName
    If pc_es_tablet Then
        frmLoginTablet.Show
        Unload Me
        Exit Sub
    End If
'    If App.PrevInstance = True Then
'        MsgBox "El programa Geslab ya se encuentra en ejecución.", vbInformation, App.Title
'        End
'    End If
'    On Error Resume Next
    If Dir(ReadINI(App.Path + "\config.ini", "logo", "logo")) <> "" Then
        Set imagen.Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
    End If
    If glogin = 1 Then
        txtDatos(0) = USUARIO.getUSUARIO
        lblLabels(2).visible = True
        txtDatos(2).visible = True
        Exit Sub
    End If
'Q    If glogin = 2 Then
'Q        txtDatos(0) = USUARIO.getUSUARIO
'Q        txtDatos(0).Enabled = False
'Q        Exit Sub
'Q    End If
    ' Para Pruebas
    If ReadINI(App.Path + "\config.ini", "Otros", "usuario") <> "" Then
        txtDatos(0) = ReadINI(App.Path + "\config.ini", "Otros", "usuario")
    End If
    If ReadINI(App.Path + "\config.ini", "Otros", "password") <> "" Then
        txtDatos(1) = ReadINI(App.Path + "\config.ini", "Otros", "password")
    End If
    If txtDatos(0) <> "" And txtDatos(1) <> "" Then
'         cmdok_Click
    End If
    ' Quitar la copia del version.exe
'Q    Dim f1 As String
'Q    Dim f2 As String
'Q    f1 = ReadINI(App.Path + "\config.ini", "version", "ruta") & "\version.exe"
'Q    f2 = App.Path & "\version.exe"
'Q    FileCopy f1, f2

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmLogin"
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
    Do While cIT.CodigoError = -1 And intentos < 3
        Espera (1)
        intentos = intentos + 1
        DoEvents
    Loop
    If cIT.CodigoError = 0 Then
'        MsgBox Format(Time, "hh:mm")
'        MsgBox Format(Time, "hh:mm")
        If Format(Date, "dd-mm-yyyy") <> Format(cIT.fecha, "dd-mm-yyyy") Then
'           Format(Time, "hh:mm") <> Format(cIT.HORA, "hh:mm") Then
            Me.Height = 5160
'        Label1 = cIT.TextoError
'        lblfecha = cIT.Fecha
            Label3.Caption = "Nuevos: " & cIT.FechayHora
'        lblhora = cIT.Hora
            cIT.ActualizarFechaSistema
        End If
    End If
    Set cIT = Nothing

   On Error GoTo 0
   Exit Sub

verificar_fecha_sistema_Error:

    log ("Error " & Err.Number & " (" & Err.Description & ") in procedure verificar_fecha_sistema of Formulario frmLogin")
End Sub
