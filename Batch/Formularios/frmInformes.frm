VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInformes 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Generador de Informes E-AVISA"
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9675
   Icon            =   "frmInformes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   9675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   6120
      Top             =   8325
   End
   Begin VB.TextBox txtmysql 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   6
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   29
      Top             =   7650
      Width           =   9420
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Servidor Impresión"
      Height          =   960
      Left            =   90
      TabIndex        =   28
      Top             =   8145
      Width           =   1185
   End
   Begin VB.TextBox txtmysql 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   7
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   6480
      Width           =   9420
   End
   Begin VB.TextBox txtmysql 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   24
      Top             =   5895
      Width           =   9420
   End
   Begin VB.TextBox txtproceso 
      Height          =   1725
      IMEMode         =   3  'DISABLE
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   3915
      Width           =   9555
   End
   Begin VB.CheckBox chkPrueba 
      Caption         =   "Modo Prueba"
      Height          =   240
      Left            =   6615
      TabIndex        =   22
      Top             =   2565
      Width           =   1590
   End
   Begin VB.TextBox txtmysql 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   20
      Top             =   7065
      Width           =   9420
   End
   Begin MSComCtl2.DTPicker hora_ejecucion 
      Height          =   420
      Left            =   8010
      TabIndex        =   16
      Top             =   1215
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16449538
      CurrentDate     =   40856
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1215
      Top             =   8370
   End
   Begin VB.CheckBox chkInicio 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Iniciar la aplicación al ejecutar"
      Height          =   195
      Left            =   6615
      TabIndex        =   15
      Top             =   2295
      Width           =   2940
   End
   Begin VB.CheckBox chklog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generar log"
      Height          =   195
      Left            =   6615
      TabIndex        =   14
      Top             =   2025
      Width           =   2355
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1710
      Top             =   8325
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   1005
      Left            =   8280
      Picture         =   "frmInformes.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8100
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Base de datos MYSQL"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1950
      Left            =   45
      TabIndex        =   1
      Top             =   1935
      Width           =   6360
      Begin VB.CheckBox chkConectado 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Conectado"
         Enabled         =   0   'False
         Height          =   240
         Left            =   4905
         TabIndex        =   11
         Top             =   0
         Width           =   1140
      End
      Begin VB.CommandButton cmdConectar 
         Caption         =   "Conectar"
         Height          =   1005
         Left            =   4815
         Picture         =   "frmInformes.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   495
         Width           =   1320
      End
      Begin VB.TextBox txtmysql 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1170
         PasswordChar    =   "*"
         TabIndex        =   9
         Text            =   "Aer0p0lis2016*"
         Top             =   1440
         Width           =   3390
      End
      Begin VB.TextBox txtmysql 
         Height          =   330
         Index           =   2
         Left            =   1170
         TabIndex        =   8
         Text            =   "geslab"
         Top             =   1080
         Width           =   3390
      End
      Begin VB.TextBox txtmysql 
         Height          =   330
         Index           =   1
         Left            =   1170
         TabIndex        =   7
         Top             =   720
         Width           =   3390
      End
      Begin VB.TextBox txtmysql 
         Height          =   330
         Index           =   0
         Left            =   1170
         TabIndex        =   6
         Top             =   360
         Width           =   3390
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         Height          =   330
         Index           =   3
         Left            =   225
         TabIndex        =   5
         Top             =   1485
         Width           =   870
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario"
         Height          =   330
         Index           =   2
         Left            =   225
         TabIndex        =   4
         Top             =   1125
         Width           =   870
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Base Datos"
         Height          =   330
         Index           =   1
         Left            =   225
         TabIndex        =   3
         Top             =   765
         Width           =   870
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Servidor"
         Height          =   330
         Index           =   0
         Left            =   225
         TabIndex        =   2
         Top             =   405
         Width           =   870
      End
   End
   Begin VB.CommandButton cmdConvertir 
      Caption         =   "Enviar"
      Height          =   1005
      Left            =   6885
      Picture         =   "frmInformes.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8100
      Width           =   1320
   End
   Begin MSComCtl2.DTPicker hora_actual 
      Height          =   420
      Left            =   8010
      TabIndex        =   18
      Top             =   765
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   741
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   16449538
      CurrentDate     =   40856
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Correo Listado de ProCNC más de tres meses abiertas"
      Height          =   330
      Index           =   6
      Left            =   45
      TabIndex        =   30
      Top             =   7425
      Width           =   3795
   End
   Begin VB.Image imagen 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   225
      Picture         =   "frmInformes.frx":2328
      Stretch         =   -1  'True
      Top             =   45
      Width           =   2580
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Correo Destino INFORMES GESLAB"
      Height          =   330
      Index           =   7
      Left            =   90
      TabIndex        =   27
      Top             =   6255
      Width           =   4965
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Correo Destino Informe con TODAS las acciones correctoras"
      Height          =   330
      Index           =   5
      Left            =   90
      TabIndex        =   25
      Top             =   5670
      Width           =   4965
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Correo Log Prueba"
      Height          =   330
      Index           =   4
      Left            =   90
      TabIndex        =   21
      Top             =   6840
      Width           =   1770
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hora Actual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   5220
      TabIndex        =   19
      Top             =   810
      Width           =   2715
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Hora de Ejecución"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   5220
      TabIndex        =   17
      Top             =   1260
      Width           =   2760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Generador de Informes GESLAB v.1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   3195
      TabIndex        =   0
      Top             =   90
      Width           =   5775
   End
   Begin VB.Menu opMenu 
      Caption         =   "Menu"
      Index           =   0
      Visible         =   0   'False
      Begin VB.Menu opRestaurar 
         Caption         =   "Restaurar"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmInformes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConectar_Click()
    chkConectado.Value = 0
    CrearConexionGlobal
End Sub
Private Sub cmdConvertir_Click()
    txtproceso = ""
    Call informeNCGenerar
    Call enviarPROCNC
End Sub
Private Sub verificarServidorImpresion()
   On Error GoTo verificarServidorImpresion_Error

    If chkConectado.Value = unchecked Then
        CrearConexionGlobal
    End If
    txtproceso = txtproceso & "Comienzo proceso verificarServidorImpresion : " & Date & " " & Time & vbNewLine
    Dim rs As ADODB.Recordset
    Dim consulta As String
    
    consulta = " select id from geslab_canagrosa.impresion " & _
               "  where date_add(concat(substr(fecha,7,4),'-',substr(fecha,4,2),'-',substr(fecha,1,2),' ',hora), interval 5 minute) < current_timestamp " & _
               "    and estado = 0 "
    Set rs = datos_bd(consulta)
    Dim existe As Boolean
    existe = False
    txtproceso = txtproceso & "Registro localizados : " & rs.RecordCount & vbNewLine
    If rs.RecordCount > 0 Then
        existe = True
        rs.MoveNext
    End If
    If existe Then ' Enviar correo
        txtproceso = txtproceso & "Enviando correo : " & txtmysql(7) & vbNewLine
        Enviar_Mail_CDO txtmysql(7), "Existen documentos en GESLAB sin generar. Fecha: " & Format(Date, "dd-mm-yyyy") & " Hora : " & Time, "", ""
    End If
    txtproceso = txtproceso & "Fin del proceso ..." & vbNewLine
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

verificarServidorImpresion_Error:

    txtproceso = txtproceso & "Error " & Err.Number & " (" & Err.Description & ") in procedure verificarServidorImpresion of Formulario frmInformes" & vbNewLine
End Sub
Private Sub cmdEliminar_Click()
    If listabd.ListItems.Count > 0 Then
        listabd.ListItems.Remove listabd.SelectedItem.Index
    End If
End Sub
Private Sub cmdSalir_Click()
    On Error Resume Next
    conn.Close
    Unload Me
    End
End Sub
Private Sub Command1_Click()
    verificarServidorImpresion
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error
    log "Iniciando aplicación....."
    txtmysql(0) = ReadINI(App.Path & "\config.ini", "server", "ip")
    txtmysql(1) = ReadINI(App.Path & "\config.ini", "server", "bd")
    txtmysql(4) = ReadINI(App.Path & "\config.ini", "Parametros", "Correo_Log_prueba")
    txtmysql(5) = ReadINI(App.Path & "\config.ini", "Parametros", "Correo_Todas_NC")
    txtmysql(6) = ReadINI(App.Path & "\config.ini", "Parametros", "Correo_Listado_PROCNC")
    txtmysql(7) = ReadINI(App.Path & "\config.ini", "Parametros", "Correo_Informes_geslab")
    chklog.Value = ReadINI(App.Path & "\config.ini", "Parametros", "log")
    chkInicio.Value = ReadINI(App.Path & "\config.ini", "Parametros", "autoinicio")
    hora_ejecucion = ReadINI(App.Path & "\config.ini", "Parametros", "hora")
   
   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmVersion"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WriteINI App.Path & "\config.ini", "Parametros", "hora", hora_ejecucion
    WriteINI App.Path & "\config.ini", "server", "ip", txtmysql(0)
    WriteINI App.Path & "\config.ini", "server", "bd", txtmysql(1)
    WriteINI App.Path & "\config.ini", "Parametros", "Correo_Log_prueba", txtmysql(4)
    WriteINI App.Path & "\config.ini", "Parametros", "Correo_Todas_NC", txtmysql(5)
    WriteINI App.Path & "\config.ini", "Parametros", "Correo_Listado_PROCNC", txtmysql(6)
    WriteINI App.Path & "\config.ini", "Parametros", "Correo_Informes_geslab", txtmysql(7)
    WriteINI App.Path & "\config.ini", "Parametros", "autoinicio", chkInicio.Value
    WriteINI App.Path & "\config.ini", "Parametros", "log", chklog.Value
End Sub

Private Sub Timer1_Timer()
   On Error GoTo Timer1_Timer_Error
    hora_actual = Time
'    If Weekday(Now) = vbSaturday Or Weekday(Now) = vbSunday Then
'        Exit Sub
'    End If
    If Weekday(Now) = vbSunday And hora_actual = hora_ejecucion Then
        CrearConexionGlobal
        DoEvents
        cmdConvertir_Click
    End If
    ' Verificar servidor de impresión geslab
'    If Right(hora_actual, 5) = "00:00" Or Right(hora_actual, 5) = "10:00" Or Right(hora_actual, 5) = "20:00" Or Right(hora_actual, 5) = "30:00" Or Right(hora_actual, 5) = "40:00" Or Right(hora_actual, 5) = "50:00" Then
'        verificarServidorImpresion
'    End If
   On Error GoTo 0
   Exit Sub

Timer1_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer1_Timer of Formulario frmVersion"
End Sub

Private Sub Timer2_Timer()
   On Error GoTo Timer2_Timer_Error

    hora_actual = Time
    ' Verificar servidor de impresión geslab
    If Right(hora_actual, 4) = "0:00" Or Right(hora_actual, 4) = "5:00" Then
        verificarServidorImpresion
    End If

   On Error GoTo 0
   Exit Sub

Timer2_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer2_Timer of Formulario frmInformes"
End Sub
