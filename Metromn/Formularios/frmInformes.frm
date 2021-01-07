VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmInformes 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Captura de datos Metrohm"
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
   Begin VB.CommandButton cmdProcesarPendientes 
      Caption         =   "Procesar Pendientes"
      Height          =   870
      Left            =   3015
      TabIndex        =   20
      Top             =   8190
      Width           =   3525
   End
   Begin VB.TextBox txtmysql 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   5
      Left            =   45
      ScrollBars      =   2  'Vertical
      TabIndex        =   18
      Top             =   3555
      Width           =   9420
   End
   Begin VB.TextBox txtmysql 
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   4
      Left            =   45
      ScrollBars      =   2  'Vertical
      TabIndex        =   16
      Top             =   2970
      Width           =   9420
   End
   Begin VB.TextBox txtproceso 
      Height          =   3975
      IMEMode         =   3  'DISABLE
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   15
      Top             =   4050
      Width           =   9555
   End
   Begin VB.CheckBox chkPrueba 
      Caption         =   "Modo Prueba"
      Height          =   240
      Left            =   6525
      TabIndex        =   14
      Top             =   1080
      Width           =   1590
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1215
      Top             =   8370
   End
   Begin VB.CheckBox chklog 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generar log"
      Height          =   195
      Left            =   6525
      TabIndex        =   13
      Top             =   810
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
      Left            =   90
      TabIndex        =   1
      Top             =   765
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
   Begin VB.Image imagen 
      Appearance      =   0  'Flat
      Height          =   660
      Left            =   225
      Picture         =   "frmInformes.frx":1A5E
      Stretch         =   -1  'True
      Top             =   45
      Width           =   2580
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ruta para los ficheros procesados"
      Height          =   330
      Index           =   7
      Left            =   90
      TabIndex        =   19
      Top             =   3330
      Width           =   4965
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ruta con los ficheros de Metrohm"
      Height          =   330
      Index           =   5
      Left            =   90
      TabIndex        =   17
      Top             =   2745
      Width           =   4965
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Captura de datos Metrohm"
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
Private Sub conectarBd()
    chkConectado.Value = 0
End Sub

Private Sub cmdProcesarPendientes_Click()
    Dim oM As New clsMetrohm
   On Error GoTo cmdProcesarPendientes_Click_Error

    oM.procesarListadoPendientes

   On Error GoTo 0
   Exit Sub

cmdProcesarPendientes_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdProcesarPendientes_Click of Formulario frmInformes"
End Sub

Private Sub cmdSalir_Click()
    On Error Resume Next
    conn.Close
    Unload Me
    End
End Sub
Private Sub Form_Load()
   On Error GoTo Form_Load_Error
    log "Iniciando aplicación....."
    txtmysql(0) = ReadINI(App.Path & "\config.ini", "server", "ip")
    txtmysql(1) = ReadINI(App.Path & "\config.ini", "server", "bd")
    txtmysql(4) = ReadINI(App.Path & "\config.ini", "Documentos", "Ruta")
    txtmysql(5) = ReadINI(App.Path & "\config.ini", "Documentos", "Procesados")
    
    chklog.Value = ReadINI(App.Path & "\config.ini", "Parametros", "log")
    
    conectarBd
    
    On Error Resume Next
    MkDir txtmysql(4)
    MkDir txtmysql(5)
    
   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmVersion"
End Sub
Private Sub Form_Unload(Cancel As Integer)
    WriteINI App.Path & "\config.ini", "server", "ip", txtmysql(0)
    WriteINI App.Path & "\config.ini", "server", "bd", txtmysql(1)
    WriteINI App.Path & "\config.ini", "Documentos", "Ruta", txtmysql(4)
    WriteINI App.Path & "\config.ini", "Documentos", "Procesados", txtmysql(5)
    WriteINI App.Path & "\config.ini", "Parametros", "log", chklog.Value
End Sub

Private Sub Timer1_Timer()
   On Error GoTo Timer1_Timer_Error
    proceso
   On Error GoTo 0
   Exit Sub

Timer1_Timer_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Timer1_Timer of Formulario frmVersion"
End Sub

Private Sub proceso()
    Dim sArchivo As String
   On Error GoTo proceso_Error
    Dim sql As String
    sArchivo = Dir(txtmysql(4).Text & "\*.csv")
    Do While sArchivo <> ""
        log ("Encontrado fichero : " & sArchivo)
        ' Procesando fichero
        Open txtmysql(4).Text & "\" & sArchivo For Input As #1
        Line Input #1, linea
        Do
            DoEvents
            If InStr(1, linea, "Sample name") = 0 Then
                sql = "INSERT INTO metrohm (REGISTRO,FICHERO) VALUES ('" & linea & "','" & sArchivo & "')"
                execute_bd sql
                
                cmdProcesarPendientes_Click
            End If
            linea = ""
            If EOF(1) = False Then
                Line Input #1, linea
            End If
        Loop Until linea = "" And EOF(1) = True
        Close #1
        ' Mover Archivo a Procesado
        FileCopy txtmysql(4).Text & "\" & sArchivo, txtmysql(5).Text & "\" & sArchivo
        Kill txtmysql(4).Text & "\" & sArchivo
        sArchivo = Dir
    Loop

    ' Contador de pendientes de procesar
    Dim oM As New clsMetrohm
    cmdProcesarPendientes.Caption = "Procesar Pendientes (" & oM.Pendientes & ")"
    
   On Error GoTo 0
   Exit Sub

proceso_Error:

    log "Error " & Err.Number & " (" & Err.Description & ") in procedure proceso of Formulario frmInformes"
End Sub

Private Sub log(texto As String)
    txtproceso = txtproceso & texto & vbNewLine
End Sub
