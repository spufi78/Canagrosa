VERSION 5.00
Begin VB.Form frmEquipos_Listados_Seleccion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   3000
   ClientLeft      =   2445
   ClientTop       =   3465
   ClientWidth     =   10110
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEquipos_Listados_Seleccion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3000
   ScaleWidth      =   10110
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdlistado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado Equipos Sujetos a NADCAP  (LI-06)"
      Height          =   1500
      Index           =   3
      Left            =   1740
      Picture         =   "frmEquipos_Listados_Seleccion.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   540
      Width           =   1635
   End
   Begin VB.CommandButton cmdlistado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mantenimientos Previstos"
      Height          =   1500
      Index           =   7
      Left            =   6780
      Picture         =   "frmEquipos_Listados_Seleccion.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   540
      Width           =   1635
   End
   Begin VB.CommandButton cmdlistado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Material de Referencia"
      Height          =   1500
      Index           =   4
      Left            =   8460
      Picture         =   "frmEquipos_Listados_Seleccion.frx":11A0
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   540
      Width           =   1635
   End
   Begin VB.CommandButton cmdlistado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Inventario (IN-05)"
      Height          =   1500
      Index           =   2
      Left            =   60
      Picture         =   "frmEquipos_Listados_Seleccion.frx":1A6A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   540
      Width           =   1635
   End
   Begin VB.CommandButton cmdlistado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Histórico de Calibraciones"
      Height          =   1500
      Index           =   1
      Left            =   3420
      Picture         =   "frmEquipos_Listados_Seleccion.frx":2334
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   540
      Width           =   1635
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9030
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2100
      Width           =   1050
   End
   Begin VB.CommandButton cmdlistado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historico de Verificaciones"
      Height          =   1500
      Index           =   0
      Left            =   5100
      Picture         =   "frmEquipos_Listados_Seleccion.frx":2BFE
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   540
      Width           =   1635
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9540
      Picture         =   "frmEquipos_Listados_Seleccion.frx":34C8
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Listados de Equipos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   90
      TabIndex        =   2
      Top             =   120
      Width           =   2145
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   10065
   End
End
Attribute VB_Name = "frmEquipos_Listados_Seleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    Public informe As String
    Public criterio As String

Private Sub cmdCancel_Click()
    Me.informe = ""
    Unload Me
End Sub

Private Sub cmdListado_Click(Index As Integer)
    Select Case Index
        Case 0 ' verificación
            'CRITERIO = "{equipos.ALTA_BAJA}=0 and {equipos.CON_VERIFICACION}=1 and {eq_verificacion_equipos.ESTADO} > 0 AND {PERIODICIDADES.CODIGO} = " & decodificadora.eq_periodicidad
            criterio = ""
            informe = "Equipos\rptEquipos_Listado_Verificacion"
            
        Case 5 ' verificación
            criterio = "{equipos.ALTA_BAJA}=0 and {equipos.CON_VERIFICACION}=1 and {eq_verificacion_equipos.ESTADO} = 0 AND {PERIODICIDADES.CODIGO} = " & decodificadora.EQ_periodicidad
            informe = "Equipos\rptEquipos_Listado_Verificacion_previstos"
            
        Case 1 ' calibración
            'CRITERIO = "{equipos.ALTA_BAJA}=0 and {equipos.CON_CALIBRACION}=1 and {eq_verificacion_equipos.ESTADO} > 0 AND {PERIODICIDADES.CODIGO} = " & decodificadora.EQ_periodicidad
            criterio = "{calibraciones.fecha_proxima} >= CDate (" & Year(Date) & "," & Format(Date, "mm") & "," & Format(Date, "dd") & ")"
            informe = "Equipos\rptEquipos_Listado_Calibracion"
            
        Case 6 ' calibración
            criterio = "{equipos.ALTA_BAJA}=0 and {equipos.CON_CALIBRACION}=1 and {eq_verificacion_equipos.ESTADO} = 0 AND {PERIODICIDADES.CODIGO} = " & decodificadora.EQ_periodicidad
            informe = "Equipos\rptEquipos_Listado_Verificacion_previstos"
            
        Case 2 ' inventario (IN-05)
            criterio = "{equipos.ALTA_BAJA}=0"
            informe = "Equipos\rptEquipos_Listado_Inventario"
            
        Case 3 ' Listado Equipos Sujetos a NADCAP (LI-06)
            criterio = "{equipos.ALTA_BAJA}=0 and {equipos.ES_NADCAP}=1"
            informe = "Equipos\rptEquipos_Listado_nadcap"
        
        Case 7 ' mantenimiento
            criterio = "{equipos.ALTA_BAJA}=0 and {equipos.CON_MANTENIMIENTO}=1 and {eq_mantenimiento_equipos.ESTADO} = 0  AND {PERIODICIDADES.CODIGO} = " & decodificadora.EQ_periodicidad
            informe = "Equipos\rptEquipos_Listado_Mantenimiento_previstos"
            
        Case 4 ' mrc
            criterio = "({tipos_bote_ex.TIPO_M_REFERENCIA_ID}= 2 or {tipos_bote_ex.TIPO_M_REFERENCIA_ID}= 3)"
            If MsgBox("¿Desea mostrar los equipos caducados?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                criterio = criterio & " and {botes_ex.FECHA_CADUCIDAD} >= CDate (" & Year(Date) & "," & Format(Date, "mm") & "," & Format(Date, "dd") & ")"
            End If
            informe = "Equipos\rptEquipos_Listado_MRC"
            
    End Select
    'frmEquipos_Listado.informe = informe
    'frmEquipos_Listado.criterio = criterio
    'Unload Me
    Me.Hide
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
End Sub
