VERSION 5.00
Begin VB.Form frmAlveograma 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Determinaciones Alveograma"
   ClientHeight    =   7275
   ClientLeft      =   8985
   ClientTop       =   1920
   ClientWidth     =   6285
   Icon            =   "frmAlveograma.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7275
   ScaleWidth      =   6285
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdObservador 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observador"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   39
      Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
      Top             =   6390
      Width           =   1095
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   4050
      Style           =   1  'Graphical
      TabIndex        =   38
      Top             =   6390
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   5175
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   6390
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3705
      Left            =   60
      TabIndex        =   27
      Top             =   2640
      Width           =   6165
      Begin VB.CommandButton cmdCalcular 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calcular"
         Height          =   375
         Left            =   4710
         TabIndex        =   16
         Top             =   3180
         Width           =   1065
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   12
         Left            =   3150
         TabIndex        =   15
         Top             =   3150
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   10
         Left            =   3150
         TabIndex        =   13
         Top             =   2730
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   11
         Left            =   4680
         TabIndex        =   14
         Top             =   2730
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   0
         Left            =   3150
         TabIndex        =   3
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   1
         Left            =   4680
         TabIndex        =   4
         Top             =   630
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   2
         Left            =   3150
         TabIndex        =   5
         Top             =   1050
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   3
         Left            =   4680
         TabIndex        =   6
         Top             =   1050
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   4
         Left            =   3150
         TabIndex        =   7
         Top             =   1470
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   5
         Left            =   4680
         TabIndex        =   8
         Top             =   1470
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   6
         Left            =   3150
         TabIndex        =   9
         Top             =   1890
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   7
         Left            =   4680
         TabIndex        =   10
         Top             =   1890
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   8
         Left            =   3150
         TabIndex        =   11
         Top             =   2310
         Width           =   1095
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   9
         Left            =   4680
         TabIndex        =   12
         Top             =   2310
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "INDICE DE DEGRADACION"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   36
         Top             =   3240
         Width           =   2460
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   35
         Top             =   2820
         Width           =   135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TENACIDAD (mm H2O)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   34
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "EXTENSIBILIDAD (mm)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   33
         Top             =   1140
         Width           =   2235
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "P/L"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   32
         Top             =   1560
         Width           =   345
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "S (cm2)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   31
         Top             =   1980
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "W (x10-4 jul)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   30
         Top             =   2400
         Width           =   1305
      End
      Begin VB.Label lbldeter 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Normal"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   2
         Left            =   3150
         TabIndex        =   29
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label lbldeter 
         Alignment       =   2  'Center
         BackColor       =   &H00C0E0FF&
         Caption         =   "Reposo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   300
         Index           =   3
         Left            =   4680
         TabIndex        =   28
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   60
      TabIndex        =   17
      Top             =   690
      Width           =   6165
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   13
         Left            =   4380
         TabIndex        =   0
         Top             =   180
         Width           =   1245
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   14
         Left            =   4380
         TabIndex        =   1
         Top             =   600
         Width           =   1245
      End
      Begin VB.TextBox val 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Index           =   15
         Left            =   4380
         TabIndex        =   2
         Top             =   1020
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "TEMPERATURA DEL LABORATORIO"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   23
         Top             =   270
         Width           =   3225
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "ºC"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   5730
         TabIndex        =   22
         Top             =   300
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "HUMEDAD RELATIVA DEL AIRE"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   690
         Width           =   2895
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   5730
         TabIndex        =   20
         Top             =   720
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "HUMEDAD DE LA HARINA"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   19
         Top             =   1110
         Width           =   2370
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   5730
         TabIndex        =   18
         Top             =   1140
         Width           =   225
      End
   End
   Begin VB.Label lblCerrada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   345
      Left            =   4185
      TabIndex        =   40
      Top             =   30
      Width           =   2085
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Valores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   1
      Left            =   60
      TabIndex        =   26
      Top             =   2340
      Width           =   6165
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Organoleptico"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   30
      TabIndex        =   25
      Top             =   30
      Width           =   6225
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resultados"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Index           =   0
      Left            =   60
      TabIndex        =   24
      Top             =   390
      Width           =   6165
   End
End
Attribute VB_Name = "frmAlveograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public alveo As Long

'Private WithEvents TecladoNumerico As frmTecladoNumerico
'Private blnTecladoNumericoPrimeraVez As Boolean
Private mvarintIndiceValor As Integer
'Private blnEsTablet As Boolean



Private Sub cmdObservador_Click()
Dim objfrm As New frmObservadorEnsayo
   
   'MANTIS-807-I
    objfrm.FORMULARIO_ORIGEN = 2
   'MANTIS-807-F
    objfrm.ES_CONTROL_EFICACIA = False
    objfrm.MUESTRA_ID = gmuestra ' Id de la muestra
    objfrm.TIPO_DETERMINACION_ENSAYO_ID = CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma")) ' tipo de la Determinacion
    objfrm.DETERMINACION_ENSAYO_ID = gdeterminacion
    objfrm.MUESTRA_CERRADA = (Not cmdok.Enabled)
    objfrm.TIPO_OBSERVACION_ID = MC_TIPOS_OBSERVACION.MCTO_DETERMINACION


    objfrm.Show vbModal
    
    Set objfrm = Nothing
    
End Sub

'Private Sub ConfigurarTablet()
'Set TecladoNumerico = New frmTecladoNumerico
'
'    TecladoNumerico.OcultarConformidad = True
'    TecladoNumerico.posX = Me.Width + 60
'    TecladoNumerico.posY = Me.top
'
'    blnEsTablet = pc_es_tablet
'
'    If blnEsTablet Then
'
'        blnTecladoNumericoPrimeraVez = True
'
'        val(0).Locked = True
'        val(1).Locked = True
'        val(2).Locked = True
'        val(3).Locked = True
'        val(4).Locked = True
'        val(5).Locked = True
'        val(6).Locked = True
'        val(7).Locked = True
'        val(8).Locked = True
'        val(9).Locked = True
'        val(10).Locked = True
'        val(11).Locked = True
'        val(12).Locked = True
'        val(13).Locked = True
'        val(14).Locked = True
'        val(15).Locked = True
'
'
'        Me.Left = 0
'    End If
'End Sub

'Private Sub MostrarTecladoNumerico()
'
'    If Not blnEsTablet Then Exit Sub
'
'    If blnTecladoNumericoPrimeraVez Then
'        blnTecladoNumericoPrimeraVez = False
'        TecladoNumerico.TextoInicial = val(mvarintIndiceValor).Text
'        TecladoNumerico.cabecera = getCabecera()
'        TecladoNumerico.Subcabecera = getSubCabecera()
'        If Not TecladoNumerico.Visible Then
'            TecladoNumerico.Show 1
'        End If
'    End If
'
'End Sub

Private Sub cmdCalcular_Click()
    On Error GoTo fallo
'    log ("A-1")
'    If (Trim(val(13) = "") Or IsNumeric(val(13)) = False) Then
'        MsgBox "Debe indicar la TEMPERATURA.", vbInformation, App.Title
'        Exit Sub
'    End If
'    log ("A-2")
'    If (Trim(val(14) = "") Or IsNumeric(val(14)) = False) Then
'        MsgBox "Debe indicar la HUMEDAD DEL AIRE.", vbInformation, App.Title
'        Exit Sub
'    End If
'    log ("A-3")
'    If (Trim(val(15) = "") Or IsNumeric(val(15)) = False) Then
'        MsgBox "Debe indicar la HUMEDAD DE LA HARINA.", vbInformation, App.Title
'        Exit Sub
'    End If
'    log ("A-4")
'    If (Trim(val(0) = "") Or IsNumeric(val(0)) = False) Then
'        MsgBox "Debe indicar la TENACIDAD.", vbInformation, App.Title
'        Exit Sub
'    End If
'    log ("A-5")
'    If (Trim(val(2) = "") Or IsNumeric(val(2)) = False) Then
'        MsgBox "Debe indicar la EXTENSIBILIDAD.", vbInformation, App.Title
'        Exit Sub
'    End If
'    log ("A-6")
'    If (Trim(val(6) = "") Or IsNumeric(val(6)) = False) Then
'        MsgBox "Debe indicar S.", vbInformation, App.Title
'        Exit Sub
'    End If
'    log ("A-7")
'    If (Trim(val(1) = "") Or IsNumeric(val(1)) = False) Then
'        MsgBox "Debe indicar la TENACIDAD en el ensayo de REPOSO.", vbInformation, App.Title
'        Exit Sub
'    End If
'    log ("A-8")
'    If (Trim(val(3) = "") Or IsNumeric(val(3)) = False) Then
'        MsgBox "Debe indicar la EXTENSIBILIDAD en el ensayo de REPOSO.", vbInformation, App.Title
'        Exit Sub
'    End If
'    log ("A-9")
'    If (Trim(val(7) = "") Or IsNumeric(val(7)) = False) Then
'        MsgBox "Debe indicar S en el ensayo de REPOSO.", vbInformation, App.Title
'        Exit Sub
'    End If
    If Trim(val(13)) = "" Then
        MsgBox "Debe indicar la TEMPERATURA.", vbInformation, App.Title
        Exit Sub
    End If
    If Trim(val(14)) = "" Then
        MsgBox "Debe indicar la HUMEDAD DEL AIRE.", vbInformation, App.Title
        Exit Sub
    End If
    If Trim(val(15)) = "" Then
        MsgBox "Debe indicar la HUMEDAD DE LA HARINA.", vbInformation, App.Title
        Exit Sub
    End If
    If Trim(val(0)) = "" Then
        MsgBox "Debe indicar la TENACIDAD.", vbInformation, App.Title
        Exit Sub
    End If
    If Trim(val(2)) = "" Then
        MsgBox "Debe indicar la EXTENSIBILIDAD.", vbInformation, App.Title
        Exit Sub
    End If
    If Trim(val(6)) = "" Then
        MsgBox "Debe indicar S.", vbInformation, App.Title
        Exit Sub
    End If
    If Trim(val(1)) = "" Then
        MsgBox "Debe indicar la TENACIDAD en el ensayo de REPOSO.", vbInformation, App.Title
        Exit Sub
    End If
    If Trim(val(3)) = "" Then
        MsgBox "Debe indicar la EXTENSIBILIDAD en el ensayo de REPOSO.", vbInformation, App.Title
        Exit Sub
    End If
    If Trim(val(7)) = "" Then
        MsgBox "Debe indicar S en el ensayo de REPOSO.", vbInformation, App.Title
        Exit Sub
    End If
    If CSng(val(2)) > 0 Then
        val(4) = formatear(CSng(val(0)) / CSng(val(2)), 5, 2)
    Else
        val(4) = formatear(CSng(val(0)), 5, 2)
    End If
    If CSng(val(3)) > 0 Then
        val(5) = formatear(CSng(val(1)) / CSng(val(3)), 5, 2)
    Else
        val(5) = formatear(CSng(val(1)), 5, 2)
    End If
    val(10) = formatear(2.22 * Sqr(CSng(val(2))), 5, 2)
    val(11) = formatear(2.22 * Sqr(CSng(val(3))), 5, 2)
    val(8) = formatear(6.54 * CSng(val(6)), 5, 2)
    val(9) = formatear(6.54 * CSng(val(7)), 5, 2)
    If CSng(val(9)) > CSng(val(8)) Then
        val(12) = 0
    Else
        val(12) = formatear((CSng(val(8)) - CSng(val(9))) / CSng(val(8)) * 100, 5, 0)
    End If
    Exit Sub
fallo:
    MsgBox "Error al calcular los resultados del alveograma.", vbCritical, App.Title
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    cmdCalcular_Click
    Dim oalveo As New clsAlveogramas
    Dim oalveoval As New clsAlveograma_valores
    Dim oDeter As New clsDeterminaciones
    Dim odd As New clsDatos_determinaciones
    Dim i As Integer
    If alveo = 0 Then  ' Nuevo alveograma
      If MsgBox("Va a introducir los datos del Alveograma. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        ' Alveograma
        With oalveo
         .setDETERMINACION_ID = gdeterminacion
         .setMUESTRA_ID = CSng(gmuestra)
         .setTEMPERATURA = Replace(val(13), ",", ".")
         .setHUMEDAD_AIRE = Replace(val(14), ",", ".")
         .setHUMEDAD_HARINA = Replace(val(15), ",", ".")
         .setINDICE_DEGRADACION = Replace(val(12), ",", ".")
         .setFECHA = Format(Date, "yyyy-mm-dd")
         alveo = .InsertarAlveograma
        End With
        ' Alveograma_Valores
        With oalveoval
         ' NORMAL
         .setALVEOGRAMA_ID = alveo
         .setTENACIDAD = Replace(val(0), ",", ".")
         .setEXTENSIBILIDAD = Replace(val(2), ",", ".")
         .setS = Replace(val(6), ",", ".")
         .setW = Replace(val(8), ",", ".")
         .setG = Replace(val(10), ",", ".")
         .setDE_REPOSO = 0
         .InsertarAlveogramaValores
         ' REPOSO
         .setALVEOGRAMA_ID = alveo
         .setTENACIDAD = Replace(val(1), ",", ".")
         .setEXTENSIBILIDAD = Replace(val(3), ",", ".")
         .setS = Replace(val(7), ",", ".")
         .setW = Replace(val(9), ",", ".")
         .setG = Replace(val(11), ",", ".")
         .setDE_REPOSO = 1
         .InsertarAlveogramaValores
        End With
        ' Datos Determinaciones
        If odd.CARGAR(gdeterminacion, 331) = True Then
            odd.setVALOR_1 = Replace(val(12), ",", ".")
            odd.Insertar_Valores
        End If
        ' Almacena determinacion (Solucion)
        oDeter.setRESULTADO = Replace(val(12), ",", ".")
        oDeter.setDIF_DUPLICADOS = ""
        oDeter.setFECHA = Format(Date, "yyyy-mm-dd")
        oDeter.setHORA = Format(Time, "hh:mm")
        oDeter.setEMPLEADO_ID = usuario.getID_EMPLEADO
        oDeter.InsertarSolucion (gdeterminacion)
        Unload Me
      End If
    Else    ' Modificacion alveograma
      If MsgBox("Va a modificar los datos del Alveograma. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        ' Alveograma
        With oalveo
         .setTEMPERATURA = Replace(val(13), ",", ".")
         .setHUMEDAD_AIRE = Replace(val(14), ",", ".")
         .setHUMEDAD_HARINA = Replace(val(15), ",", ".")
         .setINDICE_DEGRADACION = Replace(val(12), ",", ".")
         .setFECHA = Format(Date, "yyyy-mm-dd")
         Call .ModificarAlveograma(alveo)
        End With
        ' Alveograma_Valores
        With oalveoval
         ' NORMAL
         .setTENACIDAD = Replace(val(0), ",", ".")
         .setEXTENSIBILIDAD = Replace(val(2), ",", ".")
         .setS = Replace(val(6), ",", ".")
         .setW = Replace(val(8), ",", ".")
         .setG = Replace(val(10), ",", ".")
         Call .ModificarAlveogramaValores(alveo, 0)
         ' REPOSO
         .setTENACIDAD = Replace(val(1), ",", ".")
         .setEXTENSIBILIDAD = Replace(val(3), ",", ".")
         .setS = Replace(val(7), ",", ".")
         .setW = Replace(val(9), ",", ".")
         .setG = Replace(val(11), ",", ".")
         Call .ModificarAlveogramaValores(alveo, 1)
        End With
        ' Datos Determinaciones
        If odd.CARGAR(gdeterminacion, 331) = True Then
            odd.setVALOR_1 = Replace(val(12), ",", ".")
            odd.Insertar_Valores
        End If
        ' Almacena determinacion (Solucion)
        oDeter.setRESULTADO = Replace(val(12), ",", ".")
        oDeter.setFECHA = Format(Date, "yyyy-mm-dd")
        oDeter.setHORA = Format(Time, "hh:mm")
        oDeter.setEMPLEADO_ID = usuario.getID_EMPLEADO
        oDeter.InsertarSolucion (gdeterminacion)
        Unload Me
      End If
    End If
    Exit Sub
fallo:
    MsgBox "Error al insertar los datos (granulometria)", vbCritical, Err.Description
End Sub

Private Sub Form_Activate()
    ' Comprobar si ya existe
    Dim oalveo As New clsAlveogramas
   On Error GoTo Form_Activate_Error

    alveo = oalveo.ComprobarAlveograma(gmuestra, gdeterminacion)
    If alveo <> 0 Then
        ' Alveograma
        With oalveo
         .CargarAlveograma (alveo)
         val(13) = .getTEMPERATURA
         val(14) = .getHUMEDAD_AIRE
         val(15) = .getHUMEDAD_HARINA
         val(12) = .getINDICE_DEGRADACION
        End With
        ' Alveograma_valores
        Dim oalveo_val As New clsAlveograma_valores
        With oalveo_val
         If .CargarAlveogramaValores(alveo, 0) = True Then
            val(0) = .getTENACIDAD
            val(2) = .getEXTENSIBILIDAD
            val(6) = .getS
            val(8) = .getW
            val(10) = .getG
         End If
         If .CargarAlveogramaValores(alveo, 1) = True Then
            val(1) = .getTENACIDAD
            val(3) = .getEXTENSIBILIDAD
            val(7) = .getS
            val(9) = .getW
            val(11) = .getG
         End If
        End With
        ' Formatear valores
        Dim i As Integer
        For i = 0 To 15
             If val(i).Text <> "" Then
                 val(i) = formatear(val(i), 5, 2)
             End If
        Next
        cmdCalcular_Click
        cmdok.Visible = True
    End If
    Set oalveo = Nothing
    Set oalveo_val = Nothing

    mvarintIndiceValor = 13
    'TecladoNumerico_Change val(mvarintIndiceValor).Text
    
'    MostrarTecladoNumerico
    Dim oMuestra As New clsMuestra
    If oMuestra.CargaMuestra(gmuestra) Then
        proteger_campos oMuestra.getCERRADA
    End If

   On Error GoTo 0
   Exit Sub

Form_Activate_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Activate of Formulario frmAlveograma"
End Sub


Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    ' Título
    Dim oMuestra As New clsMuestra
    lbltitulo = "Alveograma : " & Trim(str(gmuestra)) & " (" & oMuestra.CodigoParticular(CLng(gmuestra)) & ")"
    Me.Caption = lbltitulo
    
'    ConfigurarTablet
    
End Sub

'Private Sub TecladoNumerico_Change(ByVal res As String)
'    Dim iCont As Integer
'    For iCont = 0 To 15
'        val(iCont).BackColor = vbWhite
'    Next iCont
'
'    val(mvarintIndiceValor).BackColor = &H80C0FF
'    val(mvarintIndiceValor).Text = res
'End Sub
'
'Private Sub TecladoNumerico_Salir()
'    blnTecladoNumericoPrimeraVez = False
'    cmdCalcular_Click
'    cmdCalcular.SetFocus
'End Sub
'
'Private Sub TecladoNumerico_SiguienteElemento(cabecera As String, Subcabecera As String, RESULTADO As String, fecha As String, CONFORME As Integer, Cerrar As Boolean, desestimarEvento As Boolean)
'If mvarintIndiceValor = 15 Then
'    mvarintIndiceValor = 0
'ElseIf mvarintIndiceValor = 11 Then
'    Cerrar = True
'    cmdCalcular_Click
'Else
'    mvarintIndiceValor = mvarintIndiceValor + 1
'    cabecera = getCabecera
'    Subcabecera = getSubCabecera
'    RESULTADO = val(mvarintIndiceValor).Text
'End If
'
'
'End Sub


Private Sub val_GotFocus(Index As Integer)
    val(Index).BackColor = &H80C0FF
    val(Index).SelStart = 0
    val(Index).SelLength = Len(val(Index))
    
    mvarintIndiceValor = Index
    
'    blnTecladoNumericoPrimeraVez = True
'    MostrarTecladoNumerico
    
End Sub

Private Sub val_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40
       If Index = 12 Then
        val(13).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 13 Then
        val(12).SetFocus
       Else
        SendKeys "+{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 27
        cmdcancel_Click
    End Select
End Sub

Private Sub val_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       If Index = 12 Then
        val(13).SetFocus
        KeyAscii = 0 ' Para evitar el "bip" del sistema
       Else
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
       End If
    End If
    ' Escribir ',' al pulsar '.'
    If KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

Private Sub val_LostFocus(Index As Integer)
    val(Index).BackColor = vbWhite
    If val(Index).Text <> "" Then
        val(Index) = formatear(val(Index), 5, 2)
    End If
End Sub


Private Function getCabecera() As String
Select Case mvarintIndiceValor
    Case 0
        getCabecera = Label1(6)
    Case 1
        getCabecera = Label1(6)
    Case 2
        getCabecera = Label1(7)
    Case 3
        getCabecera = Label1(7)
    Case 4
        getCabecera = Label1(8)
    Case 5
        getCabecera = Label1(8)
    Case 6
        getCabecera = Label1(9)
    Case 7
        getCabecera = Label1(9)
    Case 8
        getCabecera = Label1(10)
    Case 9
        getCabecera = Label1(10)
    Case 10
        getCabecera = Label1(3)
    Case 11
        getCabecera = Label1(3)
    Case 13
        getCabecera = Label1(0)
    Case 14
        getCabecera = Label1(1)
    Case 15
        getCabecera = Label1(2)
    Case Else
        getCabecera = ""
    
    End Select
End Function

Private Function getSubCabecera() As String

Select Case mvarintIndiceValor
    Case 0
        getSubCabecera = lbldeter(2)
    Case 1
        getSubCabecera = lbldeter(3)
    Case 2
        getSubCabecera = lbldeter(2)
    Case 3
        getSubCabecera = lbldeter(3)
    Case 4
        getSubCabecera = lbldeter(2)
    Case 5
        getSubCabecera = lbldeter(3)
    Case 6
        getSubCabecera = lbldeter(2)
    Case 7
        getSubCabecera = lbldeter(3)
    Case 8
        getSubCabecera = lbldeter(2)
    Case 9
        getSubCabecera = lbldeter(3)
    Case 10
        getSubCabecera = lbldeter(2)
    Case 11
        getSubCabecera = lbldeter(3)
    Case 13
        getSubCabecera = Label3(0)
    Case 14
        getSubCabecera = Label3(1)
    Case 15
        getSubCabecera = Label3(2)
    Case Else
        getSubCabecera = ""
    
    End Select


End Function
Private Sub proteger_campos(CERRADA As Integer)
    Select Case CERRADA
        Case 0
            lblCerrada = "ABIERTA"
        Case 1
            lblCerrada = "CERRADA"
        Case 2
            lblCerrada = "PTE. CIERRE"
        Case 3
            lblCerrada = "C.SIN INFORME"
    End Select
    If CERRADA = 1 Then
        cmdok.Enabled = False
        cmdCalcular.Enabled = False
    Else
        cmdok.Enabled = True
        cmdCalcular.Enabled = True
    End If
End Sub

