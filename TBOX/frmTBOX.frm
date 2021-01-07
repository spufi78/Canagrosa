VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Begin VB.Form frmTBOX 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Captura de datos T-BOX"
   ClientHeight    =   11070
   ClientLeft      =   5475
   ClientTop       =   3915
   ClientWidth     =   13710
   DrawWidth       =   10
   Icon            =   "frmTBOX.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   11070
   ScaleWidth      =   13710
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtIdEquipo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   11520
      TabIndex        =   18
      Top             =   1440
      Width           =   2085
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10710
      Picture         =   "frmTBOX.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   10035
      Width           =   1410
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   4950
      Top             =   10305
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12195
      Picture         =   "frmTBOX.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10035
      Width           =   1410
   End
   Begin VB.TextBox txtMedidas 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8730
      TabIndex        =   8
      Top             =   1440
      Width           =   1005
   End
   Begin VB.TextBox txtPuntos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   6255
      TabIndex        =   5
      Top             =   1440
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Configuración del puerto Serie"
      Height          =   1350
      Left            =   90
      TabIndex        =   1
      Top             =   9450
      Visible         =   0   'False
      Width           =   3525
      Begin VB.TextBox txtconf 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   945
         TabIndex        =   22
         Top             =   765
         Width           =   2445
      End
      Begin VB.TextBox txtcom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   945
         TabIndex        =   20
         Top             =   315
         Width           =   2445
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "CONF."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   45
         TabIndex        =   23
         Top             =   810
         Width           =   1005
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "NºCOM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   45
         TabIndex        =   21
         Top             =   360
         Width           =   1005
      End
   End
   Begin MSComctlLib.ListView listaPrecargas 
      Height          =   5775
      Left            =   90
      TabIndex        =   0
      Top             =   3015
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   10186
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin XtremeSuiteControls.PushButton cmdPrecargas 
      Height          =   555
      Left            =   90
      TabIndex        =   6
      Top             =   2430
      Width           =   2535
      _Version        =   851970
      _ExtentX        =   4471
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "PRECARGAS"
      BackColor       =   8421631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdPrecargaEliminar 
      Height          =   555
      Left            =   90
      TabIndex        =   12
      Top             =   8820
      Width           =   2535
      _Version        =   851970
      _ExtentX        =   4471
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "Eliminar Última"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "frmTBOX.frx":149E
   End
   Begin MSComctlLib.ListView listaMedidas 
      Height          =   5775
      Left            =   2655
      TabIndex        =   13
      Top             =   3015
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   10186
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin XtremeSuiteControls.PushButton cmdMedidas 
      Height          =   555
      Left            =   2655
      TabIndex        =   14
      Top             =   2430
      Width           =   3255
      _Version        =   851970
      _ExtentX        =   5741
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "MEDIDAS"
      BackColor       =   8421631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdMedidasEliminar 
      Height          =   555
      Left            =   2655
      TabIndex        =   15
      Top             =   8820
      Width           =   3255
      _Version        =   851970
      _ExtentX        =   5741
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "Eliminar Última"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "frmTBOX.frx":174A
   End
   Begin VB.TextBox txtCodigo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2205
      TabIndex        =   3
      Text            =   "PM4A"
      Top             =   1440
      Width           =   2625
   End
   Begin MSComctlLib.ListView listaBrazo 
      Height          =   5775
      Left            =   5940
      TabIndex        =   25
      Top             =   3015
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   10186
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin XtremeSuiteControls.PushButton cmdBrazo 
      Height          =   555
      Left            =   5940
      TabIndex        =   26
      Top             =   2430
      Width           =   2535
      _Version        =   851970
      _ExtentX        =   4471
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "VAR. BRAZO"
      BackColor       =   8421631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdBrazoEliminar 
      Height          =   555
      Left            =   5940
      TabIndex        =   27
      Top             =   8820
      Width           =   2535
      _Version        =   851970
      _ExtentX        =   4471
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "Eliminar Última"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "frmTBOX.frx":19F6
   End
   Begin MSComctlLib.ListView listaTiempo 
      Height          =   5775
      Left            =   8505
      TabIndex        =   28
      Top             =   3015
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   10186
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin XtremeSuiteControls.PushButton cmdTiempo 
      Height          =   555
      Left            =   8505
      TabIndex        =   29
      Top             =   2430
      Width           =   2535
      _Version        =   851970
      _ExtentX        =   4471
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "VAR. TIEMPO"
      BackColor       =   8421631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmbTiempoEliminar 
      Height          =   555
      Left            =   8505
      TabIndex        =   30
      Top             =   8820
      Width           =   2535
      _Version        =   851970
      _ExtentX        =   4471
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "Eliminar Última"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "frmTBOX.frx":1CA2
   End
   Begin MSComctlLib.ListView listaRepro 
      Height          =   5775
      Left            =   11070
      TabIndex        =   31
      Top             =   3015
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   10186
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin XtremeSuiteControls.PushButton cmdRepro 
      Height          =   555
      Left            =   11070
      TabIndex        =   32
      Top             =   2430
      Width           =   2535
      _Version        =   851970
      _ExtentX        =   4471
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "REPRODUCIBILIDAD"
      BackColor       =   8421631
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdReproEliminar 
      Height          =   555
      Left            =   11070
      TabIndex        =   33
      Top             =   8820
      Width           =   2535
      _Version        =   851970
      _ExtentX        =   4471
      _ExtentY        =   979
      _StockProps     =   79
      Caption         =   "Eliminar Última"
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      UseVisualStyle  =   -1  'True
      Picture         =   "frmTBOX.frx":1F4E
   End
   Begin VB.Label Label9 
      BackColor       =   &H00FFFFFF&
      Caption         =   "F11-Configurar el puerto serie"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   90
      TabIndex        =   24
      Top             =   10800
      Width           =   4920
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FFFFFF&
      Caption         =   "ID EQUIPO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   10170
      TabIndex        =   19
      Top             =   1530
      Width           =   1365
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      Caption         =   "v.1.00"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   8685
      TabIndex        =   16
      Top             =   495
      Width           =   4920
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CAPTURA DE DATOS T-BOX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   5805
      TabIndex        =   10
      Top             =   90
      Width           =   7800
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   315
      TabIndex        =   9
      Top             =   2025
      Visible         =   0   'False
      Width           =   10680
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "MEDIDAS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   7515
      TabIndex        =   7
      Top             =   1530
      Width           =   1185
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "PUNTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   5130
      TabIndex        =   4
      Top             =   1530
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "CÓDIGO EQUIPO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   135
      TabIndex        =   2
      Top             =   1530
      Width           =   3075
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   135
      Picture         =   "frmTBOX.frx":21FA
      Top             =   -45
      Width           =   5250
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
Attribute VB_Name = "frmTBOX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTiempoEliminar_Click()
    If listaTiempo.ListItems.Count > 0 Then
        listaTiempo.ListItems.Remove listaTiempo.ListItems.Count
    End If
End Sub

Private Sub cmdBrazo_Click()
    deshabilitarTodo
    listaBrazo.Enabled = True
    cmdBrazo.BackColor = vbGreen
    cmdBrazo.Enabled = False
    lblMsg = "ESPERANDO VARIACIÓN DE BRAZO Nº " & listaBrazo.ListItems.Count + 1
    lblMsg.Visible = True
    cierra_comm
    inicia_comm
End Sub

Private Sub cmdBrazoEliminar_Click()
    If listaBrazo.ListItems.Count > 0 Then
        listaBrazo.ListItems.Remove listaBrazo.ListItems.Count
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdMedidas_Click()
    deshabilitarTodo
    listaMedidas.Enabled = True
    cmdMedidas.BackColor = vbGreen
    cmdMedidas.Enabled = False
    If listaMedidas.ListItems.Count = 0 Then
        lblMsg = "ESPERANDO PUNTO 1, MEDIDA 1"
    End If
    lblMsg.Visible = True
    cierra_comm
    inicia_comm
End Sub
Private Sub deshabilitarTodo()
    listaPrecargas.Enabled = False
    listaMedidas.Enabled = False
    listaBrazo.Enabled = False
    listaTiempo.Enabled = False
    listaRepro.Enabled = False
    
    cmdPrecargas.Enabled = True
    cmdMedidas.Enabled = True
    cmdBrazo.Enabled = True
    cmdTiempo.Enabled = True
    cmdRepro.Enabled = True
    
    cmdPrecargas.BackColor = &H8080FF
    cmdMedidas.BackColor = &H8080FF
    cmdBrazo.BackColor = &H8080FF
    cmdTiempo.BackColor = &H8080FF
    cmdRepro.BackColor = &H8080FF
    
End Sub

Private Sub cmdMedidasEliminar_Click()
    If listaMedidas.ListItems.Count > 0 Then
        listaMedidas.ListItems.Remove listaMedidas.ListItems.Count
    End If

End Sub

Private Sub cmdok_Click()
    Dim oTT As New clsTorque_tbox
   On Error GoTo cmdok_Click_Error
    
    If Not IsNumeric(txtPuntos) Or Not IsNumeric(txtMedidas) Or Not IsNumeric(txtIdEquipo) Or listaPrecargas.ListItems.Count = 0 Or listaMedidas.ListItems.Count = 0 Then
        MsgBox "Rellene todos los datos!!!", vbCritical, App.Title
        Exit Sub
    End If

    Dim precargas As String
    Dim msj As String
    Dim brazo As String
    Dim tiempo As String
    Dim repro As String
    Dim i As Integer
    For i = 1 To listaPrecargas.ListItems.Count
        precargas = precargas & listaPrecargas.ListItems(i).SubItems(1) & ";"
    Next
    For i = 1 To listaMedidas.ListItems.Count
        msj = msj & listaMedidas.ListItems(i).SubItems(2) & ";"
    Next
    For i = 1 To listaBrazo.ListItems.Count
        brazo = brazo & listaBrazo.ListItems(i).SubItems(1) & ";"
    Next
    For i = 1 To listaTiempo.ListItems.Count
        tiempo = tiempo & listaTiempo.ListItems(i).SubItems(1) & ";"
    Next
    For i = 1 To listaRepro.ListItems.Count
        repro = repro & listaRepro.ListItems(i).SubItems(1) & ";"
    Next
    ' Reemplazar , por .
    precargas = Replace(precargas, ",", ".")
    msj = Replace(msj, ",", ".")
    brazo = Replace(brazo, ",", ".")
    tiempo = Replace(tiempo, ",", ".")
    repro = Replace(repro, ",", ".")
    
    With oTT
        .setEQUIPO_ID = txtIdEquipo
        .setN_PUNTOS = txtPuntos
        .setN_MEDIDAS = txtMedidas
        
        .setPRECARGAS = precargas
        .setMSJ = msj
        .setVARIACION_BRAZO = brazo
        .setVARIACION_TIEMPO = tiempo
        .setREPRODUCIBILIDAD = repro
        
        .Insertar
        MsgBox "Registro insertado correctamente.", vbOKOnly + vbInformation, App.Title
    End With
    limpiarCampos
    limpiarlistas
    txtCodigo = ""
    txtCodigo.SetFocus
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmTBOX"
End Sub

Private Sub cmdPrecargaEliminar_Click()
    If listaPrecargas.ListItems.Count > 0 Then
        listaPrecargas.ListItems.Remove listaPrecargas.ListItems.Count
    End If
End Sub

Private Sub cmdPrecargas_Click()
    deshabilitarTodo
    listaPrecargas.Enabled = True
    cmdPrecargas.BackColor = vbGreen
    cmdPrecargas.Enabled = False
    lblMsg = "ESPERANDO PRECARGA Nº " & listaPrecargas.ListItems.Count + 1
    lblMsg.Visible = True
    cierra_comm
    inicia_comm
End Sub

Private Sub cmdRepro_Click()
    deshabilitarTodo
    listaRepro.Enabled = True
    cmdRepro.BackColor = vbGreen
    cmdRepro.Enabled = False
    lblMsg = "ESPERANDO REPRODUCIBILIDAD Nº " & listaRepro.ListItems.Count + 1
    lblMsg.Visible = True
    cierra_comm
    inicia_comm

End Sub

Private Sub cmdReproEliminar_Click()
    If listaRepro.ListItems.Count > 0 Then
        listaRepro.ListItems.Remove listaRepro.ListItems.Count
    End If

End Sub

Private Sub cmdTiempo_Click()
    deshabilitarTodo
    listaTiempo.Enabled = True
    cmdTiempo.BackColor = vbGreen
    cmdTiempo.Enabled = False
    lblMsg = "ESPERANDO VARIACIÓN DE TIEMPO Nº " & listaTiempo.ListItems.Count + 1
    lblMsg.Visible = True
    cierra_comm
    inicia_comm

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 122 Then ' Tecla F11
        Frame1.Visible = Not Frame1.Visible
    End If
End Sub

Private Sub Form_Load()
    If CrearConexionGlobal = False Then
        MsgBox "Error al crear la conexión global. Contacte con mantenimiento.", vbCritical, App.Title
        End
    End If
    Me.Caption = Me.Caption & " (Host: " & ReadINI(App.Path + "\config.ini", "server", "ip") & " -> BD: " & database & ")"
    txtcom = ReadINI(App.Path + "\config.ini", "config", "COM")
    txtconf = ReadINI(App.Path + "\config.ini", "config", "SETTING")
    cabecera
'    cargar_lista
End Sub

Private Sub cabecera()
    With listaPrecargas.ColumnHeaders
        .Add , , "Nº", 400, lvwColumnLeft
        .Add , , "Valor", listaPrecargas.Width - 700, lvwColumnRight
    End With
    With listaMedidas.ColumnHeaders
        .Add , , "Punto", 400, lvwColumnLeft
        .Add , , "Medida", 400, lvwColumnCenter
        .Add , , "Valor", listaMedidas.Width - 1100, lvwColumnRight
    End With
    With listaBrazo.ColumnHeaders
        .Add , , "Nº", 400, lvwColumnLeft
        .Add , , "Valor", listaBrazo.Width - 700, lvwColumnRight
    End With
    With listaTiempo.ColumnHeaders
        .Add , , "Nº", 400, lvwColumnLeft
        .Add , , "Valor", listaTiempo.Width - 700, lvwColumnRight
    End With
    With listaRepro.ColumnHeaders
        .Add , , "Nº", 400, lvwColumnLeft
        .Add , , "Valor", listaRepro.Width - 700, lvwColumnRight
    End With
End Sub
Private Sub txtCodigo_Change()
    limpiarCampos
End Sub
Private Sub limpiarCampos()
    txtPuntos = ""
    txtMedidas = ""
    txtIdEquipo = ""
    lblMsg.Visible = False
    listaPrecargas.ListItems.Clear
End Sub
Private Sub limpiarlistas()
    listaPrecargas.ListItems.Clear
    listaMedidas.ListItems.Clear
    listaBrazo.ListItems.Clear
    listaTiempo.ListItems.Clear
    listaRepro.ListItems.Clear
End Sub

Private Sub txtCodigo_GotFocus()
    txtCodigo.SelStart = 0
    txtCodigo.SelLength = Len(txtCodigo)
End Sub

Private Sub txtCodigo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        If Trim(txtCodigo) <> "" Then
            Dim c As String
            c = "select a.id_equipo,b.VALOR as PUNTOS,c.VALOR as MEDIDAS " & _
                "  from equipos a,eq_campos_valores b, eq_campos_valores c " & _
                " Where a.ID_EQUIPO = b.EQUIPO_ID and b.CAMPO_ID = " & PARAMETRO_PUNTOS & _
                "   and a.ID_EQUIPO = c.EQUIPO_ID and c.CAMPO_ID = " & PARAMETRO_MEDIDAS & _
                "   and a.NUMERO_EQUIPO_CLIENTE = '" & Trim(txtCodigo) & "'"
            Dim rs As ADODB.Recordset
            Set rs = datos_bd(c)
            If rs.RecordCount > 0 Then
                Do
                    txtIdEquipo = rs(0)
                    txtPuntos = rs(1)
                    txtMedidas = rs(2)
                    
                    cmdPrecargas.Enabled = True
                    cmdPrecargas_Click
                    
                    rs.MoveNext
                Loop Until rs.EOF
            Else
                lblMsg = "NO EXISTE EL EQUIPO."
                lblMsg.Visible = True
            End If
            txtCodigo.SetFocus
        End If
    End If
End Sub
Private Sub inicia_comm()
   On Error GoTo inicia_comm_Error

    With MSComm1
        .InputLen = 0 ' El valor 0 hace que se lea todo
        .RThreshold = 5 ' al recibir uno o mas caracteres
        .SThreshold = 0 ' al enviar uno o mas caracteres
'        .InputMode = comInputModeText 'Los datos se dan en modo texto
'        .Handshaking = 0
        .CommPort = txtcom 'Paso 1: elijo el puerto COM 1
        .Settings = txtconf ' Vel. 1200, paridad odd, 8 bits
        .PortOpen = True 'Abro el puerto
    End With

   On Error GoTo 0
   Exit Sub

inicia_comm_Error:

    lblMsg = "Error " & Err.Number & " (" & Err.Description & ") in procedure inicia_comm of Formulario frmTBOX"
    lblMsg.Visible = True
End Sub
Private Sub cierra_comm()
    On Error Resume Next
    MSComm1.PortOpen = False 'Puede haber error si
End Sub
Private Sub MSComm1_OnComm()
    If Not IsNumeric(txtPuntos) Or Not IsNumeric(txtMedidas) Or Not IsNumeric(txtIdEquipo) Then
        Exit Sub
    End If
    If MSComm1.CommEvent = comEvReceive Then
        Dim ID As Integer
        Dim cad As String
        Dim punto As Integer
        ' Precargas
        If listaPrecargas.Enabled = True Then
            ID = listaPrecargas.ListItems.Count + 1
            With listaPrecargas.ListItems.Add(, , ID)
                 .SubItems(1) = Replace(MSComm1.Input, Chr(32) & Chr(13), "")
            End With
            cierra_comm
            inicia_comm
            lblMsg = "ESPERANDO PRECARGA Nº " & ID + 1
            lblMsg.Visible = True
            
            punto = listaPrecargas.ListItems.Count
            If punto >= 6 Then
                listaPrecargas.Enabled = False
            End If
        End If
        ' Medidas
        If listaMedidas.Enabled = True Then
            Dim medida As Integer
            If listaMedidas.ListItems.Count = 0 Then
                punto = 1
                medida = 1
            Else
                punto = listaMedidas.ListItems(listaMedidas.ListItems.Count).Text
                medida = CInt(listaMedidas.ListItems(listaMedidas.ListItems.Count).SubItems(1)) + 1
                If medida > CInt(txtMedidas) Then
                    If punto < CInt(txtPuntos) Then
                        punto = punto + 1
                        medida = 1
                    Else
                        lblMsg = "CAPTURA DE MEDIDAS FINALIZADA"
                        deshabilitarTodo
                        Exit Sub
                    End If
                End If
            End If
            With listaMedidas.ListItems.Add(, , punto)
                 .SubItems(1) = medida
                 .SubItems(2) = Replace(MSComm1.Input, Chr(32) & Chr(13), "")
            End With
            cierra_comm
            inicia_comm
            listaMedidas.ListItems(listaMedidas.ListItems.Count).EnsureVisible
            If medida > CInt(txtMedidas) Then
                If punto < CInt(txtPuntos) Then
                    punto = punto + 1
                    medida = 1
                End If
            End If
            lblMsg = "ESPERANDO PUNTO " & punto & ", MEDIDA " & medida
            
                punto = listaMedidas.ListItems(listaMedidas.ListItems.Count).Text
                medida = CInt(listaMedidas.ListItems(listaMedidas.ListItems.Count).SubItems(1)) + 1
                If medida > CInt(txtMedidas) Then
                    If punto < CInt(txtPuntos) Then
                        punto = punto + 1
                        medida = 1
                    Else
                        lblMsg = "CAPTURA DE MEDIDAS FINALIZADA"
                        deshabilitarTodo
                        Exit Sub
                    End If
                End If
            
            lblMsg.Visible = True
        End If
        ' Var.Brazo
        If listaBrazo.Enabled = True Then
            ID = listaBrazo.ListItems.Count + 1
            With listaBrazo.ListItems.Add(, , ID)
                 .SubItems(1) = Replace(MSComm1.Input, Chr(32) & Chr(13), "")
            End With
            cierra_comm
            inicia_comm
            lblMsg = "ESPERANDO VARIACIÓN DE BRAZO Nº " & ID + 1
            lblMsg.Visible = True
        End If
        ' Var.Tiempo
        If listaTiempo.Enabled = True Then
            ID = listaTiempo.ListItems.Count + 1
            With listaTiempo.ListItems.Add(, , ID)
                 .SubItems(1) = Replace(MSComm1.Input, Chr(32) & Chr(13), "")
            End With
            cierra_comm
            inicia_comm
            lblMsg = "ESPERANDO VARIACIÓN DE TIEMPO Nº " & ID + 1
            lblMsg.Visible = True
        End If
        ' Reproducibilidad
        If listaRepro.Enabled = True Then
            ID = listaRepro.ListItems.Count + 1
            With listaRepro.ListItems.Add(, , ID)
                 .SubItems(1) = Replace(MSComm1.Input, Chr(32) & Chr(13), "")
            End With
            cierra_comm
            inicia_comm
            lblMsg = "ESPERANDO REPRODUCIBILIDAD Nº " & ID + 1
            lblMsg.Visible = True
        End If
    End If
End Sub

Private Sub txtcom_Change()
    WriteINI App.Path + "\config.ini", "config", "COM", txtcom
End Sub

Private Sub txtconf_Change()
    WriteINI App.Path + "\config.ini", "config", "SETTING", txtconf
End Sub
