VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmSoluciones_Etiqueta 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Imprimir Etiquetas de Soluciones"
   ClientHeight    =   4620
   ClientLeft      =   4590
   ClientTop       =   2880
   ClientWidth     =   9045
   Icon            =   "frmSoluciones_Etiqueta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   9045
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1005
      Left            =   45
      TabIndex        =   12
      Top             =   2385
      Width           =   8745
      Begin VB.TextBox txtSp 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   6795
         TabIndex        =   18
         Top             =   360
         Width           =   1500
      End
      Begin MSComCtl2.DTPicker fechaFabricacion 
         Height          =   330
         Left            =   1620
         TabIndex        =   13
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   60096513
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaCaducidad 
         Height          =   330
         Left            =   4635
         TabIndex        =   15
         Top             =   360
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   582
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   60096513
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SP : "
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   6345
         TabIndex        =   17
         Top             =   405
         Width           =   420
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Caducidad"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   3150
         TabIndex        =   16
         Top             =   405
         Width           =   1545
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Fabricación"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   135
         TabIndex        =   14
         Top             =   405
         Width           =   1635
      End
   End
   Begin VB.OptionButton optSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pequeña"
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   0
      Left            =   1080
      TabIndex        =   7
      Top             =   1080
      Width           =   972
   End
   Begin VB.OptionButton optSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mediana"
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   1
      Left            =   2250
      TabIndex        =   6
      Top             =   1080
      Value           =   -1  'True
      Width           =   972
   End
   Begin VB.OptionButton optSize 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Grande"
      ForeColor       =   &H80000008&
      Height          =   192
      Index           =   2
      Left            =   3375
      TabIndex        =   5
      Top             =   1080
      Width           =   972
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   915
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3555
      Width           =   1275
   End
   Begin MSComCtl2.UpDown UpDown1 
      Height          =   330
      Left            =   2640
      TabIndex        =   3
      Top             =   1695
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   582
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtDatos(0)"
      BuddyDispid     =   196611
      BuddyIndex      =   0
      OrigLeft        =   2820
      OrigTop         =   960
      OrigRight       =   3060
      OrigBottom      =   1395
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   1965
      TabIndex        =   2
      Text            =   "1"
      Top             =   1695
      Width           =   675
   End
   Begin pryCombo.miCombo cmbEtiqueta 
      Height          =   345
      Left            =   810
      TabIndex        =   10
      Top             =   540
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   609
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   915
      Left            =   6300
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3555
      Width           =   1275
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Etiqueta"
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
      Index           =   1
      Left            =   60
      TabIndex        =   11
      Top             =   585
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Etiquetas de Soluciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   60
      Width           =   8850
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tamaño"
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
      Index           =   6
      Left            =   90
      TabIndex        =   8
      Top             =   1080
      Width           =   690
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Número de etiquetas"
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
      Index           =   2
      Left            =   90
      TabIndex        =   4
      Top             =   1770
      Width           =   1770
   End
End
Attribute VB_Name = "frmSoluciones_Etiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdImprimir_Click()
   On Error GoTo cmdImprimir_Click_Error

    If cmbEtiqueta.getTEXTO = "" Then
        MsgBox "Indique el tipo de etiqueta.", vbCritical, App.Title
        Exit Sub
    End If
    ' Almacenar datos etiqueta
    Dim oMe As New clsMuestras_soluciones
    With oMe
        .setMUESTRA_ID = PK
        .setFECHA_CADUCIDAD = fechaCaducidad
        .setFECHA_FABRICACION = fechaFabricacion
        .setETIQUETA_ID = cmbEtiqueta.getPK_SALIDA
        .setTAMANO = 0
        If optSize(1).Value = True Then
            .setTAMANO = 1
        ElseIf optSize(2).Value = True Then
            .setTAMANO = 2
        End If
        .setSP = txtSp
        .Insertar
    End With
    ' Imprimir la etiqueta
    Dim oSe As New clsSoluciones_etiqueta
    oSe.ImprimirEtiquetas PK, txtDatos(0)
    Set oset = Nothing

   On Error GoTo 0
   Exit Sub

cmdImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmSoluciones_Etiqueta"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
    Case 27
        cmdSalir_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    llenar_combo cmbEtiqueta, New clsSoluciones_etiqueta, 0, frmSoluciones_Etiquetas_Detalle, ""
    fechaCaducidad = Date
    fechaFabricacion = Date
    txtSp = ""
    CARGAR
End Sub

Private Sub CARGAR()
    Dim oMe As New clsMuestras_soluciones
    With oMe
        If .Carga(PK) Then
            cmbEtiqueta.MostrarElemento .getETIQUETA_ID
            fechaCaducidad = .getFECHA_CADUCIDAD
            fechaFabricacion = .getFECHA_FABRICACION
            optSize(.getTAMANO).Value = True
            txtSp = .getSP
        End If
    End With
End Sub
