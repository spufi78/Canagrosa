VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmAlodine_Etiquetas 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Personalización de Etiquetas"
   ClientHeight    =   10785
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11895
   Icon            =   "frmAlodine_Etiquetas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFrases 
      Height          =   330
      Index           =   4
      Left            =   1935
      ScrollBars      =   2  'Vertical
      TabIndex        =   17
      Top             =   10350
      Width           =   7980
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Componentes"
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
      Index           =   3
      Left            =   45
      TabIndex        =   15
      Top             =   9135
      Width           =   10050
      Begin VB.TextBox txtFrases 
         Height          =   870
         Index           =   3
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   225
         Width           =   9780
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Pictogramas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7305
      Left            =   10170
      TabIndex        =   14
      Top             =   450
      Width           =   1680
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   5
         Left            =   135
         TabIndex        =   21
         Top             =   6570
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   7
         Top             =   5490
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   6
         Top             =   4230
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   5
         Top             =   3060
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   4
         Top             =   1890
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   240
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1095
         Index           =   5
         Left            =   450
         Picture         =   "frmAlodine_Etiquetas.frx":08CA
         Stretch         =   -1  'True
         Top             =   6165
         Width           =   1095
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   4
         Left            =   450
         Picture         =   "frmAlodine_Etiquetas.frx":1530
         Stretch         =   -1  'True
         Top             =   3825
         Width           =   1050
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   3
         Left            =   450
         Picture         =   "frmAlodine_Etiquetas.frx":3F35
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1050
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   2
         Left            =   450
         Picture         =   "frmAlodine_Etiquetas.frx":71E5
         Stretch         =   -1  'True
         Top             =   2655
         Width           =   1050
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   1
         Left            =   450
         Picture         =   "frmAlodine_Etiquetas.frx":10546
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   1050
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   0
         Left            =   450
         Picture         =   "frmAlodine_Etiquetas.frx":13436
         Stretch         =   -1  'True
         Top             =   315
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frases Etiqueta Grande"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Index           =   2
      Left            =   45
      TabIndex        =   13
      Top             =   6255
      Width           =   10050
      Begin VB.TextBox txtFrases 
         Height          =   2535
         Index           =   2
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   225
         Width           =   9780
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frases Etiqueta Mediana"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Index           =   1
      Left            =   45
      TabIndex        =   12
      Top             =   3330
      Width           =   10050
      Begin VB.TextBox txtFrases 
         Height          =   2535
         Index           =   1
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   225
         Width           =   9780
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frases Etiqueta Pequeña"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2850
      Index           =   0
      Left            =   45
      TabIndex        =   11
      Top             =   450
      Width           =   10050
      Begin VB.TextBox txtFrases 
         Height          =   2535
         Index           =   0
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   225
         Width           =   9780
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   9405
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8460
      Width           =   1050
   End
   Begin MSDataListLib.DataCombo cmbIdioma 
      Height          =   330
      Left            =   9720
      TabIndex        =   20
      Top             =   45
      Width           =   2085
      _ExtentX        =   3678
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Idioma : "
      Height          =   240
      Left            =   9000
      TabIndex        =   19
      Top             =   90
      Width           =   825
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Advertencia/Peligro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   18
      Top             =   10395
      Width           =   1860
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Personalización Etiqueta : "
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
      Left            =   30
      TabIndex        =   10
      Top             =   45
      Width           =   8820
   End
End
Attribute VB_Name = "frmAlodine_Etiquetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmbIdioma_Change()
    CARGAR
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = True Then
      ' Alodine
      Dim oalodine_etiqueta As New clsAlodine_etiqueta
      Dim LOTE As Long
      Dim i As Integer
      With oalodine_etiqueta
        .setALODINE_ID = PK
        Dim oD As New clsDecodificadora
        oD.Carga_valor DECODIFICADORA.IDIOMAS, cmbIdioma.BoundText
        .setIDIOMA = oD.getPARAMETROS
        .setFRASES_PEQ = txtFrases(0)
        .setFRASES_MED = txtFrases(1)
        .setFRASES_GRA = txtFrases(2)
        .setCOMPONENTES = txtFrases(3)
        .setADVERTENCIA = txtFrases(4)
        .setPIC1 = chkPictograma(0).Value
        .setPIC2 = chkPictograma(1).Value
        .setPIC3 = chkPictograma(2).Value
        .setPIC4 = chkPictograma(3).Value
        .setPIC5 = chkPictograma(4).Value
        .setPIC6 = chkPictograma(5).Value
        .Insertar
      End With
      MsgBox "Datos almacenados correctamente.", vbOKOnly + vbInformation, App.Title
      Unload Me
    End If
    Exit Sub
fallo:
    error_grave (Err.Description)
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Dim oD As New clsDecodificadora
    oD.cargar_combo cmbIdioma, DECODIFICADORA.IDIOMAS
    cmbIdioma.BoundText = 0
    CARGAR
End Sub

Private Sub CARGAR()
    On Error GoTo fallo
    Dim oD As New clsDecodificadora
    oD.Carga_valor DECODIFICADORA.IDIOMAS, IIf(cmbIdioma.Text = "", "ES", cmbIdioma.BoundText)
    Dim oalodine As New clsAlodine
    oalodine.Carga PK
    lbltitulo = lbltitulo & oalodine.getDESCRIPCION
    Set oalodine = Nothing
    Dim oalodine_etiqueta As New clsAlodine_etiqueta
    limpiar
    With oalodine_etiqueta
        If .Carga(PK, oD.getPARAMETROS) Then
            txtFrases(0) = .getFRASES_PEQ
            txtFrases(1) = .getFRASES_MED
            txtFrases(2) = .getFRASES_GRA
            txtFrases(3) = .getCOMPONENTES
            txtFrases(4) = .getADVERTENCIA
            chkPictograma(0).Value = .getPIC1
            chkPictograma(1).Value = .getPIC2
            chkPictograma(2).Value = .getPIC3
            chkPictograma(3).Value = .getPIC4
            chkPictograma(4).Value = .getPIC5
            chkPictograma(5).Value = .getPIC6
        End If
    End With
    Set oalodine_etiqueta = Nothing
    Exit Sub
fallo:
    error_grave (Err.Description)
End Sub
Private Function limpiar()
    txtFrases(0) = ""
    txtFrases(1) = ""
    txtFrases(2) = ""
    txtFrases(3) = ""
    txtFrases(4) = ""
    chkPictograma(0).Value = Unchecked
    chkPictograma(1).Value = Unchecked
    chkPictograma(2).Value = Unchecked
    chkPictograma(3).Value = Unchecked
    chkPictograma(4).Value = Unchecked
    chkPictograma(5).Value = Unchecked
End Function
Public Function validar() As Boolean
    validar = True
    If cmbIdioma.Text = "" Then
        MsgBox "Por favor, indique el idioma.", vbCritical, App.Title
        validar = False
    End If
End Function
