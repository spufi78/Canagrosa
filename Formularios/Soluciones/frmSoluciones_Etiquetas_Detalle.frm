VERSION 5.00
Begin VB.Form frmSoluciones_Etiquetas_Detalle 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Personalización de Etiquetas de Soluciones Preparadas"
   ClientHeight    =   11700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13650
   Icon            =   "frmSoluciones_Etiquetas_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11700
   ScaleWidth      =   13650
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFrases 
      Alignment       =   2  'Center
      Height          =   330
      Index           =   6
      Left            =   2115
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   675
      Width           =   11445
   End
   Begin VB.TextBox txtFrases 
      Alignment       =   2  'Center
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
      ForeColor       =   &H00FF0000&
      Height          =   420
      Index           =   5
      Left            =   2115
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   135
      Width           =   11400
   End
   Begin VB.TextBox txtFrases 
      Height          =   555
      Index           =   4
      Left            =   1170
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   11070
      Width           =   8925
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
      Height          =   1230
      Index           =   3
      Left            =   45
      TabIndex        =   17
      Top             =   9765
      Width           =   10050
      Begin VB.TextBox txtFrases 
         Height          =   915
         Index           =   3
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
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
      TabIndex        =   16
      Top             =   1080
      Width           =   3390
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   5
         Left            =   1800
         TabIndex        =   23
         Top             =   720
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   6
         Left            =   1800
         TabIndex        =   22
         Top             =   1890
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   7
         Left            =   1800
         TabIndex        =   21
         Top             =   3060
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   8
         Left            =   1800
         TabIndex        =   20
         Top             =   4230
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   10
         Top             =   5490
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   3
         Left            =   135
         TabIndex        =   9
         Top             =   4230
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   3060
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   1890
         Width           =   240
      End
      Begin VB.CheckBox chkPictograma 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check1"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   720
         Width           =   240
      End
      Begin VB.Image Image1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1050
         Index           =   8
         Left            =   2115
         Picture         =   "frmSoluciones_Etiquetas_Detalle.frx":08CA
         Stretch         =   -1  'True
         Top             =   3825
         Width           =   1050
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   7
         Left            =   2115
         Picture         =   "frmSoluciones_Etiquetas_Detalle.frx":3D56
         Stretch         =   -1  'True
         Top             =   2655
         Width           =   1050
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   6
         Left            =   2115
         Picture         =   "frmSoluciones_Etiquetas_Detalle.frx":6AA5
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   1050
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1095
         Index           =   5
         Left            =   2115
         Picture         =   "frmSoluciones_Etiquetas_Detalle.frx":7E3E
         Stretch         =   -1  'True
         Top             =   315
         Width           =   1095
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   4
         Left            =   450
         Picture         =   "frmSoluciones_Etiquetas_Detalle.frx":8AA4
         Stretch         =   -1  'True
         Top             =   3825
         Width           =   1050
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   3
         Left            =   450
         Picture         =   "frmSoluciones_Etiquetas_Detalle.frx":B4A9
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   1050
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   2
         Left            =   450
         Picture         =   "frmSoluciones_Etiquetas_Detalle.frx":E759
         Stretch         =   -1  'True
         Top             =   2655
         Width           =   1050
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   1
         Left            =   450
         Picture         =   "frmSoluciones_Etiquetas_Detalle.frx":17ABA
         Stretch         =   -1  'True
         Top             =   1485
         Width           =   1050
      End
      Begin VB.Image Image1 
         Enabled         =   0   'False
         Height          =   1050
         Index           =   0
         Left            =   450
         Picture         =   "frmSoluciones_Etiquetas_Detalle.frx":1A9AA
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
      TabIndex        =   15
      Top             =   6885
      Width           =   10050
      Begin VB.TextBox txtFrases 
         Height          =   2535
         Index           =   2
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
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
      TabIndex        =   14
      Top             =   3960
      Width           =   10050
      Begin VB.TextBox txtFrases 
         Height          =   2535
         Index           =   1
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
      TabIndex        =   13
      Top             =   1080
      Width           =   10050
      Begin VB.TextBox txtFrases 
         Height          =   2535
         Index           =   0
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   225
         Width           =   9780
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10710
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11385
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10710
      Width           =   1050
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Subtitulo Etiqueta"
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
      Left            =   225
      TabIndex        =   26
      Top             =   720
      Width           =   1860
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para poner en negrita, etiquete la palabra o frase de la siguiente forma: <negritra>PALABRA</negrita>"
      ForeColor       =   &H80000008&
      Height          =   690
      Left            =   10170
      TabIndex        =   24
      Top             =   8415
      Width           =   3390
   End
   Begin VB.Shape Shape1 
      Height          =   600
      Left            =   90
      Top             =   45
      Width           =   13470
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Titulo Etiqueta"
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
      Left            =   225
      TabIndex        =   19
      Top             =   225
      Width           =   1860
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pie Etiqueta"
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
      Left            =   45
      TabIndex        =   18
      Top             =   11250
      Width           =   1095
   End
End
Attribute VB_Name = "frmSoluciones_Etiquetas_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = True Then
      ' Alodine
      Dim oSe As New clsSoluciones_etiqueta
      Dim LOTE As Long
      Dim i As Integer
      With oSe
        .setDESCRIPCION = txtFrases(5)
        .setSUBTITULO = txtFrases(6)
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
        .setPIC7 = chkPictograma(6).Value
        .setPIC8 = chkPictograma(7).Value
        .setPIC9 = chkPictograma(8).Value
        If PK = 0 Then
            .Insertar
        Else
            .Modificar PK
        End If
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
    CARGAR
End Sub

Private Sub CARGAR()
    On Error GoTo fallo
    Dim oSe As New clsSoluciones_etiqueta
    With oSe
        If .Carga(PK) Then
            txtFrases(5) = .getDESCRIPCION
            txtFrases(6) = .getSUBTITULO
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
            chkPictograma(6).Value = .getPIC7
            chkPictograma(7).Value = .getPIC8
            chkPictograma(8).Value = .getPIC9
        End If
    End With
    Set oSe = Nothing
    Exit Sub
fallo:
    error_grave (Err.Description)
End Sub
Public Function validar() As Boolean
    validar = True
End Function
