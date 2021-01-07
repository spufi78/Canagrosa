VERSION 5.00
Begin VB.Form frmCE_Materiales 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Nuevo Material/Pintura"
   ClientHeight    =   4155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9915
   Icon            =   "frmCE_Materiales.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   9915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rangos"
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
      Height          =   795
      Left            =   45
      TabIndex        =   12
      Top             =   3240
      Width           =   5370
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   2
         Left            =   945
         TabIndex        =   2
         Top             =   285
         Width           =   1440
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   3
         Left            =   3780
         TabIndex        =   3
         Top             =   270
         Width           =   1395
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   3780
         TabIndex        =   5
         Top             =   855
         Width           =   1395
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   4
         Left            =   945
         TabIndex        =   4
         Top             =   855
         Width           =   1440
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mínimo"
         Height          =   195
         Index           =   25
         Left            =   90
         TabIndex        =   16
         Top             =   375
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Maximo"
         Height          =   195
         Index           =   24
         Left            =   2880
         TabIndex        =   15
         Top             =   375
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Texto Max."
         Height          =   195
         Index           =   20
         Left            =   2880
         TabIndex        =   14
         Top             =   900
         Width           =   795
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Texto Mín."
         Height          =   195
         Index           =   21
         Left            =   90
         TabIndex        =   13
         Top             =   945
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7650
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3240
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8775
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3240
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Lote"
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
      Height          =   2580
      Left            =   45
      TabIndex        =   8
      Top             =   630
      Width           =   9810
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1860
         Index           =   1
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   585
         Width           =   8235
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1440
         TabIndex        =   0
         Top             =   225
         Width           =   8235
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Criterio"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   1440
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Material"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   300
         Width           =   555
      End
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo Material/Pintura"
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
      Height          =   285
      Left            =   90
      TabIndex        =   11
      Top             =   90
      Width           =   9210
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9405
      Picture         =   "frmCE_Materiales.frx":2AFA
      Top             =   45
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   11790
   End
End
Attribute VB_Name = "frmCE_Materiales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      Dim MATERIAL As Long
      Dim oCE_Mat As New clsCe_banos_materiales
      With oCE_Mat
        .setMATERIAL = txtDatos(0)
        .setCRITERIO = txtDatos(1)
        .setMINIMO = txtDatos(2)
        .setMAXIMO = txtDatos(3)
        .setMINIMO_TEXTO = txtDatos(4)
        .setMAXIMO_TEXTO = txtDatos(5)
      End With
      If PK = 0 Then
        If MsgBox("Va a introducir un nuevo Material. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            MATERIAL = oCE_Mat.Insertar
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el Material. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            oCE_Mat.Modificar (PK)
        Else
            Exit Sub
        End If
      End If
      If PK = 0 Then
          MsgBox "El Material se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "El Material se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmCE_Materiales"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    If PK <> 0 Then
        lbltitulo.Caption = "Modificación de Material/Pintura"
        Me.Caption = lbltitulo
        cargar_ficha
    End If
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Or Index = 3 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub cargar_ficha()
    Dim oCE_Mat As New clsCe_banos_materiales
    With oCE_Mat
        If .Carga(PK) = True Then
            txtDatos(0) = .getMATERIAL
            txtDatos(1) = .getCRITERIO
            txtDatos(2) = .getMINIMO
            txtDatos(3) = .getMAXIMO
            txtDatos(4) = .getMINIMO_TEXTO
            txtDatos(5) = .getMAXIMO_TEXTO
        End If
    End With
End Sub
Public Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle una descripción al material.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(2)) <> "" Then
        If Not IsNumeric(txtDatos(2)) Then
            MsgBox "El valor mínimo debe ser numérico.", vbInformation, App.Title
            txtDatos(2).SetFocus
            validar = False
            Exit Function
        End If
    End If
    If Trim(txtDatos(3)) <> "" Then
        If Not IsNumeric(txtDatos(3)) Then
            MsgBox "El valor máximo debe ser numérico.", vbInformation, App.Title
            txtDatos(3).SetFocus
            validar = False
            Exit Function
        End If
    End If
End Function
