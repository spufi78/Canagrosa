VERSION 5.00
Begin VB.Form frmCA_Valoracion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuestionario sobre procedimiento / Questionnarie about procedures"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10020
   ControlBox      =   0   'False
   Icon            =   "frmCA_Valoracion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10020
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   8775
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   5175
      Width           =   1140
   End
   Begin VB.TextBox txtObservaciones 
      Appearance      =   0  'Flat
      Height          =   1500
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   25
      Top             =   4545
      Width           =   8565
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
      Height          =   870
      Left            =   8775
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   5175
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Valoración"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   45
      TabIndex        =   7
      Top             =   2160
      Width           =   9915
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   5490
         TabIndex        =   21
         Top             =   1350
         Width           =   4290
         Begin VB.OptionButton opt3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   3465
            TabIndex        =   29
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton opt3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2025
            TabIndex        =   23
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton opt3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   585
            TabIndex        =   22
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   465
         Left            =   5490
         TabIndex        =   15
         Top             =   900
         Width           =   4290
         Begin VB.OptionButton opt2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   3
            Left            =   3465
            TabIndex        =   28
            Top             =   225
            Width           =   195
         End
         Begin VB.OptionButton opt2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   2025
            TabIndex        =   20
            Top             =   225
            Width           =   195
         End
         Begin VB.OptionButton opt2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   1
            Left            =   585
            TabIndex        =   19
            Top             =   225
            Width           =   195
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   420
         Left            =   5490
         TabIndex        =   14
         Top             =   540
         Width           =   4290
         Begin VB.OptionButton opt1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   3
            Left            =   3465
            TabIndex        =   18
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton opt1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   2
            Left            =   2025
            TabIndex        =   17
            Top             =   180
            Width           =   195
         End
         Begin VB.OptionButton opt1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H0000C000&
            Height          =   195
            Index           =   1
            Left            =   585
            TabIndex        =   16
            Top             =   180
            Width           =   195
         End
      End
      Begin VB.Shape Shape1 
         Height          =   1410
         Left            =   5445
         Top             =   495
         Width           =   4380
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 3. Orden del procedimiento / Order of the procedure"
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
         Height          =   240
         Left            =   135
         TabIndex        =   13
         Top             =   1485
         Width           =   5145
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 2. Estructura adecuada / Adecuate structure"
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
         Height          =   240
         Left            =   135
         TabIndex        =   12
         Top             =   1080
         Width           =   5145
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " 1. Claridad en la redacción / Clarity in the wording"
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
         Height          =   240
         Left            =   135
         TabIndex        =   11
         Top             =   675
         Width           =   5145
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Mal/bad"
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
         Left            =   8370
         TabIndex        =   10
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Regular/normal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   240
         Left            =   6930
         TabIndex        =   9
         Top             =   225
         Width           =   1410
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Bien/good"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   240
         Left            =   5490
         TabIndex        =   8
         Top             =   225
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos generales"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1590
      Left            =   45
      TabIndex        =   0
      Top             =   540
      Width           =   9915
      Begin VB.TextBox txtdescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   3195
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   315
         Width           =   6495
      End
      Begin VB.TextBox txtAsistente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   720
         Width           =   8295
      End
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   315
         Width           =   1770
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   330
         Left            =   1395
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1125
         Width           =   1770
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   240
         Left            =   225
         TabIndex        =   5
         Top             =   1170
         Width           =   915
      End
      Begin VB.Label lblAsistente 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Asistente"
         Height          =   240
         Left            =   225
         TabIndex        =   3
         Top             =   765
         Width           =   915
      End
      Begin VB.Label lblCurso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Documento"
         Height          =   240
         Left            =   225
         TabIndex        =   1
         Top             =   360
         Width           =   1005
      End
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cuestionario sobre procedimiento / Questionnarie about procedures"
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
      Height          =   300
      Left            =   225
      TabIndex        =   27
      Top             =   90
      Width           =   8145
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   10020
   End
   Begin VB.Label Label19 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comentarios"
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
      TabIndex        =   26
      Top             =   4275
      Width           =   1320
   End
End
Attribute VB_Name = "frmCA_Valoracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_DOCUMENTO_ID As Long
Public PK_USUARIO_ID As Long

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Dim oCA_Documento As New clsCa_documentos
    With oCA_Documento
        .Carga PK_DOCUMENTO_ID
        txtCodigo = .getCODIGO
        txtdescripcion = .getNOMBRE
        txtAsistente = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
        txtFecha.Text = Date
    End With
    Set oCA_Documento = Nothing
    Dim oCDV As New clsCa_documentos_val
    If oCDV.Carga(PK_DOCUMENTO_ID, PK_USUARIO_ID) Then
        If oCDV.getCUMPLIMENTADO = 0 Then
            cmdok.Visible = True
            cmdSalir.Visible = False
        Else
            cmdok.Visible = False
            cmdSalir.Visible = True
            carga_formulario
        End If
    End If
    Set oCDV = Nothing
End Sub
Private Sub carga_formulario()
    Dim oCDV As New clsCa_documentos_val
   On Error GoTo carga_formulario_Error
    Frame2.Enabled = False
    txtObservaciones.Enabled = False
    With oCDV
        .Carga PK_DOCUMENTO_ID, PK_USUARIO_ID
        txtObservaciones.Text = Trim(.getCOMENTARIOS)
        opt1.Item(.getP1).value = True
        opt2.Item(.getP2).value = True
        opt3.Item(.getP3).value = True
    End With

   On Error GoTo 0
   Exit Sub

carga_formulario_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure carga_formulario of Formulario frmCA_Valoracion"
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo cmdok_Click_Error
    Dim i As Integer
    If comprueba_formulario Then
        Dim oCDV As New clsCa_documentos_val
        With oCDV
            For i = 1 To 3
                If opt1(i).value = True Then
                    .setP1 = i
                End If
                If opt2(i).value = True Then
                    .setP2 = i
                End If
                If opt3(i).value = True Then
                    .setP3 = i
                End If
            Next
            .setCOMENTARIOS = Trim(txtObservaciones.Text)
            .Modificar PK_DOCUMENTO_ID
        End With
        MsgBox "El cuestionario se ha guardado con éxito.", vbOKOnly + vbInformation
        Unload Me
    Else
        MsgBox "Debe completar todos los valores", vbExclamation, App.Title
    End If
    Exit Sub
cmdok_Click_Error:
  MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmFormacion_Evalucion Procedure cmdok_Click"
End Sub
Private Function comprueba_formulario() As Boolean
     If opt1(1).value = False And opt1(2).value = False And opt1(3).value = False Then
        comprueba_formulario = False
        Exit Function
     End If
     If opt2(1).value = False And opt2(2).value = False And opt2(3).value = False Then
        comprueba_formulario = False
        Exit Function
     End If
     If opt3(1).value = False And opt3(2).value = False And opt3(3).value = False Then
        comprueba_formulario = False
        Exit Function
     End If
     comprueba_formulario = True
End Function
