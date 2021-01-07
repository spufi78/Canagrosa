VERSION 5.00
Begin VB.Form frmREX_evaluacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Evaluación de certificado de material de Referencia"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Certificado"
      Height          =   870
      Left            =   90
      Picture         =   "frmREX_evaluacion.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   52
      Top             =   8775
      Width           =   1005
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7425
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8775
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8505
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8775
      Width           =   1050
   End
   Begin VB.TextBox texto 
      Appearance      =   0  'Flat
      Height          =   240
      Index           =   17
      Left            =   1260
      MaxLength       =   100
      TabIndex        =   19
      Top             =   7380
      Width           =   8205
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Evaluación del certificado según el PNTA002"
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
      Height          =   7035
      Left            =   45
      TabIndex        =   23
      Top             =   1710
      Width           =   9510
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   17
         Left            =   5445
         TabIndex        =   51
         Top             =   5985
         Width           =   1590
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   15
         Left            =   4005
         MaxLength       =   100
         TabIndex        =   18
         Top             =   4725
         Width           =   5415
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   14
         Left            =   4230
         MaxLength       =   100
         TabIndex        =   17
         Top             =   4455
         Width           =   5190
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   13
         Left            =   4005
         MaxLength       =   100
         TabIndex        =   16
         Top             =   4140
         Width           =   5415
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   12
         Left            =   4005
         MaxLength       =   100
         TabIndex        =   15
         Top             =   3870
         Width           =   5415
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   11
         Left            =   4005
         MaxLength       =   100
         TabIndex        =   14
         Top             =   3600
         Width           =   5415
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   10
         Left            =   4365
         MaxLength       =   100
         TabIndex        =   13
         Top             =   3285
         Width           =   5055
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   9
         Left            =   4365
         MaxLength       =   100
         TabIndex        =   12
         Top             =   3015
         Width           =   5055
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   8
         Left            =   4005
         MaxLength       =   100
         TabIndex        =   11
         Top             =   2745
         Width           =   5415
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   7
         Left            =   4770
         MaxLength       =   100
         TabIndex        =   10
         Top             =   2475
         Width           =   4650
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   6
         Left            =   4005
         MaxLength       =   100
         TabIndex        =   9
         Top             =   2205
         Width           =   5415
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   5
         Left            =   4320
         MaxLength       =   100
         TabIndex        =   8
         Top             =   1935
         Width           =   5100
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   4
         Left            =   4005
         MaxLength       =   100
         TabIndex        =   7
         Top             =   1665
         Width           =   5415
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   3
         Left            =   4005
         MaxLength       =   100
         TabIndex        =   6
         Top             =   1395
         Width           =   5415
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   2
         Left            =   4005
         MaxLength       =   100
         TabIndex        =   5
         Top             =   1125
         Width           =   5415
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   240
         Index           =   1
         Left            =   4005
         MaxLength       =   100
         TabIndex        =   4
         Top             =   855
         Width           =   5415
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "RESULTADO (Marque si es Conforme) "
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   16
         Left            =   315
         TabIndex        =   41
         Top             =   4995
         Width           =   4200
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Trazabilidad : "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   15
         Left            =   315
         TabIndex        =   40
         Top             =   4725
         Width           =   3300
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estabilidad y Período de validez, si es apropiado : "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   14
         Left            =   315
         TabIndex        =   39
         Top             =   4455
         Width           =   4200
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad : "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   13
         Left            =   315
         TabIndex        =   38
         Top             =   4185
         Width           =   3795
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Presentación o estado físico : "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   12
         Left            =   315
         TabIndex        =   37
         Top             =   3915
         Width           =   3795
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "La matriz (si procede) : "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   11
         Left            =   315
         TabIndex        =   36
         Top             =   3645
         Width           =   3795
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concentración o valor numérico de la característica certificada y su incertidumbre : "
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   10
         Left            =   315
         TabIndex        =   35
         Top             =   3240
         Width           =   4380
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nivel de homogeneidad : "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   9
         Left            =   315
         TabIndex        =   34
         Top             =   3015
         Width           =   3795
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Situación peligrosa : "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   8
         Left            =   315
         TabIndex        =   33
         Top             =   2745
         Width           =   3795
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Instrucciones para el uso correcto del MRC, si proceden : "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   315
         TabIndex        =   32
         Top             =   2475
         Width           =   4380
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Utilización prevista : "
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   315
         TabIndex        =   31
         Top             =   2205
         Width           =   3795
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código y número del lote del material de referencia : "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   315
         TabIndex        =   30
         Top             =   1935
         Width           =   4110
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de certificación : "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   315
         TabIndex        =   29
         Top             =   1665
         Width           =   3795
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabricante y código de fabricación del material : "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   315
         TabIndex        =   28
         Top             =   1395
         Width           =   3795
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre del material : "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   315
         TabIndex        =   27
         Top             =   1125
         Width           =   3795
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Título del documento :"
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   315
         TabIndex        =   26
         Top             =   855
         Width           =   3795
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   240
         Index           =   0
         Left            =   4005
         MaxLength       =   100
         TabIndex        =   3
         Top             =   585
         Width           =   5415
      End
      Begin VB.CheckBox op 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre y dirección del organismo que certifica : "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   315
         TabIndex        =   25
         Top             =   585
         Width           =   3795
      End
      Begin VB.Label lblresultado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         Caption         =   "NO CONFORME"
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
         Height          =   375
         Index           =   2
         Left            =   3060
         TabIndex        =   46
         Top             =   6570
         Width           =   3255
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "3. Evaluación Final del Material y su Certificado: CONFORME/NO CONFORME"
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
         Index           =   2
         Left            =   90
         TabIndex        =   45
         Top             =   6300
         Width           =   9240
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "¿Es conforme la propiedad certificada a este uso? Marque si es Conforme"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   17
         Left            =   135
         TabIndex        =   44
         Top             =   5985
         Width           =   5640
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Uso previsto:"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   43
         Top             =   5670
         Width           =   1590
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "2. En cuanto a la propiedad certificada y su uso en el laboratorio:"
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
         Index           =   1
         Left            =   135
         TabIndex        =   42
         Top             =   5310
         Width           =   9240
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0E0FF&
         Caption         =   "1. En cuanto a su corrección según lo especificado por la guía ISO 31:2000, en el certificado aparece:"
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
         Index           =   0
         Left            =   180
         TabIndex        =   24
         Top             =   270
         Width           =   9240
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1050
      Left            =   45
      TabIndex        =   22
      Top             =   585
      Width           =   9510
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   21
         Left            =   2745
         MaxLength       =   100
         TabIndex        =   2
         Top             =   720
         Width           =   6675
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   20
         Left            =   2745
         MaxLength       =   100
         TabIndex        =   1
         Top             =   450
         Width           =   6675
      End
      Begin VB.TextBox texto 
         Appearance      =   0  'Flat
         Height          =   240
         Index           =   19
         Left            =   2745
         MaxLength       =   100
         TabIndex        =   0
         Top             =   180
         Width           =   6675
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de Identificación en inventario : "
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   49
         Top             =   720
         Width           =   2445
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entidad que emite el certificado :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   48
         Top             =   450
         Width           =   2445
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de certificado evaluado :"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   47
         Top             =   180
         Width           =   2265
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Evaluación de certificado de material de Referencia/Material de referencia Certificado"
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
      Height          =   480
      Left            =   90
      TabIndex        =   50
      Top             =   30
      Width           =   7605
      WordWrap        =   -1  'True
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   9540
   End
End
Attribute VB_Name = "frmREX_evaluacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BOTE_EX_ID As Long
Public consulta As Boolean
Private Sub cmdcancel_Click()
'    If consulta = True Then
'        Unload Me
'    Else
        If MsgBox("¿Esta seguro de salir sin informar la evaluación del material de referencia?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Unload Me
        End If
'    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
        Dim oEvaluacion As New clsRex_botes_certificados
        With oEvaluacion
            .setBOTE_EX_ID = BOTE_EX_ID
            .setC01_NUMERO_CERTIFICADO = texto(19)
            .setC02_ENTIDAD = texto(20)
            .setC03_INVENTARIO = texto(21)
            .setC04_ORGANISMO = texto(0)
            .setC05_TITULO_DOCUMENTO = texto(1)
            .setC06_NOMBRE_MATERIAL = texto(2)
            .setC07_FABRICANTE = texto(3)
            .setC08_FECHA_CERTIFICACION = texto(4)
            .setC09_CODIGO_LOTE = texto(5)
            .setC10_UTILIZACION_PREVISTA = texto(6)
            .setC11_INSTRUCCIONES_USO = texto(7)
            .setC12_SITUACION_PELIGROSA = texto(8)
            .setC13_NIVEL_HOMOGENEIDAD = texto(9)
            .setC14_CONCENTRACION = texto(10)
            .setC15_MATRIZ = texto(11)
            .setC16_PRESENTACION = texto(12)
            .setC17_CANTIDAD = texto(13)
            .setC18_ESTABILIDAD = texto(14)
            .setC19_TRAZABILIDAD = texto(15)
            If op(16).value = Checked Then
                .setC20_RESULTADO = 1
            Else
                .setC20_RESULTADO = 0
            End If
            .setC21_USO_PREVISTO = texto(17)
            If op(17).value = Checked Then
                .setC22_CONFORME_PROPIEDAD = 1
            Else
                .setC22_CONFORME_PROPIEDAD = 0
            End If
            .setC23_TECNICO_RESPONSABLE = usuario.getID_EMPLEADO
            .setC24_FECHA_EVALUACION = Format(Date, "dd-mm-yyyy")
            If .Insertar = 0 Then
                If .Modificar(BOTE_EX_ID) = True Then
                    MsgBox "Los datos de la certificación se han modificado correctamente.", vbInformation, App.Title
                    Unload Me
                End If
            Else
                MsgBox "Los datos de la certificación se han insertado correctamente.", vbInformation, App.Title
                Unload Me
            End If
        End With
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmREX_evaluacion"
End Sub

Private Sub Command1_Click()
        Dim consulta As String
        ' Copiar la firma del usuario a la imagen temporal
'        Dim oUsuario As New clsUsuarios
'        Dim oEvaluacion As New clsRex_botes_certificados
   On Error GoTo Command1_Click_Error

'        oEvaluacion.Carga BOTE_EX_ID
'        oUsuario.CARGAR oEvaluacion.getC23_TECNICO_RESPONSABLE
'        If oUsuario.getFIRMA <> "" Then
'            If Dir(oUsuario.getFIRMA) <> "" Then
'                FileCopy oUsuario.getFIRMA, "c:\imagen.bmp"
'            End If
'        End If
'        consulta = "SELECT * FROM REX_BOTES_CERTIFICADOS WHERE BOTE_EX_ID = " & BOTE_EX_ID
        frmReport.iniciar
        frmReport.criterio = "{REX_BOTES_CERTIFICADOS.BOTE_EX_ID} = " & BOTE_EX_ID
        frmReport.informe = "\REX\rptcertificado"
        frmReport.consulta = consulta
        Dim destino As String
        destino = App.Path & "\certificado.pdf"
        frmReport.pdf = destino
        frmReport.imprimir = False
        frmReport.generar
        frmReport.Visible = False
        If Dir(destino) <> "" Then
            R = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
        End If

   On Error GoTo 0
   Exit Sub

Command1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command1_Click of Formulario frmREX_evaluacion"
End Sub

Private Sub Form_Activate()
'    If consulta = False Then
'        texto(21) = BOTE_EX_ID
'        Command1.Enabled = False
'    Else
'        cargar_certificado
'    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    If BOTE_EX_ID = 0 Then
        texto(21) = BOTE_EX_ID
        lblresultado(2).Caption = "NO REALIZADO"
    Else
        cargar_certificado
    End If
End Sub

Private Sub op_Click(Index As Integer)
'    If consulta = False Then
    On Error Resume Next
        If Index <> 16 And Index <> 17 Then
            If op(Index).value = Checked Then
                texto(Index).Enabled = True
                texto(Index).BackColor = &HC0FFFF
                texto(Index).SetFocus
            Else
                texto(Index).Enabled = False
                texto(Index) = ""
            End If
        End If
        comprobar_resultado
'    End If
End Sub
Private Sub texto_LostFocus(Index As Integer)
    texto(Index).BackColor = vbWhite
End Sub

Public Function validar() As Boolean
    validar = True
    If Trim(texto(19)) = "" Or Trim(texto(20)) = "" Or Trim(texto(21)) = "" Then
        MsgBox "Rellene todos los datos marcados en azul.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If
    If Trim(texto(0)) = "" Or Trim(texto(2)) = "" Or _
       Trim(texto(3)) = "" Or Trim(texto(4)) = "" Or _
       Trim(texto(5)) = "" Or Trim(texto(10)) = "" Or _
       Trim(texto(14)) = "" Or Trim(texto(15)) = "" Or _
       Trim(texto(17)) = "" Then
        MsgBox "Rellene todos los datos marcados en azul.", vbExclamation, App.Title
        validar = False
        Exit Function
    End If

End Function

Public Sub comprobar_resultado()
    If op(0).value = Checked And op(2).value = Checked And _
       op(3).value = Checked And op(4).value = Checked And _
       op(5).value = Checked And op(10).value = Checked And _
       op(14).value = Checked And op(15).value = Checked Then
        op(16).value = Checked
    Else
        op(16).value = Unchecked
    End If
    If texto(19).Text = "" Then
        lblresultado(2).Caption = "NO REALIZADO"
    Else
        If op(16).value = Checked And op(17).value = Checked Then
            lblresultado(2).Caption = "CONFORME"
            lblresultado(2).BackColor = &HC0FFC0
        Else
            lblresultado(2).Caption = "NO CONFORME"
            lblresultado(2).BackColor = &H8080FF
        End If
    End If
End Sub

Public Sub cargar_certificado()
'    Frame1.Enabled = False
'    Frame2.Enabled = False
'    cmdok.Enabled = False
    Dim oEvaluacion As New clsRex_botes_certificados
    With oEvaluacion
        .Carga BOTE_EX_ID
        If .getC01_NUMERO_CERTIFICADO = "" Then
            lblresultado(2).Caption = "NO REALIZADO"
        End If
        texto(19) = .getC01_NUMERO_CERTIFICADO
        texto(20) = .getC02_ENTIDAD
        texto(21) = .getC03_INVENTARIO
        If .getC04_ORGANISMO <> "" Then
            texto(0) = .getC04_ORGANISMO
            op(0).value = Checked
        End If
        If .getC05_TITULO_DOCUMENTO <> "" Then
            texto(1) = .getC05_TITULO_DOCUMENTO
            op(1).value = Checked
        End If
        If .getC06_NOMBRE_MATERIAL <> "" Then
            texto(2) = .getC06_NOMBRE_MATERIAL
            op(2).value = Checked
        End If
        If .getC07_FABRICANTE <> "" Then
            texto(3) = .getC07_FABRICANTE
            op(3).value = Checked
        End If
        If .getC08_FECHA_CERTIFICACION <> "" Then
            texto(4) = .getC08_FECHA_CERTIFICACION
            op(4).value = Checked
        End If
        If .getC09_CODIGO_LOTE <> "" Then
            texto(5) = .getC09_CODIGO_LOTE
            op(5).value = Checked
        End If
        If .getC10_UTILIZACION_PREVISTA <> "" Then
            texto(6) = .getC10_UTILIZACION_PREVISTA
            op(6).value = Checked
        End If
        If .getC11_INSTRUCCIONES_USO <> "" Then
            texto(7) = .getC11_INSTRUCCIONES_USO
            op(7).value = Checked
        End If
        If .getC12_SITUACION_PELIGROSA <> "" Then
            texto(8) = .getC12_SITUACION_PELIGROSA
            op(8).value = Checked
        End If
        If .getC13_NIVEL_HOMOGENEIDAD <> "" Then
            texto(9) = .getC13_NIVEL_HOMOGENEIDAD
            op(9).value = Checked
        End If
        If .getC14_CONCENTRACION <> "" Then
            texto(10) = .getC14_CONCENTRACION
            op(10).value = Checked
        End If
        If .getC15_MATRIZ <> "" Then
            texto(11) = .getC15_MATRIZ
            op(11).value = Checked
        End If
        If .getC16_PRESENTACION <> "" Then
            texto(12) = .getC16_PRESENTACION
            op(12).value = Checked
        End If
        If .getC17_CANTIDAD <> "" Then
            texto(13) = .getC17_CANTIDAD
            op(13).value = Checked
        End If
        If .getC18_ESTABILIDAD <> "" Then
            texto(14) = .getC18_ESTABILIDAD
            op(14).value = Checked
        End If
        If .getC19_TRAZABILIDAD <> "" Then
            texto(15) = .getC19_TRAZABILIDAD
            op(15).value = Checked
        End If
        If .getC21_USO_PREVISTO <> "" Then
            texto(17) = .getC21_USO_PREVISTO
        End If
        If .getC22_CONFORME_PROPIEDAD = 1 Then
            op(17).value = Checked
        End If
        comprobar_resultado
    End With
End Sub
