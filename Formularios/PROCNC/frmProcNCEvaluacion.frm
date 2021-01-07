VERSION 5.00
Begin VB.Form frmProcNCEvaluacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Evaluacion Final"
   ClientHeight    =   6000
   ClientLeft      =   1890
   ClientTop       =   3360
   ClientWidth     =   6780
   Icon            =   "frmProcNCEvaluacion.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   6780
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   5670
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5040
      Width           =   1020
   End
   Begin VB.Frame fraResultado 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Resultado"
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
      Height          =   1035
      Left            =   2160
      TabIndex        =   14
      Top             =   4950
      Visible         =   0   'False
      Width           =   3825
      Begin VB.OptionButton optEval_res_incidencia 
         BackColor       =   &H00C0C0C0&
         Caption         =   "INCIDENCIA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   450
         TabIndex        =   15
         Top             =   210
         Width           =   2925
      End
      Begin VB.OptionButton optEval_res_nc 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO CONFORMIDAD"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   450
         TabIndex        =   16
         Top             =   570
         Value           =   -1  'True
         Width           =   2925
      End
   End
   Begin VB.Frame fraEvidencias 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Evidencias"
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
      Height          =   2085
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6735
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   4
         Left            =   150
         TabIndex        =   2
         Top             =   330
         Width           =   6345
         Begin VB.OptionButton optEvidencias_si 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   1
            Left            =   5610
            TabIndex        =   4
            Top             =   60
            Width           =   315
         End
         Begin VB.OptionButton optEvidencias_no 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   1
            Left            =   6030
            TabIndex        =   3
            Top             =   60
            Width           =   315
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Las Acciones Correctivas han sido puestas en marcha en plazo?"
            Height          =   495
            Index           =   4
            Left            =   0
            TabIndex        =   25
            Top             =   60
            Width           =   4665
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   285
         Index           =   5
         Left            =   150
         TabIndex        =   5
         Top             =   630
         Width           =   6345
         Begin VB.OptionButton optEvidencias_no 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   2
            Left            =   6030
            TabIndex        =   7
            Top             =   30
            Width           =   315
         End
         Begin VB.OptionButton optEvidencias_si 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   2
            Left            =   5610
            TabIndex        =   6
            Top             =   30
            Width           =   315
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Son Efectivas?"
            Height          =   195
            Index           =   5
            Left            =   0
            TabIndex        =   24
            Top             =   60
            Width           =   1170
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   465
         Index           =   6
         Left            =   150
         TabIndex        =   8
         Top             =   960
         Width           =   6345
         Begin VB.OptionButton optEvidencias_no 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   3
            Left            =   6030
            TabIndex        =   10
            Top             =   30
            Width           =   315
         End
         Begin VB.OptionButton optEvidencias_si 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   3
            Left            =   5610
            TabIndex        =   9
            Top             =   30
            Width           =   315
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Disponemos de las evidencias de todas las acciones tomadas para la corrección de la incidencia ?"
            Height          =   585
            Index           =   6
            Left            =   0
            TabIndex        =   23
            Top             =   30
            Width           =   5205
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   495
         Index           =   7
         Left            =   150
         TabIndex        =   11
         Top             =   1440
         Width           =   6345
         Begin VB.OptionButton optEvidencias_no 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   4
            Left            =   6030
            TabIndex        =   13
            Top             =   30
            Width           =   315
         End
         Begin VB.OptionButton optEvidencias_si 
            BackColor       =   &H00C0C0C0&
            Height          =   255
            Index           =   4
            Left            =   5610
            TabIndex        =   12
            Top             =   30
            Width           =   255
         End
         Begin VB.Label Label9 
            BackColor       =   &H00C0C0C0&
            Caption         =   "¿Se han comunicado las modificaciones a todos los departamentos?"
            Height          =   495
            Index           =   7
            Left            =   30
            TabIndex        =   26
            Top             =   60
            Width           =   5025
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sí  -  No"
         Height          =   195
         Left            =   5790
         TabIndex        =   22
         Top             =   150
         Width           =   600
      End
   End
   Begin VB.Frame fraSolucion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Solución Aceptable"
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
      Height          =   975
      Left            =   45
      TabIndex        =   17
      Top             =   4905
      Visible         =   0   'False
      Width           =   3825
      Begin VB.OptionButton optEval_res_si 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   450
         TabIndex        =   18
         Top             =   240
         Value           =   -1  'True
         Width           =   795
      End
      Begin VB.OptionButton optEval_res_no 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   450
         TabIndex        =   19
         Top             =   570
         Width           =   885
      End
   End
   Begin VB.Frame fraObservaciones 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones a la Evaluación"
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
      Height          =   2865
      Left            =   30
      TabIndex        =   20
      Top             =   2130
      Width           =   6675
      Begin VB.TextBox txtObservaciones 
         Appearance      =   0  'Flat
         Height          =   2475
         Left            =   90
         MaxLength       =   65000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   0
         Top             =   270
         Width           =   6465
      End
   End
End
Attribute VB_Name = "frmProcNCEvaluacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private mvarobjProcNC As New clsProcNc
Private rs As ADODB.RecordSet
Private strSql As String
Private mvarblnEditable As Boolean

Private Sub cargar_datos()

    mvarobjProcNC.Carga PK

    With mvarobjProcNC
        If .getES_NO_CONFORMIDAD > -1 Then
            If .getES_NO_CONFORMIDAD = 0 Then optEval_res_incidencia.value = True
            If .getES_NO_CONFORMIDAD = 1 Then optEval_res_nc.value = True
        End If
        If .getES_SOLUCION_ACEPTABLE > -1 Then
            If .getES_SOLUCION_ACEPTABLE = 1 Then optEval_res_si.value = True
            If .getES_SOLUCION_ACEPTABLE = 0 Then optEval_res_no.value = True
        End If
        
        If .getEVIDENCIAS_EN_PLAZO > -1 Then
            If .getEVIDENCIAS_EN_PLAZO = 1 Then optEvidencias_si(1).value = True
            If .getEVIDENCIAS_EN_PLAZO = 0 Then optEvidencias_no(1).value = True
        End If
        
        If .getEVIDENCIAS_EFECTIVAS > -1 Then
            If .getEVIDENCIAS_EFECTIVAS = 1 Then optEvidencias_si(2).value = True
            If .getEVIDENCIAS_EFECTIVAS = 0 Then optEvidencias_no(2).value = True
        End If
        
        If .getEVIDENCIAS_EVIDENCIAS > -1 Then
            If .getEVIDENCIAS_EVIDENCIAS = 1 Then optEvidencias_si(3).value = True
            If .getEVIDENCIAS_EVIDENCIAS = 0 Then optEvidencias_no(3).value = True
        End If
        
        If .getEVIDENCIAS_COMUNICADO_MODIFICACIONES > -1 Then
            If .getEVIDENCIAS_COMUNICADO_MODIFICACIONES = 1 Then optEvidencias_si(4).value = True
            If .getEVIDENCIAS_COMUNICADO_MODIFICACIONES = 0 Then optEvidencias_no(4).value = True
        End If
                    
        txtObservaciones.Text = Trim(.getOBSERVACIONES_RESULTADO)
                    
    End With
    

End Sub




    
Public Property Get Editable() As Boolean

    Editable = mvarblnEditable

End Property

Public Property Let Editable(ByVal blnEditable As Boolean)

    mvarblnEditable = blnEditable

End Property

Private Function guardar_datos() As Boolean

On Error GoTo guardar_datos_Error

Dim es_no_conformidad As Integer
Dim es_solucion_aceptable As Integer
Dim evidencia_en_plazo As Integer
Dim evidencia_efectiva As Integer
Dim evidencia_evidencia As Integer
Dim evidencia_comunicado_modificaciones As Integer
        
        
    es_no_conformidad = -1
    es_solucion_aceptable = -1
    evidencia_en_plazo = -1
    evidencia_efectiva = -1
    evidencia_evidencia = -1
    evidencia_comunicado_modificaciones = -1
        
        With mvarobjProcNC
            If optEval_res_incidencia.value Then es_no_conformidad = 0
            If optEval_res_nc.value Then es_no_conformidad = 1
            If optEval_res_si.value Then es_solucion_aceptable = 1
            If optEval_res_no.value Then es_solucion_aceptable = 0

            If optEvidencias_si(1).value Then evidencia_en_plazo = 1
            If optEvidencias_no(1).value Then evidencia_en_plazo = 0
        
            If optEvidencias_si(2).value Then evidencia_efectiva = 1
            If optEvidencias_no(2).value Then evidencia_efectiva = 0
        
            If optEvidencias_si(3).value Then evidencia_evidencia = 1
            If optEvidencias_no(3).value Then evidencia_evidencia = 0
        
            If optEvidencias_si(4).value Then evidencia_comunicado_modificaciones = 1
            If optEvidencias_no(4).value Then evidencia_comunicado_modificaciones = 0
                            
        End With
        
        mvarobjProcNC.guardar_datos_evaluacion es_no_conformidad, es_solucion_aceptable, evidencia_en_plazo, evidencia_efectiva, evidencia_evidencia, evidencia_comunicado_modificaciones, txtObservaciones.Text
        
        guardar_datos = True
    
On Error GoTo 0
    Exit Function
guardar_datos_Error:
    guardar_datos = False
End Function

Private Sub opciones_edicion()

    fraEvidencias.Enabled = mvarblnEditable
    fraResultado.Enabled = mvarblnEditable
    fraSolucion.Enabled = mvarblnEditable
    fraObservaciones.Enabled = mvarblnEditable
        
End Sub

Private Sub cmdcancel_Click()

If Not mvarblnEditable Then Unload Me

If Not guardar_datos Then Exit Sub

Unload Me

End Sub

Private Sub Form_Activate()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_Load()

    cargar_botones Me
    
    cargar_datos
    
    opciones_edicion

End Sub


Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = 0 Then cmdcancel_Click
End Sub


