VERSION 5.00
Begin VB.Form frmOferta_Requisitos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Requisitos de la Oferta"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14505
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   14505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame frmUsuario 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   9585
      TabIndex        =   58
      Top             =   5940
      Visible         =   0   'False
      Width           =   4830
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1125
         TabIndex        =   62
         Top             =   585
         Width           =   3615
      End
      Begin VB.TextBox txtUsuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1125
         TabIndex        =   61
         Top             =   180
         Width           =   3615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   60
         Top             =   675
         Width           =   885
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   59
         Top             =   270
         Width           =   885
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frmTitulo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "ENSAYOS/TESTS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5730
      Left            =   45
      TabIndex        =   12
      Top             =   90
      Width           =   14460
      Begin VB.CheckBox chkSI 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI/Yes"
         Height          =   195
         Index           =   8
         Left            =   11745
         TabIndex        =   56
         Top             =   5130
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtDatos1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   8
         Left            =   3510
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   55
         Top             =   4995
         Width           =   8175
      End
      Begin VB.CheckBox chkNO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   195
         Index           =   8
         Left            =   12780
         TabIndex        =   54
         Top             =   5130
         Value           =   1  'Checked
         Width           =   600
      End
      Begin VB.CheckBox chkNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.A."
         Height          =   195
         Index           =   8
         Left            =   13500
         TabIndex        =   53
         Top             =   5130
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkSI 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI/Yes"
         Height          =   195
         Index           =   0
         Left            =   11745
         TabIndex        =   44
         Top             =   450
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtDatos1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   0
         Left            =   3510
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   43
         Top             =   315
         Width           =   8175
      End
      Begin VB.CheckBox chkNO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   195
         Index           =   0
         Left            =   12780
         TabIndex        =   42
         Top             =   450
         Value           =   1  'Checked
         Width           =   600
      End
      Begin VB.CheckBox chkNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.A."
         Height          =   195
         Index           =   0
         Left            =   13500
         TabIndex        =   41
         Top             =   450
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkSI 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI/Yes"
         Height          =   195
         Index           =   1
         Left            =   11745
         TabIndex        =   40
         Top             =   1035
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtDatos1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   1
         Left            =   3510
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   39
         Top             =   900
         Width           =   8175
      End
      Begin VB.CheckBox chkNO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   195
         Index           =   1
         Left            =   12780
         TabIndex        =   38
         Top             =   1035
         Value           =   1  'Checked
         Width           =   600
      End
      Begin VB.CheckBox chkNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.A."
         Height          =   195
         Index           =   1
         Left            =   13500
         TabIndex        =   37
         Top             =   1035
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkSI 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI/Yes"
         Height          =   195
         Index           =   2
         Left            =   11745
         TabIndex        =   36
         Top             =   1620
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtDatos1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   2
         Left            =   3510
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   35
         Top             =   1485
         Width           =   8175
      End
      Begin VB.CheckBox chkNO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   195
         Index           =   2
         Left            =   12780
         TabIndex        =   34
         Top             =   1620
         Value           =   1  'Checked
         Width           =   600
      End
      Begin VB.CheckBox chkNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.A."
         Height          =   195
         Index           =   2
         Left            =   13500
         TabIndex        =   33
         Top             =   1620
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkSI 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI/Yes"
         Height          =   195
         Index           =   3
         Left            =   11745
         TabIndex        =   32
         Top             =   2205
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtDatos1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   3
         Left            =   3510
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   31
         Top             =   2070
         Width           =   8175
      End
      Begin VB.CheckBox chkNO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   195
         Index           =   3
         Left            =   12780
         TabIndex        =   30
         Top             =   2205
         Value           =   1  'Checked
         Width           =   600
      End
      Begin VB.CheckBox chkNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.A."
         Height          =   195
         Index           =   3
         Left            =   13500
         TabIndex        =   29
         Top             =   2205
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkSI 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI/Yes"
         Height          =   195
         Index           =   4
         Left            =   11745
         TabIndex        =   28
         Top             =   2790
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtDatos1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   4
         Left            =   3510
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   2655
         Width           =   8175
      End
      Begin VB.CheckBox chkNO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   195
         Index           =   4
         Left            =   12780
         TabIndex        =   26
         Top             =   2790
         Value           =   1  'Checked
         Width           =   600
      End
      Begin VB.CheckBox chkNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.A."
         Height          =   195
         Index           =   4
         Left            =   13500
         TabIndex        =   25
         Top             =   2790
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkSI 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI/Yes"
         Height          =   195
         Index           =   5
         Left            =   11745
         TabIndex        =   24
         Top             =   3375
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtDatos1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   5
         Left            =   3510
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   3240
         Width           =   8175
      End
      Begin VB.CheckBox chkNO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   195
         Index           =   5
         Left            =   12780
         TabIndex        =   22
         Top             =   3375
         Value           =   1  'Checked
         Width           =   600
      End
      Begin VB.CheckBox chkNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.A."
         Height          =   195
         Index           =   5
         Left            =   13500
         TabIndex        =   21
         Top             =   3375
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkSI 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI/Yes"
         Height          =   195
         Index           =   6
         Left            =   11745
         TabIndex        =   20
         Top             =   3960
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtDatos1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   6
         Left            =   3510
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   3825
         Width           =   8175
      End
      Begin VB.CheckBox chkNO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   195
         Index           =   6
         Left            =   12780
         TabIndex        =   18
         Top             =   3960
         Value           =   1  'Checked
         Width           =   600
      End
      Begin VB.CheckBox chkNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.A."
         Height          =   195
         Index           =   6
         Left            =   13500
         TabIndex        =   17
         Top             =   3960
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.CheckBox chkSI 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SI/Yes"
         Height          =   195
         Index           =   7
         Left            =   11745
         TabIndex        =   16
         Top             =   4545
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.TextBox txtDatos1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   540
         Index           =   7
         Left            =   3510
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   4410
         Width           =   8175
      End
      Begin VB.CheckBox chkNO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   195
         Index           =   7
         Left            =   12780
         TabIndex        =   14
         Top             =   4545
         Value           =   1  'Checked
         Width           =   600
      End
      Begin VB.CheckBox chkNA 
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.A."
         Height          =   195
         Index           =   7
         Left            =   13500
         TabIndex        =   13
         Top             =   4545
         Value           =   1  'Checked
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros/Others"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   57
         Top             =   5160
         Width           =   3390
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Norma de proceso/ Process Specification"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   52
         Top             =   435
         Width           =   3315
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Norma de ensayo/Test  Specification"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   51
         Top             =   975
         Width           =   3315
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Criterios de aceptación( p.e. rangos/Acceptance criteria (f.e.range)"
         Height          =   465
         Index           =   2
         Left            =   135
         TabIndex        =   50
         Top             =   1515
         Width           =   3330
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Matriz/Matrix"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   49
         Top             =   2235
         Width           =   3345
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº y tipo de probetas/ Test specimens number and type"
         Height          =   480
         Index           =   4
         Left            =   135
         TabIndex        =   48
         Top             =   2730
         Width           =   3330
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cantidad de muestra (baños)/ Sample amount (baths)"
         Height          =   435
         Index           =   5
         Left            =   135
         TabIndex        =   47
         Top             =   3360
         Width           =   3315
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Toma de muestra/ Sampling"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   46
         Top             =   3990
         Width           =   3345
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Plazo/ Lead Time"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   45
         Top             =   4575
         Width           =   3330
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Aprobaciones/Approvals"
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
      Left            =   4815
      TabIndex        =   6
      Top             =   5940
      Width           =   4695
      Begin VB.CheckBox chkAprobacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No especificado/Not specified"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   11
         Top             =   1395
         Width           =   3255
      End
      Begin VB.CheckBox chkAprobacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bombardier"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   10
         Top             =   1125
         Width           =   1500
      End
      Begin VB.CheckBox chkAprobacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Airbus D&S"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   9
         Top             =   315
         Width           =   1500
      End
      Begin VB.CheckBox chkAprobacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Airbus"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   8
         Top             =   585
         Width           =   1500
      End
      Begin VB.CheckBox chkAprobacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Boeing"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   855
         Width           =   1500
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Acreditaciones/Accreditations"
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
      TabIndex        =   2
      Top             =   5940
      Width           =   4695
      Begin VB.CheckBox chkAcreditacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No especificado/Not specified"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   5
         Top             =   855
         Width           =   3480
      End
      Begin VB.CheckBox chkAcreditacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENAC"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   4
         Top             =   585
         Width           =   1500
      End
      Begin VB.CheckBox chkAcreditacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nadcap   "
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   3
         Top             =   315
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   13185
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7110
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   915
      Left            =   11835
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7110
      Width           =   1275
   End
End
Attribute VB_Name = "frmOferta_Requisitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public TIPO_OFERTA As Integer
Public PK As Long

Private Sub chkNA_Click(Index As Integer)
    If chkNA(Index).Value = Checked Then
        chkNO(Index).Value = Unchecked
        chkSI(Index).Value = Unchecked
    End If

End Sub

Private Sub chkNO_Click(Index As Integer)
    If chkNO(Index).Value = Checked Then
        chkSI(Index).Value = Unchecked
        chkNA(Index).Value = Unchecked
    End If

End Sub

Private Sub chkSI_Click(Index As Integer)
    If chkSI(Index).Value = Checked Then
        chkNO(Index).Value = Unchecked
        chkNA(Index).Value = Unchecked
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    Dim oOR As New clsOfertas_requisitos
   On Error GoTo cmdok_Click_Error

   On Error GoTo cmdok_Click_Error
    ' Recorrer textos
    Dim texto As String
    Dim sn As String
    Dim acreditaciones As String
    Dim aprobaciones As String
    For i = 0 To 8
        If texto <> "" Then
            texto = texto & "###"
        End If
        texto = texto & IIf(txtDatos1(i) = "", " ", txtDatos1(i))
        If chkSI(i).Value = Checked Then
            sn = sn & "1"
        ElseIf chkNO(i).Value = Checked Then
            sn = sn & "2"
        ElseIf chkNA(i).Value = Checked Then
            sn = sn & "3"
        Else
            sn = sn & "0"
        End If
    Next
        ' Acreditaciones
        For j = 0 To 2
            If chkAcreditacion(j).Value = Checked Then
               acreditaciones = acreditaciones & "1"
            Else
               acreditaciones = acreditaciones & "0"
            End If
        Next
        ' Aprobaciones
        For j = 0 To 4
            If chkAprobacion(j).Value = Checked Then
               aprobaciones = aprobaciones & "1"
            Else
               aprobaciones = aprobaciones & "0"
            End If
        Next
    With oOR
        .setOFERTA_ID = PK
        .setTEXTO = texto
        .setSN = sn
        .setACREDITACIONES = acreditaciones
        .setAPROBACIONES = aprobaciones
        .Insertar
    End With
    MsgBox "Requisitos almacenados correctamente.", vbInformation, App.Title
    Unload Me
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmOferta_Requisitos"

   On Error GoTo 0
   Exit Sub
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    configurarVentana
    cargarDatos
End Sub
Private Sub cargarDatos()
    Dim oOR As New clsOfertas_requisitos
    If oOR.Carga(PK) Then
        frmUsuario.visible = True
        With oOR
            Dim lista() As String
            Dim i As Integer
            lista = Split(.getTEXTO, "###")
            ' TEXTO
            For i = LBound(lista) To UBound(lista)
                txtDatos1(i) = lista(i)
                ' SN
                If Mid(.getSN, i + 1, 1) = "1" Then
                    chkSI(i).Value = Checked
                ElseIf Mid(.getSN, i + 1, 1) = "2" Then
                    chkNO(i).Value = Checked
                ElseIf Mid(.getSN, i + 1, 1) = "3" Then
                    chkNA(i).Value = Checked
                End If
            Next
            ' ACREDITACIONES
            For i = 0 To Len(.getACREDITACIONES) - 1
                If Mid(.getACREDITACIONES, i + 1, 1) = "1" Then
                    chkAcreditacion(i).Value = Checked
                Else
                    chkAcreditacion(i).Value = Unchecked
                End If
            Next
            ' APROBACIONES
            For i = 0 To Len(.getAPROBACIONES) - 1
                If Mid(.getAPROBACIONES, i + 1, 1) = "1" Then
                    chkAprobacion(i).Value = Checked
                Else
                    chkAprobacion(i).Value = Unchecked
                End If
            Next
            Dim oUsuario As New clsUsuarios
            oUsuario.cargar .getUSUARIO_ID
            txtUsuario = oUsuario.getUSUARIO
            txtFecha = .getTS
        End With
    End If
    Set oOR = Nothing
    
End Sub
Private Sub configurarVentana()
   On Error GoTo configurarVentana_Error

    Select Case TIPO_OFERTA
    Case 1
        frmTitulo.Caption = "ENSAYOS/TESTS"
    Case 2
        frmTitulo.Caption = "CALIBRACIONES/CALIBRATIONS"
    Case 3
        frmTitulo.Caption = "SUMINISTROS/SUPPLY"
    Case 4
        frmTitulo.Caption = "OTROS SERVICIOS/OTHERS SERVICES"
    End Select
    Dim oDeco As New clsDecodificadora
    oDeco.Carga_valor DECODIFICADORA_OFERTAS_REQUISITOS, CLng(TIPO_OFERTA)
    Dim lista() As String
    Dim i As Integer
    For i = 0 To 7
        chkSI(i).Value = Unchecked
        chkNO(i).Value = Unchecked
        chkNA(i).Value = Unchecked
        mostrar i, False
    Next
    chkSI(8).Value = Unchecked
    chkNO(8).Value = Unchecked
    chkNA(8).Value = Unchecked
    lista = Split(oDeco.getDESCRIPCION, ";")
    For i = LBound(lista) To UBound(lista)
        lblCampos(i).Caption = lista(i)
        mostrar i, True
    Next
    ' Acreditaciones/Aprobaciones
    lista = Split(oDeco.getPARAMETROS, ";")
    For i = 0 To 2
        chkAcreditacion(i).visible = Mid(lista(0), i + 1, 1)
    Next
    For i = 0 To 4
        chkAprobacion(i).visible = Mid(lista(1), i + 1, 1)
    Next

   On Error GoTo 0
   Exit Sub

configurarVentana_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure configurarVentana of Formulario frmOferta_Requisitos"
End Sub
Private Sub mostrar(linea As Integer, visible As Boolean)
    lblCampos(linea).visible = visible
    txtDatos1(linea).visible = visible
    chkSI(linea).visible = visible
    chkNO(linea).visible = visible
    chkNA(linea).visible = visible
End Sub

