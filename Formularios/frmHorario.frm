VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#39.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmHorario 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6270
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6270
   ScaleWidth      =   6465
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCambiar 
      Caption         =   "Cambiar"
      Height          =   285
      Left            =   5220
      TabIndex        =   15
      Top             =   5895
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tiempos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   135
      TabIndex        =   7
      Top             =   5040
      Width           =   6180
      Begin VB.TextBox txtJornada 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4635
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   270
         Width           =   1320
      End
      Begin VB.TextBox txtAusencia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   270
         Width           =   1050
      End
      Begin VB.TextBox txtComida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   675
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label3 
         Caption         =   "Jornada"
         Height          =   240
         Left            =   3915
         TabIndex        =   13
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Ausencia"
         Height          =   240
         Left            =   1800
         TabIndex        =   11
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Comida"
         Height          =   240
         Left            =   90
         TabIndex        =   8
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.TextBox txtmotivo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   1980
      TabIndex        =   5
      Text            =   "Indique motivo..."
      Top             =   855
      Visible         =   0   'False
      Width           =   4335
   End
   Begin XtremeSuiteControls.RadioButton opSubTipo 
      Height          =   465
      Index           =   0
      Left            =   180
      TabIndex        =   2
      Top             =   45
      Width           =   2760
      _Version        =   851970
      _ExtentX        =   4868
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "JORNADA LABORAL"
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
      Value           =   -1  'True
   End
   Begin XtremeSuiteControls.PushButton cmdEntrada 
      Height          =   960
      Left            =   135
      TabIndex        =   0
      Top             =   1260
      Width           =   2985
      _Version        =   851970
      _ExtentX        =   5265
      _ExtentY        =   1693
      _StockProps     =   79
      Caption         =   "ENTRADA"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Appearance      =   5
      Picture         =   "frmHorario.frx":0000
   End
   Begin XtremeSuiteControls.PushButton cmdSalida 
      Height          =   960
      Left            =   3285
      TabIndex        =   1
      Top             =   1260
      Width           =   3030
      _Version        =   851970
      _ExtentX        =   5345
      _ExtentY        =   1693
      _StockProps     =   79
      Caption         =   "SALIDA"
      BackColor       =   -2147483633
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Appearance      =   5
      Picture         =   "frmHorario.frx":08DA
   End
   Begin XtremeSuiteControls.RadioButton opSubTipo 
      Height          =   465
      Index           =   1
      Left            =   180
      TabIndex        =   3
      Top             =   405
      Width           =   1185
      _Version        =   851970
      _ExtentX        =   2090
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "COMIDA"
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
   Begin XtremeSuiteControls.RadioButton opSubTipo 
      Height          =   465
      Index           =   2
      Left            =   180
      TabIndex        =   4
      Top             =   765
      Width           =   1500
      _Version        =   851970
      _ExtentX        =   2646
      _ExtentY        =   820
      _StockProps     =   79
      Caption         =   "AUSENCIA"
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
   Begin MSComctlLib.ListView lista 
      Height          =   2730
      Left            =   135
      TabIndex        =   6
      Top             =   2250
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   4815
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin pryCombo.miCombo cmdUsuario 
      Height          =   330
      Left            =   135
      TabIndex        =   14
      Top             =   5850
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   582
   End
   Begin MSComCtl2.DTPicker fecha 
      Height          =   330
      Left            =   3870
      TabIndex        =   16
      Top             =   5850
      Width           =   1320
      _ExtentX        =   2328
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
      Format          =   16515073
      CurrentDate     =   38002
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   5760
      Picture         =   "frmHorario.frx":11B4
      Top             =   180
      Width           =   480
   End
End
Attribute VB_Name = "frmHorario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCambiar_Click()
    cargarEntradas
End Sub

Private Sub cmdEntrada_Click()
    Dim oH As New clsHorario
    With oH
        .setUSUARIO_ID = cmdUsuario.getPK_SALIDA
        .setTIPO = "E"
        .setMOTIVO = ""
        If opSubTipo(0).value = True Then
            .setSUBTIPO = 0
        ElseIf opSubTipo(1).value = True Then
            .setSUBTIPO = 1
        Else
            .setSUBTIPO = 2
            .setMOTIVO = txtmotivo
        End If
        .Insertar
        cargarEntradas
    End With
    Set oH = Nothing
End Sub

Private Sub cmdSalida_Click()
    Dim oH As New clsHorario
    With oH
        .setUSUARIO_ID = cmdUsuario.getPK_SALIDA
        .setTIPO = "S"
        .setMOTIVO = ""
        If opSubTipo(0).value = True Then
            .setSUBTIPO = 0
        ElseIf opSubTipo(1).value = True Then
            .setSUBTIPO = 1
        Else
            .setSUBTIPO = 2
            .setMOTIVO = txtmotivo
        End If
        .Insertar
        cargarEntradas
    End With
    Set oH = Nothing

End Sub

Private Sub cmdUsuario_change()
'    cargarEntradas
End Sub

Private Sub Form_Load()
    log Me.Name
    Me.Left = Screen.Width - Me.Width - frmMenu.ButtonBar.Width - 500
    Me.Top = 3500
    llenar_combo cmdUsuario, New clsUsuarios, 0, Me, " AND ANULADO = 0 "
    cmdUsuario.MostrarElemento USUARIO.getID_EMPLEADO
    fecha = Date
    cabecera
    cargarEntradas
    If UCase(USUARIO.getUSUARIO) <> "JULIO" Then
        cmdUsuario.Visible = False
        cmdCambiar.Visible = False
        fecha.Visible = False
    End If
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Fecha", 1700, lvwColumnLeft
        .Add , , "E/S", 1200, lvwColumnCenter
        .Add , , "Tipo", 1100, lvwColumnCenter
        .Add , , "Motivo", 1700, lvwColumnLeft
    End With
End Sub

Private Sub cargarEntradas()
    Dim rs As ADODB.RecordSet
    Dim oH As New clsHorario
    Dim subtipo As String
   On Error GoTo cargarEntradas_Error

    lista.ListItems.Clear
    Set rs = oH.Listado(cmdUsuario.getPK_SALIDA, fecha)
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                If rs(1) = "E" Then
                    .SubItems(1) = "Entrada"
                Else
                    .SubItems(1) = "Salida"
                End If
                Select Case rs(2)
                Case 0
                    subtipo = "Jornada"
                Case 1
                    subtipo = "Comida"
                Case 2
                    subtipo = "Ausencia"
                End Select
                .SubItems(2) = subtipo
                .SubItems(3) = rs(3)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    ' Bloqueo
    cmdEntrada.Enabled = False
    cmdSalida.Enabled = False
    opSubTipo(0).Enabled = False
    opSubTipo(1).Enabled = False
    opSubTipo(2).Enabled = False
    If lista.ListItems.Count = 0 Then
        cmdEntrada.Enabled = True
        opSubTipo(0).Enabled = True
        opSubTipo(0).value = True
    Else
        If lista.ListItems(1).SubItems(1) = "Entrada" Then
            cmdSalida.Enabled = True
'            If lista.ListItems(1).SubItems(2) = "Jornada" Then
                opSubTipo(0).Enabled = True
                opSubTipo(1).Enabled = True
                opSubTipo(2).Enabled = True
'            End If
'                If lista.ListItems(1).SubItems(2) = "Comida" Then
'                    opSubTipo(1).Enabled = True
'                    opSubTipo(1).value = True
'                End If
'                If lista.ListItems(1).SubItems(2) = "Ausencia" Then
'                    opSubTipo(2).Enabled = True
'                    opSubTipo(2).value = True
'                End If
'            End If
        Else
            cmdEntrada.Enabled = True
            If lista.ListItems(1).SubItems(2) = "Jornada" Then
                opSubTipo(0).Enabled = True
                opSubTipo(0).value = True
            End If
            If lista.ListItems(1).SubItems(2) = "Comida" Then
                opSubTipo(1).Enabled = True
                opSubTipo(1).value = True
            End If
            If lista.ListItems(1).SubItems(2) = "Ausencia" Then
                opSubTipo(2).Enabled = True
                opSubTipo(2).value = True
            End If
        End If
    End If
    txtComida = oH.Horas(cmdUsuario.getPK_SALIDA, fecha, 1)
    txtAusencia = oH.Horas(cmdUsuario.getPK_SALIDA, fecha, 2)
    Dim total As Date
    total = oH.Horas(cmdUsuario.getPK_SALIDA, fecha, 0)
    If txtComida <> "" And total > 0 Then
        total = total - CDate(txtComida)
    End If
    If txtAusencia <> "" And total > 0 Then
        total = total - CDate(txtAusencia)
    End If
    txtJornada = total
   On Error GoTo 0
   Exit Sub

cargarEntradas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargarEntradas of Formulario frmHorario"
End Sub

Private Sub opSubTipo_Click(Index As Integer)
    If Index = 2 Then
        txtmotivo.Visible = True
    Else
        txtmotivo.Visible = False
    End If
End Sub
