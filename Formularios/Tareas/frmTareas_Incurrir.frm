VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmTareas_Incurrir 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parte de Horas"
   ClientHeight    =   9375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10590
   ControlBox      =   0   'False
   Icon            =   "frmTareas_Incurrir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9375
   ScaleWidth      =   10590
   WindowState     =   1  'Minimized
   Begin VB.CommandButton cmdIncurridos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Incurridos"
      Height          =   855
      Left            =   2205
      Picture         =   "frmTareas_Incurrir.frx":2AFA
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8460
      Width           =   1080
   End
   Begin VB.CommandButton cmdTarea 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tareas"
      Height          =   855
      Left            =   1125
      Picture         =   "frmTareas_Incurrir.frx":33C4
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8460
      Width           =   1080
   End
   Begin VB.CommandButton cmdTipos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tipos de Tareas"
      Height          =   855
      Left            =   45
      Picture         =   "frmTareas_Incurrir.frx":3C8E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8460
      Width           =   1080
   End
   Begin VB.Frame frameUsuario 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usuario"
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
      Height          =   735
      Left            =   45
      TabIndex        =   22
      Top             =   675
      Width           =   10500
      Begin MSDataListLib.DataCombo cmbusuario 
         Height          =   315
         Left            =   1260
         TabIndex        =   23
         Top             =   270
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   24
         Top             =   315
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar Tarea"
      Height          =   855
      Left            =   8325
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8460
      Width           =   1080
   End
   Begin VB.CommandButton cmdMinimizar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Minimizar"
      Height          =   855
      Left            =   9405
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8460
      Width           =   1080
   End
   Begin VB.Frame frameTarea 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle de la Tarea"
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
      Height          =   4695
      Left            =   45
      TabIndex        =   15
      Top             =   1440
      Width           =   10485
      Begin VB.CheckBox chkOp 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FESTIVO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Index           =   4
         Left            =   2700
         TabIndex        =   39
         Top             =   360
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   1260
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1440
         Width           =   9000
      End
      Begin VB.CheckBox chkFacturable 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   135
         TabIndex        =   6
         Top             =   2565
         Width           =   1320
      End
      Begin VB.Frame frmFacturable 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   1275
         Left            =   90
         TabIndex        =   27
         Top             =   2880
         Width           =   10275
         Begin VB.CheckBox chkOp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Desplazamiento"
            Height          =   240
            Index           =   3
            Left            =   2520
            TabIndex        =   34
            Top             =   900
            Width           =   1680
         End
         Begin VB.CheckBox chkOp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Estancia"
            Height          =   240
            Index           =   2
            Left            =   2520
            TabIndex        =   33
            Top             =   675
            Width           =   1680
         End
         Begin VB.CheckBox chkOp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Urgencia (Pedido Fuera de Hora)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Index           =   1
            Left            =   4545
            TabIndex        =   32
            Top             =   630
            Width           =   3300
         End
         Begin VB.CheckBox chkOp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dieta Completa"
            Height          =   240
            Index           =   5
            Left            =   225
            TabIndex        =   31
            Top             =   855
            Width           =   1680
         End
         Begin VB.CheckBox chkOp 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Dieta Media"
            Height          =   240
            Index           =   0
            Left            =   225
            TabIndex        =   30
            Top             =   630
            Width           =   1680
         End
         Begin MSDataListLib.DataCombo cmbTipoHora 
            Height          =   315
            Left            =   5580
            TabIndex        =   7
            Top             =   225
            Width           =   3465
            _ExtentX        =   6112
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            ListField       =   ""
            BoundColumn     =   ""
            Text            =   ""
            Object.DataMember      =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSComCtl2.DTPicker hinicio 
            Height          =   330
            Left            =   1170
            TabIndex        =   36
            Top             =   225
            Width           =   1110
            _ExtentX        =   1958
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
            Format          =   60358658
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker hfin 
            Height          =   330
            Left            =   3285
            TabIndex        =   38
            Top             =   225
            Width           =   1110
            _ExtentX        =   1958
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
            Format          =   60358658
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hora Fin"
            Height          =   195
            Index           =   9
            Left            =   2520
            TabIndex        =   37
            Top             =   270
            Width           =   600
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hora Inicio"
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   35
            Top             =   270
            Width           =   765
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo Hora"
            Height          =   195
            Index           =   6
            Left            =   4590
            TabIndex        =   28
            Top             =   270
            Width           =   705
         End
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   9180
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   26
         Top             =   315
         Width           =   1080
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   6795
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   315
         Width           =   1080
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   600
         Index           =   1
         Left            =   1260
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   1845
         Width           =   9000
      End
      Begin MSDataListLib.DataCombo cmbtipo 
         Height          =   315
         Left            =   1260
         TabIndex        =   1
         Top             =   720
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbTarea 
         Height          =   315
         Left            =   1260
         TabIndex        =   2
         Top             =   1080
         Width           =   9000
         _ExtentX        =   15875
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   1260
         TabIndex        =   0
         Top             =   315
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   60358657
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin XtremeSuiteControls.PushButton cmdOk 
         Height          =   435
         Left            =   8010
         TabIndex        =   40
         Top             =   4185
         Width           =   2355
         _Version        =   851970
         _ExtentX        =   4154
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmTareas_Incurrir.frx":4558
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1530
         TabIndex        =   41
         Top             =   2520
         Width           =   8790
         _ExtentX        =   15505
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Referencia"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   29
         Top             =   1530
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Total horas día"
         Height          =   195
         Index           =   4
         Left            =   7965
         TabIndex        =   25
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   20
         Top             =   405
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tiempo (horas)"
         Height          =   195
         Index           =   1
         Left            =   5535
         TabIndex        =   19
         Top             =   360
         Width           =   1050
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Left            =   135
         TabIndex        =   18
         Top             =   2025
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tarea"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   17
         Top             =   1170
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   16
         Top             =   810
         Width           =   315
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   2250
      Left            =   45
      TabIndex        =   21
      Top             =   6165
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   3969
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Inserte las horas de dedicación a los distintos módulos"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   315
      Width           =   3825
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9900
      Picture         =   "frmTareas_Incurrir.frx":ADBA
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Horas Trabajadas"
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
      Height          =   240
      Index           =   0
      Left            =   90
      TabIndex        =   13
      Top             =   45
      Width           =   3060
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   -450
      Top             =   -45
      Width           =   11115
   End
End
Attribute VB_Name = "frmTareas_Incurrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub chkFacturable_Click()
'    If chkFacturable.value = Checked Then
'        frmFacturable.Enabled = True
'    Else
'        frmFacturable.Enabled = False
'        cmbClientes.Limpiar
'        cmbTipoHora.Text = ""
'    End If
End Sub

Private Sub cmbTipo_change()
    If cmbTipo.Text <> "" Then
     If IsNumeric(cmbTipo.BoundText) Then
        Dim oTarea As New clsTareas
        Set cmbTarea.RowSource = oTarea.Listado_Combo_Filtro(CLng(cmbTipo.BoundText), fecha)
        cmbTarea.ListField = "descripcion" 'campo que veo
        cmbTarea.DataField = "descripcion" 'campo asociado
        cmbTarea.BoundColumn = "id_tarea" 'lo que realmente envia
        Set oTarea = Nothing
     End If
    End If
End Sub
Private Sub cmdEliminar_Click()
   On Error GoTo cmdEliminar_Click_Error

    If lista.ListItems.Count > 0 Then
        Dim oTarea As New clsTareas_incurridos
        oTarea.Eliminar lista.ListItems(lista.selectedItem.Index).SubItems(3)
        cargar_lista
    End If

   On Error GoTo 0
   Exit Sub

cmdEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminar_Click of Formulario frmTareas_Incurrir"
End Sub

Private Sub cmdIncurridos_Click()
    frmTareas_Incurridos.Show
End Sub

Private Sub cmdMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar Then
        Dim oTarea As New clsTareas_incurridos
        With oTarea
            .setUSUARIO_ID = cmbUsuario.BoundText
            '.setFECHA = Format(Date, "yyyy-mm-dd")
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setTAREA_ID = cmbTarea.BoundText
            .setHORAS = Replace(txtDatos(0), ",", ".")
            .setHORA_INICIO = Format(hinicio.value, "hh:mm:ss")
            .setHORA_FIN = Format(hfin.value, "hh:mm:ss")
            .setREFERENCIA = txtDatos(3)
            .setOBSERVACIONES = txtDatos(1)
            If chkFacturable.value = Checked Then
                .setFACTURABLE = 1
                .setCLIENTE_ID = cmbclientes.getPK_SALIDA
            Else
                .setFACTURABLE = 0
                .setCLIENTE_ID = 0
            End If
            If cmbTipoHora.Text = "" Then
                .setTIPO_HORA = 0
            Else
                .setTIPO_HORA = cmbTipoHora.BoundText
            End If
            .setDOC_ID = 0
            .setDIETA_MEDIA = chkOp(0).value
            .setDIETA_COMPLETA = chkOp(5).value
            .setESTANCIA = chkOp(2).value
            .setDESPLAZAMIENTO = chkOp(3).value
            .setURGENCIA = chkOp(1).value
            .Insertar
        End With
        ' Enviar correo tarea facturable
        If chkFacturable.value = Checked Then
           On Error Resume Next
           Dim ASUNTO As String
           Dim DETALLE As String
           ASUNTO = "Alta tarea facturable en Geslab"
           DETALLE = "El usuario " & cmbUsuario.Text & " ha dado de alta una nueva tarea facturable en geslab." & vbNewLine & vbNewLine
           DETALLE = DETALLE & " Fecha : " & Format(fecha, "dd-mm-yyyy") & vbNewLine
           DETALLE = DETALLE & " Tarea : " & cmbTipo.Text & vbNewLine
           DETALLE = DETALLE & " Tiempo : " & txtDatos(0) & " h." & vbNewLine
           DETALLE = DETALLE & " Cliente : " & cmbclientes.getTEXTO & vbNewLine
           DETALLE = DETALLE & " Tipo de Hora : " & cmbTipoHora.Text & vbNewLine
           DETALLE = DETALLE & " Referencia : " & txtDatos(3) & vbNewLine & vbNewLine
           DETALLE = DETALLE & " Observaciones : " & txtDatos(1) & vbNewLine & vbNewLine
           DETALLE = DETALLE & " Dieta Media : " & IIf(chkOp(0).value = Checked, "Si", "No") & vbNewLine
           DETALLE = DETALLE & " Dieta Completa : " & IIf(chkOp(5).value = Checked, "Si", "No") & vbNewLine
           DETALLE = DETALLE & " Estancia : " & IIf(chkOp(2).value = Checked, "Si", "No") & vbNewLine
           DETALLE = DETALLE & " Desplazamiento : " & IIf(chkOp(3).value = Checked, "Si", "No") & vbNewLine
           DETALLE = DETALLE & " Urgencia : " & IIf(chkOp(1).value = Checked, "Si", "No") & vbNewLine & vbNewLine
           DETALLE = DETALLE & " Fecha y hora de Alta de la tarea : " & Date & " " & Time
           ret = Enviar_Mail_CDO("salvador.alarcon@canagrosa.com", ASUNTO, DETALLE, vbNullString)
           ret = Enviar_Mail_CDO(BUZON_CORREO_LOG, ASUNTO, DETALLE, vbNullString)
           On Error GoTo cmdok_Click_Error
        End If
        cmbTarea.Text = ""
        txtDatos(0) = ""
        txtDatos(1) = ""
        txtDatos(3) = ""
        chkFacturable.value = Unchecked
        cargar_lista
        cmbTarea.SetFocus
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmTareas_Incurrir"
End Sub

Private Sub cmdTarea_Click()
    frmTareas_Listado.Show
End Sub

Private Sub cmdTipos_Click()
    Dim oform As New frmDecodificadora
    oform.CODIGO = DECODIFICADORA.TAREAS_MODULOS
    oform.Show
End Sub

Private Sub fecha_Change()
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 50
    Me.top = 50
    cargar_botones Me
    cabecera
    cargar_combos
    fecha = Date
    cmbUsuario.BoundText = usuario.getID_EMPLEADO
    cargar_lista
    If usuario.getPER_INCURRIDOS = True Then
        cmbUsuario.Enabled = True
    End If
End Sub

Private Sub hfin_Change()
    calcular_horas
End Sub
Private Sub hinicio_Change()
    hfin = hinicio
    calcular_horas
 '   txtDatos(0) = Format(TimeValue(Format(hfin.value, "hh:mm:ss")) - TimeValue(Format(hinicio.value, "hh:mm:ss")), "hh:mm:ss")
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe indicar un numero de horas.", vbExclamation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    Else
        If IsNumeric(txtDatos(0)) = False Then
            MsgBox "El número de horas debe ser numérico.", vbExclamation, App.Title
            txtDatos(0).SetFocus
            validar = False
            Exit Function
        End If
    End If
    If cmbTipo.Text = "" Then
        MsgBox "Debe indicar un tipo de tarea.", vbExclamation, App.Title
        cmbTipo.SetFocus
        validar = False
        Exit Function
    End If
    If cmbTarea.Text = "" Then
        MsgBox "Debe indicar la tarea.", vbExclamation, App.Title
        cmbTarea.SetFocus
        validar = False
        Exit Function
    End If
    If chkFacturable.value = Checked Then
        If cmbclientes.getTEXTO = "" Then
            MsgBox "Debe indicar un cliente para la facturación.", vbExclamation, App.Title
            cmbclientes.SetFocus
            validar = False
            Exit Function
        End If
        If cmbTipoHora.Text = "" Then
            MsgBox "Debe indicar un Tipo de Hora para la facturación.", vbExclamation, App.Title
            cmbTipoHora.SetFocus
            validar = False
            Exit Function
        End If
    End If
End Function
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, DECODIFICADORA.TAREAS_MODULOS
    oDeco.cargar_combo cmbTipoHora, DECODIFICADORA.TAREAS_TIPOS_HORAS
    cargar_combo cmbUsuario, New clsUsuarios
    llenar_combo cmbclientes, New clsCliente, 0, frmClientes, ""
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Tarea", 3200, lvwColumnLeft
        .Add , , "Observaciones", 3200, lvwColumnLeft
        .Add , , "Horas", 1000, lvwColumnCenter
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "Tipo Hora", 1500, lvwColumnLeft
        .Add , , "Fact.", 700, lvwColumnCenter
    End With
End Sub
Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    lista.ListItems.Clear
    Dim oTareas As New clsTareas_incurridos
    Set rs = oTareas.Listado(fecha, cmbUsuario.BoundText)
    Dim total As Single
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4) ' TIPO HORA
             If rs(5) = 0 Then
                 .SubItems(5) = "" ' NO FACTURABLE
             Else
                 .SubItems(5) = "X" ' FACTURABLE
             End If
             total = total + rs(2)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oTareas = Nothing
    txtDatos(2) = Format(total, "0.00") & "h."
End Sub

Private Sub calcular_horas()
    Dim min As Integer
    Dim horas As Integer
    Dim resto As Integer
    Dim r As Single
    min = Format(DateDiff("n", hinicio.value, hfin.value), "#.##0")
    horas = min / 60
    resto = min Mod 60
    r = 0
    If resto > 0 Then
        r = resto / 60
    End If
    txtDatos(0) = horas + r
'    txtDatos(0) = Format(DateDiff("h", hinicio.value, hfin.value), "#.##0")
End Sub
