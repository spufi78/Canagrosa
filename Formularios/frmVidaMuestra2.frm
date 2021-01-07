VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmVidaMuestra 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Vida de la Muestra"
   ClientHeight    =   8910
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12135
   Icon            =   "frmVidaMuestra2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton opTipo 
      Caption         =   "Plasma"
      Height          =   240
      Index           =   5
      Left            =   8145
      TabIndex        =   36
      Top             =   8550
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.OptionButton opTipo 
      Caption         =   "Sellante"
      Height          =   240
      Index           =   2
      Left            =   8145
      TabIndex        =   34
      Top             =   8325
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.OptionButton opTipo 
      Caption         =   "CE"
      Height          =   240
      Index           =   1
      Left            =   8145
      TabIndex        =   33
      Top             =   8100
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.OptionButton opTipo 
      Caption         =   "Determinaciones"
      Height          =   240
      Index           =   0
      Left            =   8145
      TabIndex        =   32
      Top             =   7875
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.CheckBox chkModificar 
      Caption         =   "Permiso para modificar"
      Height          =   330
      Left            =   9225
      TabIndex        =   23
      Top             =   7695
      Visible         =   0   'False
      Width           =   1950
   End
   Begin Geslab.ControlPanelXP ControlPanelXP1 
      Height          =   3165
      Left            =   6840
      TabIndex        =   3
      Top             =   900
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   5583
      Caption         =   "Nuevas Ediciones Generadas"
      BackColor       =   16777215
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   3165
      Begin MSComctlLib.ListView listaEdiciones 
         Height          =   2550
         Left            =   90
         TabIndex        =   4
         Top             =   450
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   4498
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
   End
   Begin Geslab.ControlPanelXP ControlPanelXP2 
      Height          =   3165
      Left            =   6840
      TabIndex        =   5
      Top             =   1440
      Width           =   5280
      _ExtentX        =   9313
      _ExtentY        =   5583
      Caption         =   "Archivos Adjuntos"
      BackColor       =   16777215
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   3165
      Begin MSComctlLib.ListView listaAdjuntos 
         Height          =   2550
         Left            =   90
         TabIndex        =   6
         Top             =   450
         Width           =   5070
         _ExtentX        =   8943
         _ExtentY        =   4498
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   1005
      Left            =   11025
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7830
      Width           =   1050
   End
   Begin Geslab.ControlPanelXP panel 
      Height          =   5775
      Left            =   45
      TabIndex        =   7
      Top             =   1980
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   10186
      Caption         =   "Lista de Resultados"
      BackColor       =   16777215
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   5775
      Begin VB.Frame frmModificar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modificar Datos"
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
         Height          =   2685
         Left            =   2925
         TabIndex        =   24
         Top             =   1215
         Visible         =   0   'False
         Width           =   6570
         Begin VB.CommandButton cmdSalir 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Salir"
            Height          =   1005
            Left            =   5355
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   1575
            Width           =   1050
         End
         Begin VB.CommandButton cmdModificarDatosEspeciales 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Modificar"
            Height          =   1005
            Left            =   4185
            Picture         =   "frmVidaMuestra2.frx":08CA
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1575
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker fecha 
            Height          =   330
            Left            =   945
            TabIndex        =   26
            Top             =   675
            Width           =   1470
            _ExtentX        =   2593
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
            Format          =   51904513
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin pryCombo.miCombo cmbAnalista 
            Height          =   330
            Left            =   945
            TabIndex        =   27
            Top             =   270
            Width           =   5010
            _ExtentX        =   8837
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker hora 
            Height          =   330
            Left            =   945
            TabIndex        =   28
            Top             =   1125
            Width           =   1470
            _ExtentX        =   2593
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
            Format          =   51904514
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   28
            Left            =   135
            TabIndex        =   31
            Top             =   720
            Width           =   645
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Analista"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   24
            Left            =   135
            TabIndex        =   30
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hora"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   30
            Left            =   135
            TabIndex        =   29
            Top             =   1170
            Width           =   645
         End
      End
      Begin MSComctlLib.ListView listaDeterminaciones 
         Height          =   5250
         Left            =   45
         TabIndex        =   8
         Top             =   405
         Width           =   11955
         _ExtentX        =   21087
         _ExtentY        =   9260
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
   End
   Begin Geslab.ControlPanelXP ControlPanelXP4 
      Height          =   1050
      Left            =   45
      TabIndex        =   9
      Top             =   900
      Width           =   6765
      _ExtentX        =   11933
      _ExtentY        =   1852
      Caption         =   "Datos de la Recepción"
      BackColor       =   16777215
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   1050
      Begin VB.TextBox txtrecepcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   2
         Left            =   5670
         MaxLength       =   512
         TabIndex        =   14
         Top             =   540
         Width           =   960
      End
      Begin VB.TextBox txtrecepcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   1
         Left            =   4095
         MaxLength       =   512
         TabIndex        =   12
         Top             =   540
         Width           =   960
      End
      Begin VB.TextBox txtrecepcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   0
         Left            =   855
         MaxLength       =   512
         TabIndex        =   10
         Top             =   540
         Width           =   2535
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hora"
         Height          =   195
         Index           =   1
         Left            =   5220
         TabIndex        =   15
         Top             =   585
         Width           =   345
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   0
         Left            =   3510
         TabIndex        =   13
         Top             =   585
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   11
         Top             =   585
         Width           =   540
      End
   End
   Begin Geslab.ControlPanelXP panelCerrada 
      Height          =   1050
      Left            =   45
      TabIndex        =   16
      Top             =   7785
      Width           =   8070
      _ExtentX        =   14235
      _ExtentY        =   1852
      Caption         =   "Datos del Cierre"
      BackColor       =   16777215
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   1050
      Begin VB.TextBox txtrecepcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   5
         Left            =   810
         MaxLength       =   512
         TabIndex        =   19
         Top             =   540
         Width           =   3255
      End
      Begin VB.TextBox txtrecepcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   4
         Left            =   4815
         MaxLength       =   512
         TabIndex        =   18
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox txtrecepcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   315
         Index           =   3
         Left            =   7155
         MaxLength       =   512
         TabIndex        =   17
         Top             =   540
         Width           =   645
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   22
         Top             =   585
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   4230
         TabIndex        =   21
         Top             =   585
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ult.Edición"
         Height          =   195
         Index           =   2
         Left            =   6255
         TabIndex        =   20
         Top             =   585
         Width           =   765
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informe de Vida de la Muestra"
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
      TabIndex        =   2
      Top             =   90
      Width           =   3120
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   11520
      Picture         =   "frmVidaMuestra2.frx":1194
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Muestra Información de fechas y analistas que han realizado el análisis de una muestra"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   1
      Top             =   405
      Width           =   6135
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   12090
   End
End
Attribute VB_Name = "frmVidaMuestra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub fEnvio_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
End Sub

Private Sub cmdModificarDatosEspeciales_Click()
    Dim cAnalista As Integer
    Dim cFecha As Integer
    Dim cHora As Integer
    Dim cAnalistaId As Integer
    If opTipo(0).Value = True Then
        Dim oDV As New clsDeterminaciones_historico
        With oDV
            .setEMPLEADO_ID = cmbAnalista.getPK_SALIDA
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setHORA = Format(hora, "hh:mm")
            .ModificarVida listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(8), listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(9), listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(10)
        End With
        cAnalista = 4
        cFecha = 5
        cHora = 6
        cAnalistaId = 7
    ElseIf opTipo(1).Value = True Then ' CE
        Dim oCEH As New clsCe_resultados_historico
        With oCEH
            .setEMPLEADO_ID = cmbAnalista.getPK_SALIDA
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setHORA = Format(hora, "hh:mm")
            .ModificarVida PK, _
                listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(9), _
                listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(10), _
                listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(11), _
                listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(12), _
                listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(13)
        End With
        cAnalista = 5
        cFecha = 6
        cHora = 7
        cAnalistaId = 8

    ElseIf opTipo(2).Value = True Then ' SELLANTE
        Dim oSR As New clsSellantes_resultados
        With oSR
            .ModificarVida PK, _
                listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(7), _
                listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(8), _
                cmbAnalista.getPK_SALIDA, _
                Format(fecha, "yyyy-mm-dd"), _
                Format(hora, "hh:mm")
        End With
        Set oSR = Nothing
        cAnalista = 3
        cFecha = 4
        cHora = 5
        cAnalistaId = 6
        
    ElseIf opTipo(5).Value = True Then ' PLASMA
        Dim oPR As New clsPlasma_resultados_historico
        With oPR
            .setEMPLEADO_ID = cmbAnalista.getPK_SALIDA
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setHORA = Format(hora, "hh:mm:ss")
            .ModificarVida PK, _
                listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(1), _
                listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(2), _
                listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(4)
        End With
        Set oSR = Nothing
        cAnalista = 7
        cFecha = 8
        cHora = 9
        cAnalistaId = 10
    End If
    ' Actualizar lista
    Dim oUsuario As New clsUsuarios
    oUsuario.CARGAR cmbAnalista.getPK_SALIDA
    listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(cAnalista) = oUsuario.getUSUARIO
    listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(cFecha) = Format(fecha, "dd-mm-yyyy")
    listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(cHora) = Format(hora, "hh:mm")
    listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(cAnalistaId) = cmbAnalista.getPK_SALIDA
    Set oUsuario = Nothing

    frmModificar.visible = False

End Sub

Private Sub cmdSalir_Click()
    frmModificar.visible = False
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.MousePointer = 0
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select

End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    log (Me.Name)
    cargar_botones Me
    
    llenar_combo cmbAnalista, New clsUsuarios, 0, frmUsuarios, ""

    cabecera_general
    cargar_adjuntos
    cargar_ediciones
    lbltitulo(0) = "Informe de Vida de la Muestra"
    If PK <> 0 Then
        ' Datos de recepcion
        Dim oMuestra As New clsMuestra
        If oMuestra.CargaMuestra(PK) Then
            Dim oUsuario As New clsUsuarios
            oUsuario.CARGAR oMuestra.getEMPLEADO_ID
            txtrecepcion(0) = oUsuario.getAPELLIDOS & "," & oUsuario.getNOMBRE
            txtrecepcion(1) = Format(oMuestra.getFECHA_RECEPCION, "dd-mm-yyyy")
            txtrecepcion(2) = oMuestra.getHORA_RECEPCION
            If oMuestra.getCERRADA = 1 Then
                oUsuario.CARGAR oMuestra.getCERRADA_USUARIO
                txtrecepcion(5) = oUsuario.getAPELLIDOS & "," & oUsuario.getNOMBRE
                If IsDate(oMuestra.getFECHA_CIERRE) Then
                    txtrecepcion(4) = Format(oMuestra.getFECHA_CIERRE, "dd-mm-yyyy")
                End If
                txtrecepcion(3) = oMuestra.getULT_EDICION_IMP
                panelCerrada.PanelOpen = True
            Else
                panelCerrada.PanelOpen = False
            End If
        End If
        
        Select Case oMuestra.getANALISIS_MODIFICADO
            Case 2 ' Control de eficacia
                opTipo(1).Value = True
                cargar_control_eficacia
            Case 3 ' Sellante
                opTipo(2).Value = True
                cargar_sellante
            Case 5 ' Plasma
                opTipo(5).Value = True
                cargar_plasma
            Case Else
                opTipo(0).Value = True
                cargar_determinaciones
        End Select
        
        Set oMuestra = Nothing
'        Set rs = Nothing
    End If
    ' Permiso para modificar la vida
    Dim op As New clsParametros
    Dim s() As String
    Dim i As Integer
    op.Carga parametros.PARAM_USUARIOS_MODIFICAN_VIDA, ""
    If op.getVALOR <> "" Then
        s = Split(op.getVALOR, ",")
        For i = LBound(s) To UBound(s)
            If USUARIO.getID_EMPLEADO = CInt(s(i)) Then
                chkModificar.Value = Checked
                Exit For
            End If
        Next
    End If
    Set op = Nothing
   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmVidaMuestra"
End Sub

Private Sub cabecera_determinaciones()
    With listaDeterminaciones.ColumnHeaders
        .Add , , "Determinación", 3000, lvwColumnLeft
        .Add , , "Campo", 3000, lvwColumnLeft
        .Add , , "Resultado(1)", 1100, lvwColumnRight
        .Add , , "Resultado(2)", 1100, lvwColumnRight
        .Add , , "Analista", 1200, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Hora", 1100, lvwColumnCenter
        .Add , , "EMPLEADO_ID", 0, lvwColumnCenter
        .Add , , "DETERMINACION_ID", 0, lvwColumnCenter
        .Add , , "ORDEN", 0, lvwColumnCenter
        .Add , , "CAMPO", 0, lvwColumnCenter
    End With
End Sub
Private Sub cabecera_control_eficacia(Formula As Integer)
    On Error Resume Next
    With listaDeterminaciones.ColumnHeaders
        If Formula > 0 Then
            .Add , , "Iden. Canagrosa", 2000, lvwColumnLeft
            .Add , , "Iden. Cliente", 2000, lvwColumnLeft
            .Add , , "Campo", 2000, lvwColumnLeft
            .Add , , "Resultado(1)", 1100, lvwColumnRight
            .Add , , "Resultado(2)", 1100, lvwColumnRight
        Else
            .Add , , "Iden. Canagrosa", 3000, lvwColumnLeft
            .Add , , "Iden. Cliente", 3000, lvwColumnLeft
            .Add , , "Resultado", 1100, lvwColumnCenter
            .Add , , "Conforme", 1100, lvwColumnCenter
            .Add , , "Vacio", 0, lvwColumnCenter
        End If
        .Add , , "Analista", 1200, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Hora", 1100, lvwColumnCenter
        
        .Add , , "EMPLEADO_ID", 0, lvwColumnCenter
        .Add , , "DESIGNACION", 0, lvwColumnCenter
        .Add , , "PROBETA", 0, lvwColumnCenter
        .Add , , "AREA", 0, lvwColumnCenter
        .Add , , "ORDEN", 0, lvwColumnCenter
        .Add , , "CAMPO_ID", 0, lvwColumnCenter
    End With
End Sub

Private Sub cabecera_sellante()
    With listaDeterminaciones.ColumnHeaders
        .Add , , "Ensayo", 3000, lvwColumnLeft
        .Add , , "Resultado", 2200, lvwColumnRight
        .Add , , "Unidad", 1200, lvwColumnLeft
        .Add , , "Analista", 1500, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Hora", 1100, lvwColumnCenter
        .Add , , "EMPLEADO_ID", 0, lvwColumnCenter
        .Add , , "SELLANTE_ID", 0, lvwColumnCenter
        .Add , , "ORDEN", 0, lvwColumnCenter
    End With
End Sub
Private Sub cargar_sellante()
        cabecera_sellante
        Dim oSe_Resultados As New clsSellantes_resultados
        Dim rs As ADODB.Recordset
        Set rs = oSe_Resultados.Listado_Resultados_Vida(PK)
        If rs.RecordCount > 0 Then
            Do
                With listaDeterminaciones.ListItems.Add(, , rs(1))
                  .SubItems(1) = rs(4)
                  .SubItems(2) = rs(5)
                  .SubItems(3) = rs(11)
                  .SubItems(4) = Format(rs(9), "DD-MM-YYYY")
                  .SubItems(5) = Format(rs(10), "hh:mm")
                  .SubItems(6) = rs(12)
                  .SubItems(7) = rs(13) 'SELLANTE_ID
                  .SubItems(8) = rs(0) 'ORDEN
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
         
End Sub

Private Sub cargar_determinaciones()
        ' Determinaciones
        cabecera_determinaciones
        Dim oDeterminaciones As New clsDeterminaciones
        Dim oTD As New clsTipos_determinacion
        Dim rs As ADODB.Recordset
        Set rs = oDeterminaciones.lista_determinaciones_vida(PK)
        Dim aux_determinacion As String
        Dim aux_campo As String
        Dim aux_resultado As String
        Dim DETERMINACION As String
        If rs.RecordCount <> 0 Then
            aux_determinacion = ""
            aux_campo = ""
            aux_resultado = ""
            Do
                ' Determinacion
                If aux_determinacion <> rs(0) Then
                    DETERMINACION = rs(0)
                    aux_determinacion = rs(0)
                Else
                   DETERMINACION = " "
                End If
                If Not IsNull(rs(1)) Then
                    With listaDeterminaciones.ListItems.Add(, , DETERMINACION)
                     ' Campo
                     If aux_campo <> rs(1) Then
                         .SubItems(1) = rs(1)
                         aux_campo = rs(1)
                     Else
                        .SubItems(1) = " "
                     End If
                     If aux_resultado <> rs(2) Then
                         .SubItems(2) = rs(2)
                         aux_resultado = rs(2)
                     Else
                        .SubItems(2) = " "
                     End If
                     If rs(3) <> "" Then
                        .SubItems(3) = rs(3)
                     Else
                        .SubItems(3) = " "
                     End If
                     .SubItems(4) = rs(4)
                     .SubItems(5) = Format(rs(5), "DD-MM-YYYY")
                     .SubItems(6) = Format(rs(6), "hh:mm")
                     .SubItems(7) = rs(9) 'ID_EMPLEADO
                     .SubItems(8) = rs(10) ' DETERMINACION
                     .SubItems(9) = rs(11) ' ORDEN
                     .SubItems(10) = rs(12) ' CAMPO
                    End With
                Else
                    ' Alveograma
                    If CLng(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma")) = rs(7) Then
                        Dim oalveo As New clsAlveogramas
                        Dim oalveo_valN As New clsAlveograma_valores
                        Dim oalveo_valR As New clsAlveograma_valores
                        Dim oDET As New clsDeterminaciones
                        oDET.CargarDeterminacion rs(8)
                        Dim oUsu As New clsUsuarios
                        oUsu.CARGAR oDET.getEMPLEADO_ID
                        Dim alveo As Long
                        alveo = oalveo.ComprobarAlveograma(PK, rs(8))
                        If alveo <> 0 Then
                            oalveo_valN.CargarAlveogramaValores alveo, 0
                            oalveo_valR.CargarAlveogramaValores alveo, 1
                            ' TENACIDAD
                            With listaDeterminaciones.ListItems.Add(, , DETERMINACION)
                                .SubItems(1) = "TENACIDAD"
                                .SubItems(2) = formatear(oalveo_valN.getTENACIDAD, 5, 2)
                                .SubItems(3) = formatear(oalveo_valR.getTENACIDAD, 5, 2)
                                .SubItems(4) = oUsu.getUSUARIO
                                .SubItems(5) = Format(oDET.getFECHA, "DD-MM-YYYY")
                                .SubItems(6) = Format(oDET.getHORA, "hh:mm")
                                .SubItems(7) = oUsu.getID_EMPLEADO
                            End With
                            ' EXTENSIBILIDAD
                            With listaDeterminaciones.ListItems.Add(, , "")
                                .SubItems(1) = "EXTENSIBILIDAD"
                                .SubItems(2) = formatear(oalveo_valN.getEXTENSIBILIDAD, 5, 2)
                                .SubItems(3) = formatear(oalveo_valR.getEXTENSIBILIDAD, 5, 2)
                                .SubItems(4) = oUsu.getUSUARIO
                                .SubItems(5) = Format(oDET.getFECHA, "DD-MM-YYYY")
                                .SubItems(6) = Format(oDET.getHORA, "hh:mm")
                                .SubItems(7) = oUsu.getID_EMPLEADO
                            End With
                            ' W
                            With listaDeterminaciones.ListItems.Add(, , "")
                                .SubItems(1) = "W"
                                .SubItems(2) = formatear(oalveo_valN.getW, 5, 2)
                                .SubItems(3) = formatear(oalveo_valR.getW, 5, 2)
                                .SubItems(4) = oUsu.getUSUARIO
                                .SubItems(5) = Format(oDET.getFECHA, "DD-MM-YYYY")
                                .SubItems(6) = Format(oDET.getHORA, "hh:mm")
                                .SubItems(7) = oUsu.getID_EMPLEADO
                            End With
                        End If
                        Set oalveo = Nothing
                        Set oalveo_valN = Nothing
                        Set oalveo_valR = Nothing
                    End If
                End If
                rs.MoveNext
            Loop Until rs.EOF
        End If

End Sub

Private Sub cargar_control_eficacia()
    
    Dim oce_recepcion As New clsCe_recepcion
    Dim oce_tipo_ensayo As New clsCe_tipos_ensayos
    If oce_recepcion.Carga(PK) Then
        oce_tipo_ensayo.Carga (oce_recepcion.getTIPO_ENSAYO_ID)
        cabecera_control_eficacia oce_tipo_ensayo.getFORMULA_ID
        Dim oCe_resultados As New clsCe_resultados
        Dim rs As ADODB.Recordset
        Set rs = oCe_resultados.Listado_por_muestra_vida(PK)
        If rs.RecordCount > 0 Then
            Dim idcanagrosa As String
            Dim aux_idcanagrosa As String
            Dim aux_idcliente As String
            Dim aux_campo As String
            aux_idcanagrosa = ""
            aux_idcliente = ""
            aux_campo = ""
            Do
                ' Id. Canagrosa
                If aux_idcanagrosa <> rs(0) Then
                   idcanagrosa = rs(0)
                   aux_idcanagrosa = rs(0)
                Else
                   idcanagrosa = " "
                End If
                With listaDeterminaciones.ListItems.Add(, , idcanagrosa)
                 ' ID. Cliente
                 If aux_idcliente <> rs(1) Then
                     .SubItems(1) = rs(1)
                     aux_idcliente = rs(1)
                 Else
                    .SubItems(1) = " "
                 End If
                 ' FORMULA
                 If oce_tipo_ensayo.getFORMULA_ID <> 0 Then
                    If aux_campo <> rs(5) Then ' Campo
                        .SubItems(2) = rs(5)
                        aux_campo = rs(5)
                    Else
                        .SubItems(2) = " "
                    End If
                    .SubItems(3) = rs(6) ' Valor 1
                    .SubItems(4) = rs(7) ' Valor 2
                    
                    .SubItems(5) = rs(8) ' Analista
                    .SubItems(6) = Format(rs(9), "dd-mm-yyyy") ' Fecha
                    .SubItems(7) = rs(10) ' Hora
                    .SubItems(8) = rs(11) ' id_empleado
                    
                    .SubItems(9) = rs(12) ' DESIGNACION
                    .SubItems(10) = rs(13) ' PROBETA
                    .SubItems(11) = rs(14) ' AREA
                    .SubItems(12) = rs(15) ' ORDEN
                    .SubItems(13) = rs(16) ' CAMPO_ID
                 Else ' NO FORMULA
                    If rs(2) = "" Then ' Resultado
                        .SubItems(2) = " "
                    Else
                        .SubItems(2) = rs(2)
                    End If
                    If rs(3) = 0 Then
                        .SubItems(3) = "No" ' Conforme
                    Else
                        .SubItems(3) = "Si" ' Conforme
                    End If
                    .SubItems(4) = "" 'Vacio
                    .SubItems(5) = rs(8) ' Analista
                    .SubItems(6) = Format(rs(9), "dd-mm-yyyy") ' Fecha
                    .SubItems(7) = rs(10) ' Hora
                    .SubItems(8) = rs(11) ' id_empleado
                 
                    .SubItems(9) = rs(12) ' DESIGNACION
                    .SubItems(10) = rs(13) ' PROBETA
                    .SubItems(11) = rs(14) ' AREA
                    .SubItems(12) = rs(15) ' ORDEN
                    .SubItems(13) = rs(16) ' CAMPO_ID
                 End If
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = Nothing
        Set oCe_resultados = Nothing
    End If
    Set oce_recepcion = Nothing
    Set oce_tipo_ensayo = Nothing
    
End Sub


Private Sub cargar_adjuntos()
'    Dim rs As ADODB.RecordSet
'    listaAdjuntos.ListItems.Clear
'    Dim oMuestra_Adjunto As New clsMuestras_adjuntos
'    Set rs = oMuestra_Adjunto.Listado(PK)
'    If rs.RecordCount > 0 Then
'        Do
'            With listaAdjuntos.ListItems.Add(, , rs(0))
'                 .SubItems(1) = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\" & rs(3) & "\" & rs(0)
'            End With
'            rs.MoveNext
'        Loop Until rs.EOF
'    End If
'    Set rs = Nothing
'    Set oMuestra_Adjunto = Nothing
End Sub

Private Sub cargar_ediciones()
    Dim rs As ADODB.Recordset
    listaEdiciones.ListItems.Clear
    Dim oMe As New clsMuestras_ediciones
    Set rs = oMe.Listado(PK)
    If rs.RecordCount > 0 Then
        Do
            With listaEdiciones.ListItems.Add(, , "Edición " & rs("EDICION"))
                 .SubItems(1) = Format(rs("FECHA"), "dd-mm-yyyy")
                 .SubItems(2) = ""
                 .SubItems(3) = rs("OBSERVACIONES")
                 .SubItems(4) = rs("USUARIO_ID")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oMe = Nothing
End Sub

Private Sub cabecera_general()
    With listaAdjuntos.ColumnHeaders
        .Add , , "Ficheros Adjuntos...", listaAdjuntos.Width, lvwColumnLeft
        .Add , , "Ruta", 1, lvwColumnCenter
    End With
    With listaEdiciones.ColumnHeaders
        .Add , , "Edición", 800, lvwColumnLeft
        .Add , , "Fecha", 800, lvwColumnCenter
        .Add , , "Archivo", 1, lvwColumnCenter
        .Add , , "Motivo", 2400, lvwColumnCenter
        .Add , , "Usuario", 1000, lvwColumnCenter
    End With
End Sub

Private Sub listaAdjuntos_DblClick()
    If listaAdjuntos.ListItems.Count > 0 Then
        Dim destino As String
        destino = listaAdjuntos.ListItems(listaAdjuntos.selectedItem.Index).SubItems(1)
        On Error GoTo fallo
        If Dir(destino) <> "" Then
            Dim r As Long
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbNormalFocus)
        End If
    End If
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento asociado.", vbCritical, App.Title

End Sub

Private Sub listaDeterminaciones_Click()
    If listaDeterminaciones.ListItems.Count > 0 Then
        Dim colAnalista As Integer
        Dim colFecha As Integer
        Dim colHora As Integer
        
        If opTipo(0).Value = True Then ' DETERMINACIONES
            colAnalista = 7
            colFecha = 5
            colHora = 6
        End If
        If opTipo(1).Value = True Then ' CE
            colAnalista = 8
            colFecha = 6
            colHora = 7
        End If
        If opTipo(2).Value = True Then ' SELLANTE
            colAnalista = 6
            colFecha = 4
            colHora = 5
        End If
        If opTipo(5).Value = True Then ' PLASMA
            colAnalista = 10
            colFecha = 8
            colHora = 9
        End If
        cmbAnalista.MostrarElemento listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(colAnalista)
        fecha = listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(colFecha)
        hora = fecha & " " & listaDeterminaciones.ListItems(listaDeterminaciones.selectedItem.Index).SubItems(colHora)
    End If
End Sub

Private Sub listaDeterminaciones_DblClick()
    If listaDeterminaciones.ListItems.Count > 0 And chkModificar.Value = Checked Then
            frmModificar.visible = True
    End If
End Sub
Private Sub cabecera_plasma()
    On Error Resume Next
    With listaDeterminaciones.ColumnHeaders
        .Add , , "MUESTRA_ID", 0, lvwColumnLeft
        .Add , , "EDICION", 500, lvwColumnCenter
        .Add , , "TIPO", 0, lvwColumnCenter
        .Add , , "COAT", 1000, lvwColumnCenter
        .Add , , "DESIGNACION", 3000, lvwColumnCenter
        .Add , , "RESULT", 2500, lvwColumnCenter
        .Add , , "PASS/FAIL", 1100, lvwColumnCenter
        .Add , , "Analista", 1200, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Hora", 1100, lvwColumnCenter
        .Add , , "AnalistaId", 0, lvwColumnCenter
    End With
End Sub
Private Sub cargar_plasma()
        cabecera_plasma
        Dim oPRH As New clsPlasma_resultados_historico
        Dim rs As ADODB.Recordset
        Set rs = oPRH.Listado(PK)
        Dim capa As String
        Dim des As String
        If rs.RecordCount > 0 Then
            Do
                With listaDeterminaciones.ListItems.Add(, , rs("MUESTRA_ID"))
                  .SubItems(1) = rs("EDICION")
                  .SubItems(2) = rs("TIPO")
                  If rs("tipo") = 1 Then
                      .SubItems(3) = "BOND"
                  ElseIf rs("tipo") = 2 Then
                      .SubItems(3) = "TOP"
                  End If
                  .SubItems(4) = rs("DESIGNACION")
                  .SubItems(5) = rs("RESULTADO")
                  Select Case rs("CONFORME")
                    Case 0
                        .SubItems(6) = "FAIL"
                    Case 1
                        .SubItems(6) = "PASS"
                    Case 2
                        .SubItems(6) = "N.R."
                  End Select
                  .SubItems(7) = rs("ANALISTA")
                  .SubItems(8) = Format(rs("FECHA"), "DD-MM-YYYY")
                  .SubItems(9) = Format(rs("HORA"), "hh:mm:ss")
                  .SubItems(10) = rs("EMPLEADO_ID")
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
         
End Sub
