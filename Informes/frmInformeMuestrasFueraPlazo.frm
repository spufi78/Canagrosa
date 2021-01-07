VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmInformeMuestrasFueraPlazo 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de muestras fuera de plazo"
   ClientHeight    =   9495
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13680
   Icon            =   "frmInformeMuestrasFueraPlazo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9495
   ScaleWidth      =   13680
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   3285
      TabIndex        =   20
      Top             =   8640
      Visible         =   0   'False
      Width           =   6450
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Generando documento EXCEL. Por favor, espere."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   330
         Index           =   1
         Left            =   675
         TabIndex        =   21
         Top             =   225
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdVerExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   11475
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8595
      Width           =   1050
   End
   Begin VB.CommandButton cmdInforme 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe Muestra"
      Height          =   870
      Left            =   10350
      Picture         =   "frmInformeMuestrasFueraPlazo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Previsualizar informe de ensayo"
      Top             =   8595
      Width           =   1095
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   870
      Left            =   30
      Picture         =   "frmInformeMuestrasFueraPlazo.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8595
      Width           =   1050
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   870
      Left            =   1102
      Picture         =   "frmInformeMuestrasFueraPlazo.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8595
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12555
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8595
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2340
      Left            =   0
      TabIndex        =   0
      Top             =   405
      Width           =   13590
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Height          =   555
         Left            =   135
         TabIndex        =   22
         Top             =   1665
         Width           =   8475
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cierre"
            Height          =   195
            Index           =   1
            Left            =   1575
            TabIndex        =   31
            Top             =   225
            Width           =   1095
         End
         Begin VB.OptionButton Option1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Recepción"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   30
            Top             =   225
            Width           =   1095
         End
         Begin VB.TextBox txtanno 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7290
            TabIndex        =   23
            Top             =   150
            Visible         =   0   'False
            Width           =   735
         End
         Begin MSComCtl2.UpDown cambiar 
            Height          =   375
            Left            =   8040
            TabIndex        =   24
            Top             =   150
            Visible         =   0   'False
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   661
            _Version        =   393216
            Value           =   2004
            BuddyControl    =   "txtanno"
            BuddyDispid     =   196619
            OrigLeft        =   1590
            OrigTop         =   6570
            OrigRight       =   1830
            OrigBottom      =   6975
            Max             =   2015
            Min             =   2004
            SyncBuddy       =   -1  'True
            BuddyProperty   =   0
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.DTPicker fdesde 
            Height          =   330
            Left            =   3315
            TabIndex        =   25
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
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
         Begin MSComCtl2.DTPicker fhasta 
            Height          =   330
            Left            =   5250
            TabIndex        =   26
            Top             =   180
            Width           =   1275
            _ExtentX        =   2249
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
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Año"
            Height          =   195
            Index           =   0
            Left            =   6750
            TabIndex        =   29
            Top             =   225
            Visible         =   0   'False
            Width           =   360
         End
         Begin VB.Label lblCampos 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "hasta"
            Height          =   195
            Index           =   4
            Left            =   4710
            TabIndex        =   28
            Top             =   225
            Width           =   405
         End
         Begin VB.Label lblCampos 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "desde"
            Height          =   195
            Index           =   6
            Left            =   2715
            TabIndex        =   27
            Top             =   225
            Width           =   465
         End
      End
      Begin VB.CheckBox chkAnalisis 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   10305
         TabIndex        =   19
         Top             =   990
         Value           =   1  'Checked
         Width           =   870
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   10305
         TabIndex        =   3
         Top             =   270
         Value           =   1  'Checked
         Width           =   870
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
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
         Height          =   255
         Left            =   10305
         TabIndex        =   2
         Top             =   630
         Value           =   1  'Checked
         Width           =   915
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   975
         Left            =   12105
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   720
         Width           =   1365
      End
      Begin pryCombo.miCombo cmbTiposMuestra 
         Height          =   330
         Left            =   1485
         TabIndex        =   4
         Top             =   630
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1485
         TabIndex        =   5
         Top             =   270
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbTA 
         Height          =   330
         Left            =   1485
         TabIndex        =   16
         Top             =   990
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbTipoEnsayo 
         Height          =   345
         Left            =   1485
         TabIndex        =   32
         Top             =   1350
         Width           =   4230
         _ExtentX        =   7461
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbCentro 
         Height          =   345
         Left            =   6345
         TabIndex        =   33
         Top             =   1350
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Ensayo"
         Height          =   195
         Index           =   19
         Left            =   135
         TabIndex        =   35
         Top             =   1380
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   5760
         TabIndex        =   34
         Top             =   1395
         Width           =   465
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de analisis"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   135
         TabIndex        =   17
         Top             =   1035
         Width           =   1995
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   690
         Width           =   1185
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5505
      Left            =   45
      TabIndex        =   12
      Top             =   3060
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   9710
      View            =   3
      Arrange         =   1
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4005
      Top             =   8775
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasFueraPlazo.frx":18E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasFueraPlazo.frx":21BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasFueraPlazo.frx":2A94
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasFueraPlazo.frx":336E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasFueraPlazo.frx":3C48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInformeMuestrasFueraPlazo.frx":4522
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "MUESTRAS FUERA DE PLAZO"
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
      Left            =   135
      TabIndex        =   15
      Top             =   45
      Width           =   3855
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione Criterio."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   14
      Top             =   2745
      Width           =   13590
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   13620
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "MUESTRAS PENDIENTES DE ENVÍO"
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
      Index           =   0
      Left            =   315
      TabIndex        =   13
      Top             =   90
      Width           =   4590
   End
End
Attribute VB_Name = "frmInformeMuestrasFueraPlazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************'
'*************** Fecha de creación del formulario: 26/02/2014 ****************'
'****************                MANTIS: 1289                  ***************'
'*****************************************************************************'

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    txtanno = Year(Date)
    fdesde = "01/" & Month(Date) & "/" & Year(Date)
    fhasta = Date
    'M1373-I
    Option1(0).value = True
    'M1373-F
    cabecera
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbTiposMuestra, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
    llenar_combo cmbTA, New clsTipos_analisis, 0, frmTA_Detalle, ""
    llenar_combo cmbCentro, New clsCentros, 0, Me, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbTipoEnsayo, DECODIFICADORA.DECODIFICADORA_TM_TIPOS_ENSAYOS
    Set oDeco = Nothing
    
    Call buscar
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 300, lvwColumnLeft
        .Add , , "Cliente", 3200, lvwColumnLeft
        .Add , , "T.Muestra", 2300, lvwColumnLeft
        .Add , , "T.Análisis", 2300, lvwColumnLeft
        .Add , , "F.Recepción", 1050, lvwColumnCenter
        .Add , , "F.Prev.Fin", 1050, lvwColumnCenter
        .Add , , "F.Cierre", 1050, lvwColumnCenter
        .Add , , "IPA", 550, lvwColumnCenter
        .Add , , "Motivo Retraso", 1200, lvwColumnLeft
        .Add , , "ID_General", 1, lvwColumnLeft
    End With
End Sub

Private Sub cmdBuscar_Click()
    Call buscar
End Sub

Private Sub cmdInforme_Click()
    If lista.ListItems.Count > 0 Then
        MostrarInforme CLng(lista.ListItems(lista.selectedItem.Index).Text)
    End If
End Sub

Private Sub chkAnalisis_Click()
    If chkAnalisis.value = Checked Then
        cmbTA.Limpiar
        cmbTA.desactivar
    Else
        cmbTA.activar
    End If
End Sub

Private Sub chkTodas_Click()
    If chkTodas.value = Checked Then
        cmbTiposMuestra.Limpiar
        cmbTiposMuestra.desactivar
    Else
        cmbTiposMuestra.activar
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.value = Checked Then
        cmbClientes.Limpiar
        cmbClientes.desactivar
    Else
        cmbClientes.activar
    End If

End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub
Private Sub buscar()
    Dim i As Integer
    Dim consulta As String
    Dim strClientes As String
    Dim strTipoMuestra As String
    Dim strTipoAnalisis As String
    Dim stranno As String
    Dim fecha_desde As String
    Dim fecha_hasta As String
    
    On Error GoTo fallo
    
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    
    ' Clientes
    strClientes = ""
    If chkTodos.value = Unchecked Then
        If cmbClientes.getPK_SALIDA = 0 Then
            MsgBox "Debe seleccionar un cliente.", vbExclamation, App.Title
            Exit Sub
        End If
        strClientes = " AND m.cliente_id = " & cmbClientes.getPK_SALIDA
    End If
    
    ' Tipo de muestra
    strTipoMuestra = ""
    If chkTodas.value = Unchecked Then
        If cmbTiposMuestra.getTEXTO = "" Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Sub
        End If
        strTipoMuestra = " AND m.tipo_muestra_id=" & cmbTiposMuestra.getPK_SALIDA
    End If
    
    ' Tipo de análisis
    strTipoAnalisis = ""
    If chkAnalisis.value = Unchecked Then
        If cmbTA.getTEXTO = "" Then
            MsgBox "Debe seleccionar un tipo de análisis.", vbExclamation, App.Title
            Exit Sub
        End If
        strTipoAnalisis = " AND m.tipo_analisis_id=" & cmbTA.getPK_SALIDA
    End If
    ' Fechas
    'M1373-I
    'fecha_desde = " AND m.fecha_recepcion>='" & Format(fdesde, "yyyy-mm-dd") & "'"
    'fecha_hasta = " AND m.fecha_recepcion<='" & Format(fhasta, "yyyy-mm-dd") & "'"
    If Option1(0).value = True Then
        fecha_desde = " AND m.fecha_recepcion>='" & Format(fdesde, "yyyy-mm-dd") & "'"
        fecha_hasta = " AND m.fecha_recepcion<='" & Format(fhasta, "yyyy-mm-dd") & "'"
    Else
        fecha_desde = " AND m.cerrada = 1 AND m.fecha_cierre>='" & Format(fdesde, "yyyy-mm-dd") & "'"
        fecha_hasta = " AND m.fecha_cierre<='" & Format(fhasta, "yyyy-mm-dd") & "'"
    End If
    Dim ANNO As String
    ANNO = Year(fdesde) & "," & Year(fhasta)
    'M1373-F
    Dim strCentro As String
    If cmbCentro.getTEXTO <> "" Then
        strCentro = " and M.centro_id = " & CInt(cmbCentro.getPK_SALIDA)
    End If
    Dim strTE As String
    If cmbTipoEnsayo.getTEXTO <> "" Then
        strTE = " and TM.TIPO_ENSAYO_ID = " & cmbTipoEnsayo.getPK_SALIDA
    End If
    
    consulta = "SELECT M.ID_MUESTRA, C.NOMBRE,TM.NOMBRE, TA.NOMBRE, M.FECHA_RECEPCION, M.FECHA_PREV_FIN, M.FECHA_CIERRE, M.IPA, D.DESCRIPCION AS MOTIVO_RETRASO, M.ID_GENERAL" & _
               " FROM MUESTRAS M, CLIENTES C, TIPOS_MUESTRA TM, TIPOS_ANALISIS TA, decodificadora D" & _
               " WHERE M.ANNO IN (" & ANNO & ") AND M.ANULADA = 0 " & _
               " AND (M.FECHA_PREV_FIN < M.FECHA_CIERRE OR M.CERRADA = 0)" & _
               " AND M.FECHA_RECEPCION < M.FECHA_PREV_FIN" & _
               " AND M.CLIENTE_ID = C.ID_CLIENTE" & _
               " AND M.TIPO_MUESTRA_ID = TM.ID_TIPO_MUESTRA" & _
               " AND M.TIPO_ANALISIS_ID = TA.ID_TIPO_ANALISIS" & _
               " AND M.MOTIVO_RETRASO_ID = D.VALOR AND D.CODIGO = 155" & _
               " AND M.FECHA_PREV_FIN <='" & Format(Date, "yyyy-mm-dd") & "'" & _
               " AND M.ULT_EDICION_IMP = 1 " & _
                fecha_desde & strCentro & _
                fecha_hasta & strTE & _
                strClientes & _
                strTipoMuestra & _
                strTipoAnalisis
                
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    Dim total As Integer
    total = rs.RecordCount
    If rs.RecordCount >= 1 Then
        Dim oMuestra As New clsMuestra
        lista.ListItems.Clear
        i = 1
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
                If rs(4) <> "" Then
                    .SubItems(4) = rs(4)
                End If
                If rs(5) <> "" Then
                    .SubItems(5) = rs(5)
                End If
                If rs(6) <> "" Then
                    .SubItems(6) = rs(6)
                End If
                .SubItems(7) = rs(7)
                .SubItems(8) = rs(8)
                .SubItems(9) = rs(9)
            End With
            lista.ListItems(lista.ListItems.Count).Checked = True
            rs.MoveNext
        Wend
        lblMsg.Caption = rs.RecordCount & " muestras entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy") & " (TOTAL : " & total & ")"
    Else
        lblMsg.Caption = "No existe ninguna muestra con esos criterios."
    End If
    Set oAnalisis = Nothing
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras.", vbCritical, Err.Description
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).Text
        frmVerMuestra.Show 1
        gmuestra = 0
    End If
End Sub

Private Sub txtp1_GotFocus()
    txtp1.SelStart = 0
    txtp1.SelLength = Len(txtp1)
End Sub

Private Sub txtp1_LostFocus()
    If txtp1 <> "" Then
        txtp2 = txtp1
    End If
End Sub

Private Sub txtp2_GotFocus()
    txtp2.SelStart = 0
    txtp2.SelLength = Len(txtp2)
    
End Sub

Private Sub cmdVerExcel_Click()
       Me.MousePointer = vbHourglass
       Frame3.Visible = True
       Dim rs As New ADODB.Recordset
       Dim fechaI As String
       Dim fechaF As String
       rs.Fields.Append "c1", adChar, 10, adFldUpdatable
       rs.Fields.Append "c2", adChar, 350, adFldUpdatable
       rs.Fields.Append "c3", adChar, 350, adFldUpdatable
       rs.Fields.Append "c4", adChar, 350, adFldUpdatable
       rs.Fields.Append "c5", adChar, 20, adFldUpdatable
       rs.Fields.Append "c6", adChar, 20, adFldUpdatable
       rs.Fields.Append "c7", adChar, 20, adFldUpdatable
       rs.Fields.Append "c8", adChar, 10, adFldUpdatable
       rs.Fields.Append "c9", adChar, 10, adFldUpdatable
       rs.Fields.Append "c10", adChar, 350, adFldUpdatable
       rs.Open
       
       Dim i As Integer

       For i = 1 To lista.ListItems.Count
           If lista.ListItems(i).Checked = True Then
                rs.AddNew
                rs("c1") = lista.ListItems(i).SubItems(9)
                rs("c2") = lista.ListItems(i).SubItems(1)
                rs("c3") = lista.ListItems(i).SubItems(2)
                rs("c4") = lista.ListItems(i).SubItems(3)
                rs("c5") = lista.ListItems(i).SubItems(4)
                rs("c6") = lista.ListItems(i).SubItems(5)
                rs("c7") = lista.ListItems(i).SubItems(6)
                rs("c9") = lista.ListItems(i).SubItems(7)
                rs("c10") = lista.ListItems(i).SubItems(8)
                rs.Update
           End If
        Next i
        
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Muestras Fuera de Plazo"
 
        'Cabecera
        With XLS.Range("A1:J1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With XLS.Range("A1:J1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:I1").Borders
            .LineStyle = vbSolid
        End With
        XLS.Range("A1:A1").ColumnWidth = 12
        XLS.Range("B1:B1").ColumnWidth = 55
        XLS.Range("C1:C1").ColumnWidth = 55
        XLS.Range("D1:D1").ColumnWidth = 70
        XLS.Range("E1:E1").ColumnWidth = 12
        XLS.Range("F1:F1").ColumnWidth = 12
        XLS.Range("G1:G1").ColumnWidth = 12
        XLS.Range("H1:H1").ColumnWidth = 10
        XLS.Range("I1:I1").ColumnWidth = 5
        XLS.Range("J1:J1").ColumnWidth = 30
        
        XLS.Cells(1, 1) = "ID General"
        XLS.Cells(1, 2) = "Cliente"
        XLS.Cells(1, 3) = "Tipo de Muestra"
        XLS.Cells(1, 4) = "Tipo de Análisis"
        XLS.Cells(1, 5) = "Fecha Recepción"
        XLS.Cells(1, 6) = "Fecha Prev.Fin"
        XLS.Cells(1, 7) = "Fecha Cierre"
        XLS.Cells(1, 8) = "Retraso"
        XLS.Cells(1, 9) = "IPA"
        XLS.Cells(1, 10) = "Motivo Retraso"
        
        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            fechaI = Trim(rs("c6"))
            fechaF = Trim(rs("c7"))
            XLS.Cells(i, 1) = CLng(rs("c1"))
            XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True)
            XLS.Cells(i, 3) = ClrStr(rs("c3"), False, True, True)
            XLS.Cells(i, 4) = ClrStr(rs("c4"), False, True, True)
            XLS.Cells(i, 5) = rs("c5")
            XLS.Cells(i, 6) = rs("c6")
            XLS.Cells(i, 7) = rs("c7")
            If fechaF <> "" Then
                XLS.Cells(i, 8) = DateDiff("d", fechaI, fechaF)
            End If
            XLS.Cells(i, 9) = CLng(rs("C9"))
            XLS.Cells(i, 10) = rs("C10")
            i = i + 1
             
            XLS.Range("A" & i).EntireRow.Insert
            rs.MoveNext
          Loop Until rs.EOF
        End If
        Frame3.Visible = False
        Me.MousePointer = vbNormal
        XLA.Visible = True
        Set rs = Nothing
End Sub

