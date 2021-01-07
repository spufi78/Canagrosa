VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmProcNC_AvisosAccCorrectivas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de acciones asociadas a Procedimientos de No Conformidad"
   ClientHeight    =   9300
   ClientLeft      =   2550
   ClientTop       =   2280
   ClientWidth     =   13680
   ClipControls    =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmProcNC_AvisosAccCorrectivas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   13680
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   3780
      TabIndex        =   34
      Top             =   4275
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
         TabIndex        =   35
         Top             =   225
         Width           =   5415
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por"
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
      Height          =   1725
      Left            =   45
      TabIndex        =   8
      Top             =   675
      Width           =   13560
      Begin VB.TextBox txtCNCf 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   10305
         MaxLength       =   255
         TabIndex        =   41
         Top             =   180
         Width           =   1395
      End
      Begin VB.TextBox txtCNCi 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   8010
         MaxLength       =   255
         TabIndex        =   38
         Top             =   180
         Width           =   1530
      End
      Begin VB.CheckBox chkResolucion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   240
         Left            =   11835
         TabIndex        =   31
         Top             =   1350
         Width           =   240
      End
      Begin VB.CheckBox chkPuesta 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   240
         Left            =   11835
         TabIndex        =   30
         Top             =   990
         Width           =   240
      End
      Begin VB.CheckBox chkPrevista 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   240
         Left            =   11835
         TabIndex        =   29
         Top             =   630
         Width           =   240
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   1050
         Left            =   12375
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   1050
      End
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1350
         MaxLength       =   255
         TabIndex        =   9
         Top             =   225
         Width           =   4005
      End
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   315
         Left            =   1350
         TabIndex        =   11
         Top             =   945
         Width           =   4035
         _ExtentX        =   7117
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
      Begin MSComCtl2.DTPicker fPrevistaDesde 
         Height          =   330
         Left            =   8010
         TabIndex        =   12
         Top             =   585
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSComCtl2.DTPicker fPrevistaHasta 
         Height          =   330
         Left            =   10305
         TabIndex        =   13
         Top             =   585
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin pryCombo.miCombo cmbResponsableImplantacion 
         Height          =   315
         Left            =   1350
         TabIndex        =   18
         Top             =   585
         Width           =   5085
         _ExtentX        =   8969
         _ExtentY        =   556
      End
      Begin MSComCtl2.DTPicker fPuestaDesde 
         Height          =   330
         Left            =   8010
         TabIndex        =   20
         Top             =   945
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSComCtl2.DTPicker fPuestaHasta 
         Height          =   330
         Left            =   10305
         TabIndex        =   21
         Top             =   945
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSComCtl2.DTPicker fResolucionDesde 
         Height          =   330
         Left            =   8010
         TabIndex        =   24
         Top             =   1305
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSComCtl2.DTPicker fResolucionHasta 
         Height          =   330
         Left            =   10305
         TabIndex        =   25
         Top             =   1305
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   1350
         TabIndex        =   32
         Top             =   1305
         Width           =   4035
         _ExtentX        =   7117
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "a  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   9855
         TabIndex        =   40
         Top             =   225
         Width           =   240
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cod. P.N.C:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   6840
         TabIndex        =   39
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   8
         Left            =   225
         TabIndex        =   33
         Top             =   1350
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   7
         Left            =   9720
         TabIndex        =   27
         Top             =   1395
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resolución"
         Height          =   195
         Index           =   6
         Left            =   7065
         TabIndex        =   26
         Top             =   1395
         Width           =   795
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   5
         Left            =   9720
         TabIndex        =   23
         Top             =   1035
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comienzo"
         Height          =   195
         Index           =   4
         Left            =   7155
         TabIndex        =   22
         Top             =   1035
         Width           =   690
      End
      Begin VB.Label lbltitulo 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   19
         Top             =   630
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   17
         Top             =   990
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Título"
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   16
         Top             =   270
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Prevista"
         Height          =   195
         Index           =   0
         Left            =   6765
         TabIndex        =   15
         Top             =   630
         Width           =   1065
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   3
         Left            =   9720
         TabIndex        =   14
         Top             =   630
         Width           =   405
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   1140
      Left            =   6165
      TabIndex        =   5
      Top             =   8100
      Width           =   3435
      Begin VB.CommandButton cmdVerExcel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "EXCEL"
         Height          =   825
         Left            =   2340
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   225
         Width           =   1005
      End
      Begin VB.CommandButton cmdVerAccion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver Acción"
         Height          =   825
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   225
         Width           =   1005
      End
      Begin VB.CommandButton cmdVerPNC 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ir a P.N.C."
         Height          =   825
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1005
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Leyenda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   45
      TabIndex        =   2
      Top             =   8145
      Width           =   3750
      Begin VB.Label lblCap 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rojo :"
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
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   4
         Top             =   270
         Width           =   645
      End
      Begin VB.Label lblCap 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Prevista Sobrepasada"
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
         Height          =   225
         Index           =   4
         Left            =   855
         TabIndex        =   3
         Top             =   270
         Width           =   2670
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5670
      Left            =   45
      TabIndex        =   0
      Top             =   2430
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   10001
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
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   12330
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8280
      Width           =   1275
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Index           =   2
      Left            =   90
      TabIndex        =   37
      Top             =   360
      Width           =   870
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13095
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de acciones asociadas a Procedimientos de No Conformidad"
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
      Index           =   1
      Left            =   90
      TabIndex        =   28
      Top             =   45
      Width           =   7215
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   -135
      Top             =   0
      Width           =   13770
   End
End
Attribute VB_Name = "frmProcNC_AvisosAccCorrectivas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarobjAccCorrectoras As New clsProcNcAccionCorrectora
'Private mvarblnSinAvisos As Boolean
'Private mvarstrMensajeAviso As String

Private Sub cabecera()
On Error GoTo cabecera_Error
    
    With lista.ColumnHeaders
        .Add , , "Nº P.N.C.", 1000, lvwColumnLeft
        .Add , , "Titulo", 3000, lvwColumnLeft
        .Add , , "Responsable", 2000, lvwColumnLeft
        .Add , , "Fecha Prevista", 1300, lvwColumnLeft
        .Add , , "id_accion", 0, lvwColumnLeft
        .Add , , "Comienzo", 1200, lvwColumnLeft
        .Add , , "Resolución", 1200, lvwColumnLeft
        .Add , , "Estado", 2400, lvwColumnLeft
        .Add , , "Tipo", 1200, lvwColumnLeft
        .Add , , "EstadoId", 1, lvwColumnLeft
    End With
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.cabecera"
    Exit Sub
cabecera_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.cabecera"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cabecera of Formulario frmProcNC_AvisosAccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Private Sub chkPrevista_Click()
    If chkPrevista.value = Unchecked Then
        fPrevistaDesde.Enabled = False
        fPrevistaHasta.Enabled = False
    Else
        fPrevistaDesde.Enabled = True
        fPrevistaHasta.Enabled = True
    End If
End Sub

Private Sub chkPuesta_Click()
    If chkPuesta.value = Unchecked Then
        fPuestaDesde.Enabled = False
        fPuestaHasta.Enabled = False
    Else
        fPuestaDesde.Enabled = True
        fPuestaHasta.Enabled = True
    End If
End Sub

Private Sub chkResolucion_Click()
    If chkResolucion.value = Unchecked Then
        fResolucionDesde.Enabled = False
        fResolucionHasta.Enabled = False
    Else
        fResolucionDesde.Enabled = True
        fResolucionHasta.Enabled = True
    End If
End Sub

Private Sub cmdVerAccion_Click()
    lista_DblClick
End Sub

Private Sub Form_Load()
On Error GoTo Form_Load_Error
    
'    mvarblnSinAvisos = False
    Me.top = 300
    Me.Left = 300
'    Me.Width = 9060
'    lista.Width = Me.ScaleWidth
'    lista.Height = cmdcancel.Top - 30
    cabecera
    cargar_botones Me
    cargar_combos
    cargar_fechas
    cargar_checks
    Carga
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.Form_Load"
    Exit Sub
Form_Load_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.Form_Load"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmProcNC_AvisosAccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Form_Unload_Error
    
    Set frmProcNC_AvisosAccCorrectivas = Nothing
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.Form_Unload"
    Exit Sub
Form_Unload_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.Form_Unload"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Form_Unload of Formulario frmProcNC_AvisosAccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub
Private Sub actualiza()
    Dim blnCaducada As Boolean, blnEnPreaviso As Boolean
    Dim oUsuario As New clsUsuarios
    Dim oDeco As New clsDecodificadora
    Dim strUsuario As String
    
On Error GoTo Actualiza_Error
    mvarobjAccCorrectoras.Carga CLng(lista.ListItems(lista.selectedItem.Index).SubItems(4))
    blnCaducada = comprobarCaducidad(mvarobjAccCorrectoras.getFECHA_PREVISTA, mvarobjAccCorrectoras.getESTADO_ID)
    
    With lista.ListItems(lista.selectedItem.Index)
    
        If blnCaducada Then .ForeColor = RGB(255, 0, 0)
        .SubItems(1) = mvarobjAccCorrectoras.getTITULO
      
        If blnCaducada Then
             .ListSubItems(1).ForeColor = RGB(255, 0, 0)
        Else
             .ListSubItems(1).ForeColor = RGB(0, 0, 0)
        End If
        oUsuario.CARGAR mvarobjAccCorrectoras.getRESPONSABLE_ID
        strUsuario = oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
        .SubItems(2) = strUsuario
        If blnCaducada Then
          .ListSubItems(2).ForeColor = RGB(255, 0, 0)
        Else
          .ListSubItems(2).ForeColor = RGB(0, 0, 0)
        End If
        
        .SubItems(3) = Format(mvarobjAccCorrectoras.getFECHA_PREVISTA, "dd/mm/yyyy")
         If blnCaducada Then
             .ListSubItems(3).ForeColor = RGB(255, 0, 0)
         Else
             .ListSubItems(3).ForeColor = RGB(0, 0, 0)
         End If
 
        .SubItems(4) = mvarobjAccCorrectoras.getID_ACCION
         If blnCaducada Then
            .ListSubItems(4).ForeColor = RGB(255, 0, 0)
         Else
            .ListSubItems(4).ForeColor = RGB(0, 0, 0)
         End If
 
        .SubItems(5) = Format(mvarobjAccCorrectoras.getFECHA_PUESTA_EN_MARCHA, "dd/mm/yyyy")
        If blnCaducada Then
             .ListSubItems(5).ForeColor = RGB(255, 0, 0)
        Else
             .ListSubItems(5).ForeColor = RGB(0, 0, 0)
        End If

        If mvarobjAccCorrectoras.getESTADO_ID = 7 Or mvarobjAccCorrectoras.getESTADO_ID = 8 Then
            .SubItems(6) = Format(mvarobjAccCorrectoras.getFECHA_RESOLUCION, "dd/mm/yyyy")
         Else
             .SubItems(6) = " - - "
         End If
         If blnCaducada Then
            
             .ListSubItems(6).ForeColor = RGB(255, 0, 0)
         Else
             .ListSubItems(6).ForeColor = RGB(0, 0, 0)
         End If
        oDeco.Carga_valor DECODIFICADORA.PROCNC_ESTADOS_ACCIONES_CORRECTIVAS, mvarobjAccCorrectoras.getESTADO_ID
         
        .SubItems(7) = oDeco.getDESCRIPCION
        If blnCaducada Then
             .ListSubItems(7).ForeColor = RGB(255, 0, 0)
        Else
             .ListSubItems(7).ForeColor = RGB(0, 0, 0)
        End If

        oDeco.Carga_valor DECODIFICADORA.PROCNC_ACCIONES_TIPOS, mvarobjAccCorrectoras.getTIPO_ID
        .SubItems(8) = oDeco.getDESCRIPCION
        If blnCaducada Then
             .ListSubItems(8).ForeColor = RGB(255, 0, 0)
        Else
            .ListSubItems(8).ForeColor = RGB(0, 0, 0)
        End If
        lista.Refresh
    End With
    
    Set oUsuario = Nothing
    Set oDeco = Nothing
Exit Sub
Actualiza_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.actualiza"
    MsgBox Err.Number & " (" & Err.Description & ") in procedure actualiza of Formulario frmProcNC_AvisosAccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR, vbExclamation
    G_TRAZABILIDAD_ERROR = ""
End Sub
Private Sub Carga()
    
    Dim rs As ADODB.Recordset, x As Integer
    Dim blnCaducada As Boolean, blnEnPreaviso As Boolean
    Dim objLitem As ListItem, objSI As ListSubItem
    
On Error GoTo Carga_Error
    Set rs = mvarobjAccCorrectoras.ListadoAcciones(Trim(txtDescripcion), cmbResponsableImplantacion.getPK_SALIDA, chkPrevista.value, fPrevistaDesde, fPrevistaHasta, chkPuesta.value, fPuestaDesde, fPuestaHasta, chkResolucion.value, fResolucionDesde, fResolucionHasta, cmbestados.BoundText, cmbTipo.BoundText, txtCNCi, txtCNCf)
    lista.ListItems.Clear
    On Error Resume Next
    leido = True
    lbltitulo(2) = "Encontrados " & rs.RecordCount & " registros."
    If rs.RecordCount > 0 Then
        Do
'            blnEnPreaviso = (CLng(rs("FECHA_AVISO")) <= CLng(Now))
            'M1331-I
'            blnCaducada = (CLng(rs("FECHA_PREVISTA")) <= CLng(Now))
            blnCaducada = comprobarCaducidad(CLng(rs("FECHA_PREVISTA")), CLng(rs("ESTADO_ID")))
            'M1331-F
            Set objLitem = lista.ListItems.Add(, , Format(rs("ID_PROCNC"), "000000"))
            With objLitem
'                objLitem.SmallIcon = IIf(blnCaducada, "caducado", IIf(blnEnPreaviso, "preaviso", "nada"))
'                If blnEnPreaviso Or blnCaducada Then
'                    .Bold = True
'                    intContador = intContador + 1
'                End If
                If blnCaducada Then .ForeColor = RGB(255, 0, 0)
                Set objSI = .ListSubItems.Add(, , rs("TITULO"))
'                objSI.Bold = blnEnPreaviso Or blnCaducada
                If blnCaducada Then objSI.ForeColor = RGB(255, 0, 0)
                
                Set objSI = .ListSubItems.Add(, , rs("RESPONSABLE"))
'                objSI.Bold = blnEnPreaviso Or blnCaducada
                If blnCaducada Then objSI.ForeColor = RGB(255, 0, 0)
                
                Set objSI = .ListSubItems.Add(, , Format(rs("FECHA_PREVISTA"), "dd/mm/yyyy"))
'                objSI.Bold = blnEnPreaviso Or blnCaducada
                If blnCaducada Then objSI.ForeColor = RGB(255, 0, 0)
                
                Set objSI = .ListSubItems.Add(, , rs("ID_Accion_correctiva"))
'                objSI.Bold = blnEnPreaviso Or blnCaducada
                If blnCaducada Then objSI.ForeColor = RGB(255, 0, 0)
                
                Set objSI = .ListSubItems.Add(, , rs("FECHA_PUESTA_EN_MARCHA"))
'                objSI.Bold = blnEnPreaviso Or blnCaducada
                If blnCaducada Then objSI.ForeColor = RGB(255, 0, 0)
                'M1331-I
                If rs("estado_id") = 7 Or rs("estado_id") = 8 Then
                    Set objSI = .ListSubItems.Add(, , rs("FECHA_RESOLUCION"))
    '                objSI.Bold = blnEnPreaviso Or blnCaducada
                Else
                    Set objSI = .ListSubItems.Add(, , " - - ")
                End If
                If blnCaducada Then objSI.ForeColor = RGB(255, 0, 0)
                'M1331-I
                Set objSI = .ListSubItems.Add(, , rs("ESTADO"))
                If blnCaducada Then objSI.ForeColor = RGB(255, 0, 0)
                
                Set objSI = .ListSubItems.Add(, , rs("TIPO"))
                If blnCaducada Then objSI.ForeColor = RGB(255, 0, 0)
            
                Set objSI = .ListSubItems.Add(, , rs("ESTADO_ID"))
                If blnCaducada Then objSI.ForeColor = RGB(255, 0, 0)
            End With
            'intContador = intContador + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.Carga"
    Exit Sub
Carga_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.Carga"
    MsgBox Err.Number & " (" & Err.Description & ") in procedure Carga of Formulario frmProcNC_AvisosAccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR, vbExclamation
    G_TRAZABILIDAD_ERROR = ""
End Sub
'M1331-I
Private Function comprobarCaducidad(FECHA_PREVISTA As Long, ESTADO As Long) As Boolean
    If ESTADO = 7 Or ESTADO = 8 Then
       comprobarCaducidad = False
       Exit Function
    End If
    comprobarCaducidad = (FECHA_PREVISTA <= CLng(Now))
End Function
'M1331-F
Private Sub cargar_checks()
'    chkPrevista.value = 1
'    chkPuesta.value = 1
'    chkResolucion.value = 1
End Sub
Private Sub cargar_fechas()
    fPrevistaDesde = "01/01/" & Year(Date)
    fPrevistaHasta = "31/12/" & Year(Date)
    fPuestaDesde = "01/01/" & Year(Date)
    fPuestaHasta = Date
    fResolucionDesde = "01/01/" & Year(Date)
    fResolucionHasta = Date
End Sub

Private Sub cargar_combos()
    Dim oDecodificadora As New clsDecodificadora
    llenar_combo cmbResponsableImplantacion, New clsUsuarios, 0, frmUsuarios, ""
    oDecodificadora.cargar_combo cmbestados, DECODIFICADORA.PROCNC_ESTADOS_ACCIONES_CORRECTIVAS
    oDecodificadora.cargar_combo cmbTipo, DECODIFICADORA.PROCNC_ACCIONES_TIPOS
    Set oDecodificadora = Nothing
End Sub
Private Sub cmdcancel_Click()
On Error GoTo cmdcancel_Click_Error
    
    Unload Me
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.cmdcancel_Click"
    Exit Sub
cmdcancel_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.cmdcancel_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdcancel_Click of Formulario frmProcNC_AvisosAccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmdVerPNC_Click()
    If lista.ListItems.Count > 0 Then
        cargar_pnc
    End If
End Sub

Public Sub cargar_mensaje()

Dim objfrm As New frmProcNCEdicion_AccionCorrectiva
Dim strIdAccCorrectora As String
Dim strFecha As String
'Dim oPnc As New clsProcNc

On Error GoTo cargar_mensaje_Error
    'M1331-I
    strIdAccCorrectora = lista.ListItems(lista.selectedItem.Index).SubItems(4)
    objfrm.PK = CLng(strIdAccCorrectora)
    objfrm.PK_PNC = lista.ListItems(lista.selectedItem.Index).Text
    objfrm.NivelAcceso = ACCESO_TOTAL
'    oPnc.Carga CLng(lista.ListItems(lista.selectedItem.Index).Text)
'    objfrm.estado_pnc = oPnc.getESTADO_ID
    objfrm.estado_pnc = lista.ListItems(lista.selectedItem.Index).SubItems(9)
    Set oPnc = Nothing
    objfrm.Show vbModal

    Unload objfrm
    Set objfrm = Nothing
    'M1331-I
    'Carga
    actualiza
    'M1331-I

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.cargar_mensaje"
    Exit Sub
cargar_mensaje_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.cargar_mensaje"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cargar_mensaje of Formulario frmProcNC_AvisosAccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cargar_pnc()
    'Dim objfrm As frmProcNC_Detalle
    Dim objfrm As frmProcNCEdicion
    Dim strIdAccCorrectora As String
    Dim strFecha As String
On Error GoTo cargar_pnc_Error
    strIdAccCorrectora = lista.ListItems(lista.selectedItem.Index)
    'Set objfrm = New frmProcNC_Detalle
    Set objfrm = New frmProcNCEdicion
    objfrm.PK = CLng(strIdAccCorrectora)
    objfrm.Show vbModal
    Unload objfrm
    Set objfrm = Nothing
    Carga
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.cargar_pnc"
    Exit Sub
cargar_pnc_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.cargar_pnc"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cargar_pnc of Formulario frmProcNC_AvisosAccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
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
On Error GoTo lista_DblClick_Error
    If lista.ListItems.Count > 0 Then
        cargar_mensaje
        'cargar_pnc
    End If
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.lista_DblClick"
    Exit Sub
lista_DblClick_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.lista_DblClick"
    MsgBox Err.Number & " (" & Err.Description & ") in procedure lista_DblClick of Formulario frmProcNC_AvisosAccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdBuscar_Click()
    Carga
End Sub

Private Sub cmdVerExcel_Click()
On Error GoTo error_excel
       Me.MousePointer = vbHourglass
       Frame4.Visible = True
       Dim rs As New ADODB.Recordset
       Dim fechaI As String
       Dim fechaF As String
       rs.Fields.Append "c1", adChar, 10, adFldUpdatable
       rs.Fields.Append "c2", adChar, 350, adFldUpdatable
       rs.Fields.Append "c3", adChar, 250, adFldUpdatable
       rs.Fields.Append "c4", adChar, 20, adFldUpdatable
       rs.Fields.Append "c5", adChar, 10, adFldUpdatable
       rs.Fields.Append "c6", adChar, 20, adFldUpdatable
       rs.Fields.Append "c7", adChar, 20, adFldUpdatable
       rs.Fields.Append "c8", adChar, 50, adFldUpdatable
       rs.Fields.Append "c9", adChar, 50, adFldUpdatable
       rs.Open
       
       Dim i As Integer

       For i = 1 To lista.ListItems.Count
           If lista.ListItems(i).Checked = True Then
                rs.AddNew
                rs("c1") = lista.ListItems(i).Text
                rs("c2") = lista.ListItems(i).SubItems(1)
                rs("c3") = lista.ListItems(i).SubItems(2)
                rs("c4") = lista.ListItems(i).SubItems(3)
                rs("c5") = lista.ListItems(i).SubItems(4)
                rs("c6") = lista.ListItems(i).SubItems(5)
                rs("c7") = lista.ListItems(i).SubItems(6)
                rs("c8") = lista.ListItems(i).SubItems(7)
                rs("c9") = lista.ListItems(i).SubItems(8)
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
        XLW.Worksheets(1).Name = "Listado de acciones"
 
        'Cabecera
        With XLS.Range("A1:I1")
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
        With XLS.Range("A1:I1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:I1").Borders
            .LineStyle = vbSolid
        End With
        XLS.Range("A1:A1").ColumnWidth = 10
        XLS.Range("B1:B1").ColumnWidth = 60
        XLS.Range("C1:C1").ColumnWidth = 35
        XLS.Range("D1:D1").ColumnWidth = 15
        XLS.Range("E1:E1").ColumnWidth = 15
        XLS.Range("F1:F1").ColumnWidth = 15
        XLS.Range("G1:G1").ColumnWidth = 15
        XLS.Range("H1:H1").ColumnWidth = 20
        XLS.Range("I1:I1").ColumnWidth = 15
        
        XLS.Cells(1, 1) = "Nº P.N.C."
        XLS.Cells(1, 2) = "Titulo"
        XLS.Cells(1, 3) = "Responsable"
        XLS.Cells(1, 4) = "Fecha Prevista"
        XLS.Cells(1, 5) = "Fecha Comienzo"
        XLS.Cells(1, 6) = "Resolución"
        XLS.Cells(1, 7) = "id_accion"
        XLS.Cells(1, 8) = "Estado"
        XLS.Cells(1, 9) = "Tipo"
         
        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Cells(i, 1) = rs("c1")
            XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True)
            XLS.Cells(i, 3) = ClrStr(rs("c3"), False, True, True)
            XLS.Cells(i, 4) = Format(Trim(rs("c5")), "dd/mm/yyyy")
            XLS.Cells(i, 5) = Format(Trim(rs("c6")), "dd/mm/yyyy")
            XLS.Cells(i, 6) = Format(Trim(rs("c7")), "dd/mm/yyyy")
            XLS.Cells(i, 7) = rs("c4")
            XLS.Cells(i, 8) = rs("c8")
            XLS.Cells(i, 9) = rs("c9")

            i = i + 1
            XLS.Range("A" & i).EntireRow.Insert
            rs.MoveNext
          Loop Until rs.EOF
        End If
        Frame4.Visible = False
        Me.MousePointer = vbNormal
        XLA.Visible = True
        Set rs = Nothing
        Exit Sub
error_excel:
    Frame4.Visible = False
    Me.MousePointer = vbNormal
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_AvisosAccCorrectivas.cmdVerExcel_Click"
    MsgBox Err.Number & " (" & Err.Description & ") in procedure cmdVerExcel_Click of Formulario frmProcNC_AvisosAccCorrectivas" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub txtCNCf_GotFocus()
    txtCNCf.BackColor = &HFFFFCC
End Sub

Private Sub txtCNCf_LostFocus()
    txtCNCf.BackColor = &HFFFFFF
End Sub

Private Sub txtCNCi_GotFocus()
    txtCNCi.BackColor = &HFFFFCC
End Sub

Private Sub txtCNCi_LostFocus()
    txtCNCi.BackColor = &HFFFFFF
    If Trim(txtCNCf.Text) = "" Then
        txtCNCf.Text = txtCNCi.Text
    End If
End Sub
