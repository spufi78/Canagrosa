VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmFormacion_PFA_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del plan de formación anual"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12900
   Icon            =   "frmFormacion_PFA_Detalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   12900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Formación (Tipo)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9090
      TabIndex        =   33
      Top             =   2205
      Width           =   3750
      Begin VB.OptionButton optFormacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Práctica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2160
         TabIndex        =   35
         Top             =   270
         Width           =   1140
      End
      Begin VB.OptionButton optFormacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Teórica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   630
         TabIndex        =   34
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nivel de formación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4095
      TabIndex        =   26
      Top             =   2205
      Width           =   4965
      Begin VB.OptionButton optNivel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Específica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3510
         TabIndex        =   29
         Top             =   270
         Width           =   1320
      End
      Begin VB.OptionButton optNivel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1980
         TabIndex        =   28
         Top             =   270
         Width           =   1140
      End
      Begin VB.OptionButton optNivel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Técnica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   27
         Top             =   270
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modalidad"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   45
      TabIndex        =   23
      Top             =   2205
      Width           =   4020
      Begin VB.OptionButton optModalidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Externa (RFI)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   2385
         TabIndex        =   36
         Top             =   270
         Width           =   1545
      End
      Begin VB.OptionButton optModalidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Interna"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   25
         Top             =   270
         Width           =   1095
      End
      Begin VB.OptionButton optModalidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Externa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1260
         TabIndex        =   24
         Top             =   270
         Width           =   1140
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones"
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
      Height          =   1050
      Left            =   45
      TabIndex        =   19
      Top             =   8055
      Width           =   6945
      Begin VB.TextBox txtObservaciones 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   135
         MaxLength       =   75
         TabIndex        =   20
         Top             =   270
         Width           =   6720
      End
   End
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Planes de Formación Interna"
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
      Height          =   5100
      Left            =   45
      TabIndex        =   11
      Top             =   2880
      Width           =   6945
      Begin MSComctlLib.ListView ListaPlanes 
         Height          =   4725
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   8334
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
   End
   Begin VB.Frame frameBotones 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   7065
      TabIndex        =   8
      Top             =   8100
      Width           =   4470
      Begin VB.CommandButton cmdCurso 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver RFI"
         Height          =   915
         Left            =   360
         TabIndex        =   13
         Top             =   90
         Width           =   1365
      End
      Begin VB.CommandButton cmdok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Height          =   915
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   90
         Width           =   1275
      End
      Begin VB.CommandButton cmdRFI 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Generar RFI"
         Height          =   915
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   90
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   45
      TabIndex        =   5
      Top             =   720
      Width           =   12795
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Left            =   1620
         MaxLength       =   500
         TabIndex        =   6
         Top             =   900
         Width           =   7080
      End
      Begin XtremeSuiteControls.DateTimePicker cmbFecha 
         Height          =   330
         Left            =   11115
         TabIndex        =   21
         Top             =   945
         Width           =   1500
         _Version        =   851970
         _ExtentX        =   2646
         _ExtentY        =   582
         _StockProps     =   68
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   1
      End
      Begin VB.Label lblIDCurso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   330
         Left            =   11115
         TabIndex        =   32
         Top             =   135
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label lblCurso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   11115
         TabIndex        =   31
         Top             =   450
         Width           =   1455
      End
      Begin VB.Label lblRFI 
         BackColor       =   &H00C0C0C0&
         Caption         =   "RFI asociado:"
         Height          =   285
         Left            =   9765
         TabIndex        =   30
         Top             =   495
         Width           =   1050
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha prevista:"
         Height          =   195
         Left            =   9720
         TabIndex        =   22
         Top             =   990
         Width           =   1185
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Plan de formación:"
         Height          =   195
         Left            =   180
         TabIndex        =   18
         Top             =   990
         Width           =   1320
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modalidad:"
         Height          =   195
         Left            =   5940
         TabIndex        =   17
         Top             =   450
         Width           =   825
      End
      Begin VB.Label lblModalidad 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   6840
         TabIndex        =   16
         Top             =   405
         Width           =   1860
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código:"
         Height          =   195
         Left            =   945
         TabIndex        =   15
         Top             =   495
         Width           =   555
      End
      Begin VB.Label lblID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   1620
         TabIndex        =   14
         Top             =   405
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documentos/PNTs del Plan de formación"
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
      Height          =   5115
      Left            =   7065
      TabIndex        =   3
      Top             =   2880
      Width           =   5775
      Begin VB.CommandButton cmdeliminar 
         BackColor       =   &H00E0E0E0&
         Height          =   480
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4545
         Width           =   660
      End
      Begin MSComctlLib.ListView listaDocs 
         Height          =   4185
         Left            =   90
         TabIndex        =   4
         Top             =   270
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   7382
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
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   915
      Left            =   11610
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8190
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   12285
      Top             =   6075
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12195
      Top             =   6570
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_PFA_Detalle.frx":08CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_PFA_Detalle.frx":11A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_PFA_Detalle.frx":1A7E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   12195
      Picture         =   "frmFormacion_PFA_Detalle.frx":2358
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del plan de formación anual"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   2490
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Plan de formación"
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
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   1890
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   12870
   End
End
Attribute VB_Name = "frmFormacion_PFA_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long


Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera

    If PK <> 0 Then   'Modificación
        cargar_campos
       ' cargar_lista_documentos
    Else
        'Alta
        If Not usuario.getPER_RFI Then
            cmdCurso.Enabled = False
            cmdRFI.Enabled = False
        Else
            cmdCurso.Enabled = True
            cmdRFI.Enabled = True
        End If
        cargar_campos_alta
    End If
    
    Dim oDeco As New clsDecodificadora
    'oDeco.cargar_combo cmbFecha, DECODIFICADORA.DECODIFICADORA_MESES
End Sub

Private Sub cabecera()
        With ListaPlanes.ColumnHeaders
            .Add , , "ID Plan", 1000, lvwColumnLeft
            .Add , , "Descripcion", 5450, lvwColumnLeft
        End With
        With listaDocs.ColumnHeaders
            .Add , , "ID", 1, lvwColumnLeft
            .Add , , "Código", 1000, lvwColumnLeft
            .Add , , "Documento", listaDocs.Width, lvwColumnLeft
        End With
        
End Sub

'BLOQUE DE FUNCIONES ALTA
Private Sub cargar_campos_alta()
    txtDescripcion.Text = ""
    txtObservaciones.Text = ""
'M1164-I
    txtDescripcion.MaxLength = 500
'M1164-F
    cmbFecha.value = Date
    optModalidad(0).value = True
    optNivel(0).value = True
    optFormacion(0).value = True
    CargarListaPlanes
    
 End Sub

Private Sub Alta()
    On Error GoTo fallo
    If txtDescripcion.Text = "" Then
        MsgBox "Indique una descripción para el plan de formación.", vbExclamation, App.Title
        Exit Sub
    End If
    If ListaPlanes.ListItems.Count = 0 Then
        MsgBox "Seleccione un curso/documentación sobre el que generar el plan", vbExclamation, App.Title
        Exit Sub
    End If
    
    Dim PLAN As Long
    Dim oPF As New clsFormacion_pfa
    With oPF
        .CrearIdPlan
        .setCURSO_ID = 0
        .setPLAN_FORMACION_ID = CLng(lblID.Caption)
        .setFECHA_PREVISTA = Format(cmbFecha.value, "yyyy-mm-dd")
        .setOBSERVACIONES = Trim(txtObservaciones.Text)
        .setDESCRIPCION = txtDescripcion
        .setANYO = Format(cmbFecha.value, "yyyy")
        If optModalidad(0).value = True Then
            .setMODALIDAD = 0
        Else
            .setMODALIDAD = 1
        End If
        If optNivel(0).value = True Then
            .setNIVEL = 0
        Else
            If optNivel(1).value = True Then
                .setNIVEL = 1
            Else
                .setNIVEL = 2
            End If
        End If
        If optFormacion(0).value = True Then
            .setFORMACION = 0
        Else
            .setFORMACION = 1
        End If
        
        PLAN = .Insertar
    End With
    
    PK = PLAN
    
    cmdRFI.Enabled = False
    Set oCA = Nothing
    Set oPlanDocs = Nothing
    MsgBox "Plan creado correctamente.", vbInformation, App.Title
    cmdRFI.Enabled = True
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Alta of Formulario frmFormacion_PlanAnual_Detalle"
End Sub

Private Sub CargarListaPlanes()

    Dim oPF As New clsFormacion_pf
    Dim rsPF As New ADODB.Recordset
    
    Set rsPF = oPF.ListadoFiltro("", "")
    
    If rsPF.RecordCount > 0 Then
        Do
            With ListaPlanes.ListItems.Add(, , rsPF("ID_PLAN_FORMACION"))
                 .SubItems(1) = rsPF("DESCRIPCION")

            End With
      
            rsPF.MoveNext
        Loop Until rsPF.EOF
    End If
    
    Set rsPF = Nothing
    Set oPF = Nothing
End Sub

Private Sub listaDocs_DblClick()
    If listaDocs.ListItems.Count = 0 Then Exit Sub
    Dim oca_documento As New clsCa_documentos
    oca_documento.mostrar listaDoc.ListItems(listaDoc.selectedItem.Index).Text, True
    Set oca_documento = Nothing
End Sub

Private Sub ListaPlanes_DblClick()
'Carga de la lista de documentos
    If ListaPlanes.ListItems.Count = 0 Then Exit Sub
    Dim oDocs As New clsFormacion_pf_docs
    Dim documento As New clsCa_documentos
    Dim rsDocs As New ADODB.Recordset
    
    Set rsDocs = oDocs.Listado_Plan(CLng(ListaPlanes.ListItems(ListaPlanes.selectedItem.Index).Text))
    lblID.Caption = ListaPlanes.ListItems(ListaPlanes.selectedItem.Index).Text
    txtDescripcion.Text = ListaPlanes.ListItems(ListaPlanes.selectedItem.Index).SubItems(1)
    
    If rsDocs.RecordCount > 0 Then
        listaDocs.ListItems.Clear
        Do
            With listaDocs.ListItems.Add(, , rsDocs("DOCUMENTO_ID"))
                 documento.Carga rsDocs("DOCUMENTO_ID")
                 .SubItems(1) = "(" & documento.getCODIGO & ")"
                 .SubItems(2) = documento.getNOMBRE
            End With
      
            rsDocs.MoveNext
        Loop Until rsDocs.EOF
    End If
    Set rsDocs = Nothing
    Set documento = Nothing
    Set oDocs = Nothing

End Sub

'BLOQUE DE FUNCIONES MODIFICACION
Private Sub cargar_campos()

    Dim oPlanFA As New clsFormacion_pfa
    Dim oPlan As New clsFormacion_pf
    Dim i As Integer

    oPlanFA.Carga PK
    oPlan.Carga oPlanFA.getPLAN_FORMACION_ID
    
    lblID.Caption = oPlanFA.getID_PFA
    cmbFecha.value = oPlanFA.getFECHA_PREVISTA
    txtDescripcion.Text = Trim(oPlanFA.getDESCRIPCION)
    txtObservaciones.Text = Trim(oPlanFA.getOBSERVACIONES)

    ListaPlanes.ListItems.Clear
    With ListaPlanes.ListItems.Add(, , oPlan.getID_PLAN_FORMACION)
        .SubItems(1) = oPlan.getDESCRIPCION
    End With
    
    ListaPlanes.Enabled = False
    listaDocs.Enabled = True
    
    If oPlanFA.getCURSO_ID > 0 Then
        Dim oCurso As New clsFormacion_cursos
        oCurso.Carga oPlanFA.getCURSO_ID
        strCurso = "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & oCurso.getANYO
        lblCurso.Caption = strCurso
        lblIDCurso.Caption = oPlanFA.getCURSO_ID
        cmdCurso.Enabled = True
        cmdRFI.Enabled = False
        cmbFecha.value = oPlanFA.getFECHA_PREVISTA
    Else
        cmdCurso.Enabled = False
        cmdRFI.Enabled = True
    End If
    
    If oPlanFA.getMODALIDAD = 0 Then
        lblModalidad.Caption = "Interna"
        optModalidad(0).value = True
    Else
        lblModalidad.Caption = "Externa"
        optModalidad(1).value = True
    End If
    
    If oPlanFA.getNIVEL = 0 Then
        optNivel(0).value = True
    Else
        If oPlanFA.getNIVEL = 1 Then
            optNivel(1).value = True
        Else
            optNivel(2).value = True
        End If
    End If
    
    If oPlanFA.getFORMACION = 0 Then
        optFormacion(0).value = True
    Else
        optFormacion(1).value = True
    End If
    
    ' Carga de la lista de documentos
    Dim oDocs As New clsFormacion_pf_docs
    Dim oCADoc As New clsCa_documentos
    Dim rsDocs As New ADODB.Recordset
    Set rsDocs = oDocs.Listado_Plan(oPlan.getID_PLAN_FORMACION)
    
    listaDocs.ListItems.Clear
    
    If rsDocs.RecordCount > 0 Then
        Do
            With listaDocs.ListItems.Add(, , rsDocs("DOCUMENTO_ID"))
                      oCADoc.Carga rsDocs("DOCUMENTO_ID")
                     .SubItems(1) = "(" & oCADoc.getCODIGO & ") "
                     .SubItems(2) = oCADoc.getNOMBRE
            End With
            rsDocs.MoveNext
        Loop Until rsDocs.EOF
    End If
    
    Set oCADoc = Nothing
    Set oDocs = Nothing
    Set rsDocs = Nothing
    
End Sub

Private Sub MODIFICACION()
    On Error GoTo fallo
    If txtDescripcion.Text = "" Then
        MsgBox "Indique una descripción para el plan de formación.", vbExclamation, App.Title
        Exit Sub
    End If
    
    Dim PLAN As Long
    Dim oPF As New clsFormacion_pfa
    With oPF
        .Carga PK
        .setFECHA_PREVISTA = Format(cmbFecha.value, "yyyy-mm-dd")
        .setANYO = Format(cmbFecha.value, "yyyy")
        .setOBSERVACIONES = Trim(txtObservaciones.Text)
        .setDESCRIPCION = txtDescripcion
        If optModalidad(0).value = True Then
           .setMODALIDAD = 0
        Else
            .setMODALIDAD = 1
        End If
        If optNivel(0).value = True Then
           .setNIVEL = 0
        Else
            If optNivel(1).value = True Then
                .setNIVEL = 1
            Else
                .setNIVEL = 2
            End If
        End If
        If optFormacion(0).value = True Then
           .setFORMACION = 0
        Else
           .setFORMACION = 1
        End If
        .Modificar PK
    End With

    Dim oPlanDocs As New clsFormacion_pf_docs
    oPlanDocs.Eliminar PLAN
    If PLAN > 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            oPlanDocs.setDOCUMENTO_ID = CLng(lista.ListItems(i).Text)
            oPlanDocs.setPLAN_FORMACION_ID = PLAN
            oPlanDocs.Insertar
        Next
    End If
    Set oPlanDocs = Nothing
    MsgBox "Plan modificado correctamente.", vbInformation, App.Title
    cmdRFI.Enabled = True
    Exit Sub
fallo:
         MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Modificacion of Formulario frmFormacion_PlanAnual_Detalle"
End Sub

'FUNCIONES/EVENTOS
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCurso_Click()
   If lblIDCurso.Caption <> "" Then
    frmFormacion_Curso.PK = CLng(lblIDCurso.Caption)
    frmFormacion_Curso.PLAN = PK
    frmFormacion_Curso.Show 1
   End If
End Sub

Private Sub cmdEliminar_Click()
    If listaDocs.ListItems.Count > 0 Then
        listaDocs.ListItems.Remove listaDocs.selectedItem.Index
    End If
End Sub

Private Sub cmdok_Click()
    If PK = 0 Then
        Alta
    Else
        MODIFICACION
    End If
    Unload Me
End Sub

Private Sub cmdRFI_Click()
    With frmFormacion_Curso
     .PK = 0
     .PLAN = PK
     .txtDescripcion.Text = txtDescripcion.Text
    
     If optNivel(0).value = True Then
            .optNivel(0).value = True
     Else
         If optNivel(1).value = True Then
             .optNivel(1).value = True
         Else
             .optNivel(2).value = True
         End If
     End If
     If optFormacion(0).value = True Then
         .optModalidad(0).value = True
     Else
         .optModalidad(1).value = True
     End If
     .fechaPrevistaI = cmbFecha.value
     .fechaPrevistaF = cmbFecha.value + 1
     .fechaRealI = cmbFecha.value
     .fechaRealF = cmbFecha.value + 1
     .Frame2.Enabled = False
     .Frame4.Enabled = False
     .lblExterna.Visible = True
     
     Select Case True
     Case optModalidad(0).value
          .lblExterna.Caption = "F. Interna"
          .chkExterno.value = 0
     Case optModalidad(1).value
          .lblExterna.Caption = "F. Externa"
          .chkExterno.value = 1
     Case optModalidad(2).value
          .lblExterna.Caption = "F. Externa"
          .chkExterno.value = 1
     End Select
     
     .chkExterno.Enabled = False
     .Show 1
    End With
    If lblIDCurso <> "" Then
        Dim oPF As New clsFormacion_pfa
        oPF.Actualizar_Curso PK, CLng(lblIDCurso.Caption)
        Set oPF = Nothing
        MsgBox "El curso " & lblCurso.Caption & " se ha vinculado correctamente al Plan Nº: " & PK, vbInformation + vbOKOnly, App.Title
        cmdCurso.Enabled = True
        cmdRFI.Enabled = False
    End If
End Sub

Private Sub optModalidad_Click(Index As Integer)

     Select Case True
     Case optModalidad(0).value
          lblModalidad.Caption = "F. Interna"
          cmdRFI.Enabled = True
     Case optModalidad(1).value
          If Trim(lblCurso.Caption) <> "" Then
             optModalidad(2).value = True
             cmdRFI.Enabled = True
          Else
            cmdRFI.Enabled = False
          End If
          lblModalidad.Caption = "F. Externa"
          
     Case optModalidad(2).value
          lblModalidad.Caption = "F. Externa"
          cmdRFI.Enabled = True
     End Select

End Sub
