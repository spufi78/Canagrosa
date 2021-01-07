VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#38.0#0"; "miCombo.ocx"
Begin VB.Form frmObservaciones_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Observaciones de Ensayos"
   ClientHeight    =   9090
   ClientLeft      =   2715
   ClientTop       =   1785
   ClientWidth     =   14475
   Icon            =   "frmObservaciones_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9090
   ScaleWidth      =   14475
   Begin VB.CommandButton cmdImprimir_usuario_en_formacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clasificado por Usuario en Formación"
      Height          =   870
      Left            =   3540
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8190
      Width           =   1710
   End
   Begin VB.CommandButton cmdImprimir_usuario_formador 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clasificado por Formador Cualificado"
      Height          =   870
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8190
      Width           =   1710
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
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
      Height          =   1560
      Left            =   30
      TabIndex        =   5
      Top             =   855
      Width           =   14400
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recualificaciones"
         Height          =   885
         Left            =   11520
         TabIndex        =   21
         Top             =   240
         Width           =   1545
         Begin VB.OptionButton optRecual_Todos 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cualquiera"
            Height          =   255
            Left            =   90
            TabIndex        =   24
            Top             =   540
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optRecual_No 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            Height          =   255
            Left            =   660
            TabIndex        =   23
            Top             =   240
            Width           =   525
         End
         Begin VB.OptionButton optRecual_SI 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sí"
            Height          =   255
            Left            =   90
            TabIndex        =   22
            Top             =   240
            Width           =   765
         End
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   990
         Left            =   13230
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   330
         Width           =   1050
      End
      Begin MSComCtl2.DTPicker txtDesde 
         Height          =   345
         Left            =   9690
         TabIndex        =   14
         Top             =   330
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   53608449
         CurrentDate     =   40357
      End
      Begin pryCombo.miCombo cmbFormador 
         Height          =   315
         Left            =   2100
         TabIndex        =   7
         Top             =   210
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   556
      End
      Begin pryCombo.miCombo cmbUsuarioEnFormacion 
         Height          =   315
         Left            =   2100
         TabIndex        =   8
         Top             =   540
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   556
      End
      Begin pryCombo.miCombo cmbPNT 
         Height          =   315
         Left            =   2100
         TabIndex        =   11
         Top             =   870
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   556
      End
      Begin MSComCtl2.DTPicker txtHasta 
         Height          =   345
         Left            =   9690
         TabIndex        =   16
         Top             =   690
         Width           =   1725
         _ExtentX        =   3043
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CheckBox        =   -1  'True
         Format          =   53608449
         CurrentDate     =   40357
      End
      Begin pryCombo.miCombo cmbTipoMuestra 
         Height          =   315
         Left            =   2100
         TabIndex        =   17
         Top             =   1200
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   556
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Muestra"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   18
         Top             =   1260
         Width           =   930
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hasta"
         Height          =   195
         Index           =   5
         Left            =   9120
         TabIndex        =   15
         Top             =   840
         Width           =   420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Documento Calidad"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   12
         Top             =   930
         Width           =   1395
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   2
         Left            =   9120
         TabIndex        =   10
         Top             =   450
         Width           =   465
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario En Formacion"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1560
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Formador"
         Height          =   195
         Index           =   3
         Left            =   105
         TabIndex        =   6
         Top             =   300
         Width           =   660
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Clasificado por Doc. Calidad"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8190
      Width           =   1710
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13440
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8160
      Width           =   990
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5670
      Left            =   45
      TabIndex        =   0
      Top             =   2445
      Width           =   14385
      _ExtentX        =   25374
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Observaciones de Ensayo"
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
      TabIndex        =   4
      Top             =   135
      Width           =   3915
   End
   Begin VB.Image imagen 
      Height          =   720
      Left            =   13650
      Picture         =   "frmObservaciones_Listado.frx":1272
      Top             =   60
      Width           =   720
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "En la lista existen un total de 0 registros"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   405
      Width           =   2775
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   14430
   End
End
Attribute VB_Name = "frmObservaciones_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



Private Sub cmdBuscar_Click()

    cargar_lista
    
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub


Private Sub cmdImprimir_Click()
    With frmReport
        .iniciar
        .informe = "/MC/rptMC_informe_pnt"
        .CRITERIO = "" ' "{procnc.ID_PROCNC} = " & id_pnc & " and {decodificadora.CODIGO}=110" '"{decodificadora.codigo}=11 and {nc.ID_NC} in [" & Left(strIncidencias, Len(strIncidencias) - 1) & "]"
        .imprimir = False
        .generar
        .Visible = True
    End With
End Sub



Private Sub cmdImprimir_usuario_formador_Click()
    With frmReport
        .iniciar
        .informe = "/MC/rptMC_informe_formador"
        .CRITERIO = "" ' "{procnc.ID_PROCNC} = " & id_pnc & " and {decodificadora.CODIGO}=110" '"{decodificadora.codigo}=11 and {nc.ID_NC} in [" & Left(strIncidencias, Len(strIncidencias) - 1) & "]"
        .imprimir = False
        .generar
        .Visible = True
    End With
End Sub


Private Sub cmdImprimir_usuario_en_formacion_Click()
    With frmReport
        .iniciar
        .informe = "/MC/rptMC_informe_en_formacion"
        .CRITERIO = "" ' "{procnc.ID_PROCNC} = " & id_pnc & " and {decodificadora.CODIGO}=110" '"{decodificadora.codigo}=11 and {nc.ID_NC} in [" & Left(strIncidencias, Len(strIncidencias) - 1) & "]"
        .imprimir = False
        .generar
        .Visible = True
    End With
End Sub



Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub
Private Sub Form_Load()
    log (Me.Name)
    
    cargar_botones Me

    Set cmdImprimir_usuario_formador.Picture = cmdImprimir.Picture
    Set cmdImprimir_usuario_en_formacion.Picture = cmdImprimir.Picture
    
    
    Me.Left = 100
    Me.Top = 100
    
    cabecera
    
    llenar_combo cmbFormador, USUARIO, 0, frmUsuarios, ""
    llenar_combo cmbUsuarioEnFormacion, USUARIO, 0, frmUsuarios, ""
    llenar_combo cmbPNT, New clsCa_documentos, 0, frmCA_Documento, ""
    llenar_combo cmbTipoMuestra, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
    
    
    cargar_lista
    
    
End Sub
Private Sub cargar_lista()
    
    Dim rs As New ADODB.RecordSet
    Dim oMo As New clsMuestras_observadores
    
    
    
    Dim fd As String, fh As String
    
    If IsNull(txtDesde.value) Then
        fd = ""
    Else
        fd = txtDesde.value
    End If
    
    If IsNull(txtHasta.value) Then
        fh = ""
    Else
        fh = txtHasta.value
    End If
    
    
'    Set rs = oMo.Listado_Filtrado(cmbFormador.getPK_SALIDA, cmbUsuarioEnFormacion.getPK_SALIDA, cmbPNT.getPK_SALIDA, cmbTipoMuestra.getPK_SALIDA, fd, fh, optRecual_Todos, optRecual_SI)
    lista.ListItems.Clear
    lbltitulo(1) = "En la lista existen un total de " & rs.RecordCount & " registros."
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("MUESTRA"))
                .SubItems(1) = Format(rs("FECHA"), "dd/mm/yyyy")
                .SubItems(2) = rs("DOCUMENTO")
                If CInt(rs("ACTOR_ES_FORMADOR")) = 1 Then
                    .SubItems(3) = rs("USUARIO")
                    .SubItems(4) = rs("USUARIO_OBSERVADOR")
                Else
                    .SubItems(3) = rs("USUARIO_OBSERVADOR")
                    .SubItems(4) = rs("USUARIO")
                End If
                .SubItems(5) = rs("USUARIO")
                If CInt(rs("ES_RECUALIFICACION")) = 1 Then
                    .SubItems(6) = "[X]"
                'Else
                '    .SubItems(6) = ""
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oMo = Nothing
    
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


Private Sub cabecera()
    
    With lista.ColumnHeaders.Add(, , "Muestra", 800, lvwColumnLeft)
        .Tag = "Muestra"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 800, lvwColumnLeft)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Doc. Calidad", 3000, lvwColumnLeft)
        .Tag = "Doc. Calidad"
    End With
    With lista.ColumnHeaders.Add(, , "Formador Cualificado", 2600, lvwColumnLeft)
        .Tag = "Formador Culaificado"
    End With
    With lista.ColumnHeaders.Add(, , "En Formación", 2600, lvwColumnLeft)
        .Tag = "Formador Culaificado"
    End With
    With lista.ColumnHeaders.Add(, , "Realiza Ensayo", 2600, lvwColumnLeft)
        .Tag = "Realiza Ensayo"
    End With
    With lista.ColumnHeaders.Add(, , "Recual.", 800, lvwAutoLeft)
        .Tag = "Recual."
    End With
    
    txtDesde.value = ""
    txtHasta.value = ""
    
    
    
    
End Sub

