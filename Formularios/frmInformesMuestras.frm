VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#39.0#0"; "miCombo.ocx"
Begin VB.Form frmInformesMuestras 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informes de Muestras"
   ClientHeight    =   8970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12420
   Icon            =   "frmInformesMuestras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8970
   ScaleWidth      =   12420
   Begin VB.CommandButton cmdCerrarSinInforme 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar Sin Informe"
      Height          =   870
      Left            =   3381
      Picture         =   "frmInformesMuestras.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   8055
      Width           =   1185
   End
   Begin VB.CommandButton cmdMarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar Todas"
      Height          =   870
      Left            =   30
      Picture         =   "frmInformesMuestras.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdDesmarcar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Desmarcar Todas"
      Height          =   870
      Left            =   1102
      Picture         =   "frmInformesMuestras.frx":15D6
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdDeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registro"
      Height          =   870
      Left            =   4590
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8055
      Width           =   1185
   End
   Begin VB.CommandButton cmdCerrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar Muestras"
      Enabled         =   0   'False
      Height          =   870
      Left            =   2174
      Picture         =   "frmInformesMuestras.frx":18E0
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8055
      Width           =   1185
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   11310
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8085
      Width           =   1050
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
      Height          =   1920
      Left            =   60
      TabIndex        =   18
      Top             =   360
      Width           =   12330
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
         Left            =   4380
         TabIndex        =   5
         Top             =   1050
         Width           =   780
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1245
         Left            =   6030
         TabIndex        =   25
         Top             =   585
         Width           =   4365
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Muestras cerradas sin informe"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   0
            Left            =   180
            TabIndex        =   36
            Top             =   900
            Width           =   3810
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Muestras cerradas"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   9
            Top             =   675
            Width           =   3810
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Muestras sin terminar"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   8
            Top             =   450
            Width           =   3075
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Muestras terminadas pdtes. de cerrar"
            ForeColor       =   &H80000008&
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   7
            Top             =   225
            Value           =   -1  'True
            Width           =   3975
         End
      End
      Begin VB.TextBox txtcopias 
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
         Height          =   360
         Left            =   1890
         TabIndex        =   6
         Text            =   "1"
         Top             =   1440
         Width           =   810
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
         Left            =   10125
         TabIndex        =   0
         Top             =   315
         Width           =   1095
      End
      Begin VB.TextBox txtp1 
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
         Height          =   360
         Left            =   1890
         TabIndex        =   3
         Top             =   1050
         Width           =   810
      End
      Begin VB.TextBox txtp2 
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
         Height          =   360
         Left            =   3090
         TabIndex        =   4
         Top             =   1050
         Width           =   705
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Height          =   975
         Left            =   10740
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   690
         Width           =   1410
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   375
         Left            =   5160
         TabIndex        =   30
         Top             =   1050
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   661
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196615
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
         Left            =   1890
         TabIndex        =   1
         Top             =   660
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   73007105
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   4140
         TabIndex        =   2
         Top             =   660
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   73007105
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbTiposMuestra 
         Height          =   330
         Left            =   1890
         TabIndex        =   35
         Top             =   270
         Width           =   8160
         _ExtentX        =   14393
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepcionadas desde"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   33
         Top             =   750
         Width           =   1635
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   4
         Left            =   3660
         TabIndex        =   32
         Top             =   750
         Width           =   405
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   195
         Index           =   0
         Left            =   3930
         TabIndex        =   31
         Top             =   1140
         Width           =   585
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número de copias"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   29
         Top             =   1500
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   24
         Top             =   330
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo, desde"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   23
         Top             =   1110
         Width           =   1065
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
         Height          =   195
         Index           =   7
         Left            =   2850
         TabIndex        =   22
         Top             =   1140
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar Informes"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   10470
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6630
      Visible         =   0   'False
      Width           =   2085
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5415
      Left            =   45
      TabIndex        =   11
      Top             =   2610
      Width           =   12360
      _ExtentX        =   21802
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
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
   Begin VB.Frame Frame 
      BackColor       =   &H00C0C0C0&
      Height          =   795
      Left            =   5895
      TabIndex        =   26
      Top             =   8085
      Visible         =   0   'False
      Width           =   5055
      Begin MSComctlLib.ProgressBar pb 
         Height          =   315
         Left            =   1725
         TabIndex        =   27
         Top             =   270
         Width           =   3120
         _ExtentX        =   5503
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Generando informes"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   28
         Top             =   330
         Width           =   1455
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Informes de Muestras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   4
      Left            =   45
      TabIndex        =   21
      Top             =   0
      Width           =   12345
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
      Left            =   60
      TabIndex        =   20
      Top             =   2295
      Width           =   12330
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse sobre el análisis para ver el detalle"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   3
      Left            =   5895
      TabIndex        =   19
      Top             =   8370
      Width           =   4095
   End
End
Attribute VB_Name = "frmInformesMuestras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub chkTodas_Click()
    If chkTodas.value = Checked Then
        cmbTiposMuestra.Limpiar
        cmbTiposMuestra.desactivar
    Else
        cmbTiposMuestra.activar
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCerrar_Click()
    If MsgBox("Va a cerrar las muestras marcadas. ¿Desea Continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    Dim i As Integer
    Me.MousePointer = 11
    Dim oMuestra As New clsMuestra
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            oMuestra.Cerrar (lista.ListItems(i).SubItems(7))
        End If
    Next
    Set oMuestra = Nothing
    Call buscar
    Me.MousePointer = 0
End Sub

Private Sub cmdCerrarSinInforme_Click()
    If MsgBox("Va a cerrar las muestras marcadas SIN INFORME. ¿Desea Continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    Dim i As Integer
    Me.MousePointer = 11
    Dim oMuestra As New clsMuestra
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If (oMuestra.CerrarSinInforme(lista.ListItems(i).SubItems(7)) = False) Then
                Exit For
            End If
        End If
    Next
    Set oMuestra = Nothing
    Call buscar
    Me.MousePointer = 0

End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdDeter_Click()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.SelectedItem.Index).SubItems(7)
        Dim oMuestra As New clsMuestra
        oMuestra.CargaMuestra (gmuestra)
        Select Case oMuestra.getANALISIS_MODIFICADO
            Case 2 ' Control de eficacia
                With frmCE_Resultados
                    .PK_ID_MUESTRA = gmuestra
                    .Show 1
                End With
            Case 3 ' Sellante
                frmSE_Resultados.Show 1
            Case Else
                frmDeterminaciones.Show 1
        End Select
        gmuestra = 0
    End If
End Sub

Private Sub cmdListado_Click()
    Dim i As Integer
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim c As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            c = c + 1
        End If
    Next
    If MsgBox("Va a generar " & c & " informes de muestras. ¿Desea Continuar?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
    End If
    pb.min = 1
    If c = pb.min Then
        pb.Max = pb.min + 1
    Else
        pb.Max = c
    End If
    pb.value = 1
    Me.MousePointer = 11
    Dim COPIAS As Integer
    If IsNumeric(txtcopia) = False Then
        MsgBox "El numero de copias debe ser numérico.", vbCritical, App.Title
        txtcopias.SetFocus
        Exit Sub
    End If
    If txtcopias = "" Then
        COPIAS = 1
    Else
        COPIAS = CInt(txtcopias)
    End If
    For i = 1 To lista.ListItems.Count
     lblCampos(1).Caption = "Generando informes"
     Frame.Visible = True
     If lista.ListItems(i).Checked = True Then
'          If generar_informe(CLng(lista.ListItems(i).SubItems(6)), 1, copias) = True Then
'              omuestra.aumentar_edicion_impresa (CLng(lista.ListItems(i).SubItems(6)))
'          End If
     End If
     If pb.value < pb.Max Then
        pb.value = pb.value + 1
     Else
        pb.value = pb.Max
     End If
    Next
    Frame.Visible = False
    Me.MousePointer = 0
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
        cmdCancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.Top = 50
    txtanno = Year(Date)
    fdesde = Date
    fhasta = Date
    cabecera
    llenar_combo cmbTiposMuestra, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Código", 1100, lvwColumnLeft)
        .Tag = "Código"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 2500, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Analisis", 2400, lvwColumnLeft)
        .Tag = "Analisis"
    End With
    With lista.ColumnHeaders.Add(, , "Ref.Cliente", 2500, lvwColumnLeft)
        .Tag = "Ref.Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1300, lvwColumnCenter)
        .Tag = "Fecha"
    End With
        With lista.ColumnHeaders.Add(, , "Precio", 1300, lvwColumnCenter)
        .Tag = "Precio"
    End With
    With lista.ColumnHeaders.Add(, , "General", 800, lvwColumnCenter)
        .Tag = "General"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim i As Integer
    Dim consulta As String
    Dim strMuestra As String
    Dim strpar As String
    Dim stranno As String
    Dim strTipo As String
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim RS As New ADODB.RecordSet
    cmdCerrar.Enabled = False
    ' Tipo de muestra
    strMuestra = ""
    If chkTodas.value = Unchecked Then
        If cmbTiposMuestra.getTEXTO = "" Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Sub
        End If
        strMuestra = " AND mu.tipo_muestra_id=" & cmbTiposMuestra.getPK_SALIDA
    End If
    ' Fechas
    Dim FECHA_DESDE As String
    FECHA_DESDE = " AND mu.fecha_recepcion>='" & Format(fdesde, "yyyy-mm-dd") & "'"
    Dim FECHA_HASTA As String
    FECHA_HASTA = " AND mu.fecha_recepcion<='" & Format(fhasta, "yyyy-mm-dd") & "'"
   
    ' Particular
    strpar = ""
    If txtp1 <> "" Or txtp2 <> "" Then
        If txtp1 = "" Or txtp2 = "" Then
            MsgBox "Debe completar los codigos de búsqueda.", vbInformation, App.Title
            Exit Sub
        Else
            If IsNumeric(txtp1) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp1.SetFocus
                Exit Sub
            End If
            If IsNumeric(txtp2) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp2.SetFocus
                Exit Sub
            End If
            strpar = " AND mu.id_particular between " & CLng(txtp1) & " and " & CLng(txtp2)
        End If
    End If
    If strpar <> "" Then
        stranno = " and mu.anno = " & CInt(txtanno)
    End If
    strTipo = ""
    If opTipo(1).value = True Then
        strTipo = " and mu.cerrada = 2 "
    ElseIf opTipo(2).value = True Then
        strTipo = " and mu.cerrada = 0 "
    ElseIf opTipo(0).value = True Then
        strTipo = " and mu.cerrada = 3 "
    Else
        strTipo = " and mu.cerrada = 1 "
    End If
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "ta.nombre, " & _
               "mu.id_general " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                     "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                     "mu.tipo_analisis_id=ta.id_tipo_analisis AND " & _
                     "mu.anulada = 0 " & _
                     FECHA_DESDE & FECHA_HASTA & _
                     strMuestra & _
                     FECHA_DESDE & FECHA_HASTA & _
                     strpar & stranno & _
                     strTipo & _
                     " order by mu.id_muestra desc"
    Me.MousePointer = 11
    Set RS = datos_bd(consulta)
    If RS.RecordCount >= 1 Then
        Dim oMuestra As New clsMuestra
        lista.ListItems.Clear
        i = 1
        While Not RS.EOF
            With lista.ListItems.Add(, , RS.Fields(1))
            .SubItems(1) = RS.Fields(2)
            .SubItems(2) = RS.Fields(8)
            .SubItems(3) = RS.Fields(4)
            If Not IsNull(RS.Fields(5)) Then
            .SubItems(4) = RS.Fields(5)
            End If
            If Not IsNull(RS.Fields(7)) Then
            .SubItems(5) = Format(RS.Fields(7), "currency")
            End If
            If Not IsNull(RS.Fields(9)) Then
            .SubItems(6) = Format(RS.Fields(9), "00000")
            End If
            If Not IsNull(RS.Fields(6)) Then
            .SubItems(7) = RS.Fields(6)
            End If
            End With
'            lista.ListItems(i).Checked = True
            i = i + 1
            RS.MoveNext
        Wend
        lblmsg.Caption = "Muestras entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy")
        If opTipo(1).value = True And usuario.getPER_CIERRE = True Then
            cmdCerrar.Enabled = True
        End If
    Else
        lblmsg.Caption = "No existe ninguna muestra con esos criterios."
    End If
    Set oAnalisis = Nothing
    Me.MousePointer = 0
    Set RS = Nothing
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
        gmuestra = lista.ListItems(lista.SelectedItem.Index).SubItems(7)
        frmVerMuestra.Show 1
        gmuestra = 0
    End If
End Sub
