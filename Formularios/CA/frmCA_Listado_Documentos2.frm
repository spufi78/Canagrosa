VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCA_Listado_Documentos2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Documentos de Calidad"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13695
   Icon            =   "frmCA_Listado_Documentos2.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   13695
   Begin VB.CheckBox chkEnvioCorreos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Envio de Correos Activado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   11115
      TabIndex        =   37
      Top             =   180
      Width           =   2385
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1185
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8190
      Width           =   1050
   End
   Begin VB.Frame Frame1 
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
      Height          =   1770
      Left            =   45
      TabIndex        =   5
      Top             =   630
      Width           =   13605
      Begin VB.CheckBox chkSinTocar 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sólo los documentos 5 años sin tocar"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10125
         TabIndex        =   39
         Top             =   1530
         Width           =   3255
      End
      Begin VB.CheckBox chkMTL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "MTL"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10125
         TabIndex        =   38
         Top             =   630
         Width           =   960
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1140
         MaxLength       =   255
         TabIndex        =   15
         Top             =   960
         Width           =   4020
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   870
         Left            =   11385
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   180
         Width           =   1050
      End
      Begin VB.CheckBox chkuso 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sin uso"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10125
         TabIndex        =   13
         Top             =   1080
         Width           =   1005
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   870
         Left            =   12510
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   180
         Width           =   1050
      End
      Begin VB.CheckBox chkNADCAP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "NADCAP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10125
         TabIndex        =   11
         Top             =   405
         Width           =   960
      End
      Begin VB.CheckBox chkENAC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENAC"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10125
         TabIndex        =   10
         Top             =   180
         Width           =   810
      End
      Begin VB.CheckBox chkEQA 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "EQA"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10125
         TabIndex        =   9
         Top             =   855
         Width           =   750
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   6345
         MaxLength       =   255
         TabIndex        =   8
         Top             =   945
         Width           =   3660
      End
      Begin VB.CheckBox chkcopia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar sólo copias controladas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10125
         TabIndex        =   7
         Top             =   1305
         Width           =   2850
      End
      Begin VB.CheckBox chkFechas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   285
         Left            =   180
         TabIndex        =   6
         Top             =   1305
         Width           =   285
      End
      Begin MSDataListLib.DataCombo cmbfamilias 
         Height          =   315
         Left            =   1140
         TabIndex        =   16
         Top             =   225
         Width           =   4050
         _ExtentX        =   7144
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
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   315
         Left            =   1140
         TabIndex        =   17
         Top             =   600
         Width           =   4050
         _ExtentX        =   7144
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
      Begin MSDataListLib.DataCombo cmbSubfamilia 
         Height          =   315
         Left            =   6345
         TabIndex        =   18
         Top             =   225
         Width           =   3690
         _ExtentX        =   6509
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
      Begin MSDataListLib.DataCombo cmbResponsable 
         Height          =   315
         Left            =   6345
         TabIndex        =   19
         Top             =   585
         Width           =   3690
         _ExtentX        =   6509
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
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1125
         TabIndex        =   20
         Top             =   1305
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   60293121
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3195
         TabIndex        =   21
         Top             =   1305
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   60293121
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   195
         TabIndex        =   29
         Top             =   960
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   1
         Left            =   195
         TabIndex        =   28
         Top             =   645
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   0
         Left            =   195
         TabIndex        =   27
         Top             =   285
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "SubFamilia"
         Height          =   195
         Index           =   3
         Left            =   5355
         TabIndex        =   26
         Top             =   285
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   4
         Left            =   5355
         TabIndex        =   25
         Top             =   975
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   5
         Left            =   5355
         TabIndex        =   24
         Top             =   645
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   6
         Left            =   2655
         TabIndex        =   23
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   7
         Left            =   465
         TabIndex        =   22
         Top             =   1350
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdMostrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar"
      Height          =   870
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8190
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir la Lista"
      Height          =   870
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8190
      Width           =   1830
   End
   Begin VB.CommandButton cmdListadoAuditoria 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listados Auditoria"
      Height          =   870
      Left            =   10260
      Picture         =   "frmCA_Listado_Documentos2.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8190
      Width           =   1830
   End
   Begin VB.CommandButton cmdVigor 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Documentos Vigor LI-01"
      Height          =   870
      Left            =   6345
      Picture         =   "frmCA_Listado_Documentos2.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8190
      Width           =   1875
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copias Controladas LI-03"
      Height          =   870
      Left            =   8280
      Picture         =   "frmCA_Listado_Documentos2.frx":149E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8190
      Width           =   1920
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5670
      Left            =   45
      TabIndex        =   34
      Top             =   2430
      Width           =   13605
      _ExtentX        =   23998
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   45
      TabIndex        =   36
      Top             =   315
      Width           =   45
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Documentos de Calidad"
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
      Left            =   45
      TabIndex        =   35
      Top             =   45
      Width           =   3660
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   13995
   End
End
Attribute VB_Name = "frmCA_Listado_Documentos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public VINCULAR As Boolean
Option Explicit

Private Sub chkEnvioCorreos_Click()
    Dim oParametros As New clsParametros
    oParametros.Modificar_Valor parametros.ENVIO_CORREO_PNT, "", chkEnvioCorreos.Value
    Set oParametros = Nothing
End Sub

Private Sub chkfechas_Click()
    If chkFechas.Value = Checked Then
        fdesde.Enabled = True
        fhasta.Enabled = True
    Else
        fdesde.Enabled = False
        fhasta.Enabled = False
    End If
End Sub

Private Sub chkMTL_Click()
    cmdBuscar_Click
End Sub

Private Sub cmdMostrar_Click()
   On Error GoTo CMDMOSTRAR_Click_Error

    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oCA_Documento As New clsCa_documentos
    oCA_Documento.mostrar lista.ListItems(lista.selectedItem.Index).Text, False
    Set oCA_Documento = Nothing

   On Error GoTo 0
   Exit Sub

CMDMOSTRAR_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMostrar_Click of Formulario frmCA_Listado_Documentos"
End Sub

Private Sub cmdVigor_Click()
    cmdLimpiar_Click
    imprimir_lista (2)
End Sub

Private Sub Command1_Click()
    cmdLimpiar_Click
    chkcopia.Value = Checked
    imprimir_lista (3)
End Sub

Private Sub chkcopia_Click()
    cmdBuscar_Click
End Sub

Private Sub cmdListadoAuditoria_Click()
    frmCa_Filtro_Listado.TIPO_LLAMADA = 1
    frmCa_Filtro_Listado.Show 1
End Sub
Private Sub chkENAC_Click()
    cmdBuscar_Click
End Sub

Private Sub chkEQA_Click()
    cmdBuscar_Click
End Sub

Private Sub chkNADCAP_Click()
    cmdBuscar_Click
End Sub

Private Sub chkUso_Click()
    cmdBuscar_Click
End Sub

Private Sub cmbestados_Change()
    cmdBuscar_Click
End Sub
Private Sub cmbfamilias_Change()
    cmdBuscar_Click
End Sub

Private Sub cmbResponsable_Change()
cmdBuscar_Click
End Sub

Private Sub cmbSubfamilia_Change()
    cmdBuscar_Click
End Sub

Private Sub cmdAnadir_Click()
    frmCA_Documento.PK = 0
    frmCA_Documento.Show 1
    cargar_lista
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar el documento de calidad : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oCA_Documento As New clsCa_documentos
            If oCA_Documento.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdImprimir_Click()
    imprimir_lista (1)
End Sub

Private Sub cmdLimpiar_Click()
    txtDatos(1) = ""
    cmbfamilias.Text = ""
    cmbSubfamilia.Text = ""
    cmbestados.Text = ""
    cmbResponsable.Text = ""
    chkuso.Value = Unchecked
    cmdBuscar_Click
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        If Not USUARIO.getPER_DOCUMENTACION_CALIDAD Then
            Exit Sub
        End If
        frmCA_Documento.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmCA_Documento.Show 1
        actualizar_lista
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_botones Me
    cargar_combos
    
    Dim oParametros As New clsParametros
    oParametros.Carga parametros.ENVIO_CORREO_PNT, ""
    If oParametros.getVALOR = "1" Then
        chkEnvioCorreos.Value = Checked
    End If
    Set oParametros = Nothing
    
    cargar_lista
    permisos
    fdesde = Date - 180
    fhasta = Date
    
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Documento", 3900, lvwColumnLeft
        .Add , , "Familia", 2300, lvwColumnLeft
        .Add , , "Código", 1300, lvwColumnCenter
        .Add , , "Edición", 800, lvwColumnCenter
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Estado", 1100, lvwColumnCenter
        .Add , , "En Uso", 700, lvwColumnCenter
        .Add , , "Copia", 700, lvwColumnCenter
        .Add , , "Laboratorio", 1400, lvwColumnCenter
        .Add , , "ROJO", 1, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oca_documentos As New clsCa_documentos
    lista.ListItems.Clear
    Dim familia As String
    Dim subfamilia As String
    Dim ESTADO As String
    Dim nombre As String
    Dim CODIGO As String
    Dim responsable As String
    If cmbfamilias.Text = "" Then
        familia = 0
    Else
        familia = cmbfamilias.BoundText
    End If
    If cmbSubfamilia.Text = "" Then
        subfamilia = 0
    Else
        subfamilia = cmbSubfamilia.BoundText
    End If
    If cmbestados.Text = "" Then
        ESTADO = 0
    Else
        ESTADO = cmbestados.BoundText
    End If
    If cmbResponsable.Text = "" Then
        responsable = 0
    Else
        responsable = cmbResponsable.BoundText
    End If
    nombre = txtDatos(1)
    CODIGO = txtDatos(0)
    Set rs = oca_documentos.Listado(familia, subfamilia, ESTADO, nombre, CODIGO, chkuso.Value, chkENAC.Value, chkNADCAP.Value, chkMTL.Value, chkEQA.Value, responsable, chkcopia.Value, chkFechas.Value, fdesde, fhasta, chkSinTocar.Value)
    Dim incluir As Boolean
    
    If rs.RecordCount <> 0 Then
        Do
            incluir = True
            
            If chkFechas.Value = Checked Then
                If IsDate(rs(5)) Then
                    If CDate(rs(5)) < CDate(fdesde) Or CDate(rs(5)) > CDate(fhasta) Then
                        incluir = False
                    End If
                Else
                    incluir = False
                End If
            End If
            If incluir Then
                With lista.ListItems.Add(, , Format(rs(0), "0000"))
                 .SubItems(1) = rs(1)
                 .SubItems(2) = rs(2)
                 .SubItems(3) = rs(3)
                 If rs(10) = 1 Then
                    .SubItems(4) = "N.A."
                 Else
                    .SubItems(4) = rs(4)
                 End If
                 If IsDate(rs(5)) Then
                    .SubItems(5) = Format(rs(5), "dd-mm-yyyy")
                 Else
                    .SubItems(5) = rs(5)
                 End If
                 .SubItems(6) = rs(6)
                 If rs(7) = 1 Then
                    .SubItems(7) = "SI"
                 Else
                    .SubItems(7) = "NO"
                 End If
                 If rs(8) = 1 Then
                    .SubItems(8) = "SI"
                 Else
                    .SubItems(8) = "NO"
                 End If
                 .SubItems(9) = rs(9)
                 .SubItems(10) = rs(11) 'ROJO
                End With
                ' Si esta en vigor y lleva mas de 5 años, poner la fila en rojo
                If rs(11) = 1 Then
                    lista_colorear lista, lista.ListItems.Count, vbRed
                End If
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    lblsubtitulo = "Documentos listados : " & lista.ListItems.Count
    Set oca_documentos = Nothing
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
    cmdModificar_Click
End Sub

Private Sub actualizar_lista()
    Dim rs As ADODB.Recordset
    Dim oca_documentos As New clsCa_documentos
    Set rs = oca_documentos.Listado_por_Codigo(lista.ListItems(lista.selectedItem.Index).Text)
    If rs.RecordCount <> 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs(3)
        If rs(9) = 1 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(4) = "N.A."
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs(4)
        End If
'        If rs(9) = 1 Then
'            lista.ListItems(lista.SelectedItem.Index).SubItems(5) = " "
'        Else
            If IsDate(rs(5)) Then
                lista.ListItems(lista.selectedItem.Index).SubItems(5) = Format(rs(5), "dd-mm-yyyy")
            Else
                lista.ListItems(lista.selectedItem.Index).SubItems(5) = rs(5)
            End If
'        End If
        lista.ListItems(lista.selectedItem.Index).SubItems(6) = rs(6)
        If rs(7) = 1 Then
           lista.ListItems(lista.selectedItem.Index).SubItems(7) = "SI"
        Else
           lista.ListItems(lista.selectedItem.Index).SubItems(7) = "NO"
        End If
        If rs(8) = 1 Then
           lista.ListItems(lista.selectedItem.Index).SubItems(8) = "SI"
        Else
           lista.ListItems(lista.selectedItem.Index).SubItems(8) = "NO"
        End If
    End If
    Set oca_documentos = Nothing
End Sub

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbfamilias, DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS
    oDeco.cargar_combo cmbSubfamilia, DECODIFICADORA.CA_DOCUMENTOS_SUBFAMILIAS
    oDeco.cargar_combo cmbestados, DECODIFICADORA.CA_DOCUMENTOS_ESTADOS
    oDeco.cargar_combo cmbResponsable, DECODIFICADORA.CA_DOCUMENTOS_RESPONSABLES
    Set oDeco = Nothing
End Sub
Private Sub permisos()
    If Not USUARIO.getPER_DOCUMENTACION_CALIDAD Then
        cmdModificar.Enabled = False
    End If
    If Not USUARIO.getPER_ADMIN_PNT Then
        chkEnvioCorreos.Visible = False
        chkuso.Visible = False ' SOLICITADO POR MARGA 10-04-2014
        cmdAnadir.Enabled = False
        cmdEliminar.Enabled = False
    End If
End Sub
Private Sub txtDatos_Change(Index As Integer)
    cmdBuscar_Click
End Sub
Private Sub genera_listado(ByVal rs As ADODB.Recordset, fecha_edicion As Date)
    Dim Listado As New rptListadoModal
    With Listado.Sections("cabecera")
            .Controls("titulo").Caption = "Listado de Documentos"
            .Controls("etiqueta4").Left = 170
            .Controls("etiqueta4").Width = 5800
            .Controls("etiqueta4").Caption = "Documento"
            .Controls("etiqueta5").Left = 6000
            .Controls("etiqueta5").Width = 1500
            .Controls("etiqueta5").Caption = "Código"
            .Controls("etiqueta10").Left = 7800
            .Controls("etiqueta10").Width = 1500
            .Controls("etiqueta10").Caption = "Edición"
            .Controls("etiqueta11").Left = 9400
            .Controls("etiqueta11").Width = 1500
            .Controls("etiqueta11").Caption = "Fecha"
    End With
    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
    'Detalle
    With Listado.Sections("detalle")
            .Controls("d1").Left = 170
            .Controls("d1").Width = 5800
            .Controls("d1").CanGrow = True
            .Controls("d1").Alignment = 0
            .Controls("d1").DataField = rs.Fields("c1").Name
            .Controls("d2").Left = 6000
            .Controls("d2").Width = 1500
            .Controls("d2").CanGrow = True
            .Controls("d2").Alignment = 2
            .Controls("d2").DataField = rs.Fields("c2").Name
            .Controls("d3").Left = 7800
            .Controls("d3").Width = 1500
            .Controls("d3").Alignment = 2
            .Controls("d3").DataField = rs.Fields("c3").Name
            .Controls("d4").Left = 9400
            .Controls("d4").Width = 1500
            .Controls("d4").Alignment = 2
            .Controls("d4").DataField = rs.Fields("c4").Name
    End With
    ' Pie de Pagina
    With Listado.Sections("pie")
        .Controls("pie1").Caption = "Fecha Impresión: " & Format(Date, "dd-mm-yyyy")
'        .Controls("pie2").Caption = "Fecha Ult.Edición : " & Format(fecha_edicion, "dd-mm-yyyy")
        .Controls("pie3").Caption = "Firmado, Margarita Halcón"
        .Controls("pie3").Visible = True
        .Controls("firma").Visible = True
    End With
    Set Listado.DataSource = rs
    Listado.Caption = "Listado de Documentos de calidad."
    Listado.Show 1
    Set rs = Nothing
End Sub

Private Sub genera_listado_LI01(ByVal rs As ADODB.Recordset, fecha_edicion As Date)
    Dim Listado As New rptListadoModal
    With Listado.Sections("cabecera")
            .Controls("titulo").Caption = "Lista de Documentos en Vigor (LI-01)"
            .Controls("etiqueta4").Left = 170
            .Controls("etiqueta4").Width = 5800
            .Controls("etiqueta4").Caption = "Documento"
            .Controls("etiqueta5").Left = 6000
            .Controls("etiqueta5").Width = 1500
            .Controls("etiqueta5").Caption = "Código"
            .Controls("etiqueta10").Left = 7800
            .Controls("etiqueta10").Width = 1500
            .Controls("etiqueta10").Caption = "Edición"
            .Controls("etiqueta11").Left = 9400
            .Controls("etiqueta11").Width = 1500
            .Controls("etiqueta11").Caption = "Fecha"
    End With
    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
    'Detalle
    With Listado.Sections("detalle")
            .Controls("d1").Left = 170
            .Controls("d1").Width = 5800
            .Controls("d1").CanGrow = True
            .Controls("d1").Alignment = 0
            .Controls("d1").DataField = rs.Fields("c1").Name
            .Controls("d2").Left = 6000
            .Controls("d2").Width = 1500
            .Controls("d2").CanGrow = True
            .Controls("d2").Alignment = 2
            .Controls("d2").DataField = rs.Fields("c2").Name
            .Controls("d3").Left = 7800
            .Controls("d3").Width = 1500
            .Controls("d3").Alignment = 2
            .Controls("d3").DataField = rs.Fields("c3").Name
            .Controls("d4").Left = 9400
            .Controls("d4").Width = 1500
            .Controls("d4").Alignment = 2
            .Controls("d4").DataField = rs.Fields("c4").Name
    End With
    ' Pie de Pagina
    With Listado.Sections("pie")
        .Controls("pie1").Caption = "Fecha Impresión: " & Format(Date, "dd-mm-yyyy")
        .Controls("pie2").Caption = "Fecha Ult.Edición : " & Format(fecha_edicion, "dd-mm-yyyy")
        .Controls("pie3").Caption = "Firmado, Margarita Halcón"
        .Controls("pie3").Visible = True
        .Controls("firma").Visible = True
    End With
    Set Listado.DataSource = rs
    Listado.Caption = "Listado de Documentos de calidad."
    Listado.Show 1
    Set rs = Nothing
End Sub
Private Sub genera_listado_LI03(ByVal rs As ADODB.Recordset, fecha_edicion As Date)
    Dim Listado As New rptLI03
    With Listado.Sections("cabecera")
            .Controls("titulo").Caption = "Lista para el Control de distribución de Documentos (LI-03)"
    End With
    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
    'Detalle
    With Listado.Sections("detalle")
            .Controls("d1").DataField = rs.Fields("c1").Name
            .Controls("d2").DataField = rs.Fields("c2").Name
            .Controls("d3").DataField = rs.Fields("c3").Name
            .Controls("d4").DataField = rs.Fields("c4").Name
            .Controls("d5").DataField = rs.Fields("c5").Name
    End With
    ' Pie de Pagina
    With Listado.Sections("pie")
        .Controls("pie1").Caption = "Fecha Impresión: " & Format(Date, "dd-mm-yyyy")
        .Controls("pie2").Caption = "Fecha Ult.Edición : " & Format(fecha_edicion, "dd-mm-yyyy")
        .Controls("pie3").Caption = "Firmado, Margarita Halcón"
        .Controls("pie3").Visible = True
        .Controls("firma").Visible = True
    End With
    Set Listado.DataSource = rs
    Listado.Caption = "Listado de Documentos de calidad."
    Listado.Show 1
    Set rs = Nothing
End Sub


Private Sub imprimir_lista(tipo As Integer)
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim i As Integer
    Dim fecha As Date
    fecha = "01-01-1900"
    ' Generamos los datos del listado
    Dim rs As New ADODB.Recordset
    rs.Fields.Append "c1", adChar, 250, adFldUpdatable
    rs.Fields.Append "c2", adChar, 50, adFldUpdatable
    rs.Fields.Append "c3", adChar, 50, adFldUpdatable
    rs.Fields.Append "c4", adChar, 50, adFldUpdatable
    rs.Fields.Append "c5", adChar, 50, adFldUpdatable
    rs.Open
    For i = 1 To lista.ListItems.Count
        rs.AddNew
        rs("c1") = Trim(Left(lista.ListItems(i).SubItems(1), 250))
        rs("c2") = Left(lista.ListItems(i).SubItems(3), 50)
        ' Edicion
        rs("c3") = "N.A."
        If lista.ListItems(i).SubItems(4) <> "" Then
            If IsNumeric(lista.ListItems(i).SubItems(4)) Then
                rs("c3") = Left(lista.ListItems(i).SubItems(4), 50)
            End If
        End If
        rs("c4") = Left(lista.ListItems(i).SubItems(5), 50)
        rs("c5") = Left(lista.ListItems(i).SubItems(9), 50)
        If IsDate(lista.ListItems(i).SubItems(5)) Then
            If Format(fecha, "yyyy-mm-dd") < Format(lista.ListItems(i).SubItems(5), "yyyy-mm-dd") Then
                fecha = Format(lista.ListItems(i).SubItems(5), "yyyy-mm-dd")
            End If
        End If
        rs.Update
    Next
    ' Generar Listado
    Select Case tipo
    Case 1
        genera_listado rs, fecha
    Case 2
        genera_listado_LI01 rs, fecha
    Case 3
        genera_listado_LI03 rs, fecha
    End Select
    Exit Sub
fallo:
    MsgBox "Error al generar el listado.", vbCritical, Err.Description

End Sub


