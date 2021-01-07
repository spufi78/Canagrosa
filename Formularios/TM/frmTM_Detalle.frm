VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmTM_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de Muestras"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   7785
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmTM_Detalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   7785
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Acreditaciones"
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
      Height          =   1365
      Left            =   45
      TabIndex        =   21
      Top             =   5085
      Width           =   7665
      Begin VB.CheckBox chkNadcap 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NADCAP"
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   6390
         TabIndex        =   25
         Top             =   225
         Width           =   1140
      End
      Begin VB.OptionButton opENAC 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENAC PARCIAL (Algún Ensayo no esta certificado por ENAC)"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   945
         Width           =   6945
      End
      Begin VB.OptionButton opENAC 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENAC COMPLETA (Todos los ensayos estan certificados por ENAC)"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   23
         Top             =   630
         Width           =   6945
      End
      Begin VB.OptionButton opENAC 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO ENAC"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   22
         Top             =   315
         Value           =   -1  'True
         Width           =   6945
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   5565
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6510
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   6690
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   6510
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   4305
      Left            =   45
      TabIndex        =   10
      Top             =   720
      Width           =   7695
      Begin VB.CheckBox chkProducto 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Requiere descripción del producto"
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
         Left            =   90
         TabIndex        =   20
         Top             =   3765
         Width           =   6465
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1605
         TabIndex        =   0
         Top             =   330
         Width           =   5940
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1605
         TabIndex        =   1
         Top             =   720
         Width           =   2115
      End
      Begin MSDataListLib.DataCombo cmbsectores 
         Bindings        =   "frmTM_Detalle.frx":000C
         Height          =   360
         Left            =   1605
         TabIndex        =   3
         Top             =   1530
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbfamilias 
         Bindings        =   "frmTM_Detalle.frx":003C
         Height          =   360
         Left            =   1605
         TabIndex        =   4
         Top             =   1935
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbTipos 
         Bindings        =   "frmTM_Detalle.frx":006C
         Height          =   360
         Left            =   1605
         TabIndex        =   2
         Top             =   1110
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbInforme 
         Bindings        =   "frmTM_Detalle.frx":009C
         Height          =   360
         Left            =   1605
         TabIndex        =   5
         Top             =   2370
         Width           =   5940
         _ExtentX        =   10478
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin pryCombo.miCombo cmbUsuario 
         Height          =   330
         Left            =   1605
         TabIndex        =   7
         Top             =   3240
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbTipoEnsayo 
         Height          =   330
         Left            =   1605
         TabIndex        =   6
         Top             =   2835
         Width           =   6000
         _ExtentX        =   10583
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Ensayo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   26
         Top             =   2850
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   19
         Top             =   3330
         Width           =   1080
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Informe"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   16
         Top             =   2430
         Width           =   1380
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   15
         Top             =   1965
         Width           =   600
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sector"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   1560
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo (Req.Baño)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   1170
         Width           =   1395
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   12
         Top             =   360
         Width           =   705
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   11
         Top             =   750
         Width           =   615
      End
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de los tipos de Muestras"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   18
      Top             =   375
      Width           =   2250
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Tipos de Muestras"
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
      TabIndex        =   17
      Top             =   45
      Width           =   3090
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   7830
   End
End
Attribute VB_Name = "frmTM_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmbsectores_Change()
    If cmbsectores.Text <> "" Then
        cargar_familias (cmbsectores.BoundText)
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    If validar = True Then
      ' analisis
      Dim otm As New clsTipos_muestra
      With otm
        .setNOMBRE = txtDatos(0)
        .setCODIGO = txtDatos(1)
        If cmbtipos.Text <> "" Then
            .setTIPO_ESPECIAL_ID = cmbtipos.BoundText
        End If
        If cmbsectores.Text <> "" Then
            .setSECTOR_ID = cmbsectores.BoundText
        End If
        If cmbfamilias.Text <> "" Then
            .setFAMILIA_ID = cmbfamilias.BoundText
        End If
        If cmbInforme.Text = "" Then
            .setTIPO_DOCUMENTO_ID = 1
        Else
            .setTIPO_DOCUMENTO_ID = cmbInforme.BoundText
        End If
        .setTIPO_ENSAYO_ID = cmbTipoEnsayo.getPK_SALIDA
        .setRESPONSABLE_ID = cmbUsuario.getPK_SALIDA
        .setREQUIERE_PRODUCTO = chkProducto.Value
        ' ENAC
        If opENAC(0).Value = True Then
            .setENAC = 0
        ElseIf opENAC(1).Value = True Then
            .setENAC = 1
        Else
            .setENAC = 2
        End If
        .setNADCAP = chkNADCAP
        If PK = 0 Then
            If MsgBox("Va a introducir un tipo de muestra. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
            .Insertar
        Else
            If MsgBox("Va a modificar un tipo de muestra. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
            .Modificar (PK)
        End If
      End With
      If PK = 0 Then
          MsgBox "El tipo de muestra se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "El tipo de muestra se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    If PK <> 0 Then
        lbltitulo = "Modificación del Tipo de Muestra"
        Call cargar_familias(0)
        cargar_muestra
    Else
        lbltitulo = "Alta de nuevo Tipo de Muestra"
    End If
    'M1377-I
    If USUARIO.getPER_DES_PRODUCTO = 0 Then
        chkProducto.Enabled = False
    End If
    'M1377-F
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub cargar_muestra()
    Dim otm As New clsTipos_muestra
    If otm.CARGAR(PK) = True Then
        txtDatos(1) = otm.getCODIGO
        txtDatos(0) = otm.getNOMBRE
        cmbtipos.BoundText = otm.getTIPO_ESPECIAL_ID
        cmbsectores.BoundText = otm.getSECTOR_ID
        cmbfamilias.BoundText = otm.getFAMILIA_ID
        cmbInforme.BoundText = otm.getTIPO_DOCUMENTO_ID
        cmbUsuario.MostrarElemento otm.getRESPONSABLE_ID
        cmbTipoEnsayo.MostrarElemento otm.getTIPO_ENSAYO_ID
        chkProducto.Value = otm.getREQUIERE_PRODUCTO
        opENAC(otm.getENAC).Value = True
        chkNADCAP = otm.getNADCAP
        
    End If
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un nombre al tipo de muestra.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(1)) = "" Then
        MsgBox "Debe darle un código al tipo de muestra.", vbInformation, App.Title
        txtDatos(1).SetFocus
        validar = False
        Exit Function
    End If
    If cmbUsuario.getPK_SALIDA = 0 Then
        MsgBox "Debe asignar un responsable al tipo de muestra.", vbInformation, App.Title
        cmbUsuario.SetFocus
        validar = False
        Exit Function
    End If
    If cmbTipoEnsayo.getTEXTO = "" Then
        MsgBox "Debe asignar un Tipo de Ensayo.", vbInformation, App.Title
        cmbTipoEnsayo.SetFocus
        validar = False
        Exit Function
    End If
End Function
Private Sub cargar_familias(sec As Integer)
    Dim otipos As New clsFamilias
    If sec = 0 Then
        Set cmbfamilias.RowSource = otipos.Listado_completo
    Else
        Set cmbfamilias.RowSource = otipos.Listado(sec)
    End If
    cmbfamilias.ListField = "nombre"
    cmbfamilias.BoundColumn = "id_familia" 'lo que realmente
    Set otipos = Nothing
End Sub

Private Sub cargar_combos()
    cargar_combo cmbtipos, New clsTipos_especial
    cargar_combo cmbsectores, New clsSectores
    cargar_combo cmbInforme, New clsTipos_documentos
    llenar_combo cmbUsuario, New clsUsuarios, 0, frmUsuarios, ""
    
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbTipoEnsayo, DECODIFICADORA.DECODIFICADORA_TM_TIPOS_ENSAYOS
    Set oDeco = Nothing
    
End Sub
