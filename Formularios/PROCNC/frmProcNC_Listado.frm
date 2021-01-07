VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmProcNC_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Incidencias"
   ClientHeight    =   9405
   ClientLeft      =   585
   ClientTop       =   1890
   ClientWidth     =   13680
   Icon            =   "frmProcNC_Listado.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9405
   ScaleWidth      =   13680
   Begin VB.CommandButton cmdInformeParcial 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe Parcial"
      Height          =   870
      Left            =   5535
      Picture         =   "frmProcNC_Listado.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8460
      Width           =   1020
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe Completo"
      Height          =   870
      Left            =   4455
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8460
      Width           =   1020
   End
   Begin VB.CommandButton cmdListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8460
      Width           =   1020
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
      Height          =   2010
      Left            =   30
      TabIndex        =   15
      Top             =   630
      Width           =   13605
      Begin VB.CheckBox chkFCierre 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Left            =   5805
         TabIndex        =   34
         Top             =   1305
         Width           =   240
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   795
         MaxLength       =   255
         TabIndex        =   28
         Top             =   1305
         Width           =   1530
      End
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   7305
         MaxLength       =   255
         TabIndex        =   3
         Top             =   585
         Width           =   3600
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   1050
         Left            =   11190
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   405
         Width           =   1050
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   1050
         Left            =   12285
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   405
         Width           =   1050
      End
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   315
         Left            =   7305
         TabIndex        =   1
         Top             =   225
         Width           =   3585
         _ExtentX        =   6324
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
      Begin MSDataListLib.DataCombo cmborigen 
         Height          =   315
         Left            =   795
         TabIndex        =   0
         Top             =   240
         Width           =   4635
         _ExtentX        =   8176
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
      Begin MSDataListLib.DataCombo cmbAuditoria 
         Height          =   315
         Left            =   795
         TabIndex        =   2
         Top             =   585
         Visible         =   0   'False
         Width           =   4635
         _ExtentX        =   8176
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
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   795
         TabIndex        =   4
         Top             =   930
         Width           =   4635
         _ExtentX        =   8176
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
      Begin pryCombo.miCombo cmbClientes 
         Height          =   420
         Left            =   795
         TabIndex        =   22
         Top             =   585
         Visible         =   0   'False
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   741
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   7290
         TabIndex        =   23
         Top             =   900
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   9495
         TabIndex        =   24
         Top             =   900
         Width           =   1410
         _ExtentX        =   2487
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
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbProveedores 
         Height          =   420
         Left            =   795
         TabIndex        =   27
         Top             =   585
         Visible         =   0   'False
         Width           =   4635
         _ExtentX        =   8176
         _ExtentY        =   741
      End
      Begin MSComCtl2.DTPicker fcdesde 
         Height          =   330
         Left            =   7290
         TabIndex        =   30
         Top             =   1260
         Width           =   1500
         _ExtentX        =   2646
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
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fchasta 
         Height          =   330
         Left            =   9495
         TabIndex        =   31
         Top             =   1260
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
         Format          =   16449537
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbDESVIACION_ID 
         Height          =   315
         Left            =   7290
         TabIndex        =   35
         Top             =   1620
         Width           =   3585
         _ExtentX        =   6324
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
         Caption         =   "Estudio"
         Height          =   195
         Index           =   9
         Left            =   6030
         TabIndex        =   36
         Top             =   1665
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Cierre"
         Height          =   195
         Index           =   8
         Left            =   6045
         TabIndex        =   33
         Top             =   1305
         Width           =   900
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   7
         Left            =   8955
         TabIndex        =   32
         Top             =   1305
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         Height          =   195
         Index           =   6
         Left            =   150
         TabIndex        =   29
         Top             =   1350
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   3
         Left            =   8955
         TabIndex        =   26
         Top             =   945
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Alta"
         Height          =   195
         Index           =   0
         Left            =   6045
         TabIndex        =   25
         Top             =   990
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   4
         Left            =   150
         TabIndex        =   21
         Top             =   990
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desc./Resumen"
         Height          =   195
         Index           =   2
         Left            =   6045
         TabIndex        =   18
         Top             =   630
         Width           =   1170
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Origen"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   17
         Top             =   285
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   1
         Left            =   6045
         TabIndex        =   16
         Top             =   270
         Width           =   495
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2265
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8460
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1155
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8460
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8460
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8460
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5745
      Left            =   45
      TabIndex        =   14
      Top             =   2655
      Width           =   13605
      _ExtentX        =   23998
      _ExtentY        =   10134
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8910
      Top             =   8505
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcNC_Listado.frx":1194
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProcNC_Listado.frx":162B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Procedimientos de No Conformidad"
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
      TabIndex        =   20
      Top             =   45
      Width           =   4860
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13005
      Picture         =   "frmProcNC_Listado.frx":1AC2
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rellene los datos básicos y las acciones inmediatas para generar una nueva no conformidad"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   19
      Top             =   360
      Width           =   6540
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   13635
   End
End
Attribute VB_Name = "frmProcNC_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarobjProcNC_List As New clsProcNc
Private mvarobjProcNC As clsProcNc

Private Sub cabecera()
On Error GoTo cabecera_Error

    With lista.ColumnHeaders
        .Add , , "NºIncidencia", 1000, lvwColumnLeft
        .Add , , "Origen", 1350, lvwColumnCenter
        .Add , , "", 2200, lvwColumnLeft
        .Add , , "Tipo", 1300, lvwColumnCenter
        .Add , , "Resp.Apertura", 1000, lvwColumnCenter
        .Add , , "Resumen", 2600, lvwColumnCenter
        .Add , , "F.Apert.", 1000, lvwColumnLeft
        .Add , , "F.Ult.Modif.", 0, lvwColumnCenter
        .Add , , "Estado", 1000, lvwColumnCenter
        .Add , , "F.Cierre", 1000, lvwColumnCenter
        .Add , , "NºAcc.", 800, lvwColumnCenter
'        .Add , , "Rev.", 300, lvwColumnCenter
    End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cabecera"
    Exit Sub
cabecera_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cabecera"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cabecera of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub


Private Sub chkFCierre_Click()
    If chkFCierre.Value = Checked Then
        fcdesde.Enabled = True
        fchasta.Enabled = True
    Else
        fcdesde.Enabled = False
        fchasta.Enabled = False
    End If
End Sub

Private Sub cmbDESVIACION_ID_Change()
    cmdBuscar_Click
End Sub
Private Sub Form_Load()
On Error GoTo Form_Load_Error

    log (Me.Name)
    Me.top = 100
    Me.Left = 100
    fdesde = Date - 365
    fhasta = Date
    fcdesde = fdesde
    fchasta = fhasta
    cabecera
    cargar_botones Me
    cargar_combos
    permisos
'    cargar_lista
'    If USUARIO.getUSUARIO = "julio" Then
'        cmdCargar.Visible = True
'    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.Form_Load"
    Exit Sub
Form_Load_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.Form_Load"
    error_grave Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmbestados_Change()
    cmdBuscar_Click
End Sub


Private Sub cmborigen_Change()
    Dim oDeco As New clsDecodificadora
    cmbProveedores.visible = False
    cmbClientes.visible = False
    cmbAuditoria.visible = False
    If cmbOrigen.Text <> "" Then
        Select Case cmbOrigen.BoundText
            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_RECLAMACION_CLIENTE
            Case ENUM_PNC_ORIGEN_INCIDENCIA_MENOR_CLIENTE
                llenar_combo cmbClientes, New clsCliente, 0, Me, " ANULADO = 0 "
                cmbClientes.visible = True
            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_PROVEEDOR
                llenar_combo cmbProveedores, New clsProveedor, 0, Me, " ANULADO = 0 "
                cmbProveedores.visible = True
            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_AUDITORIA_INTERNA
                cmbAuditoria.visible = True
                oDeco.cargar_combo cmbAuditoria, DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_INTERNA
            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_AUDITORIA_EXTERNA
                cmbAuditoria.visible = True
                oDeco.cargar_combo cmbAuditoria, DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_EXTERNA
            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_DETECCION_INTERNA
                cmbAuditoria.visible = True
                oDeco.cargar_combo cmbAuditoria, DECODIFICADORA.PROCNC_ORIGEN_AUDITORIA_DETECCION
        End Select
            
            
    End If
End Sub
Private Sub cmdAnadir_Click()
    
Dim objfrm As New frmProcNCEdicion

On Error GoTo cmdAnadir_Click_Error

    If Not cmdAnadir.Enabled Then
        Exit Sub
    End If
    
    objfrm.PK = 0
    objfrm.Show vbModal
    
    cargar_lista

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdAnadir_Click"
    Exit Sub
cmdAnadir_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdAnadir_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub
Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
Dim objPnc As clsProcNc
Dim permiso_eliminar As Boolean
On Error GoTo cmdEliminar_Click_Error

permiso_eliminar = False

    If lista.ListItems.Count > 0 Then
        Set objPnc = New clsProcNc
        Call objPnc.Carga(CLng(lista.ListItems(lista.selectedItem.Index).Text))
        If objPnc.getESTADO_ID > C_PROCNC_ESTADOS.ABIERTA Then
            MsgBox "No se puede eliminar la Incidencia/PNC " & lista.ListItems(lista.selectedItem.Index).Text & " por que ha entrado en proceso de gestión. Sólo pueden ser eliminadas aquellas que estén abiertas"
            Set objPnc = Nothing
            Exit Sub
        Else
            ' Comprueba que si está abierta, seas responsable de calidad o bien el creador de la incidencia
            If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.Calidad) = 1 Then
                permiso_eliminar = True
            ElseIf objPnc.getRESPONSABLE_ID = USUARIO.getID_EMPLEADO Then
                permiso_eliminar = True
            End If
        
            If permiso_eliminar Then
                If MsgBox("Va a eliminar la INCIDENCIA : " & lista.ListItems(lista.selectedItem.Index).Text, vbQuestion + vbYesNo, App.Title) = vbYes Then
                    If mvarobjProcNC_List.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                        cargar_lista
                    End If
                End If
            Else
                MsgBox "Sólo puede eliminar una Incidencia el Responsable de su apertura o un Responsable del Dpto. de Calidad", vbInformation, "Eliminar Incidencia"
            End If
        End If
    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdEliminar_Click"
    Exit Sub
cmdEliminar_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdEliminar_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdEliminar_Click of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmdListado_Click()

    Dim objfrm As New frmReport
    
On Error GoTo cmdListado_Click_Error

    With objfrm
        .iniciar
        .informe = "/NC/rptProcNC_Listado"
        .criterio = "{decodificadora.CODIGO}=110" '"{decodificadora.codigo}=11 and {nc.ID_NC} in [" & Left(strIncidencias, Len(strIncidencias) - 1) & "]"
        .imprimir = False
        .generar
        .visible = True
    End With
    
    Set objfrm = Nothing
    

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdListado_Click"
    Exit Sub
cmdListado_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdListado_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdListado_Click of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

Private Sub cmdLimpiar_Click()
On Error GoTo cmdLimpiar_Click_Error

    txtdescripcion.Text = ""
    cmbestados.BoundText = ""
    cmbOrigen.Text = ""
    cmbOrigen.BoundText = ""
    cmbAuditoria.Text = ""
    cmbAuditoria.BoundText = ""
    cmbDESVIACION_ID.Text = ""
    cmbDESVIACION_ID.BoundText = ""
    cmbClientes.limpiar
    cmbTipo.Text = ""
    cmbTipo.BoundText = ""
    txtNumero = ""
    
    Call cmdBuscar_Click

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdLimpiar_Click"
    Exit Sub
cmdLimpiar_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdLimpiar_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdLimpiar_Click of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

Private Sub cmdModificar_Click()
    
Dim objfrm As New frmProcNCEdicion
Dim lng_id As Long

On Error GoTo cmdModificar_Click_Error

    If Not cmdModificar.Enabled Then
        Exit Sub
    End If

    If lista.selectedItem Is Nothing Then Exit Sub
    If lista.selectedItem.Index < 0 Then Exit Sub
    lng_id = lista.ListItems(lista.selectedItem.Index).Text
    
    objfrm.PK = lng_id
    objfrm.Show vbModal
    actualizar_lista
'    cargar_lista lng_id
'    Dim i As Integer
'    For i = 1 To lista.ListItems.Count
'        If lista.ListItems(i).Text = lng_id Then
'            lista.ListItems(i).Selected = True
'            lista.ListItems(i).EnsureVisible
'            Exit For
'        End If
'    Next
 
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdModificar_Click"
    Exit Sub
cmdModificar_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdModificar_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub

'Private Sub cmdModificar_Click_old()
'
'Dim objfrm As New frmProcNC_Detalle
'Dim objEH As clsGenericClass
'Dim permiso_modificar As Boolean
'
'
'On Error GoTo cmdModificar_Click_Error
'
'    If Not cmdModificar.Enabled Then
'        Exit Sub
'    End If
'
'    If lista.SelectedItem Is Nothing Then Exit Sub
'    If lista.SelectedItem.Index < 0 Then Exit Sub
'
'
'    objfrm.TipoEdicion = enumTipoEdicion.EDICION
'    Set mvarobjProcNC = New clsProcNc
'    Call mvarobjProcNC.carga(lista.ListItems(lista.SelectedItem.Index).Text)
'
'    ' Comprueba si es del equipo humano, responsable de la apertura o responsable de calidad
'    permiso_modificar = False
'
'    For x = 1 To 10
'        If USUARIO.getRESPONSABLE_DEPARTAMENTOS(x) = 1 Then
'            permiso_modificar = True
'            Exit For
'        End If
'    Next x
'    If mvarobjProcNC.EquipoHumano.Count > 0 Then
'        For Each objEH In mvarobjProcNC.EquipoHumano.Iterator
'            If objEH.getID = USUARIO.getID_EMPLEADO Then
'                permiso_modificar = True
'                Exit For
'            End If
'        Next objEH
'    End If
'
'
'    If Not permiso_modificar Then
'        MsgBox "Solo pueden Ver/Modificar la Incidencia en uno de los siguientes casos: " & vbCrLf & _
'        " - Responsables de Apertura de la Incidencia" & vbCrLf & _
'        " - Responsables de Calidad" & vbCrLf & _
'        " - Responsables de Departamento" & vbCrLf & _
'        " - Usuarios asignados a la gestión de la Incidencia", vbInformation, "Ver/Modificar Incidencia"
'        Set objfrm = Nothing
'        Set objEH = Nothing
'        Exit Sub
'    End If
'
'
'
'    Set objfrm.ProcNC = mvarobjProcNC
'
'    objfrm.Show vbModal
'
'    If objfrm.Resultado = True Then
'        cargar_lista
'    End If
'
'
'
''    If lista.ListItems.Count > 0 Then
''        frmProcNC_Detalle.PK = lista.ListItems(lista.SelectedItem.Index).Text
''        frmProcNC_Detalle.Show 1
''        actualizar_lista
''    End If
'
'On Error GoTo 0
'    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdModificar_Click"
'    Exit Sub
'cmdModificar_Click_Error:
'    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdModificar_Click"
'    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
'    G_TRAZABILIDAD_ERROR = ""
'End Sub


Private Sub cmdImprimir_Click()
    
    Dim id_pnc As String
    
On Error GoTo cmdImprimir_Click_Error

    If lista.selectedItem Is Nothing Then Exit Sub
    If lista.selectedItem.Index < 0 Then Exit Sub
    
    id_pnc = lista.ListItems(lista.selectedItem.Index).Text
    
    With frmReport
        .iniciar
        .informe = "/NC/rptProcNCCompleto"
        .criterio = "{procnc.ID_PROCNC} = " & id_pnc & " and {decodificadora.CODIGO}=110" '"{decodificadora.codigo}=11 and {nc.ID_NC} in [" & Left(strIncidencias, Len(strIncidencias) - 1) & "]"
        .imprimir = False
        .generar
        .visible = True
    End With
    
    

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdImprimir_Click"
    Exit Sub
cmdImprimir_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdImprimir_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub

Private Sub cmdInformeParcial_Click()
    
    Dim id_pnc As String
    
On Error GoTo cmdInformeParcial_Click_Error

    If lista.selectedItem Is Nothing Then Exit Sub
    If lista.selectedItem.Index < 0 Then Exit Sub
    
    id_pnc = lista.ListItems(lista.selectedItem.Index).Text
    
    Dim c As String
    c = "{procnc.ID_PROCNC} = " & id_pnc
    c = c & " and {decodificadora.CODIGO}=110"
    c = c & " and {decodificadora_tipos.CODIGO}=119"
    With frmReport
        .iniciar
        .informe = "/NC/rptProcNC"
        .criterio = c
        .imprimir = False
        .generar
        .visible = True
    End With

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdInformeParcial_Click"
    Exit Sub
cmdInformeParcial_Click_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cmdInformeParcial_Click"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cmdInformeParcial_Click of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""

End Sub


Public Sub cargar_lista(Optional ByVal prmSel_Fila As Long = -1)
    Dim rs As ADODB.Recordset
On Error GoTo cargar_lista_Error
    If cmbOrigen.BoundText = CStr(ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_RECLAMACION_CLIENTE) Or _
       cmbOrigen.BoundText = CStr(ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_INCIDENCIA_MENOR_CLIENTE) Then
        Set rs = mvarobjProcNC_List.Listado(cmbOrigen.BoundText, cmbClientes.getPK_SALIDA, cmbTipo.BoundText, cmbestados.BoundText, txtdescripcion.Text, fdesde, fhasta, txtNumero, chkFCierre.Value, fcdesde, fchasta, cmbDESVIACION_ID.BoundText)
    Else
        If cmbOrigen.BoundText = CStr(ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_PROVEEDOR) Then
            Set rs = mvarobjProcNC_List.Listado(cmbOrigen.BoundText, cmbProveedores.getPK_SALIDA, cmbTipo.BoundText, cmbestados.BoundText, txtdescripcion.Text, fdesde, fhasta, txtNumero, chkFCierre.Value, fcdesde, fchasta, cmbDESVIACION_ID.BoundText)
        Else
            Set rs = mvarobjProcNC_List.Listado(cmbOrigen.BoundText, cmbAuditoria.BoundText, cmbTipo.BoundText, cmbestados.BoundText, txtdescripcion.Text, fdesde, fhasta, txtNumero, chkFCierre.Value, fcdesde, fchasta, cmbDESVIACION_ID.BoundText)
        End If
    End If

    
    lista.ListItems.Clear
    Dim objLitem As ListItem
    lblsubtitulo = "Registros listados : " & rs.RecordCount
    If rs.RecordCount <> 0 Then
        
        rs.MoveFirst
        While Not rs.EOF
        
            With lista.ListItems.Add(, , Format(rs("ID_PROCNC"), "00000"))
            .SubItems(1) = rs("ORIGEN")
            If rs("ORIGEN") = "Proveedores" Then
                .SubItems(2) = rs("auditoria5")
            Else
                If Not IsNull(rs("auditoria1")) Then
                    .SubItems(2) = rs("auditoria1")
                End If
                If Not IsNull(rs("auditoria2")) Then
                    .SubItems(2) = rs("auditoria2")
                End If
                If Not IsNull(rs("auditoria3")) Then
                    .SubItems(2) = rs("auditoria3")
                End If
                If Not IsNull(rs("auditoria4")) Then
                    .SubItems(2) = rs("auditoria4")
                End If
            End If
'            Select Case rs("ORIGEN_ID")
'            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_RECLAMACION_CLIENTE
'            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_AUDITORIA_INTERNA
'            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_AUDITORIA_EXTERNA
'            Case ENUM_PNC_ORIGEN.ENUM_PNC_ORIGEN_DETECCION_INTERNA
'            End Select
            .SubItems(3) = rs("TIPO")
            .SubItems(4) = rs("RESPONSABLE")
            .SubItems(5) = rs("RESUMEN")
            .SubItems(6) = Format(rs("FECHA_ALTA"), "dd-mm-yyyy")
            .SubItems(7) = Format(rs("FECHA_ULT_MOVIMIENTO"), "dd-mm-yyyy")
            .SubItems(8) = rs("estado")
            .SubItems(9) = Format(rs("fecha_cierre"), "dd-mm-yyyy")
               If prmSel_Fila = CLng(rs("ID_PROCNC")) Then
                   .Selected = True
               End If
            .SubItems(10) = rs("N_ACCIONES")
            If rs("estado") = "Cerrado" And rs("flimite") = 1 Then
                Set objLitem = lista.ListItems(lista.ListItems.Count)
                If rs("revisada_usuario_id") = 0 Then
                    objLitem.SmallIcon = 2
'                    objLitem.ListSubItems.Add , , "", 2
                Else
                    objLitem.SmallIcon = 1
'                    objLitem.ListSubItems.Add , , "", 1
                End If
            End If
            End With
            rs.MoveNext
        Wend
    End If
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cargar_lista"
    Exit Sub
cargar_lista_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cargar_lista"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    
    G_TRAZABILIDAD_ERROR = ""
    
End Sub


Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error GoTo lista_ColumnClick_Error

   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.lista_ColumnClick"
    Exit Sub
lista_ColumnClick_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.lista_ColumnClick"
    error_grave Err.Number & " (" & Err.Description & ") in procedure lista_ColumnClick of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub
Public Sub cargar_combos()
    Dim oDecodificadora As New clsDecodificadora
    
On Error GoTo cargar_combos_Error

    oDecodificadora.cargar_combo cmbOrigen, DECODIFICADORA.PROCNC_ORIGEN
'    oDECODIFICADORA.cargar_combo cmbAuditoria, DECODIFICADORA.PROCNC_AUDITORIAS
    oDecodificadora.cargar_combo cmbTipo, DECODIFICADORA.PROCNC_TIPOS_NO_CONFORMIDAD
    oDecodificadora.cargar_combo cmbDESVIACION_ID, DECODIFICADORA.NC_DESVIACIONES
    llenar_combo cmbClientes, New clsCliente, 0, Me, ""
    oDecodificadora.cargar_combo cmbestados, DECODIFICADORA.PROCNC_ESTADOS
    
    'Fin Jonathan
    

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cargar_combos"
    Exit Sub
cargar_combos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cargar_combos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cargar_combos of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub
Public Sub permisos()
Dim x As Integer
Dim es_resp_dpto As Boolean
On Error GoTo permisos_Error

    If Not USUARIO.getPER_NC Then
        cmdAnadir.Enabled = False
        cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
    End If
    
    ' Comprueba si es responsable de algún departamento
    es_resp_dpto = False
    For x = 1 To 10
        If USUARIO.getRESPONSABLE_DEPARTAMENTOS(x) = 1 Then
            es_resp_dpto = True
            Exit For
        End If
    Next x
    
    If Not es_resp_dpto Then
        cmdAnadir.Enabled = False
        'cmdModificar.Enabled = False
        cmdEliminar.Enabled = False
    End If

On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.permisos"
    Exit Sub
permisos_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.permisos"
    error_grave Err.Number & " (" & Err.Description & ") in procedure permisos of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    
End Sub
Private Sub txtDescripcion_Change()
 cmdBuscar_Click
End Sub


Private Sub actualizar_lista()
    Dim rs As ADODB.Recordset
    Dim objLitem As ListItem
On Error GoTo cargar_lista_Error
    Set rs = mvarobjProcNC_List.ListadoId(lista.ListItems(lista.selectedItem.Index).Text)
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
            With lista.ListItems(lista.selectedItem.Index)
            .SubItems(1) = rs("ORIGEN")
            If rs("ORIGEN") = "Proveedores" Then
                .SubItems(2) = rs("auditoria5")
            Else
                If Not IsNull(rs("auditoria1")) Then
                    .SubItems(2) = rs("auditoria1")
                End If
                If Not IsNull(rs("auditoria2")) Then
                    .SubItems(2) = rs("auditoria2")
                End If
                If Not IsNull(rs("auditoria3")) Then
                    .SubItems(2) = rs("auditoria3")
                End If
                If Not IsNull(rs("auditoria4")) Then
                    .SubItems(2) = rs("auditoria4")
                End If
            End If
            .SubItems(3) = rs("TIPO")
            .SubItems(4) = rs("RESPONSABLE")
            .SubItems(5) = rs("RESUMEN")
            .SubItems(6) = Format(rs("FECHA_ALTA"), "dd-mm-yyyy")
            .SubItems(7) = Format(rs("FECHA_ULT_MOVIMIENTO"), "dd-mm-yyyy")
            .SubItems(8) = rs("estado")
            .SubItems(9) = Format(rs("fecha_cierre"), "dd-mm-yyyy")
               If prmSel_Fila = CLng(rs("ID_PROCNC")) Then
                   .Selected = True
               End If
            .SubItems(10) = rs("N_ACCIONES")
            Set objLitem = lista.ListItems(lista.selectedItem.Index)
            objLitem.SmallIcon = vbNothing
            If rs("estado") = "Cerrado" And rs("flimite") = 1 Then
                If rs("revisada_usuario_id") = 0 Then
                    objLitem.SmallIcon = 2
                Else
                    objLitem.SmallIcon = 1
                End If
            End If
            End With
            rs.MoveNext
        Wend
    End If
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cargar_lista"
    Exit Sub
cargar_lista_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> frmProcNC_Listado.cargar_lista"
    error_grave Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmProcNC_Listado" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    
    G_TRAZABILIDAD_ERROR = ""
    
End Sub

