VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmHenkel_Precios 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TABLAS DE PRECIOS HENKEL"
   ClientHeight    =   11865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15900
   Icon            =   "frmHenkel_Precios.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11865
   ScaleWidth      =   15900
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSETS 
      BackColor       =   &H00C0C0C0&
      Height          =   2490
      Left            =   10440
      TabIndex        =   27
      Top             =   5085
      Visible         =   0   'False
      Width           =   5415
      Begin VB.TextBox txtCFM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   960
      End
      Begin VB.TextBox txtCFM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   1080
         TabIndex        =   8
         Top             =   1080
         Width           =   1320
      End
      Begin VB.TextBox txtCFM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   1080
         TabIndex        =   9
         Top             =   1440
         Width           =   1320
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   435
         Left            =   1665
         TabIndex        =   11
         Top             =   1890
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Modificar"
         Appearance      =   5
         Picture         =   "frmHenkel_Precios.frx":08CA
      End
      Begin XtremeSuiteControls.PushButton PushButton2 
         Height          =   435
         Left            =   90
         TabIndex        =   10
         Top             =   1890
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmHenkel_Precios.frx":712C
      End
      Begin XtremeSuiteControls.PushButton PushButton3 
         Height          =   435
         Left            =   3240
         TabIndex        =   12
         Top             =   1890
         Width           =   1620
         _Version        =   851970
         _ExtentX        =   2857
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmHenkel_Precios.frx":D98E
      End
      Begin pryCombo.miCombo cmbDimension 
         Height          =   375
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   661
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio ABCD"
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   34
         Top             =   1485
         Width           =   1050
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio E"
         Height          =   240
         Index           =   1
         Left            =   90
         TabIndex        =   33
         Top             =   1125
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Dimensión"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   32
         Top             =   765
         Width           =   735
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   240
         Left            =   90
         TabIndex        =   31
         Top             =   405
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1230
      Left            =   45
      TabIndex        =   26
      Top             =   6345
      Width           =   10365
      Begin VB.TextBox txtProbetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Width           =   960
      End
      Begin VB.TextBox txtProbetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1035
         TabIndex        =   1
         Top             =   180
         Width           =   1095
      End
      Begin VB.TextBox txtProbetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2115
         TabIndex        =   2
         Top             =   180
         Width           =   1230
      End
      Begin VB.TextBox txtProbetas 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3330
         TabIndex        =   3
         Top             =   180
         Width           =   6945
      End
      Begin XtremeSuiteControls.PushButton cmdModificarProbeta 
         Height          =   435
         Left            =   7065
         TabIndex        =   5
         Top             =   630
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Modificar"
         Appearance      =   5
         Picture         =   "frmHenkel_Precios.frx":141F0
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirProbeta 
         Height          =   435
         Left            =   5490
         TabIndex        =   4
         Top             =   630
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmHenkel_Precios.frx":1AA52
      End
      Begin XtremeSuiteControls.PushButton cmdEliminarProbeta 
         Height          =   435
         Left            =   8640
         TabIndex        =   6
         Top             =   630
         Width           =   1620
         _Version        =   851970
         _ExtentX        =   2857
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmHenkel_Precios.frx":212B4
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1230
      Left            =   45
      TabIndex        =   25
      Top             =   10575
      Width           =   5145
      Begin VB.TextBox txtCFM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   3
         Left            =   3330
         TabIndex        =   16
         Top             =   180
         Width           =   1320
      End
      Begin VB.TextBox txtCFM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2025
         TabIndex        =   15
         Top             =   180
         Width           =   1320
      End
      Begin VB.TextBox txtCFM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   1035
         TabIndex        =   14
         Top             =   180
         Width           =   1005
      End
      Begin VB.TextBox txtCFM 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   180
         Width           =   960
      End
      Begin XtremeSuiteControls.PushButton cmdModificar 
         Height          =   435
         Left            =   1665
         TabIndex        =   18
         Top             =   630
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Modificar"
         Appearance      =   5
         Picture         =   "frmHenkel_Precios.frx":27B16
      End
      Begin XtremeSuiteControls.PushButton cmdAnadir 
         Height          =   435
         Left            =   90
         TabIndex        =   17
         Top             =   630
         Width           =   1515
         _Version        =   851970
         _ExtentX        =   2672
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmHenkel_Precios.frx":2E378
      End
      Begin XtremeSuiteControls.PushButton cmdEliminar 
         Height          =   435
         Left            =   3240
         TabIndex        =   19
         Top             =   630
         Width           =   1620
         _Version        =   851970
         _ExtentX        =   2857
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmHenkel_Precios.frx":34BDA
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   900
      Left            =   14310
      Style           =   1  'Graphical
      TabIndex        =   20
      ToolTipText     =   "Salir"
      Top             =   10755
      Width           =   1230
   End
   Begin MSComctlLib.ListView cfm 
      Height          =   2610
      Left            =   45
      TabIndex        =   22
      Top             =   7965
      Width           =   5145
      _ExtentX        =   9075
      _ExtentY        =   4604
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
      NumItems        =   0
   End
   Begin MSComctlLib.ListView probetas 
      Height          =   5940
      Left            =   45
      TabIndex        =   23
      Top             =   360
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   10478
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView listaSETS 
      Height          =   4725
      Left            =   10440
      TabIndex        =   29
      Top             =   360
      Width           =   5415
      _ExtentX        =   9551
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
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   9270
      Top             =   10485
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHenkel_Precios.frx":3B43C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHenkel_Precios.frx":41C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHenkel_Precios.frx":48500
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "SETS"
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
      Left            =   10455
      TabIndex        =   30
      Top             =   45
      Width           =   5400
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "COSTE PROBETAS"
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
      Left            =   45
      TabIndex        =   24
      Top             =   45
      Width           =   10335
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "COSTE FIJO MENSUAL (CFM)"
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
      Left            =   45
      TabIndex        =   21
      Top             =   7650
      Width           =   5160
   End
End
Attribute VB_Name = "frmHenkel_Precios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cfm_Click()
    If cfm.ListItems.Count = 0 Then Exit Sub
    With cfm.ListItems(cfm.selectedItem.Index)
        txtCFM(0) = .SubItems(1)
        txtCFM(1) = .SubItems(2)
        txtCFM(2) = .SubItems(3)
        txtCFM(3) = .SubItems(4)
    End With
End Sub

Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If txtCFM(0) = "" Or txtCFM(0) = "" Or txtCFM(0) = "" Or txtCFM(0) = "" Then
        MsgBox "Debe completar todos los campos.", vbCritical, App.Title
    Else
        Dim oCFM As New clsHenkel_cfm
        With oCFM
            .setP_INICIO = txtCFM(0)
            .setP_FIN = txtCFM(1)
            .setPRECIO = moneda_bd(txtCFM(2))
            .setTASA = moneda_bd(txtCFM(3))
            .Insertar
        End With
        cargarListadoCFM
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmHenkel_Precios"
End Sub

Private Sub cmdAnadirProbeta_Click()
   On Error GoTo cmdAnadirProbeta_Click_Error

    If txtProbetas(0) = "" Or txtProbetas(0) = "" Or txtProbetas(0) = "" Or txtProbetas(0) = "" Then
        MsgBox "Debe completar todos los campos.", vbCritical, App.Title
    Else
        Dim op As New clsHenkel_precios
        With op
            .setCODIGO = txtProbetas(0)
            .setPRECIO_E = moneda_bd(txtProbetas(1))
            .setPRECIO_ABCD = moneda_bd(txtProbetas(2))
            .setCOMENTARIOS = txtProbetas(3)
            .Insertar
        End With
        cargarListadoProbetas
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadirProbeta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadirProbeta_Click of Formulario frmHenkel_Precios"

End Sub

Private Sub cmdEliminar_Click()
   On Error GoTo cmdEliminar_Click_Error

    If cfm.ListItems.Count = 0 Then Exit Sub
    If MsgBox("¿Esta seguro de eliminar?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oCFM As New clsHenkel_cfm
        With oCFM
            .Eliminar cfm.ListItems(cfm.selectedItem.Index).Text
        End With
        cargarListadoCFM
    End If

   On Error GoTo 0
   Exit Sub

cmdEliminar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminar_Click of Formulario frmHenkel_Precios"
End Sub

Private Sub cmdEliminarProbeta_Click()
   On Error GoTo cmdEliminarProbeta_Click_Error

    If probetas.ListItems.Count = 0 Then Exit Sub
    If MsgBox("¿Esta seguro de eliminar?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim op As New clsHenkel_precios
        With op
            .Eliminar probetas.ListItems(probetas.selectedItem.Index).Text
        End With
        cargarListadoProbetas
    End If

   On Error GoTo 0
   Exit Sub

cmdEliminarProbeta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminarProbeta_Click of Formulario frmHenkel_Precios"

End Sub

Private Sub cmdModificar_Click()
   On Error GoTo cmdModificar_Click_Error

    If txtCFM(0) = "" Or txtCFM(0) = "" Or txtCFM(0) = "" Or txtCFM(0) = "" Then
        MsgBox "Debe completar todos los campos.", vbCritical, App.Title
    Else
        Dim oCFM As New clsHenkel_cfm
        With oCFM
            .setP_INICIO = txtCFM(0)
            .setP_FIN = txtCFM(1)
            .setPRECIO = moneda_bd(txtCFM(2))
            .setTASA = moneda_bd(txtCFM(3))
            .Modificar cfm.ListItems(cfm.selectedItem.Index).Text
        End With
        cargarListadoCFM
    End If

   On Error GoTo 0
   Exit Sub

cmdModificar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificar_Click of Formulario frmHenkel_Precios"

End Sub

Private Sub cmdModificarProbeta_Click()
   On Error GoTo cmdModificarProbeta_Click_Error

    If txtProbetas(0) = "" Or txtProbetas(0) = "" Or txtProbetas(0) = "" Or txtProbetas(0) = "" Then
        MsgBox "Debe completar todos los campos.", vbCritical, App.Title
    Else
        Dim op As New clsHenkel_precios
        With op
            .setCODIGO = txtProbetas(0)
            .setPRECIO_E = moneda_bd(txtProbetas(1))
            .setPRECIO_ABCD = moneda_bd(txtProbetas(2))
            .setCOMENTARIOS = txtProbetas(3)
            .Modificar probetas.ListItems(probetas.selectedItem.Index).Text
        End With
        cargarListadoProbetas
    End If

   On Error GoTo 0
   Exit Sub

cmdModificarProbeta_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificarProbeta_Click of Formulario frmHenkel_Precios"

End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargarListadoCFM
    cargarListadoProbetas
    
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbDimension, DECODIFICADORA.DECODIFICADORA_DIMENSIONES
End Sub
Private Sub cabecera()
    With cfm.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Mínimo", 1000, lvwColumnCenter
        .Add , , "Máximo", 1000, lvwColumnCenter
        .Add , , "Coste (€)", 1300, lvwColumnCenter
        .Add , , "Tasa (%)", 1200, lvwColumnCenter
    End With
    With probetas.ColumnHeaders
        .Add , , "Código", 1000, lvwColumnLeft
        .Add , , "Precio E (No primer)", 1100, lvwColumnCenter
        .Add , , "Precio A,B,C,D (Primer)", 1100, lvwColumnCenter
        .Add , , "Comentarios", probetas.Width - 3600, lvwColumnLeft
    End With
    With listaSETS.ColumnHeaders
        .Add , , "Código", 800, lvwColumnLeft
        .Add , , "Dimensión", 2000, lvwColumnCenter
        .Add , , "Precio E (No primer)", 1100, lvwColumnCenter
        .Add , , "Precio A,B,C,D (Primer)", 1100, lvwColumnCenter
        .Add , , "DIMENSION_ID", 0, lvwColumnCenter
    End With
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cargarListadoCFM()
    ' CFM
    Dim rs As ADODB.Recordset
    Dim oCFM As New clsHenkel_cfm
    cfm.ListItems.Clear
    Set rs = oCFM.Listado()
    If rs.RecordCount > 0 Then
        Do
            With cfm.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = moneda(rs(3))
                .SubItems(4) = rs(4)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
                
End Sub
Private Sub cargarListadoProbetas()
    Dim rs As ADODB.Recordset
    Dim op As New clsHenkel_precios
    probetas.ListItems.Clear
    Set rs = op.Listado()
    If rs.RecordCount > 0 Then
        Do
            With probetas.ListItems.Add(, , rs(0))
                .SubItems(1) = moneda(rs(1))
                .SubItems(2) = moneda(rs(2))
                .SubItems(3) = rs(3)
                If Mid(rs(0), 3, 1) = "2" Then
                    .SmallIcon = 3
                End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub
Private Sub cargarListadoSets()
    Dim rs As ADODB.Recordset
    Dim oh As New clsHenkel_sets
    listaSETS.ListItems.Clear
    Set rs = oh.Listado(txtCFM(7))
    If rs.RecordCount > 0 Then
        Do
            With listaSETS.ListItems.Add(, , rs(0))
                .SubItems(1) = IIf(IsNull(rs(1)), "", rs(1))
                .SubItems(2) = moneda(rs(2))
                .SubItems(3) = moneda(rs(3))
                .SubItems(4) = rs(4)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub



Private Sub listaSETS_Click()
   On Error GoTo listaSETS_Click_Error

    If listaSETS.ListItems.Count = 0 Then Exit Sub
    With listaSETS.ListItems(listaSETS.selectedItem.Index)
        cmbDimension.MostrarElemento .SubItems(4)
        txtCFM(5) = .SubItems(2)
        txtCFM(4) = .SubItems(3)
    End With

   On Error GoTo 0
   Exit Sub

listaSETS_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure listaSETS_Click of Formulario frmHenkel_Precios"

End Sub

Private Sub probetas_Click()
    If probetas.ListItems.Count = 0 Then Exit Sub
    With probetas.ListItems(probetas.selectedItem.Index)
        txtProbetas(0) = .Text
        txtProbetas(1) = .SubItems(1)
        txtProbetas(2) = .SubItems(2)
        txtProbetas(3) = .SubItems(3)
    End With
    listaSETS.ListItems.Clear
    limpiarSets
    If Mid(txtProbetas(0), 3, 1) = "2" Then
        frmSETS.visible = True
        txtCFM(7) = txtProbetas(0)
        cargarListadoSets
    Else
        frmSETS.visible = False
    End If
End Sub
Private Sub limpiarSets()
    cmbDimension.limpiar
    txtCFM(7) = ""
    txtCFM(4) = ""
    txtCFM(5) = ""
End Sub
Private Sub PushButton1_Click()
    If txtCFM(4) = "" Or txtCFM(5) = "" Or cmbDimension.getTEXTO = "" Then
        MsgBox "Debe completar todos los campos.", vbCritical, App.Title
    Else
        Dim op As New clsHenkel_sets
        With op
            .setPRECIO_E = moneda_bd(txtCFM(5))
            .setPRECIO_ABCD = moneda_bd(txtCFM(4))
            .setDIMENSION_ID = cmbDimension.getPK_SALIDA
            .setDIMENSION = cmbDimension.getTEXTO
            .Modificar txtCFM(7), listaSETS.ListItems(listaSETS.selectedItem.Index).SubItems(4)
        End With
        cargarListadoSets
    End If

End Sub

Private Sub PushButton2_Click()
   On Error GoTo PushButton2_Click_Error

    If txtCFM(4) = "" Or txtCFM(5) = "" Or cmbDimension.getTEXTO = "" Then
        MsgBox "Debe completar todos los campos.", vbCritical, App.Title
    Else
        Dim op As New clsHenkel_sets
        With op
            .setCODIGO = txtCFM(7)
            .setPRECIO_E = moneda_bd(txtCFM(5))
            .setPRECIO_ABCD = moneda_bd(txtCFM(4))
            .setDIMENSION_ID = cmbDimension.getPK_SALIDA
            .setDIMENSION = cmbDimension.getTEXTO
            .Insertar
        End With
        cargarListadoSets
    End If

   On Error GoTo 0
   Exit Sub

PushButton2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PushButton2_Click of Formulario frmHenkel_Precios"

End Sub

Private Sub PushButton3_Click()
   On Error GoTo PushButton3_Click_Error

    If listaSETS.ListItems.Count = 0 Then Exit Sub
    If MsgBox("¿Esta seguro de eliminar?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim op As New clsHenkel_sets
        With op
            .Eliminar listaSETS.ListItems(listaSETS.selectedItem.Index).Text, listaSETS.ListItems(listaSETS.selectedItem.Index).SubItems(4)
        End With
        cargarListadoSets
    End If

   On Error GoTo 0
   Exit Sub

PushButton3_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PushButton3_Click of Formulario frmHenkel_Precios"

End Sub

Private Sub txtCFM_GotFocus(Index As Integer)
    txtCFM(Index).SelStart = 0
    txtCFM(Index).SelLength = Len(txtCFM(Index))

End Sub

Private Sub txtCFM_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 2 Or Index = 4 Or Index = 5 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
End Sub
Private Sub txtCFM_LostFocus(Index As Integer)
    If Index = 2 Or Index = 4 Or Index = 5 Then
        txtCFM(Index) = moneda(txtCFM(Index))
    End If
End Sub

Private Sub txtProbetas_GotFocus(Index As Integer)
    txtProbetas(Index).SelStart = 0
    txtProbetas(Index).SelLength = Len(txtProbetas(Index))
End Sub

Private Sub txtProbetas_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 1 Or Index = 2 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
End Sub
Private Sub txtProbetas_LostFocus(Index As Integer)
    If Index = 1 Or Index = 2 Then
        txtProbetas(Index) = moneda(txtProbetas(Index))
    End If
End Sub
