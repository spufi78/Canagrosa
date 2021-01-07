VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmEmpleados_Matriz 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Matriz de cualificaciones"
   ClientHeight    =   10905
   ClientLeft      =   2790
   ClientTop       =   1035
   ClientWidth     =   14925
   Icon            =   "frmEmpleados_Matriz.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   10905
   ScaleWidth      =   14925
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   735
      Left            =   10890
      TabIndex        =   35
      Top             =   10170
      Visible         =   0   'False
      Width           =   2580
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   13050
      Top             =   9900
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   107
      ImageHeight     =   144
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.Frame frmUsuarios 
      BackColor       =   &H00FFFFFF&
      Height          =   9690
      Left            =   1305
      TabIndex        =   29
      Top             =   630
      Visible         =   0   'False
      Width           =   12930
      Begin VB.CommandButton cmdDesmarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desmarcar Todas"
         Height          =   600
         Left            =   1530
         Picture         =   "frmEmpleados_Matriz.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   9000
         Width           =   1410
      End
      Begin VB.CommandButton cmdMarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marcar Todas"
         Height          =   600
         Left            =   90
         Picture         =   "frmEmpleados_Matriz.frx":711C
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   9000
         Width           =   1410
      End
      Begin MSComctlLib.ListView listaUsuarios 
         Height          =   8760
         Left            =   90
         TabIndex        =   30
         Top             =   180
         Width           =   12720
         _ExtentX        =   22437
         _ExtentY        =   15452
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         Checkboxes      =   -1  'True
         PictureAlignment=   4
         TextBackground  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial Black"
            Size            =   9.75
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "F5 - CERRAR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   11340
         TabIndex        =   33
         Top             =   9210
         Width           =   1410
      End
   End
   Begin VB.Frame frmLeyenda 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   90
      TabIndex        =   14
      Top             =   10035
      Width           =   13200
      Begin VB.Image Image1 
         Height          =   480
         Index           =   4
         Left            =   2205
         Picture         =   "frmEmpleados_Matriz.frx":D96E
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "EN FORMACION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2700
         TabIndex        =   27
         Top             =   360
         Width           =   1530
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "FORMADOR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   11565
         TabIndex        =   18
         Top             =   360
         Width           =   1170
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   3
         Left            =   10980
         Picture         =   "frmEmpleados_Matriz.frx":DD6B
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "RE-CUALIFICADO FUERA FECHA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   7605
         TabIndex        =   17
         Top             =   360
         Width           =   3105
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   2
         Left            =   7110
         Picture         =   "frmEmpleados_Matriz.frx":DFCD
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "RE-CUALIFICADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   5130
         TabIndex        =   16
         Top             =   360
         Width           =   1650
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   1
         Left            =   4635
         Picture         =   "frmEmpleados_Matriz.frx":E3F5
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "CUALIFICADO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   585
         TabIndex        =   15
         Top             =   360
         Width           =   1305
      End
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   90
         Picture         =   "frmEmpleados_Matriz.frx":E80E
         Top             =   225
         Width           =   480
      End
   End
   Begin MSFlexGridLib.MSFlexGrid glista 
      Height          =   7245
      Left            =   45
      TabIndex        =   8
      Top             =   2700
      Width           =   14865
      _ExtentX        =   26220
      _ExtentY        =   12779
      _Version        =   393216
      BackColor       =   16777215
      BackColorFixed  =   12640511
      BackColorSel    =   8553090
      BackColorBkg    =   12632256
      HighLight       =   2
      Appearance      =   0
   End
   Begin MSComctlLib.ImageList imglst 
      Left            =   13230
      Top             =   10125
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   100
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Matriz.frx":EC2D
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Matriz.frx":EEA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Matriz.frx":F20E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Matriz.frx":F694
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Matriz.frx":FA02
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Matriz.frx":FE86
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Matriz.frx":1022C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Matriz.frx":10496
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   13725
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   9960
      Width           =   1125
   End
   Begin VB.Frame fraDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   45
      TabIndex        =   5
      Top             =   675
      Width           =   14835
      Begin VB.CheckBox chkReducida 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reducida"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   8910
         TabIndex        =   36
         Top             =   1485
         Width           =   1110
      End
      Begin VB.Frame frmTipo 
         BackColor       =   &H00C0C0C0&
         Height          =   1770
         Left            =   10575
         TabIndex        =   22
         Top             =   135
         Width           =   2715
         Begin VB.CheckBox chkBaja 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mostrar empleados de Baja"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   180
            TabIndex        =   40
            Top             =   1485
            Width           =   2370
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "En formación"
            Height          =   195
            Index           =   4
            Left            =   180
            TabIndex        =   28
            Top             =   1125
            Width           =   2310
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Formadores"
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   26
            Top             =   900
            Width           =   2310
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Recualificaciones Pendientes"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   25
            Top             =   675
            Width           =   2490
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Recualificaciones"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   24
            Top             =   450
            Width           =   2310
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todas"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   23
            Top             =   225
            Value           =   -1  'True
            Width           =   2310
         End
      End
      Begin VB.CheckBox chkENAC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENAC"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4275
         TabIndex        =   21
         Top             =   1485
         Width           =   810
      End
      Begin VB.CheckBox chkNADCAP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "NADCAP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   5265
         TabIndex        =   20
         Top             =   1485
         Width           =   960
      End
      Begin VB.CheckBox chkEQA 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "EQA"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   6480
         TabIndex        =   19
         Top             =   1485
         Width           =   750
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   1095
         Left            =   13320
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   495
         Width           =   1410
      End
      Begin VB.TextBox txtFiltroNombre 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1530
         TabIndex        =   3
         Top             =   1260
         Width           =   2475
      End
      Begin pryCombo.miCombo cmbPNT 
         Height          =   315
         Left            =   1530
         TabIndex        =   2
         Top             =   900
         Width           =   8955
         _ExtentX        =   15796
         _ExtentY        =   556
      End
      Begin MSDataListLib.DataCombo cmbfamilia 
         Height          =   315
         Left            =   1530
         TabIndex        =   0
         Top             =   180
         Width           =   7920
         _ExtentX        =   13970
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
         Left            =   1530
         TabIndex        =   1
         Top             =   540
         Width           =   7920
         _ExtentX        =   13970
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
      Begin MSDataListLib.DataCombo cmbempresas 
         Height          =   315
         Left            =   1530
         TabIndex        =   37
         Top             =   1575
         Width           =   2475
         _ExtentX        =   4366
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
      Begin VB.CheckBox chkMTL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "MTL"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   7515
         TabIndex        =   39
         Top             =   1485
         Width           =   750
      End
      Begin VB.Label lblEmpresa 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Empresa"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   38
         Top             =   1620
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "SubFamilia P.N.T."
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia P.N.T."
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   225
         Width           =   990
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre P.N.T."
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   9
         Top             =   1305
         Width           =   1065
      End
      Begin VB.Label lblValor 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código P.N.T."
         Height          =   195
         Left            =   90
         TabIndex        =   6
         Top             =   945
         Width           =   1005
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "F5 - SELECCIONAR EMPLEADOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   9225
      TabIndex        =   34
      Top             =   180
      Width           =   3480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Matriz de Cualificaciones"
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
      TabIndex        =   11
      Top             =   45
      Width           =   2610
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   14310
      Picture         =   "frmEmpleados_Matriz.frx":1288D
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de la Matriz de Cualificaciones. Pulse sobre la cualificación para ver el detalle."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   10
      Top             =   360
      Width           =   6045
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   14895
   End
End
Attribute VB_Name = "frmEmpleados_Matriz"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public CARGA_INICIAL As Boolean

Private Sub Check1_Click()

End Sub

Private Sub chkENAC_Click()
    cargar_datos
End Sub

Private Sub chkEQA_Click()
    cargar_datos
End Sub

Private Sub chkMTL_Click()
    cargar_datos
End Sub
Private Sub chkNADCAP_Click()
    cargar_datos
End Sub
Private Sub chkReducida_Click()
    cargar_datos
End Sub

Private Sub cmbempresas_Change()
    cargar_datos
End Sub
Private Sub cmbfamilia_Change()
    cargar_datos
End Sub

Private Sub cmbPNT_change()
    cargar_datos
End Sub

Private Sub cmbSubfamilia_Change()
    cargar_datos
End Sub

Private Sub cmdBuscar_Click()
    cargar_datos
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To listaUsuarios.ListItems.Count
        listaUsuarios.ListItems(i).Checked = False
    Next

End Sub

Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To listaUsuarios.ListItems.Count
        listaUsuarios.ListItems(i).Checked = True
    Next
End Sub

Private Sub Command1_Click()
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    Dim destino As String
    destino = App.Path & "\MATRIZ " & Format(Date, "dd-mm-yyyy") & "-" & Format(Time, "hhmmss") & ".xls"
    FileCopy ReadINI(App.Path & "\config.ini", "Documentos", "Plantillas") & "\MATRIZ.xls", destino
    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Open(destino)
    Set XLS = XLW.Worksheets(1)
    Dim i As Integer
    Dim fila As Integer
    fila = 2
    ' Cabecera
    For i = 2 To glista.COLS - 1
        If i Mod 2 = 0 Then
            XLS.Cells(1, (i / 2) + 1) = glista.TextMatrix(0, i)
        End If
    Next
    For i = 1 To glista.Rows - 1
        XLS.Cells(fila, 1) = glista.TextMatrix(i, 1) ' PNT
'        XLS.Cells(fila, 2) = glista.TextMatrix(i, 2)
        fila = fila + 1
    Next
    XLS.Columns(1).AutoFit
    XLA.visible = True
'    XLW.Close
'    XLA.Quit
End Sub

Private Sub Form_Activate()
    If CARGA_INICIAL Then
'        cargar_datos
        CARGA_INICIAL = False
    End If
End Sub
Private Sub cargar_empresas()
    Dim ooe As New clsEmpleados_Empresas
    Set cmbempresas.RowSource = ooe.Listado
    cmbempresas.ListField = "DESCRIPCION"
    cmbempresas.BoundColumn = "ID_EMPRESA"
    Set ooe = Nothing
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   On Error GoTo Form_KeyUp_Error

    Me.MousePointer = 0
    Select Case KeyCode
        Case 116 ' F5 Datos especiales
            If frmUsuarios.visible = False Then
                frmUsuarios.visible = True
            Else
                frmUsuarios.visible = False
                Dim i As Integer
                Dim s As String
                For i = 1 To listaUsuarios.ListItems.Count
                    If listaUsuarios.ListItems(i).Checked = True Then
                        s = s & listaUsuarios.ListItems(i).Tag & ","
                    End If
                Next
'                MsgBox s
                If s <> "" Then
                    cargar_datos Left(s, Len(s) - 1)
                End If
            End If
'            frmUsuarios.Visible = Not frmUsuarios.Visible
    End Select

   On Error GoTo 0
   Exit Sub

Form_KeyUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_KeyUp of Formulario frmEmpleados_Matriz"

End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cargar_combos
    CARGA_INICIAL = True
'    cargar_datos
End Sub


Private Sub Form_Resize()

    glista.Width = Me.ScaleWidth
    glista.Left = 0
    
    cmdcancel.top = Me.ScaleHeight - 60 - cmdcancel.Height
    cmdcancel.Left = Me.ScaleWidth - 60 - cmdcancel.Width
    
    glista.Height = Me.ScaleHeight - glista.top - cmdcancel.Height - 120
    
    fraDatos.Width = Me.ScaleWidth
    cmdBuscar.Left = fraDatos.Width - cmdBuscar.Width - 100
    frmTipo.Left = cmdBuscar.Left - frmTipo.Width - 100
    
    fondo.Width = Me.ScaleWidth
    imagen.Left = fondo.Width - imagen.Width - 100
    frmLeyenda.top = Me.Height - frmLeyenda.Height - 600
    
    glista.ColWidth(1) = glista.Width * 0.4

    frmUsuarios.Left = (Screen.Width / 2) - (frmUsuarios.Width / 2)
    frmUsuarios.top = (Screen.Height / 2) - (frmUsuarios.Height / 2)

End Sub

Private Sub glista_DblClick()
    If glista.Col > 1 And glista.Row > 0 And glista.TextMatrix(glista.Row, glista.Col + 1) <> "" Then
        frmEmpleados_Cualificaciones_Nueva.EMPLEADO_ID = glista.TextMatrix(0, glista.Col + 1)
        frmEmpleados_Cualificaciones_Nueva.ID_CUALIFICACION = glista.TextMatrix(glista.Row, glista.Col + 1)
        frmEmpleados_Cualificaciones_Nueva.Show 1
    End If
End Sub
Private Sub cargar_datos(Optional empleados As String)
    Dim rs As ADODB.Recordset
    Dim oEC As New clsEmpleados_cualificaciones
    Dim tipo As Integer
    If opTipo(0).Value = True Then
        tipo = 0
    ElseIf opTipo(1).Value = True Then
        tipo = 1
    ElseIf opTipo(2).Value = True Then
        tipo = 2
    ElseIf opTipo(3).Value = True Then
        tipo = 3
    Else
        tipo = 4
    End If
    Dim emple As String
    If empleados <> "" Then
        emple = empleados
    End If
    If cmbPNT.getTEXTO = "" Then
        Set rs = oEC.MatrizCualificaciones(cmbFamilia.BoundText, cmbSubfamilia.BoundText, 0, txtFiltroNombre, chkENAC.Value, chkNADCAP.Value, chkEQA.Value, tipo, emple, cmbempresas.BoundText, chkMTL.Value, chkBaja.Value)
    Else
        Set rs = oEC.MatrizCualificaciones(cmbFamilia.BoundText, cmbSubfamilia.BoundText, cmbPNT.getPK_SALIDA, txtFiltroNombre, chkENAC.Value, chkNADCAP.Value, chkEQA.Value, tipo, emple, cmbempresas.BoundText, chkMTL.Value, chkBaja.Value)
    End If
    
        With glista
            .Clear
            .Rows = 1
            .COLS = 2
            .TextMatrix(0, 0) = "id_documento"
            .TextMatrix(0, 1) = "Documento Calidad"
            .ColWidth(0) = 0
            .AllowUserResizing = flexResizeColumns
            .RowHeightMin = (32 * Screen.TwipsPerPixelX) + (Screen.TwipsPerPixelX * 2)
            If chkReducida.Value = Checked Then
                .ColWidth(1) = 1200
            Else
                .ColWidth(1) = 7000
            End If
            .FixedCols = 1
'            .Gridlines = flexGridNone
        End With
    
    If rs.RecordCount > 0 Then
        Dim j As Integer
        Dim columna As Integer
        Dim fila As Integer
        Dim faux As Date
        Dim ES_FORMADOR As Boolean
        fila = 1
        glista.Rows = glista.Rows + 1
        glista.Redraw = False
        Set listaUsuarios.Icons = Nothing
        ImageList1.ListImages.Clear
        listaUsuarios.ListItems.Clear
       Do
            With glista
                .TextMatrix(fila, 0) = rs(0) ' ID_DOCUMENTO
                If chkReducida.Value = Checked Then
                    .TextMatrix(fila, 1) = rs(10) 'CODIGO DOCUMENTO
                Else
                    .TextMatrix(fila, 1) = rs(2) 'NOMBRE DOCUMENTO
                End If
                ' Buscamos el usuario
                columna = 0
                For j = 2 To .COLS - 1
                    If chkReducida.Value = Checked Then
                        If .TextMatrix(0, j) = Left(rs(3), 15) & "..." Then
                            columna = j
                            Exit For
                        End If
                    Else
                        If .TextMatrix(0, j) = rs(3) Then
                            columna = j
                            Exit For
                        End If
                    End If
                Next
                If columna = 0 Then
                    columna = .COLS
                    .COLS = .COLS + 2
                    If chkReducida.Value = Checked Then
                        .TextMatrix(0, columna) = Left(rs(3), 15) & "..." ' NOMBRE EMPLEADO
                        .ColWidth(columna) = 1500
                    Else
                        .TextMatrix(0, columna) = rs(3) ' NOMBRE EMPLEADO
                        .ColWidth(columna) = 2500
                    End If
                    .TextMatrix(0, columna + 1) = rs(1) ' ID EMPLEADO
                    .ColWidth(columna + 1) = 0
                    ' Imagen
                    On Error Resume Next
                    If Dir(Replace(rs(9), "/", "\")) = "" Then
                        ImageList1.ListImages.Add (.COLS / 2) - 1, , LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "no"))
                    Else
                        If rs(9) <> "" Then
                            ImageList1.ListImages.Add (.COLS / 2) - 1, , LoadPicture(Replace(rs(9), "/", "\"))
                        Else
                            ImageList1.ListImages.Add (.COLS / 2) - 1, , LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "no"))
                        End If
                    End If
                    listaUsuarios.ListItems.Add , , rs(3)
                    listaUsuarios.ListItems(listaUsuarios.ListItems.Count).Checked = True
                    listaUsuarios.ListItems(listaUsuarios.ListItems.Count).Tag = rs(1)
                End If
                .ColAlignment(columna) = flexAlignCenterCenter
                .Col = columna
                .Row = fila
                ' Imagen
                ' 1. A     ' Solo cualificado
                ' 2. A V   ' Cualificado y recualificado en el ultimo año
                ' 3. A V P ' Cualificado, recualificado en el ultimo año y formador
                ' 4. A R   ' Cualificado y no recualificado en el ultimo año
                ' 5. A R P ' Cualificado, no recualificado en el ultimo año y formador
                ' 6. A   P ' Cualificado y formador
                If rs(8) = 1 Then
                    ES_FORMADOR = True
                Else
                    ES_FORMADOR = False
                End If
                If rs(11) = 1 Then ' FORMADOR NO RECUALIFICADO
                    Set .CellPicture = imglst.ListImages(8).Picture
                Else
                    Dim numero_anos_recualificacion As Integer
                    Dim numero_anos_recualificacion_c As Integer
                    numero_anos_recualificacion = 2
                    numero_anos_recualificacion_c = 3
                    If Format(rs(6), "yyyy-mm-dd") = "1900-01-01" Then ' En formacion
                        Set .CellPicture = imglst.ListImages(7).Picture
                    Else
                        If Format(rs(4), "yyyy-mm-dd") = "1900-01-01" Then ' No recualificado
                            If Left(rs(2), 5) = "PNT C" Then
                                faux = DateAdd("yyyy", numero_anos_recualificacion_c, rs(6))
                            Else
                                faux = DateAdd("yyyy", numero_anos_recualificacion, rs(6))
                            End If
                            If Format(faux, "yyyy-mm-dd") < Format(Date, "yyyy-mm-dd") Then ' Cualificado y formador
                                If ES_FORMADOR Then
                                    If rs(5) = 1 Then ' Si en historico
                                        Set .CellPicture = imglst.ListImages(3).Picture
                                    Else
                                        Set .CellPicture = imglst.ListImages(5).Picture
                                    End If
                                Else
                                    If rs(5) = 1 Then ' Si en historico
                                        Set .CellPicture = imglst.ListImages(2).Picture
                                    Else
                                        Set .CellPicture = imglst.ListImages(4).Picture
                                    End If
                                End If
                            Else ' VERDE
                                If ES_FORMADOR Then
                                    If rs(5) = 1 Then ' Si en historico
                                        Set .CellPicture = imglst.ListImages(3).Picture
                                    Else
                                        Set .CellPicture = imglst.ListImages(6).Picture
                                    End If
                                Else
                                    If rs(5) = 1 Then ' Si en historico
                                        Set .CellPicture = imglst.ListImages(2).Picture
                                    Else
                                        Set .CellPicture = imglst.ListImages(1).Picture
                                    End If
                                End If
                            End If
                        Else
                            If Left(rs(2), 5) = "PNT C" Then
                                faux = DateAdd("yyyy", numero_anos_recualificacion_c, rs(4))
                            Else
                                faux = DateAdd("yyyy", numero_anos_recualificacion, rs(4))
                            End If
                            If Format(faux, "yyyy-mm-dd") < Format(Date, "yyyy-mm-dd") Then ' Cualificado y formador
                                If ES_FORMADOR Then
                                    Set .CellPicture = imglst.ListImages(5).Picture
                                Else
                                    Set .CellPicture = imglst.ListImages(4).Picture
                                End If
                            Else ' VERDE
                                If ES_FORMADOR Then
                                    Set .CellPicture = imglst.ListImages(3).Picture
                                Else
                                    Set .CellPicture = imglst.ListImages(2).Picture
                                End If
                            End If
                        End If
                    End If
                End If
                .CellPictureAlignment = 4
                .TextMatrix(fila, columna + 1) = rs(7)
            End With
            rs.MoveNext
            If Not rs.EOF Then
                If glista.TextMatrix(fila, 0) <> rs(0) Then
                    fila = fila + 1
                    glista.Rows = glista.Rows + 1
                End If
            End If
        Loop Until rs.EOF
        glista.Redraw = True
        Set listaUsuarios.Icons = ImageList1
    End If
    If glista.COLS > 2 Then
        glista.FixedCols = 2
        Dim i As Integer
        For i = 1 To (glista.COLS / 2) - 1
            listaUsuarios.ListItems(i).Icon = i
        Next
    End If
    Set rs = Nothing
    Set oEC = Nothing
End Sub

Private Sub cargar_combos()
    llenar_combo cmbPNT, New clsCa_documentos, 0, frmCA_Documento, " ANULADO = 0 "
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbFamilia, DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS
    oDeco.cargar_combo cmbSubfamilia, DECODIFICADORA.CA_DOCUMENTOS_SUBFAMILIAS
    cargar_empresas
'    cmbSubfamilia.BoundText = 4
End Sub

Private Sub opTipo_Click(Index As Integer)
    cargar_datos
End Sub
