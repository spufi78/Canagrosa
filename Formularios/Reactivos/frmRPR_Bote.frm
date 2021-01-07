VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmRPR_Bote 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Preparación de bote de reactivo propio"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13275
   Icon            =   "frmRPR_Bote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   13275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   780
      Left            =   10995
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7665
      Width           =   1050
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   780
      Left            =   9945
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   7665
      Width           =   1050
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   780
      Left            =   12060
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7665
      Width           =   1050
   End
   Begin MSComctlLib.ListView lReactivosAux 
      Height          =   1275
      Left            =   3060
      TabIndex        =   39
      Top             =   6300
      Visible         =   0   'False
      Width           =   9180
      _ExtentX        =   16193
      _ExtentY        =   2249
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
      NumItems        =   0
   End
   Begin VB.TextBox txtdatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   3150
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   7920
      Width           =   2085
   End
   Begin VB.TextBox txtdatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   1605
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   7920
      Width           =   1500
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo"
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
      Height          =   1905
      Left            =   45
      TabIndex        =   33
      Top             =   675
      Width           =   2400
      Begin VB.OptionButton optTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Suministro"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   270
         TabIndex        =   35
         Top             =   1080
         Width           =   1410
      End
      Begin VB.OptionButton optTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Reactivo propio"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   1
         Left            =   270
         TabIndex        =   34
         Top             =   540
         Value           =   -1  'True
         Width           =   1995
      End
   End
   Begin VB.TextBox txtdatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   555
      Index           =   2
      Left            =   5265
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   14
      Top             =   7920
      Width           =   4575
   End
   Begin VB.TextBox txtdatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   90
      MultiLine       =   -1  'True
      TabIndex        =   11
      Top             =   7920
      Width           =   1500
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos generales de preparación del reactivo"
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
      Height          =   1905
      Left            =   2475
      TabIndex        =   23
      Top             =   675
      Width           =   10725
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   555
         Index           =   5
         Left            =   1440
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   1305
         Width           =   7995
      End
      Begin VB.CheckBox chkFinalizado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Finalizado"
         Height          =   195
         Left            =   2925
         TabIndex        =   6
         Top             =   1035
         Width           =   1185
      End
      Begin VB.CheckBox chkUso 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Según Uso"
         Height          =   195
         Left            =   5490
         TabIndex        =   3
         Top             =   630
         Width           =   1185
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   7920
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   585
         Width           =   1500
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         Height          =   330
         Index           =   7
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   945
         Width           =   1275
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   585
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
         Format          =   60293121
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Index           =   1
         Left            =   4140
         TabIndex        =   2
         Top             =   585
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
         Format          =   60293121
         CurrentDate     =   38002
      End
      Begin pryCombo.miCombo cmbReactivos 
         Height          =   330
         Left            =   1440
         TabIndex        =   0
         Top             =   225
         Width           =   8970
         _ExtentX        =   15822
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fechaFinalizado 
         Height          =   330
         Left            =   4140
         TabIndex        =   7
         Top             =   945
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
         Format          =   60293121
         CurrentDate     =   38002
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
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
         TabIndex        =   40
         Top             =   1530
         Width           =   1260
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Volumen"
         Height          =   195
         Index           =   3
         Left            =   7200
         TabIndex        =   30
         Top             =   630
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducidad"
         Height          =   195
         Index           =   0
         Left            =   2925
         TabIndex        =   27
         Top             =   630
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabricación"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   26
         Top             =   630
         Width           =   825
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   25
         Top             =   990
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
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
         Index           =   8
         Left            =   90
         TabIndex        =   24
         Top             =   315
         Width           =   360
      End
   End
   Begin VB.CommandButton cmdFicha 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos Reactivo/Preparación"
      Height          =   915
      Left            =   90
      Picture         =   "frmRPR_Bote.frx":1272
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8595
      Width           =   2625
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   12165
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8640
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11070
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8640
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   2100
      Left            =   90
      TabIndex        =   10
      Top             =   5445
      Width           =   13065
      _ExtentX        =   23045
      _ExtentY        =   3704
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin MSComctlLib.ListView cp 
      Height          =   2100
      Left            =   45
      TabIndex        =   9
      Top             =   2925
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   3704
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
      NumItems        =   0
   End
   Begin MSComctlLib.ListView lReactivos 
      Height          =   2100
      Left            =   3060
      TabIndex        =   38
      Top             =   2925
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   3704
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
      NumItems        =   0
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "P. Referencia"
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
      Left            =   3150
      TabIndex        =   37
      Top             =   7695
      Width           =   1140
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Unidad"
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
      Left            =   1620
      TabIndex        =   36
      Top             =   7695
      Width           =   585
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de preparación de reactivo Propio / Suministro"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   32
      Top             =   360
      Width           =   3825
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de preparación de reactivo Propio / Suministro"
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
      TabIndex        =   31
      Top             =   45
      Width           =   5670
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observaciones"
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
      Left            =   5265
      TabIndex        =   29
      Top             =   7695
      Width           =   1260
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cantidad"
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
      TabIndex        =   28
      Top             =   7695
      Width           =   765
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Botes de Reactivos a utilizar y datos adicionales"
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
      Height          =   300
      Left            =   90
      TabIndex        =   22
      Top             =   5130
      Width           =   13080
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Componentes Reactivos Propio/Sustancia"
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
      Height          =   300
      Left            =   45
      TabIndex        =   21
      Top             =   2610
      Width           =   13155
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      Height          =   3525
      Left            =   45
      Top             =   5040
      Width           =   13155
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   13230
   End
End
Attribute VB_Name = "frmRPR_Bote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
'E0169-I
Public Reactivo As Long
Public LOTE_ID As Long
'E0169-F

Private Sub cabecera()
    With cp.ColumnHeaders
        .Add , , "Reactivo", 1900, lvwColumnLeft
        .Add , , "Cantidad", 970, lvwColumnCenter
        .Add , , "TIPO_REACTIVO_EX_ID", 1, lvwColumnCenter
        .Add , , "CANTIDAD_DEFINIDA", 1, lvwColumnCenter
        .Add , , "PROCEDIMIENTO", 1, lvwColumnCenter
        .Add , , "UNIDAD", 1, lvwColumnCenter
        .Add , , "TIPO", 1, lvwColumnCenter ' Externo o Interno
    End With
    With lReactivos.ColumnHeaders
        .Add , , "Numero", 800, lvwColumnLeft
        .Add , , "Proveedor", 1800, lvwColumnCenter
        .Add , , "Código", 1000, lvwColumnCenter
        .Add , , "Lote", 1400, lvwColumnCenter
        .Add , , "F.Apertura", 1100, lvwColumnCenter
        .Add , , "F.Caducidad", 1100, lvwColumnCenter
        .Add , , "Observaciones", 1, lvwColumnCenter
        .Add , , "TIPO_BOTE_EX_ID", 1, lvwColumnCenter
        .Add , , "Cantidad", 900, lvwColumnCenter
        .Add , , "Procedimiento", 1500, lvwColumnCenter
        .Add , , "Unidad", 1, lvwColumnCenter
    End With
    
    With lReactivosAux.ColumnHeaders
        .Add , , "Numero", 800, lvwColumnLeft
        .Add , , "Proveedor", 1100, lvwColumnCenter
        .Add , , "Código", 1000, lvwColumnCenter
        .Add , , "Lote", 1000, lvwColumnCenter
        .Add , , "F.Apertura", 1100, lvwColumnCenter
        .Add , , "F.Caducidad", 1100, lvwColumnCenter
        .Add , , "Observaciones", 1400, lvwColumnCenter
        .Add , , "TIPO_BOTE_EX_ID", 1100, lvwColumnCenter
        .Add , , "CANTIDAD_DEFINIDA", 1100, lvwColumnCenter
        .Add , , "PROCEDIMIENTO", 1000, lvwColumnCenter
        .Add , , "UNIDAD", 1000, lvwColumnCenter
        .Add , , "TIPO_REACTIVO_EX_ID", 1000, lvwColumnCenter
        .Add , , "TIPO", 1000, lvwColumnCenter
    End With

    
    With lista.ColumnHeaders
        .Add , , "Numero", 800, lvwColumnLeft
        .Add , , "Proveedor", 2500, lvwColumnCenter
        .Add , , "Código", 3000, lvwColumnCenter
        .Add , , "Lote", 1500, lvwColumnCenter
        .Add , , "F.Apertura", 1100, lvwColumnCenter
        .Add , , "F.Cierre", 1100, lvwColumnCenter
        .Add , , "F.Caducidad", 1100, lvwColumnCenter
        .Add , , "TIPO_BOTE_EX_ID", 1, lvwColumnCenter
    End With
End Sub

Private Sub chkFinalizado_Click()
    ' M1067-I
    If chkFinalizado.value = Checked Then
        fechaFinalizado.Enabled = True
    Else
        fechaFinalizado.Enabled = False
    End If
    ' M1067-F
End Sub

Private Sub chkUso_Click()
    If chkUso.value = Checked Then
        fdesde(1).Enabled = False
    Else
        fdesde(1).Enabled = True
    End If
End Sub

Private Sub cmbReactivos_change()
    If cmbReactivos.getTEXTO <> "" Then
        cargar_componentes cmbReactivos.getPK_SALIDA
    End If
End Sub

Private Sub cmdAdd_Click()
End Sub

Private Sub cmdAnadir_Click()
    If cp.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        If txtdatos(0) = "" Then
            MsgBox "Indique la cantidad empleada.", vbExclamation, App.Title
            txtdatos(0).SetFocus
            Exit Sub
        End If
        With lReactivosAux.ListItems.Add(, , lista.ListItems(lista.selectedItem.Index).Text)
            .SubItems(1) = lista.ListItems(lista.selectedItem.Index).SubItems(1)
            .SubItems(2) = lista.ListItems(lista.selectedItem.Index).SubItems(2)
            .SubItems(3) = lista.ListItems(lista.selectedItem.Index).SubItems(3)
            .SubItems(4) = lista.ListItems(lista.selectedItem.Index).SubItems(4)
            .SubItems(5) = lista.ListItems(lista.selectedItem.Index).SubItems(6)
            .SubItems(6) = txtdatos(2) & " "
            .SubItems(7) = lista.ListItems(lista.selectedItem.Index).SubItems(7)
            .SubItems(8) = txtdatos(0) & " " & txtdatos(3) ' Cantidad / Unidades
            .SubItems(9) = txtdatos(4) ' Procedim
            .SubItems(11) = cp.ListItems(cp.selectedItem.Index).SubItems(2) ' Tipo_REX_ID
            .SubItems(12) = cp.ListItems(cp.selectedItem.Index).SubItems(6) ' INTERNO/EXTERNO
        End With
        'M1096-I
        cp.ListItems(cp.selectedItem.Index).SubItems(1) = txtdatos(0) & " " & txtdatos(3)
        cp.ListItems(cp.selectedItem.Index).SubItems(3) = txtdatos(0)
        'M1096-F
        cp_Click
    Else
        MsgBox "Hay que seleccionar un reactivo con el que fabricará el producto.", vbExclamation, App.Title
    End If

End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lReactivos.ListItems.Count > 0 Then
        ' Eliminar de la AUX
        Dim i As Integer
        For i = 1 To lReactivosAux.ListItems.Count
            If CLng(lReactivos.ListItems(lReactivos.selectedItem.Index).Text) = CLng(lReactivosAux.ListItems(i).Text) Then
                lReactivosAux.ListItems.Remove i
                Exit For
            End If
        Next
        lReactivos.ListItems.Remove lReactivos.selectedItem.Index
    End If
End Sub

Private Sub cmdFicha_Click()
    If cmbReactivos.getTEXTO <> "" Then
        greactivopr = cmbReactivos.getPK_SALIDA
        frmRPR_Reactivo.Show 1
        greactivopr = 0
    End If
End Sub

Private Sub cmdModificar_Click()
    If cp.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems.Count > 0 Then
        If txtdatos(0) = "" Then
            MsgBox "Indique la cantidad empleada.", vbExclamation, App.Title
            txtdatos(0).SetFocus
            Exit Sub
        End If
        
        Dim i As Integer
        For i = 1 To lReactivosAux.ListItems.Count
            If CInt(cp.ListItems(cp.selectedItem.Index).SubItems(2)) = CInt(lReactivosAux.ListItems(i).SubItems(11)) Then
                With lReactivosAux.ListItems(i)
                    .Text = lista.ListItems(lista.selectedItem.Index).Text
                    .SubItems(1) = lista.ListItems(lista.selectedItem.Index).SubItems(1)
                    .SubItems(2) = lista.ListItems(lista.selectedItem.Index).SubItems(2)
                    .SubItems(3) = lista.ListItems(lista.selectedItem.Index).SubItems(3)
                    .SubItems(4) = lista.ListItems(lista.selectedItem.Index).SubItems(4)
                    .SubItems(5) = lista.ListItems(lista.selectedItem.Index).SubItems(6)
                    .SubItems(6) = txtdatos(2) & " "
                    .SubItems(7) = lista.ListItems(lista.selectedItem.Index).SubItems(7)
                    .SubItems(8) = txtdatos(0) & " " & txtdatos(3) ' Cantidad
                    .SubItems(9) = txtdatos(4) ' Procedim
                    .SubItems(11) = cp.ListItems(cp.selectedItem.Index).SubItems(2) ' Tipo_REX_ID
                    .SubItems(12) = cp.ListItems(cp.selectedItem.Index).SubItems(6) ' INTERNO/EXTERNO
                End With
            End If
        Next
        'M1096-I
'        cp.ListItems(cp.selectedItem.Index).SubItems(1) = txtdatos(0) & " " & txtdatos(3)
'        cp.ListItems(cp.selectedItem.Index).SubItems(3) = txtdatos(0)
        'M1096-F
        cp_Click
    Else
        MsgBox "Hay que seleccionar un reactivo con el que fabricará el producto.", vbExclamation, App.Title
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If cp.ListItems.Count = 0 Then
        Exit Sub
    End If
    If cmbReactivos.getTEXTO = "" Then
        MsgBox "Seleccione el reactivo/suministro a fabricar.", vbExclamation, App.Title
        cmbReactivos.SetFocus
        Exit Sub
    End If
    If txtdatos(1) = "" Then
        MsgBox "Indique el volumen generado.", vbExclamation, App.Title
        txtdatos(1).SetFocus
        Exit Sub
    End If
    Dim orp As New clsRpr_botes
    Dim BOTE As Long
    With orp
        If optTipo(1).value = True Then
            .setTIPO_ID = 1
        Else
            .setTIPO_ID = 2
        End If
        .setTIPO_REACTIVO_PR_ID = cmbReactivos.getPK_SALIDA
        .setVOLUMEN = txtdatos(1)
        .setOBSERVACIONES = txtdatos(5)
        .setFECHA_FABRICACION = Format(fdesde(0), "yyyy-mm-dd")
        .setFECHA_CADUCIDAD = Format(fdesde(1), "yyyy-mm-dd")
        .setSEGUN_USO = chkUso.value
        If PK = 0 Then
            BOTE = .Insertar
        Else
            BOTE = .Modificar(PK)
            'M1067-I
            If chkFinalizado.value = Checked Then
                .Terminar PK, fechaFinalizado
            Else
                .TerminarAnular PK
            End If
            'M1067-F
        End If
        If BOTE <> 0 Then
            Dim oRPC As New clsRpr_botes_componentes
            If PK <> 0 Then
                oRPC.Eliminar PK
            End If
            Dim i As Integer
            For i = 1 To lReactivosAux.ListItems.Count
                With oRPC
                    If lReactivosAux.ListItems(i).Text <> "" Then
                        If PK = 0 Then
                            .setBOTE_PR_ID = BOTE
                        Else
                            .setBOTE_PR_ID = PK
                        End If
                        .setBOTE_EX_ID = lReactivosAux.ListItems(i).Text  ' Numero de bote externo
                        .setTIPO = Trim(lReactivosAux.ListItems(i).SubItems(12)) ' INTERNO/EXTERNO
                        .setCANTIDAD = lReactivosAux.ListItems(i).SubItems(8)
                        .setPROCEDIMIENTO = lReactivosAux.ListItems(i).SubItems(9)
                        .setOBSERVACIONES = lReactivosAux.ListItems(i).SubItems(6)
                        .setORDEN = i
                        .Insertar
                    End If
                End With
            Next
        End If
    End With
    If PK = 0 Then
        MsgBox "Bote de Reactivo/Suministro generado correctamente.", vbInformation, App.Title
    Else
        MsgBox "Bote de Reactivo/Suministro modificado correctamente.", vbInformation, App.Title
    End If
    Unload Me

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmRPR_Bote"
End Sub

Private Sub cp_Click()
    If cp.ListItems.Count > 0 Then
        'M1096-I
        'txtdatos(0) = cp.ListItems(cp.selectedItem.Index).SubItems(1)
        txtdatos(0) = cp.ListItems(cp.selectedItem.Index).SubItems(3)
        'M1096-F
        txtdatos(3) = cp.ListItems(cp.selectedItem.Index).SubItems(5)
        'M1096-I
        'txtdatos(4) = cp.ListItems(cp.selectedItem.Index).SubItems(4)
        'M1096-F
        If cp.ListItems(cp.selectedItem.Index).SubItems(6) = "E" Then
            cargar_botes_ex (cp.ListItems(cp.selectedItem.Index).SubItems(2))
            cargar_botes_ex_utilizados
        Else
            cargar_botes_internos (cp.ListItems(cp.selectedItem.Index).SubItems(2))
            cargar_botes_internos_utilizados
'            lReactivos.ListItems.Clear
        End If
        'M1096-I
        If lReactivos.ListItems.Count = 0 Then
            txtdatos(4) = cp.ListItems(cp.selectedItem.Index).SubItems(4)
        Else
            txtdatos(4) = lReactivos.ListItems(lReactivos.selectedItem.Index).SubItems(9)
        End If
        'M1096-F
    End If
End Sub
Private Sub lReactivos_Click()
    If lReactivos.ListItems.Count > 0 Then
        txtdatos(2) = lReactivos.ListItems(lReactivos.selectedItem.Index).SubItems(6)
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Text = lReactivos.ListItems(lReactivos.selectedItem.Index).Text Then
                Set lista.selectedItem = lista.ListItems(i)
                lista.selectedItem.EnsureVisible
                Exit Sub
            End If
        Next
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
'    txtdatos(3) = 1 ' Por defecto el número de botes a crear es 1
    fdesde(0) = Date
    fdesde(1) = Date
    If PK <> 0 Then
        Dim oBote As New clsRpr_botes
        oBote.Carga PK
        If oBote.getTIPO_ID = 2 Then
            optTipo(2).value = True
        Else
            optTipo(1).value = True
            cargar_reactivos
        End If
        chkUso.value = oBote.getSEGUN_USO
        cmbReactivos.MostrarElemento oBote.getTIPO_REACTIVO_PR_ID
        cmbReactivos.desactivar
        'M1067-I
        chkFinalizado.value = Unchecked
        fechaFinalizado.Enabled = False
        fechaFinalizado = Date
        If Not IsNull(oBote.getFECHA_FIN) Then
            If IsDate(oBote.getFECHA_FIN) Then
                chkFinalizado.value = Checked
                fechaFinalizado.Enabled = True
                fechaFinalizado = oBote.getFECHA_FIN
            End If
        End If
        'M1067-F
    Else
        'M1067-I
        chkFinalizado.Visible = False
        fechaFinalizado.Visible = False
        'M1067-F
        cargar_reactivos
    End If
End Sub
Private Sub cargar_componentes(tipo As Long)
    Dim oReactivos_Componentes As New clsRPR_Componentes

    Dim oTipos_reactivo_pr As New clsRPR_Tipos
    Dim rs As ADODB.Recordset
    Dim orx As New clsTipos_reactivo_ex
    limpiar_datos
    With oTipos_reactivo_pr
         If .CARGAR(tipo) = True Then
             txtdatos(7) = .getCODIGO
             txtdatos(1) = .getCANTIDAD
             
             If PK = 0 Then
                ' Caducidad
                Dim oTC As New clsTipos_caducidad
                oTC.CARGAR (.getTIPO_CADUCIDAD_ID)
                fdesde(1) = Date + oTC.getDIAS
                ' Componentes
                Set rs = oReactivos_Componentes.Componentes(tipo)
                If rs.RecordCount <> 0 Then
                   Do
'                       orx.CARGAR (rs("tipo_reactivo_ex_id"))
'                       With cp.ListItems.Add(, , orx.getNOMBRE)
'                            .SubItems(1) = rs(3)  ' Cantidad
'                            .SubItems(2) = Format(rs("TIPO_REACTIVO_ex_ID"), "0000")
'                            .SubItems(3) = rs(3) ' Cantidad RPR
'                            .SubItems(4) = rs(2) ' Proc. Referencia
'                            .SubItems(5) = rs(4) ' Unidad
'                       End With
                       With cp.ListItems.Add(, , rs(0))
                            .SubItems(1) = rs(3)  ' Cantidad
                            .SubItems(2) = Format(rs(1), "0000")
                            .SubItems(3) = rs(3) ' Cantidad RPR
                            .SubItems(4) = rs(2) ' Proc. Referencia
                            .SubItems(5) = rs(4) ' Unidad
                            .SubItems(6) = rs(6) ' Interno o Externo
                       End With
                       rs.MoveNext
                   Loop Until rs.EOF
                   cp_Click
                End If
            End If
         End If
    End With
    Set clsRPR_Tipos = Nothing
    If PK <> 0 Then
        Dim oBote As New clsRpr_botes
        With oBote
            .Carga PK
            txtdatos(1) = .getVOLUMEN
            txtdatos(5) = .getOBSERVACIONES
            fdesde(0) = Format(.getFECHA_FABRICACION, "dd-mm-yyyy")
            fdesde(1) = Format(.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            ' Componentes bote
            Set rs = oReactivos_Componentes.Componentes(tipo)
            If rs.RecordCount <> 0 Then
              Do
'                  orx.CARGAR (rs("tipo_reactivo_ex_id"))
                  With cp.ListItems.Add(, , rs(0))
                       .SubItems(1) = rs(3) & " " & rs(4) ' Cantidad
                       .SubItems(2) = Format(rs(1), "0000")
                       .SubItems(3) = rs(3) ' Cantidad RPR
                       .SubItems(4) = rs(2) ' Proc. Referencia
                       .SubItems(5) = rs(4) ' Unidad
                       .SubItems(6) = rs(6) ' Interno o Externo
                  End With
                 rs.MoveNext
              Loop Until rs.EOF
            End If
            ' Componentes utilizados
            Dim oReactivos_Componentes_botes As New clsRpr_botes_componentes
            Dim oREX_bote As New clsBotes_ex
            Dim oREX_tipo As New clsTipos_reactivo_ex
            Dim oREX_tipo_bote As New clsTipos_bote_ex
            Dim oProveedor As New clsProveedor
            Dim ORPR_BOTE As New clsRpr_botes
            Dim oRPR_TIPO As New clsRPR_Tipos
            Dim rs2 As ADODB.Recordset
            Set rs2 = oReactivos_Componentes_botes.Listado(PK)
            If rs2.RecordCount > 0 Then
              Do
                  If rs2("TIPO") = "E" Then
                      oREX_bote.CARGAR rs2("BOTE_EX_ID")
                      oREX_tipo_bote.CARGAR oREX_bote.getTIPO_BOTE_EX_ID
                      oREX_tipo.CARGAR oREX_tipo_bote.getTIPO_REACTIVO_EX_ID
                      oProveedor.Carga oREX_tipo_bote.getPROVEEDOR_ID
                    
                      With lReactivosAux.ListItems.Add(, , Format(oREX_bote.getID_BOTE_EX, "00000"))
                            .SubItems(1) = oProveedor.getNOMBRE
                            .SubItems(2) = oREX_tipo_bote.getCODIGO
                            .SubItems(3) = oREX_bote.getLOTE
                            .SubItems(4) = Format(oREX_bote.getFECHA_APERTURA, "dd-mm-yyyy")
                            .SubItems(5) = Format(oREX_bote.getFECHA_CADUCIDAD, "dd-mm-yyyy")
                            .SubItems(6) = rs2("OBSERVACIONES")
                            .SubItems(7) = oREX_bote.getTIPO_BOTE_EX_ID
                            .SubItems(8) = rs2("CANTIDAD")
                            .SubItems(9) = rs2("PROCEDIMIENTO")
                            .SubItems(11) = oREX_tipo_bote.getTIPO_REACTIVO_EX_ID
                            .SubItems(12) = "E"
                      End With
                  Else
                      ORPR_BOTE.Carga rs2("BOTE_EX_ID")
                      oRPR_TIPO.CARGAR ORPR_BOTE.getTIPO_REACTIVO_PR_ID
                      
                      With lReactivosAux.ListItems.Add(, , Format(ORPR_BOTE.getID_BOTE_PR, "00000"))
                            .SubItems(1) = "" ' Proveedor
                            .SubItems(2) = oRPR_TIPO.getCODIGO & "-" & ORPR_BOTE.getNUMERO
                            .SubItems(3) = "" ' LOTE
                            .SubItems(4) = Format(ORPR_BOTE.getFECHA_FABRICACION, "dd-mm-yyyy")
                            .SubItems(5) = Format(ORPR_BOTE.getFECHA_CADUCIDAD, "dd-mm-yyyy")
                            .SubItems(6) = rs2("OBSERVACIONES")
                            .SubItems(7) = ORPR_BOTE.getTIPO_REACTIVO_PR_ID
                            .SubItems(8) = rs2("CANTIDAD")
                            .SubItems(9) = rs2("PROCEDIMIENTO")
                            .SubItems(11) = ORPR_BOTE.getTIPO_REACTIVO_PR_ID
                            .SubItems(12) = "I"
                      End With
                  End If
                  rs2.MoveNext
              Loop Until rs2.EOF
            End If
            Set rs2 = Nothing
        End With
        cp_Click
    End If
End Sub
Private Sub cargar_botes_ex_utilizados()
    If cp.ListItems.Count = 0 Then
        Exit Sub
    End If
    ' Reactivos utilizados
    lReactivos.ListItems.Clear
    Dim i As Integer
    Dim j As Integer
    For i = 1 To lReactivosAux.ListItems.Count
        If CInt(cp.ListItems(cp.selectedItem.Index).SubItems(2)) = CInt(lReactivosAux.ListItems(i).SubItems(11)) Then
            With lReactivos.ListItems.Add(, , lReactivosAux.ListItems(i).Text)
                For j = 1 To 10
                    .SubItems(j) = lReactivosAux.ListItems(i).SubItems(j)
                Next
            End With
        End If
    Next
End Sub
Private Sub cargar_botes_internos_utilizados()
    If cp.ListItems.Count = 0 Then
        Exit Sub
    End If
    ' Reactivos utilizados
    lReactivos.ListItems.Clear
    Dim i As Integer
    Dim j As Integer
    For i = 1 To lReactivosAux.ListItems.Count
        If CInt(cp.ListItems(cp.selectedItem.Index).SubItems(2)) = CInt(lReactivosAux.ListItems(i).SubItems(11)) Then
            With lReactivos.ListItems.Add(, , lReactivosAux.ListItems(i).Text)
                For j = 1 To 10
                    .SubItems(j) = lReactivosAux.ListItems(i).SubItems(j)
                Next
            End With
        End If
    Next
End Sub

Private Sub cargar_botes_ex(tipo As Long)
    Dim rs As ADODB.Recordset
    Dim oBote_ex As New clsBotes_ex
    Set rs = oBote_ex.Listado_Para_Propio(tipo)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "00000"))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
                .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
                .SubItems(5) = Format(rs(5), "dd-mm-yyyy")
                .SubItems(6) = Format(rs(6), "dd-mm-yyyy")
                .SubItems(7) = rs(7)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set clsBotes_ex = Nothing
End Sub
Private Sub cargar_botes_internos(tipo As Long)
    Dim rs As ADODB.Recordset
    Dim oRPR As New clsRpr_botes
    Set rs = oRPR.Listado_por_tipo(tipo)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs("ID_BOTE_PR"), "00000")) 'Num
                .SubItems(1) = "" ' Proveedor
                .SubItems(2) = rs("CODIGO") & "-" & rs("NUMERO") ' Codigo
                .SubItems(3) = "" ' Lote
                .SubItems(4) = Format(rs("FECHA_FABRICACION"), "dd-mm-yyyy") ' Apertura
                .SubItems(5) = "" ' Cierre
                .SubItems(6) = Format(rs("FECHA_CADUCIDAD"), "dd-mm-yyyy") ' Caducidad
                .SubItems(7) = rs("TIPO_REACTIVO_PR_ID")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oRPR = Nothing
End Sub

'E0178-I
Private Sub Form_Unload(Cancel As Integer)
    Reactivo = 0
    LOTE_ID = 0
End Sub

Private Sub optTipo_Click(Index As Integer)
    cargar_reactivos
End Sub

'E0178-F
Private Sub txtdatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(Trim(txtdatos(Index)))
End Sub

'E0188-I
Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 3 ' cuadro de texto número de botes
            If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And KeyAscii <> 8 Then ' si no es un número
                KeyAscii = 0 ' se anula la pulsación
            End If
    End Select
End Sub
'E0188-I

Private Sub txtdatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = vbWhite
End Sub

Private Sub limpiar_datos()
    cp.ListItems.Clear
    lista.ListItems.Clear
    txtdatos(0) = ""
    txtdatos(2) = ""
End Sub

Private Sub cargar_reactivos()
    cmbReactivos.Limpiar
    Dim anulado As String
    If PK = 0 Then
        anulado = " AND ANULADO = 0 "
    End If
    If optTipo(1).value = True Then
        llenar_combo cmbReactivos, New clsRPR_Tipos, 0, frmRPR_Reactivo, " TIPO = 1 " & anulado
    Else
        llenar_combo cmbReactivos, New clsRPR_Tipos, 0, frmRPR_Reactivo, " TIPO = 2 " & anulado
    End If
End Sub
