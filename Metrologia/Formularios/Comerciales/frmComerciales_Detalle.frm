VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmComerciales_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Comerciales"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmComerciales_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Comisiones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Index           =   0
      Left            =   60
      TabIndex        =   36
      Top             =   6120
      Width           =   6435
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Index           =   0
         Left            =   4770
         TabIndex        =   15
         Top             =   270
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comisión = Precio - portes * porcentaje asignado al comercial"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   13
         Left            =   840
         TabIndex        =   39
         Top             =   780
         Width           =   4305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Porcentaje de Comisión de Facturas Cobradas"
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
         Index           =   10
         Left            =   210
         TabIndex        =   37
         Top             =   330
         Width           =   4200
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   7890
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6240
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   915
      Left            =   6600
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6240
      Width           =   1275
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   7110
      TabIndex        =   32
      Top             =   3840
      Visible         =   0   'False
      Width           =   2085
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   360
         Left            =   90
         TabIndex        =   33
         Top             =   225
         Width           =   1890
         _ExtentX        =   3334
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
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   5760
      Top             =   6120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   2625
      Left            =   7110
      TabIndex        =   29
      Top             =   450
      Width           =   2100
      Begin VB.TextBox txtdatos 
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
         Index           =   11
         Left            =   90
         MaxLength       =   30
         TabIndex        =   13
         Top             =   2160
         Width           =   1920
      End
      Begin VB.Image img 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   1710
         Left            =   315
         Stretch         =   -1  'True
         Top             =   180
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apodo"
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
         Index           =   6
         Left            =   90
         TabIndex        =   31
         Top             =   1935
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1305
      Index           =   13
      Left            =   60
      TabIndex        =   27
      Top             =   4770
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   990
         Index           =   9
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   240
         Width           =   8985
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos del Operario "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   90
      TabIndex        =   18
      Top             =   420
      Width           =   6975
      Begin VB.TextBox txtdatos 
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
         Index           =   14
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   7
         Top             =   2400
         Width           =   1725
      End
      Begin VB.TextBox txtdatos 
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
         Index           =   13
         Left            =   1350
         MaxLength       =   25
         TabIndex        =   10
         Top             =   3240
         Width           =   5505
      End
      Begin VB.TextBox txtdatos 
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
         Index           =   12
         Left            =   4710
         MaxLength       =   6
         TabIndex        =   9
         Top             =   2820
         Width           =   2145
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   345
         Index           =   10
         Left            =   1350
         TabIndex        =   11
         Top             =   3690
         Width           =   4305
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   390
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3675
         Width           =   1095
      End
      Begin VB.TextBox txtdatos 
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
         Index           =   7
         Left            =   4710
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1980
         Width           =   2175
      End
      Begin VB.TextBox txtdatos 
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
         Index           =   6
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   5
         Top             =   1980
         Width           =   1725
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   345
         Index           =   8
         Left            =   1350
         MaxLength       =   15
         TabIndex        =   8
         Top             =   2820
         Width           =   1725
      End
      Begin VB.TextBox txtdatos 
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
         Index           =   3
         Left            =   1350
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1140
         Width           =   960
      End
      Begin VB.TextBox txtdatos 
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
         Index           =   2
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   1
         Top             =   735
         Width           =   5535
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   330
         Index           =   1
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   0
         Top             =   330
         Width           =   5535
      End
      Begin MSDataListLib.DataCombo cmbProvincia 
         Height          =   360
         Left            =   3390
         TabIndex        =   3
         Top             =   1110
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbMunicipio 
         Height          =   315
         Left            =   1350
         TabIndex        =   4
         Top             =   1575
         Width           =   5490
         _ExtentX        =   9684
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fax"
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
         Index           =   12
         Left            =   210
         TabIndex        =   38
         Top             =   2460
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.C.C."
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
         Index           =   9
         Left            =   225
         TabIndex        =   35
         Top             =   3300
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo Interno"
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
         Index           =   8
         Left            =   3330
         TabIndex        =   34
         Top             =   2880
         Width           =   1305
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imagen"
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
         Left            =   210
         TabIndex        =   30
         Top             =   3735
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
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
         Index           =   7
         Left            =   210
         TabIndex        =   26
         Top             =   1620
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Movil"
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
         Index           =   15
         Left            =   4110
         TabIndex        =   25
         Top             =   2040
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
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
         Index           =   11
         Left            =   210
         TabIndex        =   24
         Top             =   2040
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.I.F."
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
         Index           =   5
         Left            =   225
         TabIndex        =   23
         Top             =   2910
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
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
         Left            =   2430
         TabIndex        =   22
         Top             =   1170
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
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
         Left            =   210
         TabIndex        =   21
         Top             =   1170
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
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
         Left            =   225
         TabIndex        =   20
         Top             =   780
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
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
         Left            =   225
         TabIndex        =   19
         Top             =   375
         Width           =   735
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmComerciales_Detalle.frx":030A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Nuevo Comercial"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   60
      TabIndex        =   28
      Top             =   60
      Width           =   9180
   End
End
Attribute VB_Name = "frmComerciales_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbProvincia_Change()
    If cmbProvincia.Text <> "" Then
        cargar_municipios (cmbProvincia.BoundText)
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEXplorar_Click()
    cd.DialogTitle = "Abrir fichero de imagen"
    cd.InitDir = App.Path & "\recursos\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        txtdatos(10).Text = cd.FileName  ' cd.FileTitle
    End If
End Sub

'Private Sub cmdModificar_Click()
'    If lista.ListItems.Count = 0 Then
'        Exit Sub
'    End If
'    If txtdatos(0) = "" Then
'        MsgBox "Introduzca una comision.", vbExclamation, App.Title
'        Exit Sub
'    End If
'    If IsNumeric(txtdatos(0)) = False Then
'        MsgBox "Introduzca una comision numérica.", vbExclamation, App.Title
'        Exit Sub
'    End If
'    lista.ListItems(lista.SelectedItem.Index).SubItems(1) = txtdatos(0)
'End Sub

Private Sub cmdok_Click()
    If gComercial > 0 Then
        modificar_comercial
    Else
        insertar_comercial
    End If
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    Cargar_Combo cmbProvincia, New clsProvincias
    cargar_estados
    permisos
    If gComercial > 0 Then
        consulta_comercial
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmComerciales_Detalle = Nothing
End Sub

'Private Sub lista_Click()
'    If lista.ListItems.Count > 0 Then
'        txtdatos(0) = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
'        txtdatos(0).SetFocus
'    End If
'End Sub

Private Sub txtDatos_Change(Index As Integer)
    On Error Resume Next
    If Index = 10 And txtdatos(10) <> "" Then
        If Dir(txtdatos(10)) <> "" Then
            Set img.Picture = LoadPicture(txtdatos(10))
        End If
    End If
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub

Private Sub txtdatos_Keyup(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40 ' Abajo
       If Index = 9 Then
        txtdatos(1).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 1 Then
        txtdatos(9).SetFocus
       Else
        SendKeys "+{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 27
        cmdcancel_Click
     Case 121 ' F10
        cmdok_Click
    End Select
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 9 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
    If Index = 0 And KeyAscii = 46 Then
         KeyAscii = 44
    End If
    
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = &HFFFFFF
    If Index = 0 Then ' comision
        If Not IsNumeric(txtdatos(Index)) Then
            MsgBox "La comision debe ser numérica.", vbExclamation, App.Title
            txtdatos(Index) = ""
            txtdatos(Index).SetFocus
        End If
    End If
End Sub

Public Sub borrar_campos()
    Dim i As Integer
    For i = 1 To 13
        ' LP005
        If i <> 4 And i <> 5 Then
            txtdatos(i) = ""
        End If
    Next
    txtdatos(1).SetFocus
End Sub

Public Sub bloquear_campos()
    Dim i As Integer
    For i = 1 To 13
        ' LP005
        If i <> 4 And i <> 5 Then
            txtdatos(i).Locked = True
        End If
    Next
End Sub

Public Sub insertar_comercial()
    If valida_datos = False Then
        Exit Sub
    End If
    pregunta = "Va a dar de alta el comercial. ¿Esta seguro?"
    If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim aux As Long
        Set ocomercial = mover_datos
        aux = gComercial
        gComercial = ocomercial.Insertar
        If aux = -1 Then
            Unload Me
            Exit Sub
        End If
        If gComercial <> 0 Then
            borrar_campos
        End If
        Set ocomercial = Nothing
        img.Picture = LoadPicture()
    End If
End Sub

Public Sub modificar_comercial()
    If valida_datos() = False Then
        Exit Sub
    End If
    Dim pos As Integer
    Dim COMERCIAL As Integer
    pregunta = "Va a modificar los datos del comercial. ¿Esta seguro?"
    If MsgBox(pregunta, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Set ocomercial = mover_datos
        If ocomercial.Modificar(gComercial) = True Then
            Unload Me
        End If
        Set ocomercial = Nothing
    End If

End Sub
Public Function valida_datos() As Boolean
    valida_datos = True
    If txtdatos(1) = "" Then
        MsgBox "El nombre del comercial no puede estar en blanco.", vbCritical, "Error"
        txtdatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtdatos(11) = "" Then
        MsgBox "El apodo del comercial no puede estar en blanco.", vbCritical, "Error"
        txtdatos(11).SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbestados.BoundText = "" Then
        MsgBox "Debe seleccionar un estado.", vbCritical, "Error"
        cmbestados.SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtdatos(0) = "" Then
        MsgBox "La comisión no puede estar en blanco.", vbCritical, "Error"
        txtdatos(0).SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Public Sub consulta_comercial()
    On Error GoTo fallo
    Dim ocomercial As New clsComercial
    lbltitulo.Caption = "Modificacion de comercial"
    ocomercial.Cargar (gComercial)
    With ocomercial
        txtdatos(1) = .getNOMBRE
        txtdatos(2) = .getDIRECCION
        txtdatos(3) = .getCP
        cmbProvincia.BoundText = .getPROVINCIA_ID
        cargar_municipios .getPROVINCIA_ID
        cmbMunicipio.BoundText = .getMUNICIPIO_ID
        txtdatos(6) = .getTELEFONO
        txtdatos(7) = .getMOVIL
        txtdatos(8) = .getCIF
        txtdatos(9) = .getOBSERVACIONES
        txtdatos(10) = .getfoto
        txtdatos(11) = .getAPODO
        txtdatos(12) = .getCODIGO_INTERNO
        txtdatos(13) = .getCCC
        txtdatos(14) = .getFAX
        cmbestados.BoundText = .getESTADO_ID
        txtdatos(0) = Replace(.getCOMISION, ".", ",")
    End With
    Set ocomercial = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del comercial.", vbCritical, Err.Description
End Sub
Public Function mover_datos() As clsComercial
    On Error GoTo fallo
    Dim ocomercial As New clsComercial
    With ocomercial
        .setNOMBRE = txtdatos(1)
        .setDIRECCION = txtdatos(2)
        If Trim(txtdatos(3)) <> "" Then
            .setCP = CLng(txtdatos(3).Text)
        Else
            .setCP = 0
        End If
        ' LP005
'        .setPROVINCIA = txtdatos(4)
        If cmbProvincia.Text = "" Then
            .setPROVINCIA_ID = 0
        Else
            .setPROVINCIA_ID = cmbProvincia.BoundText
        End If
        If cmbMunicipio.Text = "" Then
            .setMUNICIPIO_ID = 0
        Else
            .setMUNICIPIO_ID = cmbMunicipio.BoundText
        End If
        .setCIF = txtdatos(8)
        .setTELEFONO = txtdatos(6)
        .setMOVIL = txtdatos(7)
        .setOBSERVACIONES = txtdatos(9)
        .setfoto = txtdatos(10)
        .setAPODO = txtdatos(11)
        .setESTADO_ID = cmbestados.BoundText
        .setCODIGO_INTERNO = txtdatos(12)
        .setCCC = txtdatos(13)
        .setFAX = txtdatos(14)
        .setCOMISION = moneda_bd(txtdatos(0))
    End With
    Set mover_datos = ocomercial
    Set ocomercial = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del comercial.", vbCritical, Err.Description
End Function
Public Sub cargar_estados()
'    Dim ooe As New clsOperarios_Estados
'    Set cmbestados.RowSource = ooe.Listado
'    cmbestados.ListField = "nombre"
'    cmbestados.BoundColumn = "id_estado"
'    Set ooe = Nothing
End Sub

Public Sub permisos()
    If usuario.getPER_7 = 0 Then
        cmdControl.Enabled = False
        cmdExpediente.Enabled = False
        cmdAnticipo.Enabled = False
    End If
End Sub
Public Sub cargar_municipios(PROVINCIA As Long)
    cmbMunicipio.Text = ""
    cargar_combo_FK cmbMunicipio, New clsMunicipios, PROVINCIA
End Sub
