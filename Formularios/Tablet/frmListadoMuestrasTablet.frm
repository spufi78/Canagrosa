VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListadoMuestrasTablet 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Listado de Muestras"
   ClientHeight    =   11145
   ClientLeft      =   45
   ClientTop       =   180
   ClientWidth     =   15270
   Icon            =   "frmListadoMuestrasTablet.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11145
   ScaleWidth      =   15270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   11655
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   9765
      Width           =   1770
   End
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   13455
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Muestra el informe de registrro de la muestra"
      Top             =   8460
      Width           =   1770
   End
   Begin VB.CommandButton cmdespecificas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Dat. Especificos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   11655
      Picture         =   "frmListadoMuestrasTablet.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   8460
      Width           =   1770
   End
   Begin VB.CommandButton cmdListadoDeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Determinaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   13455
      Picture         =   "frmListadoMuestrasTablet.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   5805
      Width           =   1770
   End
   Begin VB.CommandButton cmdMuestra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Muestra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   11655
      Picture         =   "frmListadoMuestrasTablet.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   4455
      Width           =   1770
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   3930
      Left            =   11655
      TabIndex        =   38
      Top             =   360
      Width           =   3525
      Begin VB.CommandButton cmdDel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2925
         Width           =   1080
      End
      Begin VB.CommandButton cmdNumero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   1
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   2025
         Width           =   1080
      End
      Begin VB.CommandButton cmdNumero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   2
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2025
         Width           =   1080
      End
      Begin VB.CommandButton cmdNumero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   3
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   2025
         Width           =   1080
      End
      Begin VB.CommandButton cmdNumero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   4
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1125
         Width           =   1080
      End
      Begin VB.CommandButton cmdNumero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   5
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1125
         Width           =   1080
      End
      Begin VB.CommandButton cmdNumero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   6
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1125
         Width           =   1080
      End
      Begin VB.CommandButton cmdNumero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   7
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   225
         Width           =   1080
      End
      Begin VB.CommandButton cmdNumero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   8
         Left            =   1215
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   225
         Width           =   1080
      End
      Begin VB.CommandButton cmdNumero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   9
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   225
         Width           =   1080
      End
      Begin VB.CommandButton cmdNumero 
         BackColor       =   &H00E0E0E0&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   14.25
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Index           =   0
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2925
         Width           =   2160
      End
   End
   Begin VB.CommandButton cmdInfRegistro 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Doc.Registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   11655
      Picture         =   "frmListadoMuestrasTablet.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Muestra el informe de registrro de la muestra"
      Top             =   7155
      Width           =   1770
   End
   Begin VB.CommandButton cmdDeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   13455
      Picture         =   "frmListadoMuestrasTablet.frx":2BF2
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   4455
      Width           =   1770
   End
   Begin VB.CommandButton cmdInforme 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   13455
      Picture         =   "frmListadoMuestrasTablet.frx":34BC
      Style           =   1  'Graphical
      TabIndex        =   24
      ToolTipText     =   "Previsualizar informe de ensayo"
      Top             =   7155
      Width           =   1770
   End
   Begin VB.CommandButton cmdVida 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vida "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Left            =   11655
      Picture         =   "frmListadoMuestrasTablet.frx":3D86
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   5805
      Width           =   1770
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   13455
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   9765
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Left            =   45
      TabIndex        =   26
      Top             =   360
      Width           =   11565
      Begin VB.TextBox txtcodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7110
         TabIndex        =   0
         Top             =   1440
         Width           =   2055
      End
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   6840
         TabIndex        =   5
         Top             =   270
         Width           =   1290
      End
      Begin VB.TextBox txtg2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4185
         TabIndex        =   4
         Top             =   810
         Width           =   1695
      End
      Begin VB.TextBox txtg1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1890
         TabIndex        =   3
         Top             =   810
         Width           =   1695
      End
      Begin VB.TextBox txtp2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4185
         TabIndex        =   2
         Top             =   270
         Width           =   1695
      End
      Begin VB.TextBox txtp1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1890
         TabIndex        =   1
         Top             =   270
         Width           =   1695
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1365
         Left            =   9900
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   495
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   465
         Left            =   1890
         TabIndex        =   6
         Top             =   1440
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   820
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
         Height          =   420
         Left            =   4185
         TabIndex        =   7
         Top             =   1485
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   741
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
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
      Begin MSComCtl2.UpDown cambiar 
         Height          =   510
         Left            =   8145
         TabIndex        =   36
         Top             =   270
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   900
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196624
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
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cod. Barras"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   0
         Left            =   7380
         TabIndex        =   39
         Top             =   1170
         Width           =   1440
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   6165
         TabIndex        =   37
         Top             =   360
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   6
         Left            =   3780
         TabIndex        =   34
         Top             =   945
         Width           =   165
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   135
         TabIndex        =   33
         Top             =   900
         Width           =   1335
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   3780
         TabIndex        =   32
         Top             =   405
         Width           =   165
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Particular"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   135
         TabIndex        =   31
         Top             =   360
         Width           =   1515
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   3780
         TabIndex        =   28
         Top             =   1530
         Width           =   165
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   135
         TabIndex        =   27
         Top             =   1530
         Width           =   1575
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   8235
      Left            =   45
      TabIndex        =   35
      Top             =   2835
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   14526
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
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Narrow"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   12555
      Top             =   9675
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestrasTablet.frx":4650
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestrasTablet.frx":4F2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestrasTablet.frx":5804
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestrasTablet.frx":60DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestrasTablet.frx":69B8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Muestras Tablet"
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
      TabIndex        =   30
      Top             =   0
      Width           =   15150
   End
   Begin VB.Label lblmsg 
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
      Left            =   45
      TabIndex        =   29
      Top             =   2520
      Width           =   11490
   End
End
Attribute VB_Name = "frmListadoMuestrasTablet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim indice As Integer


Private Sub cmdAdjuntos_Click()
'M0499-I
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_MUESTRAS
        .COBJETO = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        .Show 1
    End With
    Set frmAdjuntos = Nothing
'M0499-F
End Sub


Private Sub cmdetiqueta_Click()
    ReDim ETIQUETAS(1)
    ETIQUETAS(1) = lista.ListItems(lista.selectedItem.Index).SubItems(7)
    frmEtiquetas.Show 1
End Sub


Private Sub cmdespecificas_Click()
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra (lista.ListItems(lista.selectedItem.Index).SubItems(7))
    frmDatosEspecificos.PK_MUESTRA = CLng(lista.ListItems(lista.selectedItem.Index).SubItems(7))
    frmDatosEspecificos.PK_BANO = oMuestra.getBANO_ID
    frmDatosEspecificos.Show 1
End Sub


Private Sub cmdListadoDeter_Click()
    gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(7)
    frmVerDeterminaciones.Show 1
    gmuestra = 0
End Sub

Private Sub cmdDel_Click()
    Select Case indice
        Case 1
'            If Len(txtp1) > 1 Then
'                txtp1 = Left(txtp1, Len(txtp1) - 1)
'            Else
                txtp1 = ""
'            End If
            txtp1.SetFocus
        Case 2
'            If Len(txtp2) > 1 Then
'                txtp2 = Left(txtp2, Len(txtp2) - 1)
'            Else
                txtp2 = ""
'            End If
            txtp2.SetFocus
        Case 3
'            If Len(txtg1) > 1 Then
'                txtg1 = Left(txtg1, Len(txtg1) - 1)
'            Else
                txtg1 = ""
'            End If
            txtg1.SetFocus
        Case 4
'            If Len(txtg2) > 1 Then
'                txtg2 = Left(txtg2, Len(txtg2) - 1)
'            Else
                txtg2 = ""
'            End If
            txtg2.SetFocus
    End Select
    
End Sub

Private Sub cmdInfRegistro_Click()
    Dim oMuestra As New clsMuestra
    oMuestra.Informe_Recepcion lista.ListItems(lista.selectedItem.Index).SubItems(7), False
    Set oMuestra = Nothing
End Sub
Private Sub cmdDeter_Click()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        abrirRegistroMuestra gmuestra
'        Dim oMuestra As New clsMuestra
'        oMuestra.CargaMuestra (gmuestra)
'        Select Case oMuestra.getANALISIS_MODIFICADO
'            Case 2 ' Control de eficacia
'                With frmCE_Resultados
'                    .PK_ID_MUESTRA = gmuestra
'                    .Show 1
'                End With
'            Case 3 ' Sellante
'                frmSE_Resultados.Show 1
'            Case 5 ' Plasma
'                If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_ROCKWELL Or _
'                   oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_BRINELL Or _
'                   oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_VICKERS Then
'                    With frmPlasma_Dureza
'                        .PK = gmuestra
'                        .Show 1
'                    End With
'                Else
'                    With frmPlasma_Resultados
'                        .PK = gmuestra
'                        .Show 1
'                    End With
'                End If
'            Case Else
'                frmDeterminaciones.Show 1
'        End Select
        gmuestra = 0
    End If
End Sub

Private Sub cmdInforme_Click()
    If lista.ListItems.Count > 0 Then
'C001-I
'        Dim omuestra As New clsMuestra
'        Dim oTD As New clsTipos_documentos
'        If oTD.Nuevo_Formato(omuestra.obtener_tipo_documento(CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(7)))) Then
'            frmPrevisualizar2.PK_MUESTRA = CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(7))
'            frmPrevisualizar2.Show 1
'        Else
'            gmuestra = CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(7))
'            frmPrevisualizar.Show 1
            MostrarInforme CLng(lista.ListItems(lista.selectedItem.Index).SubItems(7))
'        End If
'C001-F
        actualizar_lista
    End If
End Sub

Private Sub cmdMuestra_Click()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        frmVerMuestra.Show 1
        actualizar_lista
    End If
End Sub

Private Sub cmdNumero_Click(Index As Integer)
    Select Case indice
        Case 1
            txtp1 = txtp1 & cmdNumero(Index).Caption
            txtp1.SetFocus
        Case 2
            txtp2 = txtp2 & cmdNumero(Index).Caption
            txtp2.SetFocus
        Case 3
            txtg1 = txtg1 & cmdNumero(Index).Caption
            txtg1.SetFocus
        Case 4
            txtg2 = txtg2 & cmdNumero(Index).Caption
            txtg2.SetFocus
    End Select
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVida_Click()
    If lista.ListItems.Count > 0 Then
        frmVidaMuestra.PK = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        frmVidaMuestra.Show 1
        gmuestra = 0
    End If
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Me.top = 0
    Me.Left = 0
    cargar_botones Me
    cabecera
    permisos
'    cargar_clientes
'    cargar_muestras
    fdesde = Date
    fhasta = Date
    txtanno = Year(Date)
    cambiar.Max = Year(Date)
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Código", 1500, lvwColumnLeft)
        .Tag = "Código"
    End With
    With lista.ColumnHeaders.Add(, , "Cliente", 1, lvwColumnLeft)
        .Tag = "Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Tipo de Analisis/Solución", 3350, lvwColumnLeft)
        .Tag = "Analisis"
    End With
    With lista.ColumnHeaders.Add(, , "Ref.Cliente", 3350, lvwColumnLeft)
        .Tag = "Ref.Cliente"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1500, lvwColumnCenter)
        .Tag = "Fecha"
    End With
        With lista.ColumnHeaders.Add(, , "Precio", 1, lvwColumnCenter)
        .Tag = "Precio"
    End With
    With lista.ColumnHeaders.Add(, , "General", 1300, lvwColumnCenter)
        .Tag = "Id"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "General"
    End With
    With lista.ColumnHeaders.Add(, , "Facturada", 1, lvwColumnCenter)
        .Tag = "Facturada"
    End With
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Sub buscar()
    Dim i As Integer
    Dim consulta As String
    Dim strMuestra As String
    Dim strClientes As String
    Dim strTipo As String
    Dim strgen As String
    Dim strpar As String
    Dim stranno As String
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    Dim f_desde As String
    Dim f_hasta As String
    f_desde = Format(fdesde, "yyyy-mm-dd")
    f_hasta = Format(fhasta, "yyyy-mm-dd")
    ' Tipo de muestra
    strMuestra = ""
    ' Clientes
    strClientes = ""
    ' Tipo
    strTipo = " AND (mu.anulada is Null or mu.anulada = 0)"
    ' Fechas
    Dim fecha_desde As String
    fecha_desde = " AND mu.fecha_recepcion>='" & f_desde & "'"
    Dim fecha_hasta As String
    fecha_hasta = " AND mu.fecha_recepcion<='" & f_hasta & "'"
    ' Particular
    strpar = ""
    If txtp1 <> "" Or txtp2 <> "" Then
        fecha_desde = ""
        fecha_hasta = ""
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
    ' General
    strgen = ""
    If txtg1 <> "" Or txtg2 <> "" Then
        fecha_desde = ""
        fecha_hasta = ""
        If IsNumeric(txtg1) = False Then
            MsgBox "El codigo debe ser numérico", vbInformation, App.Title
            txtg1.SetFocus
            Exit Sub
        End If
        If IsNumeric(txtg2) = False Then
            MsgBox "El codigo debe ser numérico", vbInformation, App.Title
            txtg2.SetFocus
            Exit Sub
        End If
        If txtg1 = "" Or txtg2 = "" Then
            MsgBox "Debe completar los codigos de búsqueda.", vbInformation, App.Title
            Exit Sub
        Else
            strgen = " AND mu.id_general between " & CLng(txtg1) & " and " & CLng(txtg2)  ' & " and anno = " & Year(Date)
        End If
    End If
    If strpar <> "" Or strgen <> "" Then
        stranno = " and mu.anno = " & CInt(txtanno)
    End If
    
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',cast(mu.id_particular as char)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "ta.nombre, " & _
               "mu.id_general, " & _
               "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.tipo_analisis_id=ta.id_tipo_analisis " & _
                      fecha_desde & fecha_hasta & _
                      strMuestra & _
                      strClientes & _
                      strTipo & _
                      strpar & strgen & stranno & _
                      " order by mu.id_general desc"
    Me.MousePointer = 11
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim oMuestra As New clsMuestra
'        Dim oAnalisis As New clsTipos_analisis
        lista.ListItems.Clear
        i = 1
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(1))
            .SubItems(1) = rs.Fields(2)
            .SubItems(2) = rs.Fields(8)
            .SubItems(3) = rs.Fields(4)
            If Not IsNull(rs.Fields(5)) Then
            .SubItems(4) = rs.Fields(5)
            End If
            If Not IsNull(rs.Fields(7)) Then
            .SubItems(5) = Format(rs.Fields(7), "currency")
            End If
            If Not IsNull(rs.Fields(9)) Then
            .SubItems(6) = Format(rs.Fields(9), "00000")
            End If
            If Not IsNull(rs.Fields(6)) Then
            .SubItems(7) = rs.Fields(6)
            End If
            .SubItems(8) = rs(10)
            End With
            lista.ListItems(i).Checked = True
            If rs.Fields(11) <> 0 Then
                lista.ListItems(i).SmallIcon = 1
                lista.ListItems(i).ToolTipText = "Enviado Correo"
            Else
                If rs(12) <> 0 Then
                    lista.ListItems(i).SmallIcon = 2
                    lista.ListItems(i).ToolTipText = "Anulada"
                Else
                    Select Case rs(13) ' Cerrada
                        Case 0 ' Abierta
                            lista.ListItems(i).SmallIcon = 5
                            lista.ListItems(i).ToolTipText = "Abierta"
                        Case 1 ' Cerrada
                            lista.ListItems(i).SmallIcon = 4
                            lista.ListItems(i).ToolTipText = "Cerrada"
                        Case 2 ' Pdte. Cierre
                            lista.ListItems(i).SmallIcon = 3
                            lista.ListItems(i).ToolTipText = "Pdte. Cierre"
                    End Select
                End If
            End If
            i = i + 1
            rs.MoveNext
        Wend
        lblMsg.Caption = "Muestras entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy")

    Else
        lblMsg.Caption = "No existe ninguna muestra con esos criterios."
    End If
    Set oAnalisis = Nothing
    Me.MousePointer = 0
    Set rs = Nothing
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
    cmdDeter_Click
End Sub

Private Sub txtcodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        buscar_codigo
    End If
End Sub

Private Sub txtg1_Change()
    txtp1 = ""
    txtp2 = ""
End Sub

Private Sub txtg1_GotFocus()
'    txtg1.SelStart = 0
'    txtg1.SelLength = Len(txtg1)
End Sub

Private Sub txtg1_LostFocus()
'    txtg2 = txtg1
    indice = 3
End Sub

Private Sub txtg2_Change()
    txtp1 = ""
    txtp2 = ""
End Sub

Private Sub txtg2_GotFocus()
'    txtg2.SelStart = 0
'    txtg2.SelLength = Len(txtg2)
End Sub

Private Sub txtg2_LostFocus()
    indice = 4
End Sub

Private Sub txtp1_Change()
    txtg1 = ""
    txtg2 = ""
End Sub

Private Sub txtp1_GotFocus()
'    txtp1.SelStart = 0
'    txtp1.SelLength = Len(txtp1)
End Sub

Private Sub txtp1_LostFocus()
'    txtp2 = txtp1
    indice = 1
End Sub

Private Sub txtp2_Change()
    txtg1 = ""
    txtg2 = ""
End Sub

Public Sub actualizar_lista()
    ' Por si se ha modificado la muestra
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.id_muestra = " & CLng(lista.ListItems(lista.selectedItem.Index).SubItems(7))
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim oMuestra As New clsMuestra
        lista.ListItems(lista.selectedItem.Index).Text = rs.Fields(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs.Fields(2)
        Dim oAnalisis As New clsTipos_analisis
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = oAnalisis.NombreAnalisis(rs.Fields(3))
        Set oAnalisis = Nothing
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs.Fields(4)
        If Not IsNull(rs.Fields(5)) Then
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs.Fields(5)
        End If
        If Not IsNull(rs.Fields(7)) Then
        lista.ListItems(lista.selectedItem.Index).SubItems(5) = Format(rs.Fields(7), "currency")
        End If
        lista.ListItems(lista.selectedItem.Index).SubItems(8) = rs(8)
        If rs(9) <> 0 Then ' Enviada por correo
            lista.ListItems(lista.selectedItem.Index).SmallIcon = 1
            lista.ListItems(lista.selectedItem.Index).ToolTipText = "Enviado Correo"
        Else
            If rs(10) <> 0 Then ' Anulada
                lista.ListItems(lista.selectedItem.Index).SmallIcon = 2
                lista.ListItems(lista.selectedItem.Index).ToolTipText = "Anulada"
            Else
                Select Case rs(11) ' Cerrada
                    Case 0 ' Abierta
                        lista.ListItems(lista.selectedItem.Index).SmallIcon = 5
                        lista.ListItems(lista.selectedItem.Index).ToolTipText = "Abierta"
                    Case 1 ' Cerrada
                        lista.ListItems(lista.selectedItem.Index).SmallIcon = 4
                        lista.ListItems(lista.selectedItem.Index).ToolTipText = "Cerrada"
                    Case 2 ' Pdte. Cierre
                        lista.ListItems(lista.selectedItem.Index).SmallIcon = 3
                        lista.ListItems(lista.selectedItem.Index).ToolTipText = "Pdte. Cierre"
                End Select
            End If
        End If

    End If
    Set rs = Nothing
End Sub

Public Sub permisos()
'    cmdVida.Visible = False
End Sub

Private Sub txtp2_GotFocus()
'    txtp2.SelStart = 0
'    txtp2.SelLength = Len(txtp2)
End Sub

Private Sub txtp2_LostFocus()
    indice = 2
End Sub
Private Sub buscar_codigo()
    If txtCodigo <> "" Then
        Select Case UCase(Left(txtCodigo, 1))
        Case "M"
            insertar_en_la_lista CLng(Mid(txtCodigo, 2, Len(txtCodigo) - 1))
        Case Else
            MsgBox "No localizo el código de barras.", vbCritical, App.Title
            txtCodigo = ""
            txtCodigo.SetFocus
        End Select
    End If
    txtCodigo = ""
End Sub

Public Sub insertar_en_la_lista(MUESTRA As Long)
    ' Por si se ha modificado la muestra
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    lista.ListItems.Clear
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada,mu.id_general " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.id_muestra = " & MUESTRA
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim oMuestra As New clsMuestra
        With lista.ListItems.Add(, , rs(1))
            .SubItems(1) = rs.Fields(2)
            Dim oAnalisis As New clsTipos_analisis
            .SubItems(2) = oAnalisis.NombreAnalisis(rs.Fields(3))
            Set oAnalisis = Nothing
            .SubItems(3) = rs.Fields(4)
            If Not IsNull(rs.Fields(5)) Then
                .SubItems(4) = rs.Fields(5)
            End If
            If Not IsNull(rs.Fields(7)) Then
                .SubItems(5) = Format(rs.Fields(7), "currency")
            End If
            .SubItems(6) = rs(12) ' ID_GENERAL
            .SubItems(7) = rs(6) ' ID
            .SubItems(8) = rs(8)
            If rs(9) <> 0 Then ' Enviada por correo
                lista.ListItems(1).SmallIcon = 1
                lista.ListItems(1).ToolTipText = "Enviado Correo"
            Else
                If rs(10) <> 0 Then ' Anulada
                    lista.ListItems(1).SmallIcon = 2
                    lista.ListItems(1).ToolTipText = "Anulada"
                Else
                    Select Case rs(11) ' Cerrada
                        Case 0 ' Abierta
                            lista.ListItems(1).SmallIcon = 5
                            lista.ListItems(1).ToolTipText = "Abierta"
                        Case 1 ' Cerrada
                            lista.ListItems(1).SmallIcon = 4
                            lista.ListItems(1).ToolTipText = "Cerrada"
                        Case 2 ' Pdte. Cierre
                            lista.ListItems(1).SmallIcon = 3
                            lista.ListItems(1).ToolTipText = "Pdte. Cierre"
                    End Select
                End If
            End If
        End With
    End If
    Set rs = Nothing
End Sub


