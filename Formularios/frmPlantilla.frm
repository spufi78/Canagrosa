VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmPlantilla 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Plantillas"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13740
   Icon            =   "frmPlantilla.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   13740
   Begin VB.TextBox text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   285
      Index           =   7
      Left            =   8850
      TabIndex        =   50
      ToolTipText     =   "NO USE LA COMA COMO SEPARADOR DE DECIMALES ,USE EL PUNTO -- EJEMPLO: 6020.85 --"
      Top             =   6930
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   5790
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   6705
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   4665
      Style           =   1  'Graphical
      TabIndex        =   48
      Top             =   6705
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   3555
      Style           =   1  'Graphical
      TabIndex        =   47
      Top             =   6705
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   46
      Top             =   6705
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11475
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   6705
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdRecepcionar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recepcionar con esta Plantilla"
      Height          =   855
      Left            =   45
      Picture         =   "frmPlantilla.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   44
      Top             =   6705
      Width           =   3435
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   645
      Left            =   3495
      TabIndex        =   40
      Top             =   510
      Width           =   10170
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         Height          =   315
         Index           =   10
         Left            =   5445
         TabIndex        =   1
         Top             =   225
         Width           =   4650
      End
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
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
         Height          =   315
         Index           =   9
         Left            =   990
         TabIndex        =   0
         Top             =   210
         Width           =   3015
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   9
         Left            =   4275
         TabIndex        =   42
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   41
         Top             =   270
         Width           =   795
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2145
      Left            =   3495
      TabIndex        =   30
      Top             =   1440
      Width           =   5475
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2100
         TabIndex        =   7
         Top             =   1620
         Width           =   3015
      End
      Begin VB.TextBox text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   4260
         TabIndex        =   2
         Text            =   "1"
         Top             =   180
         Width           =   735
      End
      Begin VB.CommandButton cmdMas 
         Caption         =   "+"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5070
         TabIndex        =   3
         Top             =   180
         Width           =   255
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   2
         Left            =   2100
         TabIndex        =   6
         Top             =   1230
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Bindings        =   "frmPlantilla.frx":1194
         Height          =   315
         Index           =   0
         Left            =   1260
         TabIndex        =   4
         Top             =   510
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   1
         Left            =   2100
         TabIndex        =   5
         Top             =   870
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de analisis"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   35
         Top             =   1260
         Width           =   1545
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de muestra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   930
         Width           =   1605
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº de muestras"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2670
         TabIndex        =   33
         Top             =   210
         Width           =   1545
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   32
         Top             =   600
         Width           =   705
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Referencia muestra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   31
         Top             =   1650
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   3555
      Left            =   8985
      TabIndex        =   24
      Top             =   1440
      Width           =   4680
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1575
         Index           =   6
         Left            =   180
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   1890
         Width           =   4395
      End
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   5
         Left            =   1230
         MaxLength       =   100
         TabIndex        =   14
         Top             =   1260
         Width           =   3345
      End
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   2115
         TabIndex        =   11
         Top             =   210
         Width           =   2460
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Bindings        =   "frmPlantilla.frx":11C4
         Height          =   315
         Index           =   4
         Left            =   1170
         TabIndex        =   12
         Top             =   540
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   5
         Left            =   1845
         TabIndex        =   13
         Top             =   900
         Width           =   2730
         _ExtentX        =   4815
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   180
         TabIndex        =   29
         Top             =   1620
         Width           =   1455
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalles"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   180
         TabIndex        =   28
         Top             =   1320
         Width           =   825
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entregada por"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   27
         Top             =   990
         Width           =   1425
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Envase"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   26
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepcionada por"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   25
         Top             =   270
         Width           =   1755
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   2700
      Left            =   3510
      TabIndex        =   20
      Top             =   3915
      Width           =   5460
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   1710
         MaxLength       =   100
         TabIndex        =   8
         Top             =   210
         Width           =   3615
      End
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1500
         Index           =   3
         Left            =   150
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   1050
         Width           =   5145
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   6
         Left            =   1710
         TabIndex        =   9
         Top             =   510
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   150
         TabIndex        =   23
         Top             =   870
         Width           =   1455
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalles"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   150
         TabIndex        =   22
         Top             =   270
         Width           =   825
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realizado por"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   150
         TabIndex        =   21
         Top             =   540
         Width           =   1365
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1290
      Left            =   9030
      TabIndex        =   18
      Top             =   5310
      Width           =   4695
      Begin VB.TextBox text1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Index           =   8
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   16
         Top             =   390
         Width           =   4485
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incidencias"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   135
         TabIndex        =   19
         Top             =   180
         Width           =   1155
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6150
      Left            =   45
      TabIndex        =   17
      Top             =   450
      Width           =   3405
      _ExtentX        =   6006
      _ExtentY        =   10848
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
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Precio"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   19
      Left            =   8100
      TabIndex        =   51
      Top             =   6990
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Plantillas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   3
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   13695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Muestreo"
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
      Height          =   270
      Index           =   0
      Left            =   3510
      TabIndex        =   39
      Top             =   3600
      Width           =   5475
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Recepción"
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
      Height          =   270
      Index           =   1
      Left            =   8985
      TabIndex        =   38
      Top             =   1170
      Width           =   4680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Otros datos"
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
      Height          =   270
      Index           =   2
      Left            =   9030
      TabIndex        =   37
      Top             =   5040
      Width           =   4680
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Datos Generales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   270
      Index           =   4
      Left            =   3495
      TabIndex        =   36
      Top             =   1170
      Width           =   5475
   End
End
Attribute VB_Name = "frmPlantilla"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim estado As String * 1
Private Sub cmdAnadir_Click()
    cmdok.Visible = True
    cmdAnadir.Enabled = True
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdRecepcionar.Enabled = False
    lista.Enabled = False
    desbloquear
    borrar_campos
    ' Usuario
    Text1(4) = USUARIO.getUSUARIO
    Label1(3) = "Nueva plantilla"
    Label1(3).BackColor = &H80C0FF
    estado = "N"
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    On Error Resume Next
    Dim consulta As String
    If lista.ListItems.Count > 0 Then
        consulta = "delete from plantilla_determinaciones where plantilla_id = " & lista.ListItems(lista.SelectedItem.Index)
        execute_bd consulta
        consulta = "delete from plantilla_banos where plantilla_id = " & lista.ListItems(lista.SelectedItem.Index)
        execute_bd consulta
        consulta = "delete from plantillas_muestras where id_plantilla = " & lista.ListItems(lista.SelectedItem.Index)
        execute_bd consulta
        MsgBox "Plantilla eliminada correctamente.", vbInformation, App.Title
        inicializar_ventana
    End If
End Sub

Private Sub cmdMas_Click()
    Text1(0) = CInt(Text1(0)) + 1
End Sub

Private Sub cmdModificar_Click()
    cmdok.Visible = True
    cmdAnadir.Enabled = False
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    cmdRecepcionar.Enabled = False
    lista.Enabled = False
    desbloquear
    Label1(3) = "Modificación de plantilla"
    Label1(3).BackColor = &H80C0FF
    estado = "M"
End Sub

Private Sub cmdok_Click()
    If validar_datos = False Then
        Exit Sub
    End If
    insertar_plantilla
End Sub

Private Sub cmdRecepcionar_Click()
    Dim tipo As New clsMuestra
    nueva_plantilla = 0
    PLANTILLA = lista.ListItems(lista.SelectedItem.Index)
    If cmbDatos(1).Text <> "" Then
        If tipo.esBano(cmbDatos(1).BoundText) Then  'es un baño Id_Espacial
            frmDetallePlantillaBano.Show 1
            Exit Sub
        End If
    End If
    Dim oform As New frmRecepcion
    oform.Show
    Set oform = Nothing
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Top = 50
    Me.Left = 50
    cabecera
    cargar_clientes
    cargar_muestras
    cargar_realizadas
    cargar_entregada
    cargar_envases
    inicializar_ventana
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "ID", 400, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Plantilla", 2750, lvwColumnLeft)
        .Tag = "Plantilla"
    End With
End Sub
Public Sub cargar_clientes()
    Dim ocliente As New clsCliente
    Set cmbDatos(0).RowSource = ocliente.Listado("", "", "") 'recorset devuelto por la funcion
    cmbDatos(0).ListField = "nombre" 'campo que veo
    cmbDatos(0).DataField = "id_cliente" 'campo asociado
    cmbDatos(0).BoundColumn = "id_cliente" 'lo que realmente envia
    Set ocliente = Nothing
End Sub

Public Sub cargar_muestras()
    Dim omuestra As New clsTipos_muestra
    Set cmbDatos(1).RowSource = omuestra.Listado
    cmbDatos(1).ListField = "nombre" 'lo que enseña
    cmbDatos(1).DataField = "id_tipo_muestra" 'campo asociado
    cmbDatos(1).BoundColumn = "id_tipo_muestra" 'lo que realmente envia
    Set omuestra = Nothing
End Sub

Public Sub cargar_realizadas()
    Dim oenti As New clsEntidades_muestreo
    Set cmbDatos(6).RowSource = oenti.Listado
    cmbDatos(6).ListField = "descripcion"
    cmbDatos(6).DataField = "id_entidad_muestreo" 'campo asociado
    cmbDatos(6).BoundColumn = "id_entidad_muestreo" 'lo que realmente
    Set oenti = Nothing
End Sub

Public Sub cargar_entregada()
    Dim oenti As New clsEntidades_Entrega
    Set cmbDatos(5).RowSource = oenti.Listado
    cmbDatos(5).ListField = "descripcion"
    cmbDatos(5).DataField = "id_entidad_entrega" 'campo asociado
    cmbDatos(5).BoundColumn = "id_entidad_entrega" 'lo que realmente
    Set oenti = Nothing
End Sub

Public Sub cargar_envases()
    Dim oFormato As New clsformatos
    Set cmbDatos(4).RowSource = oFormato.Listado
    cmbDatos(4).ListField = "descripcion"
    cmbDatos(4).DataField = "id_formato" 'campo asociado
    cmbDatos(4).BoundColumn = "id_formato" 'lo que realmente
    Set oFormato = Nothing
End Sub

Public Sub cargar_plantillas()
    Dim oplantilla As New clsPlantillas_muestras
    Dim rs As New ADODB.RecordSet
    lista.ListItems.Clear
    Set rs = oplantilla.Listado
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("id_plantilla"))
            .SubItems(1) = rs("nombre")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oplantilla = Nothing
End Sub

Private Sub lista_Click()
    If lista.ListItems(lista.SelectedItem.Index) <> "" Then
        Dim oplantilla As New clsPlantillas_muestras
        With oplantilla
         .CargaPlantilla (lista.ListItems(lista.SelectedItem.Index))
         Text1(9) = .getNOMBRE
         Text1(10) = .getDESCRIPCION
         Text1(0) = .getCANTIDAD_MUESTRAS
         If .getCLIENTE_ID > 0 Then
            Dim ocliente As New clsCliente
            ocliente.CargaCliente (.getCLIENTE_ID)
            cmbDatos(0).Text = ocliente.getNOMBRE
            Set ocliente = Nothing
         End If
         If .getTIPO_MUESTRA_ID > 0 Then
            Dim otm As New clsMuestra
            cmbDatos(1).Text = otm.NombreMuestra(.getTIPO_MUESTRA_ID)
            Set otm = Nothing
         End If
         If .getTIPO_ANALISIS_ID > 0 Then
            Dim oTA As New clsTipos_analisis
            cmbDatos(2).Text = oTA.NombreAnalisis(.getTIPO_ANALISIS_ID)
            Set oTA = Nothing
         End If
         Text1(1) = .getREFERENCIA_CLIENTE
         Text1(2) = .getDETALLE_MUESTREO
         If .getENTIDAD_MUESTREO_ID > 0 Then
            Dim oem As New clsEntidades_muestreo
            oem.CargarEntidad (.getENTIDAD_MUESTREO_ID)
            cmbDatos(6).Text = oem.getDESCRIPCION
            Set oem = Nothing
         End If
         Text1(3) = .getOBSERVACIONES_MUESTREO
         If .getFORMATO_ID > 0 Then
            Dim of As New clsformatos
            of.CargarFormato (.getFORMATO_ID)
            cmbDatos(4).Text = of.getDESCRIPCION
            Set of = Nothing
         End If
         If .getENTIDAD_ENTREGA_ID > 0 Then
            Dim oee As New clsEntidades_Entrega
            oee.CargarEntidad (.getENTIDAD_ENTREGA_ID)
            cmbDatos(6).Text = oee.getDESCRIPCION
            Set oee = Nothing
         End If
         ' Usuario
         Dim ousu As New clsUsuarios
         ousu.CARGAR (.getEMPLEADO_ID)
         Text1(4) = ousu.getUSUARIO
         Set ousu = Nothing
         
         Text1(5) = .getDETALLE_ENTREGA
         Text1(6) = .getOBSERVACIONES_ENTREGA
         Text1(7) = Format(.getPRECIO, "currency")
         Text1(8) = .getOBSERVACIONES
        End With
        cmdModificar.Enabled = True
        cmdEliminar.Enabled = True
    End If
End Sub

Public Sub borrar_campos()
    Dim i As Integer
    Text1(0) = "1"
    For i = 0 To 6
        If i <> 3 Then
            cmbDatos(i).Text = ""
        End If
    Next
    For i = 1 To 10
        Text1(i) = ""
    Next
    Text1(9).SetFocus
End Sub
Public Sub bloquear()
    Dim i As Integer
    For i = 0 To 6
        If i <> 3 Then
            cmbDatos(i).Enabled = False
        End If
    Next
    For i = 0 To 10
        If i <> 4 Then
            Text1(i).Enabled = False
        End If
    Next
    cmdMas.Enabled = False
    lista.Enabled = True
    cmdRecepcionar.Enabled = True
    cmdok.Visible = False
    cmdAnadir.Enabled = True
End Sub
Public Sub desbloquear()
    Dim i As Integer
    For i = 0 To 6
        If i <> 3 Then
            cmbDatos(i).Enabled = True
        End If
    Next
    For i = 0 To 10
        If i <> 4 Then
            Text1(i).Enabled = True
        End If
    Next
    cmdMas.Enabled = True
End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).BackColor = &H80C0FF
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
'    Dim caracter As String
'    If Index <> 7 Then
'        caracter = Chr(KeyAscii) 'devuelve el numero del caracter
'        caracter = UCase(caracter) 'conviertemayusculas todo el texto
'        KeyAscii = Asc(caracter) 'da el valor de la tecla pulsadakeyascii
'    End If
    ' Escribir ',' al pulsar '.'
    If Index = 7 And KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = &HFFFFFF
    If Index = 7 And Text1(Index) <> "" And Not IsNumeric(Text1(Index)) Then
        MsgBox "El precio debe ser numérico.", vbCritical, "Formato"
        Text1(7).SetFocus
    Else
        Text1(7) = Format(Text1(7), "currency")
    End If
End Sub

Private Sub cmbDatos_Change(Index As Integer)
'  Dim id_muestra As Integer
  On Error GoTo fallo:
  Select Case Index
  Case 1: 'muestra los analisis para el tipo de muestra seleccionado
       Dim tipo As New clsMuestra
       If cmbDatos(1).Text <> "" Then
        If Not tipo.esBano(cmbDatos(1).BoundText) Then  'es un baño Id_Espacial
        Dim oAnalisis As New clsTipos_analisis
         ' Es una determinacion de muestra
         If oAnalisis.AnalisisAsociadosMuestra(cmbDatos(1).BoundText).RecordCount <> 0 Then 'existe un registro al menos
            Set cmbDatos(2).RowSource = oAnalisis.AnalisisAsociadosMuestra(cmbDatos(1).BoundText) 'comboanalisis
            cmbDatos(2).ListField = "nombre" 'lo que enseña
            cmbDatos(2).DataField = "id_tipo_analisis" 'campo asociado
            cmbDatos(2).BoundColumn = "id_tipo_analisis" 'lo que realmente
            cmbDatos(2).Text = oAnalisis.AnalisisAsociadosMuestra(cmbDatos(1).BoundText).Fields("nombre").value
            cmbDatos(2).Enabled = True
         Else 'si no recupero ningun registro
            cmbDatos(2).Text = ""
            cmbDatos(2).Enabled = False
         End If
        Else
         ' Es un baño especial
            If cmbDatos(0).Text = "" Then
                MsgBox "Seleccione primero un cliente.", vbInformation, App.Title
                cmbDatos(1).Text = ""
                cmbDatos(0).SetFocus
                Exit Sub
            End If
            cmbDatos(2).Text = ""
            cmbDatos(2).Enabled = False
            Dim oBANO As New clsBanos
            Dim rsbano As New ADODB.RecordSet
            Set rsbano = oBANO.banos_cliente(cmbDatos(0).BoundText, cmbDatos(1).BoundText)
            If rsbano.RecordCount = 0 Then
                MsgBox "No hay baños para el cliente y tipo de muestra seleccionado.", vbInformation, App.Title
                cmbDatos(1).Text = ""
                Exit Sub
            ElseIf rsbano.RecordCount < val(Text1(0)) Then
                MsgBox "El número de baños para el cliente y tipo de Muestra " & _
                       "seleccionados es menor que el número de muestras indicado.", vbInformation, App.Title
                cmbDatos(1).Text = ""
                Exit Sub
            End If
            Set oBANO = Nothing
        End If 'fin del esBanno
        calcular_precio_analisis
        Set tipo = Nothing
        Set oAnalisis = Nothing
       End If
  Case 2: 'tipo de analisis
    calcular_precio_analisis
  End Select
  cmbDatos(2).ToolTipText = cmbDatos(2).Text
  Exit Sub
fallo:
    MsgBox "Error al decodificar los campos", vbCritical, Err.Description
End Sub
Public Sub calcular_precio_analisis()
    Dim aux As New clsTipos_analisis
    If cmbDatos(2).BoundText = "" Then
        Text1(7).Text = ""
    Else
        Text1(7) = Format(aux.PrecioDelAnalisis(cmbDatos(2).BoundText), "currency")
    End If
    Set aux = Nothing
End Sub

Public Sub insertar_plantilla()
    Dim oplantilla As New clsPlantillas_muestras
    With oplantilla
        If cmbDatos(1).Text <> "" Then
            .setTIPO_MUESTRA_ID = cmbDatos(1).BoundText
        End If
        If cmbDatos(2).Text <> "" Then
            .setTIPO_ANALISIS_ID = cmbDatos(2).BoundText
        End If
        .setANALISIS_MODIFICADO = 0
        If cmbDatos(6).Text <> "" Then
            .setENTIDAD_MUESTREO_ID = cmbDatos(6).BoundText
        End If
        .setDETALLE_MUESTREO = Text1(2)
        .setOBSERVACIONES_MUESTREO = Text1(3)
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        If cmbDatos(4).Text <> "" Then
            .setFORMATO_ID = cmbDatos(4).BoundText
        End If
        If cmbDatos(5).Text <> "" Then
            .setENTIDAD_ENTREGA_ID = cmbDatos(5).BoundText
        End If
        .setDETALLE_ENTREGA = Text1(5)
        .setOBSERVACIONES_ENTREGA = Text1(6)
        If cmbDatos(0).Text <> "" Then
            .setCLIENTE_ID = cmbDatos(0).BoundText
        End If
        .setREFERENCIA_CLIENTE = Text1(1)
        If Text1(7) <> "" Then
            .setPRECIO = moneda_bd(Text1(7))
        Else
            .setPRECIO = moneda_bd("0")
        End If
        .setOBSERVACIONES = Text1(8)
        .setANALISIS_DUPLICADO = 0
        .setCANTIDAD_MUESTRAS = CInt(Text1(0))
        .setNOMBRE = Text1(9)
        .setDESCRIPCION = Text1(10)
        If estado = "N" Then
            PLANTILLA = .Insertar
        ElseIf estado = "M" Then
            .Modificar (lista.ListItems(lista.SelectedItem.Index))
            PLANTILLA = lista.ListItems(lista.SelectedItem.Index)
        End If
        If PLANTILLA > 0 Then
           If estado = "N" Then
               nueva_plantilla = 1
           ElseIf estado = "M" Then
               nueva_plantilla = 2
           End If
           Dim tipo As New clsMuestra
           If tipo.esBano(cmbDatos(1).BoundText) Then  'es un baño Id_Espacial
              frmDetallePlantillaBano.Show 1
           End If
        End If
        inicializar_ventana
    End With
End Sub

Public Function validar_datos() As Boolean
    validar_datos = True
    If Text1(9).Text = "" Then
        MsgBox "El nombre de la plantilla no puede estar en blanco.", vbCritical, "Validación"
        Text1(9).SetFocus
        validar_datos = False
        Exit Function
    End If
    If cmbDatos(0).Text = "" Then
        MsgBox "El Cliente no puede estar en blanco.", vbCritical, "Validación"
        cmbDatos(0).SetFocus
        validar_datos = False
        Exit Function
    End If
    If cmbDatos(1).Text = "" Then
        MsgBox "El tipo de muestra no puede estar en blanco.", vbCritical, "Validación"
        cmbDatos(1).SetFocus
        validar_datos = False
        Exit Function
    End If
End Function


Public Sub inicializar_ventana()
    estado = ""
    cargar_plantillas
    bloquear
    Label1(3) = "Plantillas"
    Label1(3).BackColor = &HC0FFFF
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
