VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmFormacion_PNT 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Relación de PNTs"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   6135
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   915
      Left            =   4815
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4770
      Width           =   1275
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   915
      Left            =   3465
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   4770
      Width           =   1275
   End
   Begin VB.TextBox txtCurso 
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
      Height          =   300
      Left            =   1305
      TabIndex        =   3
      Top             =   180
      Width           =   4065
   End
   Begin pryCombo.miCombo cmbDocumentos 
      Height          =   330
      Left            =   1305
      TabIndex        =   0
      Top             =   630
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   582
   End
   Begin MSComctlLib.ListView listaDocumentos 
      Height          =   3525
      Left            =   45
      TabIndex        =   4
      Top             =   1125
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   6218
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "CURSO:"
      Height          =   240
      Left            =   585
      TabIndex        =   2
      Top             =   225
      Width           =   600
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "PNT:"
      Height          =   240
      Left            =   765
      TabIndex        =   1
      Top             =   675
      Width           =   420
   End
End
Attribute VB_Name = "frmFormacion_PNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

