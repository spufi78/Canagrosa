VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmFORMULA_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Formulas"
   ClientHeight    =   8475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmFORMULA_detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8475
   ScaleWidth      =   10380
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   720
      Left            =   9405
      Style           =   1  'Graphical
      TabIndex        =   41
      Tag             =   "Elimina el campo seleccionado"
      Top             =   4365
      Width           =   915
   End
   Begin VB.CommandButton cmdcan 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   4725
      Picture         =   "frmFORMULA_detalle.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   37
      Tag             =   "Limpia los datos para insertar nuevamente"
      Top             =   7650
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9270
      Style           =   1  'Graphical
      TabIndex        =   36
      Top             =   7560
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   8145
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   7560
      Width           =   1050
   End
   Begin VB.TextBox txtformula 
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
      Height          =   915
      Left            =   4365
      MultiLine       =   -1  'True
      TabIndex        =   33
      Top             =   6570
      Width           =   5940
   End
   Begin VB.ComboBox cmbReq 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmFORMULA_detalle.frx":0BD4
      Left            =   7965
      List            =   "frmFORMULA_detalle.frx":0BDE
      TabIndex        =   6
      Top             =   5490
      Width           =   1000
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   720
      Left            =   9405
      Style           =   1  'Graphical
      TabIndex        =   8
      Tag             =   "Elimina el campo seleccionado"
      Top             =   3600
      Width           =   915
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   720
      Left            =   9405
      Style           =   1  'Graphical
      TabIndex        =   7
      Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
      Top             =   2835
      Width           =   915
   End
   Begin VB.TextBox txtc 
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
      Height          =   315
      Index           =   2
      Left            =   4095
      TabIndex        =   4
      Top             =   5490
      Width           =   1575
   End
   Begin VB.TextBox txtc 
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
      Height          =   315
      Index           =   1
      Left            =   2610
      TabIndex        =   3
      Top             =   5490
      Width           =   1440
   End
   Begin VB.TextBox txtc 
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
      Height          =   315
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   5490
      Width           =   2500
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   45
      TabIndex        =   9
      Top             =   765
      Width           =   10290
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   615
         Index           =   1
         Left            =   1380
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   660
         Width           =   8715
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1380
         TabIndex        =   0
         Top             =   270
         Width           =   8715
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   11
         Top             =   855
         Width           =   1005
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   10
         Top             =   270
         Width           =   705
      End
   End
   Begin MSDataListLib.DataCombo cmbunidades 
      Height          =   315
      Left            =   5715
      TabIndex        =   5
      Top             =   5490
      Width           =   2220
      _ExtentX        =   3916
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
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
   Begin MSComctlLib.ListView lista 
      Height          =   2910
      Left            =   45
      TabIndex        =   13
      Top             =   2520
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   5133
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
   Begin VB.Image flecha 
      Height          =   480
      Index           =   1
      Left            =   8910
      Picture         =   "frmFORMULA_detalle.frx":0BEA
      Top             =   4185
      Width           =   480
   End
   Begin VB.Image flecha 
      Height          =   480
      Index           =   0
      Left            =   8910
      Picture         =   "frmFORMULA_detalle.frx":112A
      Top             =   3150
      Width           =   480
   End
   Begin VB.Label lblCampos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Fórmula de Cálculo"
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
      Left            =   4410
      TabIndex        =   40
      Top             =   6345
      Width           =   5850
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Fórmulas"
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
      TabIndex        =   39
      Top             =   90
      Width           =   2115
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9720
      Picture         =   "frmFORMULA_detalle.frx":1666
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de fórmulas para los tipos de determinación"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   38
      Top             =   420
      Width           =   3585
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "                  Campo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   675
      Index           =   14
      Left            =   2565
      TabIndex        =   34
      Top             =   6435
      Width           =   855
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   4
      Left            =   2430
      TabIndex        =   32
      Top             =   7950
      Width           =   555
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   2460
      Shape           =   4  'Rounded Rectangle
      Top             =   7830
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "sqrt"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   9
      Left            =   2460
      TabIndex        =   31
      Top             =   7425
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   15
      Left            =   390
      TabIndex        =   30
      Top             =   6420
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   17
      Left            =   1410
      TabIndex        =   29
      Top             =   6420
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   10
      Left            =   390
      TabIndex        =   28
      Top             =   6930
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   11
      Left            =   900
      TabIndex        =   27
      Top             =   6930
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   12
      Left            =   1410
      TabIndex        =   26
      Top             =   6930
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   5
      Left            =   390
      TabIndex        =   25
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   6
      Left            =   900
      TabIndex        =   24
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   7
      Left            =   1410
      TabIndex        =   23
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   0
      Left            =   390
      TabIndex        =   22
      Top             =   7950
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "("
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   1
      Left            =   900
      TabIndex        =   21
      Top             =   7950
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   ")"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   2
      Left            =   1410
      TabIndex        =   20
      Top             =   7950
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   18
      Left            =   1920
      TabIndex        =   19
      Top             =   6420
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   375
      Index           =   16
      Left            =   900
      TabIndex        =   15
      Top             =   6420
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Expresión Matemática para el cálculo del campo : "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   45
      TabIndex        =   14
      Top             =   5940
      Width           =   10230
   End
   Begin VB.Label lbldeter 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Campos de la fórmula"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   45
      TabIndex        =   12
      Top             =   2205
      Width           =   10245
   End
   Begin VB.Shape Shape9 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   6300
      Width           =   495
   End
   Begin VB.Shape Shape29 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   900
      Shape           =   4  'Rounded Rectangle
      Top             =   6300
      Width           =   495
   End
   Begin VB.Shape Shape11 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   1410
      Shape           =   4  'Rounded Rectangle
      Top             =   6300
      Width           =   495
   End
   Begin VB.Shape Shape21 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   6300
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   13
      Left            =   1920
      TabIndex        =   18
      Top             =   6930
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   8
      Left            =   1920
      TabIndex        =   17
      Top             =   7440
      Width           =   495
   End
   Begin VB.Label boton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Index           =   3
      Left            =   1920
      TabIndex        =   16
      Top             =   7950
      Width           =   495
   End
   Begin VB.Shape Shape22 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   6810
      Width           =   495
   End
   Begin VB.Shape Shape14 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   1410
      Shape           =   4  'Rounded Rectangle
      Top             =   6810
      Width           =   495
   End
   Begin VB.Shape Shape13 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   900
      Shape           =   4  'Rounded Rectangle
      Top             =   6810
      Width           =   495
   End
   Begin VB.Shape Shape12 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   6810
      Width           =   495
   End
   Begin VB.Shape Shape15 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape18 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   390
      Shape           =   4  'Rounded Rectangle
      Top             =   7830
      Width           =   495
   End
   Begin VB.Shape Shape16 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   900
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape19 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   900
      Shape           =   4  'Rounded Rectangle
      Top             =   7830
      Width           =   495
   End
   Begin VB.Shape Shape17 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   1410
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape20 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   1410
      Shape           =   4  'Rounded Rectangle
      Top             =   7830
      Width           =   495
   End
   Begin VB.Shape Shape23 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   7320
      Width           =   495
   End
   Begin VB.Shape Shape24 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   1920
      Shape           =   4  'Rounded Rectangle
      Top             =   7830
      Width           =   495
   End
   Begin VB.Shape Shape28 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   495
      Left            =   2460
      Shape           =   4  'Rounded Rectangle
      Top             =   7305
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H8000000D&
      Height          =   990
      Left            =   2460
      Shape           =   4  'Rounded Rectangle
      Top             =   6300
      Width           =   1125
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   735
      Left            =   0
      Top             =   0
      Width           =   10470
   End
End
Attribute VB_Name = "frmFORMULA_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

'Private Sub cmdHistorialCambios_Click()
'    If PK <> 0 Then
'        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_Formula
'        frmHistorialCambios.PK_ID = PK
'        frmHistorialCambios.PK_TITULO = "Fórmula " & txtDatos(0)
'        frmHistorialCambios.Show 1
'    End If
'End Sub
Private Sub cmdAnadir_Click()
    If validar_campo Then
        With lista.ListItems.Add(, , Trim(txtc(0)))
            .SubItems(1) = Trim(txtc(1))
            .SubItems(2) = Trim(txtc(2))
            .SubItems(3) = Trim(cmbUnidades.Text)
            .SubItems(4) = cmbReq.Text
            .SubItems(5) = cmbUnidades.BoundText
'            .SubItems(6) = "Nuevo"
        End With
        borrar_campos
        txtc(0).SetFocus
    End If
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
'        Dim odd As New clsDatos_determinaciones
'        If IsNumeric(lista.ListItems(lista.SelectedItem.Index).SubItems(6)) Then
'            If odd.Cargar_Campo(lista.ListItems(lista.SelectedItem.Index).SubItems(6)) = True Then
'                MsgBox "No puede quitar el campo de la fórmula, esta utilizado en alguna muestra registrada.", vbCritical, App.Title
'            Else
'                lista.ListItems.Remove (lista.SelectedItem.Index)
'                borrar_campos
'            End If
'         Else
            lista.ListItems.Remove (lista.selectedItem.Index)
            borrar_campos
'         End If
    End If
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If validar_campo Then
        With lista.ListItems(lista.selectedItem.Index)
            .SubItems(1) = Trim(txtc(1))
            .SubItems(2) = Trim(txtc(2))
            .SubItems(3) = Trim(cmbUnidades.Text)
            .SubItems(4) = cmbReq.Text
            .SubItems(5) = cmbUnidades.BoundText
        End With
        borrar_campos
        txtc(0).SetFocus
    End If
End Sub

Private Sub flecha_Click(Index As Integer)
    Dim aux As String
    Dim i As Integer
    If lista.ListItems.Count > 0 Then
        If Index = 0 Then 'Subir
           If lista.selectedItem.Index > 1 Then
              aux = lista.ListItems(lista.selectedItem.Index - 1).Text
              lista.ListItems(lista.selectedItem.Index - 1).Text = lista.ListItems(lista.selectedItem.Index).Text
              lista.ListItems(lista.selectedItem.Index).Text = aux
              For i = 1 To lista.ColumnHeaders.Count - 1
                  aux = lista.ListItems(lista.selectedItem.Index - 1).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index - 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
              Next
              Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index - 1)
           End If
        Else ' Bajar
           If lista.selectedItem.Index < lista.ListItems.Count Then
              aux = lista.ListItems(lista.selectedItem.Index + 1).Text
              lista.ListItems(lista.selectedItem.Index + 1).Text = lista.ListItems(lista.selectedItem.Index).Text
              lista.ListItems(lista.selectedItem.Index).Text = aux
              For i = 1 To lista.ColumnHeaders.Count - 1
                  aux = lista.ListItems(lista.selectedItem.Index + 1).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index + 1).SubItems(i) = lista.ListItems(lista.selectedItem.Index).SubItems(i)
                  lista.ListItems(lista.selectedItem.Index).SubItems(i) = aux
              Next
              Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
           End If
        End If
    End If
End Sub
Private Sub boton_Click(Index As Integer)
    Select Case Index
    Case 14 ' Campo
        If lista.ListItems.Count > 0 Then
            txtformula = txtformula & "[" & lista.ListItems(lista.selectedItem.Index) & "]"
        End If
    Case Else
        txtformula = txtformula & boton(Index).Caption
        txtformula.SelStart = Len(txtformula)
    End Select
End Sub
Private Sub cmdcan_Click()
    txtc(0) = ""
    txtc(1) = ""
    txtc(2) = ""
    cmbUnidades.Text = ""
    cmbReq.Text = ""
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = False Then
        Exit Sub
    End If
    Dim oFormula As New clsFormulas
    Dim Formula As Integer
    If PK = 0 Then
        ' Insertamos la formula
        oFormula.setNOMBRE = txtDatos(0)
        oFormula.setDESCRIPCION = txtDatos(1)
        Formula = oFormula.InsertarFormula
        If Formula = 0 Then
            Exit Sub
        End If
    Else
        If MsgBox("Va a generar una nueva versión de la fórmula. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
        ' Modificacmos la formula
        oFormula.setNOMBRE = txtDatos(0)
        oFormula.setDESCRIPCION = txtDatos(1)
        oFormula.setFORMULA_ID_ORIGEN = PK
'        FORMULA = oFormula.Modificar(PK)
        Formula = oFormula.InsertarFormula
        If Formula = 0 Then
            Exit Sub
        End If
        ' Se modifica la formula origen para ponerla como antigua
        oFormula.Pasar_Antigua PK
        ' Se modifican los Tipos de Determinacion para indicar la formula nueva
        Dim oTD As New clsTipos_determinacion
        oTD.Modificar_FORMULA CInt(PK), CInt(Formula)
        ' Borramos los campos de la formula
'        Dim ocf2 As New clsFormulas_campos
'        Dim rs2 As New ADODB.RecordSet
'        Dim consulta2 As String
'        Set rs2 = ocf2.ListaFormulas(PK)
'        If rs2.RecordCount <> 0 Then
'            Do
'                consulta2 = "delete from formulas_campos where id_campo =" & rs2("id_campo")
'                execute_bd consulta2
'                rs2.MoveNext
'            Loop Until rs2.EOF
'        End If
    End If
    ' Insertamos los campos de la formula
    Dim ocf As New clsFormulas_campos
    Dim campos(50, 2) As String
    Dim cf As Integer
    For i = 1 To lista.ListItems.Count
'        If lista.ListItems(i).SubItems(6) = "Nuevo" Then
'            oCF.CrearID
'        Else
'            oCF.setID_CAMPO = lista.ListItems(i).SubItems(6)
'        End If
        ocf.setFORMULA_ID = Formula
        ocf.setNOMBRE = lista.ListItems(i)
        ocf.setENTEROS = lista.ListItems(i).SubItems(1)
        ocf.setDECIMALES = lista.ListItems(i).SubItems(2)
        If UCase(lista.ListItems(i).SubItems(4)) = "SI" Then
            ocf.setREQUERIDO = 1
        Else
            ocf.setREQUERIDO = 0
        End If
        If Trim(lista.ListItems(i).SubItems(5)) <> "" Then
            ocf.setUNIDAD_ID = lista.ListItems(i).SubItems(5)
        Else
            ocf.setUNIDAD_ID = 0
        End If
        ocf.setFORMULA_ID_REL = 0
        If i = lista.ListItems.Count Then
            ocf.setES_SOLUCION = 1
        Else
            ocf.setES_SOLUCION = 0
        End If
        cf = ocf.InsertarCamposFormula
        campos(i, 1) = lista.ListItems(i)
        campos(i, 2) = CStr(cf)
        If cf = 0 Then
            Exit Sub
        End If
    Next
    ' Insertamos la expresion de la formula
    Dim cadena As String
    Dim CAMPO As String
    Dim EXPRESION As String
    Dim Pos As Integer
    Dim encontrado As Boolean
    cadena = txtformula
    If Trim(cadena) <> "" Then
       For i = 1 To Len(cadena)
           If Mid(cadena, i, 1) <> "[" Then
              EXPRESION = EXPRESION & Mid(cadena, i, 1)
           Else
              Pos = InStr(i + 1, cadena, "]")
              CAMPO = Mid(cadena, i + 1, (Pos) - (i + 1))
              j = 1
              encontrado = False
              Do
                 If Trim(campos(j, 1)) = Trim(CAMPO) Then
                    EXPRESION = EXPRESION & "C_" & Trim(campos(j, 2)) & "_"
                    encontrado = True
                 End If
                 j = j + 1
              Loop Until Trim(campos(j, 1)) = "" Or encontrado = True
             i = Pos
            End If
       Next
    End If
    Dim consulta As String
    consulta = "update formulas set expresion = '" & EXPRESION & "'," & _
               " campo_id_resultado=" & cf & _
               " where id_formula=" & Formula
    execute_bd consulta
    If PK = 0 Then
        MsgBox "La formula se ha insertado correctamente.", vbInformation, App.Title
    Else
        MsgBox "La formula se ha modificado correctamente.", vbInformation, App.Title
    End If
    Unload Me
    Exit Sub
fallo:
    error_grave ("Error al almacenar la formula. " & Err.Description)
End Sub


Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Call cabecera
    Call cargar_unidades
    If PK <> 0 Then
        cargar_formula
    End If
End Sub
Public Sub cabecera()
    ' Campos
    With lista.ColumnHeaders
        .Add , , "Nombre", 2500, lvwColumnLeft
        .Add , , "Enteros", 1565, lvwColumnCenter
        .Add , , "Decimales", 1565, lvwColumnCenter
        .Add , , "Unidad", 2150, lvwColumnCenter
        .Add , , "Requerido", 1000, lvwColumnCenter
        .Add , , "ID_UNIDAD", 1, lvwColumnCenter
        .Add , , "ID_CAMPO", 1, lvwColumnCenter
    End With
End Sub
Public Sub cargar_unidades()
    Dim ounidades As New clsUnidades
    Set cmbUnidades.RowSource = ounidades.Listado("")
    cmbUnidades.ListField = "nombre"
    cmbUnidades.BoundColumn = "id_unidad"
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txtc(0) = lista.ListItems(lista.selectedItem.Index).Text
        txtc(1) = lista.ListItems(lista.selectedItem.Index).SubItems(1)
        txtc(2) = lista.ListItems(lista.selectedItem.Index).SubItems(2)
        cmbUnidades.Text = lista.ListItems(lista.selectedItem.Index).SubItems(3)
        cmbReq.Text = lista.ListItems(lista.selectedItem.Index).SubItems(4)
    End If
End Sub
Private Sub txtc_GotFocus(Index As Integer)
    txtc(Index).BackColor = &H80C0FF
    txtc(Index).SelStart = 0
    txtc(Index).SelLength = Len(txtc(Index))
End Sub
Private Sub txtc_LostFocus(Index As Integer)
    txtc(Index).BackColor = vbWhite
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub borrar_campos()
    txtc(0) = ""
    txtc(1) = ""
    txtc(2) = ""
    cmbUnidades.Text = ""
    cmbReq.Text = ""
End Sub
Public Sub cargar_formula()
    Dim ofor As New clsFormulas
    ofor.CARGAR (PK)
    ' Título
    txtDatos(0) = ofor.getNOMBRE
    txtDatos(1) = ofor.getDESCRIPCION
    ' Campos de la formula
    Dim ocf As New clsFormulas_campos
    Dim ouni As New clsUnidades
    Dim rs As ADODB.Recordset
    Set rs = ocf.ListaFormulas(PK)
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Trim(rs("nombre")))
                    .SubItems(1) = Trim(rs("enteros"))
                    .SubItems(2) = Trim(rs("decimales"))
                    ouni.CARGAR (rs("unidad_id"))
                    .SubItems(3) = Trim(ouni.getNOMBRE)
                    If rs("requerido") = 0 Then
                        .SubItems(4) = "No"
                    Else
                        .SubItems(4) = "Si"
                    End If
                    .SubItems(5) = ouni.getID_UNIDAD
                    .SubItems(6) = rs("id_campo")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    ' Expresion
    Dim cadena As String
    Dim CAMPO As String
    Dim Formula As String
    Dim Pos As Integer
    Dim encontrado As Boolean
    prefijo = ""
    cadena = ofor.getEXPRESION
    If Not IsNull(cadena) Then
        For i = 1 To Len(cadena)
            If Mid(cadena, i, 1) <> "C" Then
                Formula = Formula & Mid(cadena, i, 1)
            Else
                Pos = InStr(i + 2, cadena, "_")
                If Pos > 0 Then
                    CAMPO = Mid(cadena, i + 2, (Pos) - (i + 2))
                    ocf.CARGAR (CAMPO)
                    Formula = Formula & "[" & ocf.getNOMBRE & "]"
                    i = Pos
                End If
            End If
        Next
    End If
    txtformula = Formula
End Sub

Private Function validar_campo() As Boolean
    validar_campo = True
    If cmbReq.Text = "" Then
        MsgBox "Debe seleccionar si el campo es requerido.", vbInformation, App.Title
        validar_campo = False
        cmbReq.SetFocus
        Exit Function
    End If
    If txtc(0).Text = "" Then
        MsgBox "Debe un nombre al campo.", vbInformation, App.Title
        validar_campo = False
        txtc(0).SetFocus
        Exit Function
    End If
    If txtc(1).Text = "" Then
        MsgBox "Debe anotar un numero de enteros.", vbInformation, App.Title
        validar_campo = False
        txtc(1).SetFocus
        Exit Function
    End If
    If txtc(2).Text = "" Then
        MsgBox "Debe anotar un numero de decimales.", vbInformation, App.Title
        validar_campo = False
        txtc(2).SetFocus
        Exit Function
    End If
    If UCase(cmbReq.Text) <> "SI" And UCase(cmbReq.Text) <> "NO" Then
        MsgBox "Indique correctamente si el campo es requerido (Si,No).", vbInformation, App.Title
        validar_campo = False
        cmbReq.SetFocus
        Exit Function
    End If
End Function

Private Function validar() As Boolean
    validar = True
    If txtDatos(0).Text = "" Then
        MsgBox "Debe dar un nombre a la formula.", vbInformation, App.Title
        txtDatos(0).SetFocus
        validar = False
        Exit Function
    End If
    If lista.ListItems.Count = 0 Then
        MsgBox "Debe existir algún campo para almacenar la fórmula.", vbInformation, App.Title
        txtc(0).SetFocus
        validar = False
        Exit Function
    End If
End Function
