VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#39.0#0"; "miCombo.ocx"
Begin VB.Form frmFormacion_Gestion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Formación de Empleados"
   ClientHeight    =   12765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14490
   Icon            =   "frmFormacion_Gestionl.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12765
   ScaleWidth      =   14490
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCurso 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   375
      Left            =   5040
      TabIndex        =   34
      Text            =   "txtCurso"
      Top             =   135
      Width           =   3570
   End
   Begin VB.TextBox txtObjetivos 
      Alignment       =   2  'Center
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
      Left            =   1755
      MaxLength       =   75
      TabIndex        =   28
      Top             =   1485
      Width           =   12390
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Evaluación del formador"
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
      Height          =   2460
      Left            =   6795
      TabIndex        =   27
      Top             =   9090
      Width           =   7620
      Begin VB.OptionButton optValoracion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Negativa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   3330
         TabIndex        =   44
         Top             =   585
         Width           =   1500
      End
      Begin VB.OptionButton optValoracion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Positiva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   3330
         TabIndex        =   43
         Top             =   315
         Width           =   1500
      End
      Begin RichTextLib.RichTextBox txtValoracion 
         Height          =   1410
         Left            =   135
         TabIndex        =   46
         Top             =   900
         Width           =   7395
         _ExtentX        =   13044
         _ExtentY        =   2487
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmFormacion_Gestionl.frx":08CA
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Valoración general"
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
         Index           =   13
         Left            =   1395
         TabIndex        =   45
         Top             =   450
         Width           =   1710
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Contenido del curso"
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
      Height          =   4710
      Left            =   6795
      TabIndex        =   25
      Top             =   4320
      Width           =   7620
      Begin RichTextLib.RichTextBox txtContenido 
         Height          =   4290
         Left            =   90
         TabIndex        =   26
         Top             =   270
         Width           =   7485
         _ExtentX        =   13203
         _ExtentY        =   7567
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmFormacion_Gestionl.frx":094C
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Formadores"
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
      Height          =   2460
      Left            =   45
      TabIndex        =   21
      Top             =   9090
      Width           =   6720
      Begin MSComctlLib.ListView listaFormadores 
         Height          =   1575
         Left            =   135
         TabIndex        =   22
         Top             =   765
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   2778
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14609914
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
      Begin pryCombo.miCombo cmbFormadores 
         Height          =   330
         Left            =   135
         TabIndex        =   24
         Top             =   315
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   582
      End
   End
   Begin VB.CommandButton cmdNuevo 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Nuevo Curso"
      Height          =   915
      Left            =   5445
      Picture         =   "frmFormacion_Gestionl.frx":09CE
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   11745
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Modalidad y nivel de formación"
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
      Height          =   1095
      Left            =   45
      TabIndex        =   10
      Top             =   3195
      Width           =   14370
      Begin VB.OptionButton optFormador 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Interno"
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
         Left            =   12870
         TabIndex        =   18
         Top             =   405
         Width           =   1455
      End
      Begin VB.OptionButton optFormador 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Externo"
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
         Left            =   12870
         TabIndex        =   17
         Top             =   675
         Width           =   1455
      End
      Begin VB.OptionButton optNivel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Técnico"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   7515
         TabIndex        =   15
         Top             =   405
         Width           =   1500
      End
      Begin VB.OptionButton optNivel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "General"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   7515
         TabIndex        =   14
         Top             =   675
         Width           =   1500
      End
      Begin VB.OptionButton optModalidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Teórica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   1485
         TabIndex        =   12
         Top             =   405
         Width           =   1500
      End
      Begin VB.OptionButton optModalidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Práctica"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   1485
         TabIndex        =   11
         Top             =   675
         Width           =   1500
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Formador"
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
         Left            =   11655
         TabIndex        =   19
         Top             =   540
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nivel de formación"
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
         Left            =   5535
         TabIndex        =   16
         Top             =   540
         Width           =   1680
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modalidad"
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
         Left            =   180
         TabIndex        =   13
         Top             =   540
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   915
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   11745
      Width           =   1275
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Generales"
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
      Height          =   2400
      Left            =   45
      TabIndex        =   4
      Top             =   720
      Width           =   14415
      Begin VB.TextBox txtHoras 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   465
         Left            =   8865
         MaxLength       =   75
         TabIndex        =   42
         Top             =   1800
         Width           =   1140
      End
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
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
         Left            =   1710
         MaxLength       =   75
         TabIndex        =   5
         Top             =   315
         Width           =   12390
      End
      Begin MSComCtl2.DTPicker fechaPrevistaI 
         Height          =   360
         Left            =   1710
         TabIndex        =   7
         Top             =   1215
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   635
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
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaPrevistaF 
         Height          =   360
         Left            =   3870
         TabIndex        =   33
         Top             =   1215
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   635
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
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbCalidad 
         Height          =   330
         Left            =   8865
         TabIndex        =   36
         Top             =   1260
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fechaRealI 
         Height          =   360
         Left            =   1710
         TabIndex        =   37
         Top             =   1755
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   635
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
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaRealF 
         Height          =   360
         Left            =   3870
         TabIndex        =   39
         Top             =   1755
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   635
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
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Duración (horas)"
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
         Left            =   7245
         TabIndex        =   41
         Top             =   1890
         Width           =   1485
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
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
         Left            =   3600
         TabIndex        =   40
         Top             =   1800
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Real"
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
         Left            =   135
         TabIndex        =   38
         Top             =   1800
         Width           =   1050
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable de Calidad "
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
         Left            =   6435
         TabIndex        =   35
         Top             =   1305
         Width           =   2310
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
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
         Left            =   3600
         TabIndex        =   32
         Top             =   1260
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Objetivos"
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
         Left            =   630
         TabIndex        =   31
         Top             =   810
         Width           =   855
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
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
         Left            =   405
         TabIndex        =   9
         Top             =   360
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Prevista"
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
         Left            =   135
         TabIndex        =   8
         Top             =   1260
         Width           =   1350
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Asistentes"
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
      Height          =   4710
      Left            =   45
      TabIndex        =   2
      Top             =   4320
      Width           =   6720
      Begin MSComctlLib.ListView listaAsistentes 
         Height          =   3825
         Left            =   135
         TabIndex        =   3
         Top             =   765
         Width           =   6480
         _ExtentX        =   11430
         _ExtentY        =   6747
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14609914
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
      Begin pryCombo.miCombo cmbAsistentes 
         Height          =   330
         Left            =   135
         TabIndex        =   23
         Top             =   315
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   582
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   915
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   11745
      Width           =   1275
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   720
      Top             =   10485
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   90
      Top             =   10395
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Gestionl.frx":2DA0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Gestionl.frx":367A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Gestionl.frx":3F54
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción"
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
      Left            =   495
      TabIndex        =   30
      Top             =   2025
      Width           =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción"
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
      Left            =   450
      TabIndex        =   29
      Top             =   2025
      Width           =   1080
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle del plan de formación"
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
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   225
      Width           =   2505
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13725
      Picture         =   "frmFormacion_Gestionl.frx":482E
      Top             =   135
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   14445
   End
End
Attribute VB_Name = "frmFormacion_Gestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargar_tree
    cargar_empleados
End Sub

Private Sub cargar_tree()
     Dim nodX As Node
     Tree.Nodes.Clear
     '--FAMILIA DE DOCUMENTO DE CALIDAD
     '------SUBFAMILIA DE DOCUMENTO
     '------------DOCUMENTOS
     Dim rs As ADODB.RecordSet
     Dim consulta As String
     Dim familia As Integer
     Dim subfamilia As Integer
     Dim documento As Integer
     consulta = "SELECT C.ID_DOCUMENTO,C.FAMILIA_ID,C.SUBFAMILIA_ID,D2.DESCRIPCION,D.DESCRIPCION,CONCAT('(',C.CODIGO,') ', C.NOMBRE)" & _
                " FROM CA_DOCUMENTOS C, DECODIFICADORA D, DECODIFICADORA D2 " & _
                " Where d.codigo = " & DECODIFICADORA.CA_DOCUMENTOS_SUBFAMILIAS & " And D2.codigo = " & DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS & _
                " AND C.FAMILIA_ID = D2.VALOR " & _
                " AND C.SUBFAMILIA_ID = D.VALOR " & _
                " AND C.FORMACION = 1 " & _
                " ORDER BY D2.DESCRIPCION,D.DESCRIPCION,C.NOMBRE"
     Set rs = datos_bd(consulta)
     If rs.RecordCount > 0 Then
        Do
'            Tree.Nodes(nodX.Index).Bold = True
            If familia <> rs(1) Then
                familia = rs(1)
                Set nodX = Tree.Nodes.Add(, , "ID:" & familia, rs(3), 1)
                subfamilia = rs(2)
                Set nodX = Tree.Nodes.Add("ID:" & familia, tvwChild, "ID:" & familia & "-" & subfamilia, rs(4), 2)
            End If
            If subfamilia <> rs(2) Then
                subfamilia = rs(2)
                Set nodX = Tree.Nodes.Add("ID:" & familia, tvwChild, "ID:" & familia & "-" & subfamilia, rs(4), 2)
            End If
            Set nodX = Tree.Nodes.Add("ID:" & familia & "-" & subfamilia, tvwChild, "ID:" & familia & "-" & subfamilia & "-" & rs(0), rs(5), 3)
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oDeco = Nothing
End Sub
Private Sub cabecera()
        With listaAsistentes.ColumnHeaders
            .Add , , "ID", 300, lvwColumnLeft
            .Add , , "Empleado", lista.Width - 300, lvwColumnLeft
        End With
End Sub

Private Sub cargar_empleados()
    Dim oE As New clsEmpleados
    Dim rs As ADODB.RecordSet
'    Set rs = oE.Listado
'    If rs.RecordCount > 0 Then
'        Do
'            With lista.ListItems.Add(, , rs("ID_EMPLEADO"))
'                .SubItems(1) = rs("NOMBRE")
'            End With
'            rs.MoveNext
'        Loop Until rs.EOF
'    End If
    Set oE = Nothing
    
    
    
    
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Label1_Click(Index As Integer)

End Sub

Private Sub RichTextBox1_Change()

End Sub
