VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmFormacion_Curso 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle del curso de formación"
   ClientHeight    =   9840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14295
   Icon            =   "frmFormacion_Curso.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9840
   ScaleWidth      =   14295
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDocumentacion 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "DOCUMENTACIÓN DEL CURSO"
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
      Height          =   2505
      Left            =   4095
      TabIndex        =   57
      Top             =   3240
      Visible         =   0   'False
      Width           =   6570
      Begin VB.Frame Frame6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   1050
         Left            =   3465
         TabIndex        =   64
         Top             =   180
         Width           =   3030
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Curso de INFORMACIÓN"
            Height          =   195
            Index           =   1
            Left            =   225
            TabIndex        =   66
            Top             =   630
            Width           =   2310
         End
         Begin VB.OptionButton opTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Curso de FORMACIÓN"
            Height          =   195
            Index           =   0
            Left            =   225
            TabIndex        =   65
            Top             =   315
            Value           =   -1  'True
            Width           =   1995
         End
      End
      Begin VB.CommandButton cmdSalir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cerrar"
         Height          =   870
         Left            =   4185
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   1440
         Width           =   2310
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         Height          =   870
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   60
         Top             =   1440
         Width           =   2310
      End
      Begin VB.CheckBox chkImprimirLogoTylaer 
         BackColor       =   &H00C0C0C0&
         Height          =   195
         Left            =   180
         TabIndex        =   59
         Top             =   405
         Width           =   240
      End
      Begin VB.CheckBox chkImprimirFirmas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   195
         Left            =   180
         TabIndex        =   58
         Top             =   855
         Value           =   1  'Checked
         Width           =   240
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Logo de TYLAER"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   26
         Left            =   540
         TabIndex        =   62
         Top             =   405
         Width           =   3030
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imprimir Firmas"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   36
         Left            =   540
         TabIndex        =   61
         Top             =   855
         Width           =   1230
      End
   End
   Begin VB.CommandButton cmdDuplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   915
      Left            =   10890
      Style           =   1  'Graphical
      TabIndex        =   56
      Top             =   8865
      Width           =   1080
   End
   Begin VB.TextBox txtPlan 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   255
      Left            =   1980
      Locked          =   -1  'True
      TabIndex        =   52
      Text            =   "txtPlan"
      Top             =   495
      Visible         =   0   'False
      Width           =   9465
   End
   Begin VB.Frame frmBotones 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame6"
      Height          =   1005
      Left            =   45
      TabIndex        =   44
      Top             =   8820
      Width           =   9240
      Begin VB.CommandButton cmdCualificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Cualificaciones"
         Height          =   915
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   45
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton cmdPFA 
         BackColor       =   &H00E0E0E0&
         Caption         =   "P.F.A."
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   5670
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   45
         Width           =   1080
      End
      Begin VB.CommandButton cmdParar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Parar "
         Height          =   915
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   45
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.CommandButton cmdFinalizar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Finalizar curso"
         Height          =   915
         Left            =   6795
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   45
         Width           =   1080
      End
      Begin VB.CommandButton cmdFirmas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver Firmas"
         Height          =   915
         Left            =   1170
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   45
         Width           =   1080
      End
      Begin VB.CommandButton cmdDocCurso 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Doc. Curso"
         Height          =   915
         Left            =   2295
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   45
         Width           =   1080
      End
      Begin VB.CommandButton cmdAdjuntos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntos"
         Height          =   915
         Left            =   3420
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   45
         Width           =   1080
      End
      Begin VB.CommandButton cmdHistorialCambios 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Historial Cambios"
         Height          =   915
         Left            =   4545
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   45
         Width           =   1080
      End
      Begin VB.CommandButton cmdInvitacion 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Invitaciones"
         Height          =   915
         Left            =   45
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   45
         Width           =   1080
      End
   End
   Begin Geslab.ControlPanelXP cpFormadores 
      Height          =   2805
      Left            =   9360
      TabIndex        =   24
      Top             =   945
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   4948
      Caption         =   "Formadores"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   2805
      Begin VB.CheckBox chkExterno 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Externo"
         Height          =   375
         Left            =   90
         TabIndex        =   39
         Top             =   2340
         Width           =   915
      End
      Begin MSComctlLib.ListView listaFormadores 
         Height          =   1545
         Left            =   45
         TabIndex        =   28
         Top             =   450
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   2725
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ColHdrIcons     =   "ImageList1"
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
      Begin pryCombo.miCombo cmbFormadores 
         Height          =   330
         Left            =   45
         TabIndex        =   27
         Top             =   2025
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirFormador 
         Height          =   345
         Left            =   1080
         TabIndex        =   26
         Top             =   2385
         Width           =   1380
         _Version        =   851970
         _ExtentX        =   2434
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmFormacion_Curso.frx":08CA
      End
      Begin XtremeSuiteControls.PushButton cmdEliminarFormador 
         Height          =   360
         Left            =   2610
         TabIndex        =   25
         Top             =   2385
         Width           =   1485
         _Version        =   851970
         _ExtentX        =   2619
         _ExtentY        =   635
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmFormacion_Curso.frx":712C
      End
   End
   Begin VB.TextBox txtPREFIX 
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
      Left            =   4275
      Locked          =   -1  'True
      TabIndex        =   40
      Text            =   "Curso de Formación: "
      Top             =   90
      Width           =   3075
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Nivel de formación"
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
      Height          =   600
      Left            =   45
      TabIndex        =   36
      Top             =   855
      Width           =   4875
      Begin VB.OptionButton optNivel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Específica"
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
         Index           =   2
         Left            =   3330
         TabIndex        =   43
         Top             =   270
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
         Left            =   1800
         TabIndex        =   38
         Top             =   270
         Width           =   1230
      End
      Begin VB.OptionButton optNivel 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Técnica"
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
         Left            =   315
         TabIndex        =   37
         Top             =   270
         Value           =   -1  'True
         Width           =   1140
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Objetivos"
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
      Height          =   2100
      Left            =   45
      TabIndex        =   34
      Top             =   3420
      Width           =   9240
      Begin RichTextLib.RichTextBox txtObjetivos 
         Height          =   1770
         Left            =   90
         TabIndex        =   35
         Top             =   270
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   3122
         _Version        =   393217
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmFormacion_Curso.frx":D98E
      End
   End
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
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   7515
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "txtCurso"
      Top             =   90
      Width           =   1860
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
      Height          =   3225
      Left            =   45
      TabIndex        =   11
      Top             =   5535
      Width           =   9240
      Begin RichTextLib.RichTextBox txtContenido 
         Height          =   2805
         Left            =   90
         TabIndex        =   12
         Top             =   270
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4948
         _Version        =   393217
         Enabled         =   -1  'True
         HideSelection   =   0   'False
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmFormacion_Curso.frx":DA10
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Formación (Tipo)"
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
      Height          =   600
      Left            =   4950
      TabIndex        =   8
      Top             =   855
      Width           =   4335
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
         Left            =   585
         TabIndex        =   10
         Top             =   270
         Value           =   -1  'True
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
         Left            =   2385
         TabIndex        =   9
         Top             =   270
         Width           =   1500
      End
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salvar Curso"
      Height          =   915
      Left            =   12015
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8865
      Width           =   1080
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
      Height          =   1905
      Left            =   45
      TabIndex        =   2
      Top             =   1485
      Width           =   9240
      Begin VB.TextBox txtHoras 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   330
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   23
         Top             =   1485
         Width           =   1140
      End
      Begin VB.TextBox txtDescripcion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   330
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   3
         Top             =   315
         Width           =   7755
      End
      Begin MSComCtl2.DTPicker fechaPrevistaI 
         Height          =   360
         Left            =   1350
         TabIndex        =   5
         Top             =   675
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   51773441
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaPrevistaF 
         Height          =   360
         Left            =   3060
         TabIndex        =   14
         Top             =   675
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   51773441
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbCalidad 
         Height          =   330
         Left            =   1350
         TabIndex        =   17
         Top             =   1125
         Width           =   7755
         _ExtentX        =   13679
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fechaRealI 
         Height          =   360
         Left            =   6120
         TabIndex        =   18
         Top             =   675
         Width           =   1335
         _ExtentX        =   2355
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
         Format          =   51773441
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fechaRealF 
         Height          =   360
         Left            =   7830
         TabIndex        =   20
         Top             =   675
         Width           =   1290
         _ExtentX        =   2275
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
         Format          =   51773441
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Duración (horas)"
         Height          =   195
         Index           =   12
         Left            =   135
         TabIndex        =   22
         Top             =   1575
         Width           =   1170
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
         Left            =   7560
         TabIndex        =   21
         Top             =   720
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Real"
         Height          =   195
         Index           =   9
         Left            =   5130
         TabIndex        =   19
         Top             =   765
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Resp. Calidad"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   16
         Top             =   1170
         Width           =   990
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
         Left            =   2790
         TabIndex        =   13
         Top             =   720
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   7
         Top             =   360
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Prevista"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   6
         Top             =   765
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   915
      Left            =   13140
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8865
      Width           =   1080
   End
   Begin Geslab.ControlPanelXP cpAsistentes 
      Height          =   5010
      Left            =   9360
      TabIndex        =   29
      Top             =   3780
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   8837
      Caption         =   "Asistentes"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   5010
      Begin XtremeSuiteControls.PushButton cmdEliminarAsistente 
         Height          =   345
         Left            =   2655
         TabIndex        =   33
         Top             =   4545
         Width           =   1380
         _Version        =   851970
         _ExtentX        =   2434
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmFormacion_Curso.frx":DA92
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirAsistente 
         Height          =   345
         Left            =   1125
         TabIndex        =   32
         Top             =   4545
         Width           =   1380
         _Version        =   851970
         _ExtentX        =   2434
         _ExtentY        =   609
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmFormacion_Curso.frx":142F4
      End
      Begin pryCombo.miCombo cmbAsistentes 
         Height          =   330
         Left            =   45
         TabIndex        =   31
         Top             =   4185
         Width           =   4785
         _ExtentX        =   8440
         _ExtentY        =   582
      End
      Begin MSComctlLib.ListView listaAsistentes 
         Height          =   3705
         Left            =   45
         TabIndex        =   30
         Top             =   405
         Width           =   4815
         _ExtentX        =   8493
         _ExtentY        =   6535
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7785
      Top             =   9045
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Curso.frx":1AB56
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFormacion_Curso.frx":1AD8B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblExterna 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "F. Externa"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   330
      Left            =   90
      TabIndex        =   55
      Top             =   90
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblParado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Parado"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   11790
      TabIndex        =   42
      Top             =   270
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblFinalizado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Finalizado"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   330
      Left            =   11790
      TabIndex        =   41
      Top             =   270
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   " "
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
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   90
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   870
      Left            =   0
      Top             =   0
      Width           =   14325
   End
End
Attribute VB_Name = "frmFormacion_Curso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'M0966: Formulario creado para el código MANTIS 966.

Public PK As Long
'M1110-I
Public PLAN As Long
'M1110-F
Private Sub cabecera()
    With listaFormadores.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Externo", 750, lvwColumnCenter
        .Add , , "Formador", 2790, lvwColumnLeft
        .Add , , "Firma", listaAsistentes.Width - 3830, lvwColumnCenter
    End With
    With listaAsistentes.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Empleado", 2600, lvwColumnLeft
        .Add , , "Asiste", 800, lvwColumnCenter
        .Add , , "Firma", listaAsistentes.Width - 3750, lvwColumnCenter
    End With
End Sub

Private Sub cmdCualificar_Click()
On Error GoTo cmdok_Click_Error
    If MsgBox("Se va a cualificar en el contenido del curso a todos los asistentes ¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
    
    'Carga de los datos del curso
        Dim oCurso As New clsFormacion_cursos
        oCurso.Carga PK
        'Matriz de cualificaciones
        '-------------------------
        ' Por cada documento relacionado con el curso
        ' se marcará la lista de asistentes completa
        ' Sólo se ejecuta si el curso está vinculado a un plan de formación (getPlan_id >0)
        
        If oCurso.getPLAN_ID > 0 Then
           Dim strMsg As String
           Dim rsDocumentos As New ADODB.Recordset                       'Recordset por los documentos del curso
           Dim oDocumentos As New clsFormacion_pf_docs       'Documentos del curso
           Set rsDocumentos = oDocumentos.Listado_Plan(oCurso.getPLAN_ID)
      
           If rsDocumentos.RecordCount > 0 Then
           strMsg = "Cada asistente se ha acreditado en la formación teórica de los siguientes PNTs: " & vbCrLf
           strMsg = strMsg & "---------------------------------------------------------------------------------------------- " & vbCrLf
           Do
                Dim oDetalle As New clsCa_documentos
                Dim oAsistentes As New clsFormacion_asistentes
                Dim oFormadores As New clsFormacion_Formadores
                Dim rsFormadores As New ADODB.Recordset
                Dim rsAsistentes As New ADODB.Recordset
                Dim oCualificaciones As New clsEmpleados_cualificaciones
                Dim rsCualificaciones As New ADODB.Recordset
                
                oDetalle.Carga rsDocumentos("DOCUMENTO_ID")
                strMsg = strMsg & "(" & oDetalle.getCODIGO & ") " & oDetalle.getNOMBRE & vbCrLf
                Set rsAsistentes = oAsistentes.ListadoFirmantes(oCurso.getID_CURSO)
                Set rsFormadores = oFormadores.Listado_Internos_Firmantes(oCurso.getID_CURSO)
                If rsAsistentes.RecordCount > 0 Then 'Asistentes
                    Do
                        Set rsCualificaciones = oCualificaciones.Listado_Empleado_DOC(rsAsistentes("EMPLEADO_ID"), rsDocumentos("DOCUMENTO_ID"), 0)
                        If rsCualificaciones.RecordCount = 0 Then
                            With oCualificaciones
                             .setEMPLEADO_ID = rsAsistentes("EMPLEADO_ID")
                             .setDOCUMENTO_ID = rsDocumentos("DOCUMENTO_ID")
                             .setEMPLEADO_ID_FORMADOR = rsFormadores("FORMADOR_ID")
                             .setEN_HISTORICO = 0
                             .setES_FORMADOR = 0
                             .setESTADO = 0
                             .setFECHA_FIRMA_DIRECTOR = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                             .setFECHA_FIRMA_FORMADOR = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                             .setFECHA_FIRMA_TECNICO = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                             .setFECHA_FORMACION_TEORICA = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                             .setFECHA_ULTIMA_RECUALIFICACION = "1900-01-01"
                             .setFORMADOR_NO_CUALIFICADO = 0
                             .setID_CUALIFICACION = 0
                             .setTEXTO_FORMACION_TEORICA = "Lectura del PNT y explicación por parte del formador."
                             .Insertar
                            End With
                        End If
                        rsAsistentes.MoveNext
                        Set rsCualificaciones = Nothing
                         
                    Loop Until rsAsistentes.EOF
                End If 'Fin Asistentes
                
                If rsFormadores.RecordCount > 0 Then   'Formadores
                   Do
                        Set rsCualificaciones = oCualificaciones.Listado_Empleado_DOC(rsFormadores("FORMADOR_ID"), rsDocumentos("DOCUMENTO_ID"), 1)
                        If rsCualificaciones.RecordCount = 0 Then
                            With oCualificaciones
                             .setEMPLEADO_ID = rsFormadores("FORMADOR_ID")
                             .setDOCUMENTO_ID = rsDocumentos("DOCUMENTO_ID")
                             .setEMPLEADO_ID_FORMADOR = rsFormadores("FORMADOR_ID")
                             .setEN_HISTORICO = 0
                             .setES_FORMADOR = 1
                             .setESTADO = 0
                             .setFECHA_FIRMA_DIRECTOR = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                             .setFECHA_FIRMA_FORMADOR = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                             .setFECHA_FIRMA_TECNICO = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                             .setFECHA_FORMACION_TEORICA = Format(oCurso.getFECHA_REAL_F, "yyyy-mm-dd")
                             .setFECHA_ULTIMA_RECUALIFICACION = "1900-01-01"
                             .setFORMADOR_NO_CUALIFICADO = 0
                             .setID_CUALIFICACION = 0
                             .setTEXTO_FORMACION_TEORICA = "Lectura del PNT y explicación por parte del formador."
                             .Insertar
                            End With
                        End If
                        rsFormadores.MoveNext
                        Set rsCualificaciones = Nothing
                        Set oCualificaciones = Nothing
                    Loop Until rsFormadores.EOF
                End If                                 'Fin formadores
                
                rsDocumentos.MoveNext

                Set rsAsistentes = Nothing
                Set rsFormadores = Nothing
                Set oFormadores = Nothing
                Set oAsistentes = Nothing
                Set oDetalle = Nothing
                Set oCualificaciones = Nothing
           Loop Until rsDocumentos.EOF
           MsgBox strMsg, vbInformation + vbOKOnly, App.Title
          Else
           MsgBox "El plan de formación no tiene PNTs sobre los que cualificarse", vbInformation + vbOKOnly, App.Title
          End If
      End If
      Set rsDocumentos = Nothing
    'Registro en el historial de cambios
    
         Dim ohc As New clsHistorial_cambios
        
         With ohc
             .setTIPO = HC_TIPOS.HC_CURSO
             .setIDENTIFICADOR = PK
             .setIDENTIFICADOR_TEXTO = "Curso Formación : " & txtCurso
             .setUSUARIO_ID = USUARIO.getID_EMPLEADO
             .setMOTIVO = HC_FINALIZACION
             .Insertar
         End With
                    
         Set ohc = Nothing
         
    End If

    Exit Sub
cmdok_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdCualificar of Formulario frmFormacion_Curso"
    
End Sub

Private Sub cmdduplicar_Click()
    Dim oCurso As New clsFormacion_cursos
    Dim ID As Long
    ID = oCurso.duplicarCurso(PK)
    If ID > 0 Then
        oCurso.Carga ID
        MsgBox "El curso se ha duplicado correctamente : " & "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & oCurso.getANYO, vbInformation + vbOKOnly, App.Title
    Else
        MsgBox "Error al duplicar el curso.", vbCritical, App.Title
    End If
    Set oCurso = Nothing
End Sub

Private Sub cmdFirmas_Click()
    'Listado de firmas para CURSO (17)

    With frmFormacion_Listado_Firmas
        .PK = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
    
End Sub

Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_CURSO
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Curso de Formación " & txtCurso
        frmHistorialCambios.Show 1
    End If
End Sub

Private Sub chkexterno_Click()
        cmbFormadores.limpiar
        If chkExterno.Value = 0 Then
          llenar_combo cmbFormadores, New clsEmpleados, 0, frmEmpleados_Gestion, ""
        Else
        'MXXXX-I
        ' llenar_combo cmbFormadores, New clsProveedor, 0, frmProveedores_Detalle, ""
          llenar_combo cmbFormadores, New clsProveedor, 0, frmProveedores_Detalle, "ES_FORMADOR = 1"
        'MXXXX-F
        End If
End Sub

Private Sub cmdAdjuntos_Click()
'Adjuntos para tipo CURSO (17)

    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_CURSO
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
    
    'Registro en el historial de cambios
    
        Dim ohc As New clsHistorial_cambios
   
        With ohc
            .setTIPO = HC_TIPOS.HC_CURSO
            .setIDENTIFICADOR = PK
            .setIDENTIFICADOR_TEXTO = "Curso Formación : " & txtCurso
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setMOTIVO = HC_ADJUNTOS
            .Insertar
        End With
               
        Set ohc = Nothing
        
    
End Sub

Private Sub cmdAnadirAsistente_Click()
'Subrutina para añadir a la lista de asistentes uno nuevo desde el COMBO

    If cmbAsistentes.getPK_SALIDA = 0 Then
        MsgBox "Debe seleccionar un asistente", vbOK + vbExclamation, "Añadir empleado"
        Exit Sub
    End If
    Dim i As Integer
    ' Verificar si existe el formador
    For i = 1 To listaAsistentes.ListItems.Count
        If CLng(listaAsistentes.ListItems(i).Text) = CLng(cmbAsistentes.getPK_SALIDA) Then
            MsgBox "El Asistente ya existe en la lista.", vbExclamation, App.Title
            Exit Sub
        End If
    Next
       
    With listaAsistentes.ListItems.Add(, , cmbAsistentes.getPK_SALIDA)
        .SubItems(1) = cmbAsistentes.getTEXTO()
    End With
    cmbAsistentes.limpiar
    
End Sub

Private Sub cmdAnadirFormador_Click()
'Subrutina para añadir a la lista de asistentes uno nuevo desde el COMBO

    If cmbFormadores.getPK_SALIDA = 0 Then
        MsgBox "Debe seleccionar un formador", vbOK + vbExclamation, "Añadir formador"
        Exit Sub
    End If
    
    Dim i As Integer
    Dim externo As String
    
    If chkExterno.Value = 1 Then
       externo = "Sí"
    Else
        If chkExterno.Value = 0 Then
            externo = "No"
        End If
    End If
    ' Verificar si existe el formador
    For i = 1 To listaFormadores.ListItems.Count
        If CLng(listaFormadores.ListItems(i).Text) = CLng(cmbFormadores.getPK_SALIDA) And listaFormadores.ListItems(i).SubItems(1) = externo Then
            MsgBox "El formador ya existe en la lista.", vbExclamation, App.Title
            Exit Sub
        End If
    Next
    
    With listaFormadores.ListItems.Add(, , cmbFormadores.getPK_SALIDA)
        .SubItems(1) = externo
        .SubItems(2) = cmbFormadores.getTEXTO()
    End With
    cmbFormadores.limpiar
   
End Sub

Private Sub cmdcancel_Click()
    PK = 0
    PLAN = 0
    Unload Me
End Sub

Private Sub cmdDocCurso_Click()
    frmDocumentacion.visible = True
End Sub

Private Sub cmdEliminarAsistente_Click()
    If listaAsistentes.ListItems.Count > 0 Then
        listaAsistentes.ListItems.Remove listaAsistentes.selectedItem.Index
    End If
End Sub

Private Sub cmdEliminarFormador_Click()
    If listaFormadores.ListItems.Count > 0 Then
        listaFormadores.ListItems.Remove listaFormadores.selectedItem.Index
    End If
End Sub

Private Sub cmdFinalizar_Click()

    'Finalización del curso: creación de firmas
    'TOBJETO_FIRMA = 17
    'Acción 1: Finalizar

    On Error GoTo cmdok_Click_Error
    If MsgBox("Se va a generar una solicitud de firma por cada asistente y el curso se cerrará. ¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
    
    'Modifica los datos del curso y crea una firma por cada asistente.
        
        Dim oCurso As New clsFormacion_cursos
    
         If oCurso.generar_firmas_asistencia(PK) = True Then
            'frmTelefonos.cargar_lista_firmas
            lblFinalizado.visible = True
  
            cmdFinalizar.Enabled = False
            cmdok.Enabled = False
            cmdInvitacion.Enabled = False
            'M1110-I
            'If txtPlan.Text <> "" Then
            '    cmdCualificar.Visible = True
            'End If
            'M1110-F
            Unload Me
        End If
        
    'Registro en el historial de cambios
    
        Dim ohc As New clsHistorial_cambios
   
        With ohc
            .setTIPO = HC_TIPOS.HC_CURSO
            .setIDENTIFICADOR = PK
            .setIDENTIFICADOR_TEXTO = "Curso Formación : " & txtCurso
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setMOTIVO = HC_FINALIZACION
            .Insertar
        End With
               
        Set ohc = Nothing
        
    End If
    
    Exit Sub
    
cmdok_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdOk_Click of Formulario frmFormacion_Curso"
    
End Sub


Private Sub cmdImprimir_Click()
        Dim strCad As String
        Dim arrNom() As String
        Dim arrVal() As String
        Dim objfrm As New frmReport
        
'        objfrm.iniciar
'        Dim logo As String
'        logo = ""
'        If MsgBox("¿Desea imprimirlo con el Logo de TYLAER?", vbYesNo) = vbYes Then
'        If chkImprimirLogoTylaer.value = Checked Then
'            logo = "_TYLAER"
'        End If
'        If MsgBox("¿Desea imprimirlo como CURSO DE INFORMACIÓN?", vbYesNo) = vbYes Then
'        If opTipo(1).value = True Then
'            objfrm.informe = "Formacion\rptCurso_informacion" & logo
'        Else
'            objfrm.informe = "Formacion\rptCurso" & logo
'        End If
        
        With objfrm
            .iniciar
            .informe = "Formacion\rptCurso"
        
            ReDim arrNom(3)
            ReDim arrVal(3)
            
            arrNom(1) = "TIPO"
            If opTipo(1).Value = True Then
                arrVal(1) = "2"
            Else
                arrVal(1) = "1"
            End If
            arrNom(2) = "TYLAER"
            If chkImprimirLogoTylaer.Value = Checked Then
                arrVal(2) = "S"
            Else
                arrVal(2) = "N"
            End If
            arrNom(3) = "FIRMAS"
            If chkImprimirFirmas.Value = Checked Then
                arrVal(3) = "S"
            Else
                arrVal(3) = "N"
            End If
                    
            .ParametrosNombre = arrNom
            .ParametrosValores = arrVal
            'JGM
            .criterio = "{formacion_cursos.ID_CURSO} = " & PK & "  and {firmas_1.TOBJETO} = " & TOBJETO_ASISTENCIA_CURSO_CALIDAD
            
            .imprimir = False
            .generar
            .Show 1
            
        End With

End Sub

Private Sub cmdInvitacion_Click()
    'Invitación al curso: creación de invitaciones
    'TOBJETO = 18
 
    On Error GoTo cmdok_Click_Error
    If MsgBox("Se va a generar una petición de confirmación de asistencia para cada asistente. ¿Desea continuar?", vbQuestion + vbYesNo, App.Title) = vbYes Then
    
    'Modifica los datos del curso y crea una firma por cada asistente.
        
    Dim oCurso As New clsFormacion_cursos
    
    If oCurso.generar_firmas_invitaciones(PK) = True Then
            'frmTelefonos.cargar_lista_firmas
            MsgBox "Se han enviado las invitaciones a los asistentes ", vbOKOnly
    End If
    'Registro en el historial de cambios
    
    Dim ohc As New clsHistorial_cambios
   
    With ohc
        .setTIPO = HC_TIPOS.HC_CURSO
        .setIDENTIFICADOR = PK
        .setIDENTIFICADOR_TEXTO = "Curso Formación : " & txtCurso
        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
        .setMOTIVO = HC_INVITACIONES
        .Insertar
    End With
               
    Set ohc = Nothing
    End If
    Exit Sub
    
cmdok_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdOk_Click of Formulario frmFormacion_Curso"
End Sub
    
    
Private Sub cmdok_Click()
    guardar_curso
End Sub

Private Sub cmdParar_Click()

    On Error GoTo cmdParar_Click_Error
    
    'Detenemos o reanudamos el curso en función del estado
        
    Dim oCurso As New clsFormacion_cursos
    
    oCurso.Carga PK
    
    If oCurso.getPARADO = 0 Then
       oCurso.Parar
       lblParado.visible = True
       cmdParar.Caption = "Reanudar"
    Else
       oCurso.Reanudar
       lblParado.visible = False
       cmdParar.Caption = "Parar"
    End If
   
    'Registro en el historial de cambios
    
    Dim ohc As New clsHistorial_cambios
   
    With ohc
        .setTIPO = HC_TIPOS.HC_CURSO
        .setIDENTIFICADOR = PK
        .setIDENTIFICADOR_TEXTO = "Curso Formación : " & txtCurso
        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
        .setMOTIVO = HC_PARADA
        .Insertar
    End With
               
    Set ohc = Nothing
    Set oCurso = Nothing

    Exit Sub
    
cmdParar_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdParar_Click of Formulario frmFormacion_Curso"
    
End Sub

Private Sub cmdPFA_Click()
  '  If PLAN = 0 Then
        frmFormacion_PFA_Listado_Compacto.CURSO_ID = PK
        frmFormacion_PFA_Listado_Compacto.Show 1
   
    'Registro en el historial de cambios
    
        Dim ohc As New clsHistorial_cambios
        With ohc
            .setTIPO = HC_TIPOS.HC_CURSO
            .setIDENTIFICADOR = PK
            .setIDENTIFICADOR_TEXTO = "Curso Formación : " & txtCurso
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setMOTIVO = HC_ASIGNACION
            .Insertar
        End With
        Set ohc = Nothing
        cargar_campos
   ' End If
End Sub

Private Sub cmdSalir_Click()
    frmDocumentacion.visible = False
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargar_combos
     
    If PK <> 0 Then   'Modificación
       frmBotones.visible = True
       cargar_campos
       cargar_lista_asistentes
       cargar_lista_Formadores
    Else   'Alta
       frmBotones.visible = False
       cargar_campos_alta
    End If
End Sub



Private Sub cargar_lista_asistentes()

    Dim rs As ADODB.Recordset
    Dim oAsistente As New clsFormacion_asistentes
    Dim objLitem As ListItem, objCampo As ListSubItem
     
    listaAsistentes.ListItems.Clear
    
    Set rs = oAsistente.Listado(PK)
    If rs.RecordCount <> 0 Then
        Do
            Set objLitem = listaAsistentes.ListItems.Add(, , rs(0))
            With objLitem
               .ListSubItems.Add , , rs(1) ' EMPLEADO
               If Not IsNull(rs(2)) Then
                If rs(2) = 1 Then
                   .ListSubItems.Add , , "Si", 1  ' si asiste
                Else
                   .ListSubItems.Add , , "No", 2 ' no asiste
                End If
               Else
                   .ListSubItems.Add , , " "  ' no existe la firma
               End If
               
               If Not IsNull(rs(4)) Then
                If rs(4) = 1 Then
                   .ListSubItems.Add , , "Si", 1 ' si asiste
                Else
                   .ListSubItems.Add , , "No", 2 ' no asiste
                End If
               End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oAsistente = Nothing
  
End Sub
Private Sub cargar_lista_Formadores()

    Dim rs As ADODB.Recordset
    Dim oFormador As New clsFormacion_Formadores
    
    listaFormadores.ListItems.Clear
    
    Set rs = oFormador.Listado(PK)
    
    If rs.RecordCount <> 0 Then
       Do
            Set objLitem = listaFormadores.ListItems.Add(, , rs(0))
            With objLitem
            
                If rs(1) = 0 Then
                   .ListSubItems.Add , , "No"
                Else
                   .ListSubItems.Add , , "Sí"
                End If
 
                .ListSubItems.Add , , rs(2)
               
               If Not IsNull(rs(3)) Then
                If rs(3) = 1 Then
                   .ListSubItems.Add , , "Si", 1  ' si asiste
                Else
                   .ListSubItems.Add , , "No", 2 ' no asiste
                End If
               Else
                   .ListSubItems.Add , , " "  ' no existe la firma
               End If
               
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oFicha = Nothing
End Sub

Private Sub cargar_combos()
    llenar_combo cmbCalidad, New clsEmpleados, 0, frmEmpleados_Gestion, ""
    llenar_combo cmbAsistentes, New clsEmpleados, 0, frmEmpleados_Gestion, ""
    llenar_combo cmbFormadores, New clsEmpleados, 0, frmEmpleados_Gestion, ""
End Sub

Private Sub cargar_campos()

    Dim oCurso As New clsFormacion_cursos
'M1110-I
    Dim oPlan As New clsFormacion_pfa
'M1110-F

    oCurso.Carga PK

    txtdescripcion.Text = Trim(oCurso.getDESCRIPCION)
    txtObjetivos.Text = Trim(oCurso.getOBJETIVOS)
    
    fechaPrevistaI.Value = oCurso.getFECHA_PREVISTA_I
    fechaPrevistaF.Value = oCurso.getFECHA_PREVISTA_F
    fechaRealI.Value = oCurso.getFECHA_REAL_I
    fechaRealF.Value = oCurso.getFECHA_REAL_F
    txtContenido.Text = Trim(oCurso.getCONTENIDO)
 
    txtHoras.Text = oCurso.getNHORAS

    If oCurso.getTIPO_MODALIDAD_ID = 0 Then
        
        optModalidad(0).Value = True
        
    Else
        optModalidad(1).Value = True
    End If
    
    cmbCalidad.MostrarElemento oCurso.getID_RESPONSABLE_CALIDAD
    
    If oCurso.getREALIZADO = 1 Then
        lblFinalizado.visible = True
        lblParado.visible = False
        cmdFinalizar.Enabled = False
        cmdok.Enabled = False
        cmdParar.Enabled = False
        cmdInvitacion.Enabled = False
        'M1110-I
        If oCurso.getPLAN_ID <> 0 Then
         cmdCualificar.visible = True
        End If
        'M1110-F
    Else
        lblFinalizado.visible = False
        cmdFinalizar.Enabled = True
        cmdParar.Enabled = True
        cmdInvitacion.Enabled = True
    End If
    
    If oCurso.getPARADO = 1 Then
       lblParado.visible = True
    End If
    
    If oCurso.getTIPO_NIVEL_ID = 0 Then
        optNivel(0).Value = True
    Else
        If oCurso.getTIPO_NIVEL_ID = 1 Then
            optNivel(1).Value = True
        Else
            optNivel(2).Value = True
        End If
    End If
    
    If oCurso.getTIPO_FORMADOR_ID = 0 Then
        chkExterno.Value = 0
    Else
        chkExterno.Value = 1
    End If
    
    If oCurso.getPARADO = 1 Then
        cmdParar.Caption = "Reanudar"
    Else
        cmdParar.Caption = "Parar"
    End If
    
    
    'If optModalidad(0).value = True Then
        txtCurso = "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & oCurso.getANYO
    'Else
    '    txtCurso = "0301-" & Format(oCurso.getCOD_CURSO)
    'End If
    'M1110-I
    If oCurso.getPLAN_ID > 0 Then
       txtPlan.visible = True
       oPlan.Carga oCurso.getPLAN_ID
       txtPlan.Text = "(P.F.A.) " & oPlan.getDESCRIPCION
       PLAN = oCurso.getPLAN_ID
      ' cmdPFA.Enabled = False
       If Not USUARIO.getPER_PFA Then
        cmdPFA.Enabled = False
       End If
       chkExterno.Enabled = False
       Frame4.Enabled = False
       Frame2.Enabled = False
    Else
       If Not USUARIO.getPER_PFA Then
        cmdPFA.Enabled = False
       Else
        cmdPFA.Enabled = True
       End If
       chkExterno.Enabled = True
    End If
    'M1110-F
End Sub

Private Sub cargar_campos_alta()
    
    txtObjetivos.Text = ""
    txtContenido.Text = ""
    txtHoras.Text = 1

    configurar_codigo
        
    If PLAN > 0 Then
       txtPlan.visible = True
       Dim oPlan As New clsFormacion_pfa
       oPlan.Carga PLAN
       txtPlan.Text = "(P.F.A.) " & oPlan.getDESCRIPCION
       Set oPlan = Nothing
    Else
       txtdescripcion.Text = ""
       fechaPrevistaI.Value = Date
       fechaPrevistaF.Value = Date + 1
       fechaRealI.Value = Date
       fechaRealF.Value = Date + 1
       chkExterno.Value = 0
    End If
End Sub

Private Sub configurar_codigo()
    Dim oCurso As New clsFormacion_cursos
    
    'Cálculo del máximo Cod Curso. Es un secuencial pero no la clave.
    'Los cursos llevan secuenciales diferentes en función de la modalidad.
    
    oCurso.MaxCodCurso CLng(Year(Date))
    
    'If optModalidad(0).value = True Then
        txtCurso = "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & Year(Date)
    'Else
    '    txtCurso = "0301-" & Format(oCurso.getCOD_CURSO)
    'End If
 
End Sub

Private Sub optModalidad_Click(Index As Integer)
    If PK = 0 Then
        configurar_codigo
    End If
End Sub

Private Sub guardar_curso()
     On Error GoTo cmdok_Click_Error
      
     Dim oCurso As New clsFormacion_cursos
     Dim Curso As Long
     
     If validar Then
        With oCurso
            'M1110-I
            If PLAN > 0 Then
               .setPLAN_ID = PLAN
            End If
            'M1110-F
            .setANYO = CInt(Trim(Year(Date)))
            .setDESCRIPCION = Trim(txtdescripcion.Text)
            .setOBJETIVOS = Trim(txtObjetivos.Text)
            .setCONTENIDO = Trim(txtContenido.Text)
            .setFECHA_PREVISTA_I = Format(fechaPrevistaI.Value, "yyyy-mm-dd")
            .setFECHA_PREVISTA_F = Format(fechaPrevistaF.Value, "yyyy-mm-dd")
            .setFECHA_REAL_I = Format(fechaRealI.Value, "yyyy-mm-dd")
            .setFECHA_REAL_F = Format(fechaRealF.Value, "yyyy-mm-dd")
            .setNHORAS = CInt(txtHoras.Text)
            .setID_RESPONSABLE_CALIDAD = cmbCalidad.getPK_SALIDA
            .setAPROBADO = 1
            .setPARADO = 0
            .setREALIZADO = 0
     
            If optModalidad(0).Value = True Then
               .setTIPO_MODALIDAD_ID = 0
            Else
               .setTIPO_MODALIDAD_ID = 1
            End If
            
            If optNivel(0).Value = True Then
               .setTIPO_NIVEL_ID = 0
            Else
                If optNivel(1).Value = True Then
                    .setTIPO_NIVEL_ID = 1
                Else
                    .setTIPO_NIVEL_ID = 2
                End If
            End If
            
            Dim ohc As New clsHistorial_cambios
            
            If PK = 0 Then
                Curso = .Insertar
                If Curso = 0 Then
                    MsgBox "Error al insertar el curso.", vbCritical, App.Title
                    Exit Sub
                Else
                    .generar_firmas_calidad Curso
                    With ohc
                        .setTIPO = HC_TIPOS.HC_CURSO
                        .setIDENTIFICADOR = Curso
                        .setIDENTIFICADOR_TEXTO = "Curso Formación : " & txtCurso
                        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                        .setMOTIVO = HC_CREACION
                        .Insertar
                    End With
                End If
            Else
                If MsgBox("Va a modificar el curso. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                    frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación del curso."
                    frmMotivo.Show 1
                    If Trim(MOTIVO) = "" Then
                        MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                        Exit Sub
                    End If
                    If .Modificar(PK) = False Then
                        MsgBox "Error al modificar el curso.", vbCritical, App.Title
                        Exit Sub
                    Else
                        Curso = PK
                        With ohc
                            .setTIPO = HC_TIPOS.HC_CURSO
                            .setIDENTIFICADOR = Curso
                            .setIDENTIFICADOR_TEXTO = "Curso Formación : " & txtCurso
                            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                            .setMOTIVO = Trim(MOTIVO)
                            .Insertar
                        End With
                    End If
                Else
                    Exit Sub
                End If
            
            End If
        End With
 
        ' Asistentes
        Dim i As Integer
        Dim oFAsis As New clsFormacion_asistentes
        If PK <> 0 Then
            oFAsis.Eliminar (PK)
        End If
        
        If listaAsistentes.ListItems.Count > 0 Then
           For i = 1 To listaAsistentes.ListItems.Count
               With oFAsis
                   .setCURSO_ID = Curso
                   .setEMPLEADO_ID = listaAsistentes.ListItems(i).Text
                   .Insertar
               End With
           Next i
        End If
        
        'Formadores
        Dim oFForm As New clsFormacion_Formadores
        If PK <> 0 Then
            oFForm.Eliminar (PK)
        End If
        If listaFormadores.ListItems.Count > 0 Then
            For i = 1 To listaFormadores.ListItems.Count
                With oFForm
                    .setCURSO_ID = Curso
                    .setFORMADOR_ID = listaFormadores.ListItems(i).Text
                
                    If listaFormadores.ListItems(i).SubItems(1) = "Sí" Then
                        .setTIPO_FORMADOR_ID = 1
                    Else
                        .setTIPO_FORMADOR_ID = 0
                    End If
                    
                    .Insertar
                End With
            Next i
         End If
         If PK <> 0 Then
            MsgBox "El curso se ha modificado correctamente.", vbInformation + vbOKOnly, App.Title
         Else
            Dim strCurso As String
            'If oCurso.getTIPO_MODALIDAD_ID = 0 Then
                 strCurso = "RFI-" & Format(oCurso.getCOD_CURSO, "000") & "/" & oCurso.getANYO
            'Else
            '     strCurso = "0301-" & Format(oCurso.getCOD_CURSO)
            'End If
            
            MsgBox "Se ha creado correctamente el curso " & strCurso, vbInformation + vbOKOnly, App.Title
         End If
         frmFormacion_PFA_Detalle.lblCurso = strCurso
         frmFormacion_PFA_Detalle.lblIDCurso = Curso
         Unload Me

     End If
     
     Exit Sub
cmdok_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdOk_Click of Formulario frmFormacion_Curso Procedure guardar_curso", vbCritical, App.Title
End Sub

Private Sub envia_mensajes()
'Registra un mensaje por cada usuario

    On Error GoTo envia_mensajes_Error

    Dim oMensaje As New clsMensajes
    Dim men As Integer
    With oMensaje
        .setASUNTO = "Invitación a curso de formación"
        .setTEXTO = "Se requiere respuesta a una invitación de asistencia a un curso de formación"""
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setFECHA_INICIO = Format(Date, "yyyy-mm-dd")
        .setHORA_INICIO = "00:00:00"
        .setFECHA_FIN = Format(Date + 5, "yyyy-mm-dd")
        .setHORA_FIN = "00:00:00"
        men = .Insertar
        
        If men > 0 Then
            Dim omu As New clsMensajes_usuarios
            Dim i As Integer
            For i = 1 To listaAsistentes.ListItems.Count
                If listaAsistentes.ListItems(i).Checked = True Then
                
                    Dim oempleado As New clsEmpleados
                    
                    oempleado.CARGAR listaAsistentes.ListItems(i).Text
                    omu.setEMPLEADO_ID = oempleado.getUSUARIO_ID
                    omu.setMENSAJE_ID = men
                    omu.Insertar
                    Set oempleado = Nothing
                    
                End If
            Next
        End If
    End With
    
    Exit Sub

envia_mensajes_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure envia_mensajes_Click of Formulario frmFormacion_Curso"

   
End Sub

Private Function validar() As Boolean
    
    validar = True

    If Trim(txtdescripcion.Text) = "" Then
        MsgBox "El curso no tiene descripción", vbExclamation, App.Title
        txtdescripcion.SetFocus
        validar = False
        Exit Function
    End If
    
    If cmbCalidad.getTEXTO = "" Then
        MsgBox "Informe el responsable de calidad", vbExclamation, App.Title
        cmbCalidad.SetFocus
        validar = False
        Exit Function
    End If
    
    If Not IsNumeric(Trim(txtHoras.Text)) Then
        MsgBox "Informe el número de horas por día", vbExclamation, App.Title
        txtHoras.SetFocus
        validar = False
        Exit Function
    End If
    
    If txtObjetivos.Text = "" Then
        MsgBox "Rellene los objetivos del curso", vbExclamation, App.Title
        txtObjetivos.SetFocus
        validar = False
        Exit Function
    End If
    
    If txtContenido.Text = "" Then
        MsgBox "Rellene el contenido del curso", vbExclamation, App.Title
        txtContenido.SetFocus
        validar = False
        Exit Function
    End If
    
    If listaAsistentes.ListItems.Count = 0 Then
        MsgBox "La lista de asistentes está vacía", vbExclamation, App.Title
        
        validar = False
        Exit Function
    End If
    
    If listaFormadores.ListItems.Count = 0 Then
        MsgBox "La lista de formadores está vacía", vbExclamation, App.Title
        
        validar = False
        Exit Function
    End If
    
End Function

