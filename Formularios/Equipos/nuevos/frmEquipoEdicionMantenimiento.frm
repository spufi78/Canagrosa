VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#34.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipoEdicionMantenimiento 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9210
   ClientLeft      =   3390
   ClientTop       =   1305
   ClientWidth     =   10065
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmEquipoEdicionMantenimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   4020
      Index           =   2
      Left            =   45
      TabIndex        =   51
      Top             =   4275
      Width           =   9990
      Begin VB.CheckBox chkmant 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Trimestral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   3105
         TabIndex        =   9
         Top             =   0
         Width           =   1410
      End
      Begin VB.CheckBox chkmant 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Anual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   6
         Left            =   6435
         TabIndex        =   11
         Top             =   0
         Width           =   960
      End
      Begin VB.CheckBox chkmant 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Semestral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   4725
         TabIndex        =   10
         Top             =   0
         Width           =   1410
      End
      Begin VB.CheckBox chkmant 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bimestral"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   1530
         TabIndex        =   8
         Top             =   0
         Width           =   1365
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   11
         Left            =   9135
         TabIndex        =   37
         Top             =   3150
         Width           =   240
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   12
         Left            =   9135
         TabIndex        =   39
         Top             =   3555
         Width           =   285
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   9
         Left            =   9135
         TabIndex        =   33
         Top             =   2340
         Width           =   240
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   10
         Left            =   9135
         TabIndex        =   35
         Top             =   2745
         Width           =   285
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   7
         Left            =   6120
         TabIndex        =   29
         Top             =   3150
         Width           =   240
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   8
         Left            =   6120
         TabIndex        =   31
         Top             =   3555
         Width           =   285
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   5
         Left            =   6120
         TabIndex        =   25
         Top             =   2340
         Width           =   240
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   6
         Left            =   6120
         TabIndex        =   27
         Top             =   2745
         Width           =   285
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   3
         Left            =   3105
         TabIndex        =   21
         Top             =   3150
         Width           =   240
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   4
         Left            =   3105
         TabIndex        =   23
         Top             =   3555
         Width           =   285
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   1
         Left            =   3105
         TabIndex        =   17
         Top             =   2340
         Width           =   240
      End
      Begin VB.CheckBox chkop 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   2
         Left            =   3105
         TabIndex        =   19
         Top             =   2745
         Width           =   285
      End
      Begin VB.CheckBox chkmant 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mensual"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   135
         TabIndex        =   7
         Top             =   0
         Width           =   1275
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   6
         Left            =   1530
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   1440
         Width           =   8250
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   5
         Left            =   1530
         MaxLength       =   255
         TabIndex        =   12
         Top             =   1080
         Width           =   8250
      End
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   1
         Left            =   3195
         TabIndex        =   15
         Top             =   1845
         Width           =   240
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   345
         Index           =   1
         Left            =   1530
         TabIndex        =   14
         Top             =   1800
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   1
         Left            =   1440
         TabIndex        =   16
         Top             =   2295
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   2
         Left            =   1440
         TabIndex        =   18
         Top             =   2700
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   3
         Left            =   1440
         TabIndex        =   20
         Top             =   3105
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   4
         Left            =   1440
         TabIndex        =   22
         Top             =   3510
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   5
         Left            =   4455
         TabIndex        =   24
         Top             =   2295
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   6
         Left            =   4455
         TabIndex        =   26
         Top             =   2700
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   7
         Left            =   4455
         TabIndex        =   28
         Top             =   3105
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   8
         Left            =   4455
         TabIndex        =   30
         Top             =   3510
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   9
         Left            =   7470
         TabIndex        =   32
         Top             =   2295
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   10
         Left            =   7470
         TabIndex        =   34
         Top             =   2700
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   11
         Left            =   7470
         TabIndex        =   36
         Top             =   3105
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Index           =   12
         Left            =   7470
         TabIndex        =   38
         Top             =   3510
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbMensualModalidad 
         Height          =   315
         Left            =   1530
         TabIndex        =   76
         Top             =   360
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbMensualResp_externo 
         Height          =   330
         Left            =   1530
         TabIndex        =   78
         Top             =   720
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbMensualResp_interno 
         Height          =   330
         Left            =   1530
         TabIndex        =   79
         Top             =   720
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   24
         Left            =   135
         TabIndex        =   77
         Top             =   405
         Width           =   735
      End
      Begin VB.Shape Shape1 
         Height          =   1725
         Left            =   450
         Top             =   2205
         Width           =   9195
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Noviembre"
         Height          =   195
         Index           =   21
         Left            =   6570
         TabIndex        =   67
         Top             =   3150
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Diciembre"
         Height          =   195
         Index           =   20
         Left            =   6570
         TabIndex        =   66
         Top             =   3555
         Width           =   705
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Septiembre"
         Height          =   195
         Index           =   19
         Left            =   6570
         TabIndex        =   65
         Top             =   2385
         Width           =   795
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Octubre"
         Height          =   195
         Index           =   18
         Left            =   6570
         TabIndex        =   64
         Top             =   2745
         Width           =   570
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Julio"
         Height          =   195
         Index           =   17
         Left            =   3735
         TabIndex        =   63
         Top             =   3150
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Agosto"
         Height          =   195
         Index           =   16
         Left            =   3735
         TabIndex        =   62
         Top             =   3555
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mayo"
         Height          =   195
         Index           =   15
         Left            =   3735
         TabIndex        =   61
         Top             =   2385
         Width           =   390
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Junio"
         Height          =   195
         Index           =   14
         Left            =   3735
         TabIndex        =   60
         Top             =   2745
         Width           =   375
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Marzo"
         Height          =   195
         Index           =   13
         Left            =   720
         TabIndex        =   59
         Top             =   3150
         Width           =   435
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Abril"
         Height          =   195
         Index           =   12
         Left            =   720
         TabIndex        =   58
         Top             =   3555
         Width           =   300
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Enero"
         Height          =   195
         Index           =   10
         Left            =   720
         TabIndex        =   57
         Top             =   2385
         Width           =   420
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Febrero"
         Height          =   195
         Index           =   0
         Left            =   720
         TabIndex        =   56
         Top             =   2745
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   55
         Top             =   765
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   54
         Top             =   1485
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Registro"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   53
         Top             =   1125
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Próximo"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   52
         Top             =   1845
         Width           =   1050
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   2175
      Index           =   1
      Left            =   45
      TabIndex        =   46
      Top             =   2025
      Width           =   9990
      Begin VB.CheckBox Check1 
         BackColor       =   &H00C0C0C0&
         Height          =   240
         Index           =   0
         Left            =   3330
         TabIndex        =   6
         Top             =   1755
         Width           =   240
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   1530
         MaxLength       =   255
         TabIndex        =   3
         Top             =   990
         Width           =   8250
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   3
         Left            =   1530
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1350
         Width           =   8250
      End
      Begin VB.CheckBox chkmant 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Semanal"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   2
         Top             =   0
         Width           =   1365
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   345
         Index           =   0
         Left            =   1530
         TabIndex        =   5
         Top             =   1710
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   609
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
         Format          =   70582273
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbSemanalModalidad 
         Height          =   315
         Left            =   1530
         TabIndex        =   72
         Top             =   270
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbSemanalResp_interno 
         Height          =   330
         Left            =   1530
         TabIndex        =   74
         Top             =   630
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbSemanalResp_externo 
         Height          =   330
         Left            =   1530
         TabIndex        =   75
         Top             =   630
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   23
         Left            =   135
         TabIndex        =   73
         Top             =   315
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Próximo"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   50
         Top             =   1755
         Width           =   1050
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Registro"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   49
         Top             =   1035
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   48
         Top             =   1395
         Width           =   1005
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   47
         Top             =   675
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8955
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   8325
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   8325
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
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
      Height          =   1410
      Index           =   0
      Left            =   45
      TabIndex        =   42
      Top             =   540
      Width           =   9990
      Begin VB.CheckBox chkmant 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Diario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   0
         Top             =   0
         Width           =   1050
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Index           =   0
         Left            =   1530
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   990
         Width           =   8250
      End
      Begin MSDataListLib.DataCombo cmbDiarioModalidad 
         Height          =   315
         Left            =   1530
         TabIndex        =   68
         Top             =   270
         Width           =   2925
         _ExtentX        =   5159
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbDiarioResp_interno 
         Height          =   330
         Left            =   1530
         TabIndex        =   70
         Top             =   630
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbDiarioResp_externo 
         Height          =   330
         Left            =   1530
         TabIndex        =   71
         Top             =   630
         Width           =   6150
         _ExtentX        =   10848
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   22
         Left            =   135
         TabIndex        =   69
         Top             =   315
         Width           =   735
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   45
         Top             =   675
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Procedimiento"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   43
         Top             =   1035
         Width           =   1005
      End
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9450
      Picture         =   "frmEquipoEdicionMantenimiento.frx":000C
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFC0&
      Caption         =   "Equipo"
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
      TabIndex        =   44
      Top             =   120
      Width           =   750
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   10305
   End
End
Attribute VB_Name = "frmEquipoEdicionMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Private Sub chkmant_Click(Index As Integer)
    Select Case Index
        Case 0 ' diario
            If chkmant(Index).value = Checked Then
                Call estado_marco_diario(True)
            Else
                Call estado_marco_diario(False)
            End If
        
        Case 1 ' semanal
            If chkmant(Index).value = Checked Then
                Call estado_marco_semanal(True)
                Check1(0).Enabled = True
                Check1(0).value = Checked
                fecha(0).Enabled = True
            Else
                Call estado_marco_semanal(False)
                Check1(0).Enabled = False
                Check1(0).value = Unchecked
                fecha(0).value = Format("1900-01-01", "yyyy-mm-dd")
                fecha(0).Enabled = False
                cmbSemanalModalidad.BoundText = 0
                cmbSemanalResp_interno.Limpiar
                cmbSemanalResp_externo.Limpiar
                txtDatos(3) = ""
                txtDatos(4) = ""
            End If
            
        Case 2, 3, 4, 5, 6 ' mensual
            If (chkmant(2).value = Checked Or chkmant(3).value = Checked _
               Or chkmant(4).value = Checked Or chkmant(5).value = Checked Or chkmant(6).value = Checked) Then
               
                Call estado_marco_mensual(True)

            Else
                Call estado_marco_mensual(False)
            End If
            
    End Select
End Sub

Private Sub Check1_Click(Index As Integer)
    If Check1(Index).value = Unchecked Then
        fecha(Index).value = Format("1900-01-01", "yyyy-mm-dd")
        fecha(Index).Enabled = False
    Else
        fecha(Index).Enabled = True
    End If
    
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    
    Call cargar_combos
    
    Dim titulo As String
    If PK <> 0 Then
        CARGAR
    End If
End Sub

Private Sub cargar_combos()
    Dim oDECO As New clsDecodificadora
    
    ' diario
    oDECO.Cargar_Combo cmbDiarioModalidad, decodificadora.EQ_TIPO_CALIBRACION
    llenar_combo cmbDiarioResp_interno, New clsUsuarios, 0, Me, ""
    llenar_combo cmbDiarioResp_externo, New clsProveedor, 0, frmProveedores, ""
    
    ' semanal
    oDECO.Cargar_Combo cmbSemanalModalidad, decodificadora.EQ_TIPO_CALIBRACION
    llenar_combo cmbSemanalResp_interno, New clsUsuarios, 0, Me, ""
    llenar_combo cmbSemanalResp_externo, New clsProveedor, 0, frmProveedores, ""
    
    ' mensual
    oDECO.Cargar_Combo cmbMensualModalidad, decodificadora.EQ_TIPO_CALIBRACION
    llenar_combo cmbMensualResp_interno, New clsUsuarios, 0, Me, ""
    llenar_combo cmbMensualResp_externo, New clsProveedor, 0, frmProveedores, ""
    
    Set oDECO = Nothing
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub

Private Sub chkop_Click(Index As Integer)
    If chkop(Index).value = Checked Then
        f1(Index).Enabled = True
    Else
        f1(Index).Enabled = False
    End If
End Sub

Private Sub cmbDiarioModalidad_Change()
    If UCase(cmbDiarioModalidad.Text) = "EXTERNA" Then
        cmbDiarioResp_interno.Limpiar
        cmbDiarioResp_interno.Visible = False
        cmbDiarioResp_externo.cargar_datos
        cmbDiarioResp_externo.Visible = True
        cmbDiarioResp_externo.activar
        cmbDiarioResp_externo.Limpiar
    ElseIf UCase(cmbDiarioModalidad.Text) = "INTERNA" Then
        cmbDiarioResp_externo.Limpiar
        cmbDiarioResp_externo.Visible = False
        cmbDiarioResp_interno.cargar_datos
        cmbDiarioResp_interno.Visible = True
        cmbDiarioResp_interno.activar
        cmbDiarioResp_interno.Limpiar
    End If
End Sub

Private Sub cmbSemanalModalidad_Change()
    If UCase(cmbSemanalModalidad.Text) = "EXTERNA" Then
        cmbSemanalResp_interno.Limpiar
        cmbSemanalResp_interno.Visible = False
        'cmbSemanalResp_externo.cargar_datos
        cmbSemanalResp_externo.Visible = True
        cmbSemanalResp_externo.activar
        cmbSemanalResp_externo.Limpiar
    ElseIf UCase(cmbSemanalModalidad.Text) = "INTERNA" Then
        cmbSemanalResp_externo.Limpiar
        cmbSemanalResp_externo.Visible = False
        'cmbSemanalResp_interno.cargar_datos
        cmbSemanalResp_interno.Visible = True
        cmbSemanalResp_interno.activar
        cmbSemanalResp_interno.Limpiar
    ElseIf UCase(cmbSemanalModalidad.Text) = "" Then
        cmbSemanalResp_interno.Limpiar
        cmbSemanalResp_externo.Limpiar
    End If
End Sub

Private Sub cmbMensualModalidad_Change()
    If UCase(cmbMensualModalidad.Text) = "EXTERNA" Then
        cmbMensualResp_interno.Limpiar
        cmbMensualResp_interno.Visible = False
        'cmbMensualResp_externo.cargar_datos
        cmbMensualResp_externo.Visible = True
        cmbMensualResp_externo.activar
        cmbMensualResp_externo.Limpiar
    ElseIf UCase(cmbMensualModalidad.Text) = "INTERNA" Then
        cmbMensualResp_externo.Limpiar
        cmbMensualResp_externo.Visible = False
        'cmbMensualResp_interno.cargar_datos
        cmbMensualResp_interno.Visible = True
        cmbMensualResp_interno.activar
        cmbMensualResp_interno.Limpiar
    ElseIf UCase(cmbMensualModalidad.Text) = "" Then
        cmbMensualResp_interno.Limpiar
        cmbMensualResp_externo.Limpiar
    End If
End Sub

Public Sub CARGAR()
    Dim oEquipo As New clsEquipos
    
    If oEquipo.Carga(PK) = True Then
    
        lbltitulo = "Mantenimiento del Equipo : " & oEquipo.getNOMBRE
        Me.Caption = lbltitulo
        
        Dim oEC As New clsEquipos_mantenimiento
        If oEC.Carga(PK) Then
            With oEC
                chkmant(0) = .getDIARIO_MANTENIMIENTO
                chkmant(1) = .getSEMANAL_MANTENIMIENTO
                chkmant(2) = .getMENSUAL_MANTENIMIENTO
                chkmant(3) = .getTRIMESTRAL_MANTENIMIENTO
                chkmant(4) = .getBIMESTRAL_MANTENIMIENTO
                chkmant(5) = .getSEMESTRAL_MANTENIMIENTO
                chkmant(6) = .getANUAL_MANTENIMIENTO
                
                ' Diario
                If .getDIARIO_MANTENIMIENTO = 1 Then
                    Call estado_marco_diario(True)
                    txtDatos(0) = .getDIARIO_PROCEDIMIENTO
                    cmbDiarioModalidad.BoundText = .getDIARIO_MODALIDAD_ID
                    cmbDiarioResp_interno.MostrarElemento .getDIARIO_RESPONSABLE_INTERNO_ID
                    cmbDiarioResp_externo.MostrarElemento .getDIARIO_RESPONSABLE_EXTERNO_ID
                Else
                    Call estado_marco_diario(False)
                End If
                ' -------------------------
                
                ' Semanal
                If .getSEMANAL_MANTENIMIENTO = 1 Then
                    Call estado_marco_semanal(True)
                    txtDatos(4) = .getSEMANAL_REGISTRO
                    txtDatos(3) = .getSEMANAL_PROCEDIMIENTO
                    cmbSemanalModalidad.BoundText = .getSEMANAL_MODALIDAD_ID
                    cmbSemanalResp_interno.MostrarElemento .getSEMANAL_RESPONSABLE_INTERNO_ID
                    cmbSemanalResp_externo.MostrarElemento .getSEMANAL_RESPONSABLE_EXTERNO_ID
                    
                    If Format(.getSEMANAL_FECHA, "yyyy-mm-dd") = "1900-01-01" Or Not IsDate(.getSEMANAL_FECHA) Then
                        Check1(0).value = Unchecked
                        fecha(0).Enabled = False
                    Else
                        Check1(0).value = Checked
                        fecha(0) = .getSEMANAL_FECHA
                    End If
                Else
                    Call estado_marco_semanal(False)
                End If
                ' -------------------------
                
                ' Mensual
                If .getMENSUAL_MANTENIMIENTO = 1 Or .getBIMESTRAL_MANTENIMIENTO = 1 _
                    Or .getTRIMESTRAL_MANTENIMIENTO = 1 Or .getTRIMESTRAL_MANTENIMIENTO = 1 _
                    Or .getSEMESTRAL_MANTENIMIENTO = 1 Or .getANUAL_MANTENIMIENTO = 1 Then
                    
                    Call estado_marco_mensual(True)
                    txtDatos(5) = .getMENSUAL_REGISTRO
                    txtDatos(6) = .getMENSUAL_PROCEDIMIENTO
                    cmbMensualModalidad.BoundText = .getMENSUAL_MODALIDAD_ID
                    cmbMensualResp_interno.MostrarElemento .getMENSUAL_RESPONSABLE_INTERNO_ID
                    cmbMensualResp_externo.MostrarElemento .getMENSUAL_RESPONSABLE_EXTERNO_ID
                    
                    If Format(.getMENSUAL_FECHA, "yyyy-mm-dd") = "1900-01-01" Or Not IsDate(.getMENSUAL_FECHA) Then
                        Check1(1).value = Unchecked
                        fecha(1).Enabled = False
                    Else
                        Check1(1).value = Checked
                        fecha(1) = .getMENSUAL_FECHA
                    End If
                Else
                    Call estado_marco_mensual(False)
                End If
                ' -------------------------
                
                ' Plan de mantenimiento
                Dim i As Integer
                Dim sfecha As String
                For i = 1 To 12
                    Select Case i
                    Case 1
                        sfecha = .getMENSUAL_ENERO
                    Case 2
                        sfecha = .getMENSUAL_FEBRERO
                    Case 3
                        sfecha = .getMENSUAL_MARZO
                    Case 4
                        sfecha = .getMENSUAL_ABRIL
                    Case 5
                        sfecha = .getMENSUAL_MAYO
                    Case 6
                        sfecha = .getMENSUAL_JUNIO
                    Case 7
                        sfecha = .getMENSUAL_JULIO
                    Case 8
                        sfecha = .getMENSUAL_AGOSTO
                    Case 9
                        sfecha = .getMENSUAL_SEPTIEMBRE
                    Case 10
                        sfecha = .getMENSUAL_OCTUBRE
                    Case 11
                        sfecha = .getMENSUAL_NOVIEMBRE
                    Case 12
                        sfecha = .getMENSUAL_DICIEMBRE
                    End Select
                    If Format(sfecha, "yyyy-mm-dd") = "1900-01-01" Or Not IsDate(sfecha) Then
                        chkop(i).value = Unchecked
                        f1(i).Enabled = False
                    Else
                        chkop(i).value = Checked
                        f1(i) = sfecha
                    End If
                Next
                ' -------------------------
            End With
        Else
            Call estado_marco_semanal(False)
            Call estado_marco_mensual(False)
        End If
    End If
    
    Set oEquipo = Nothing
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    If datos_mantenimiento_correctos() Then
      
        Dim oEC As New clsEquipos_mantenimiento
        With oEC
            .Eliminar PK
            
            .setEQUIPO_ID = PK
            .setDIARIO_MANTENIMIENTO = chkmant(0).value
            .setSEMANAL_MANTENIMIENTO = chkmant(1).value
            .setMENSUAL_MANTENIMIENTO = chkmant(2).value
            .setTRIMESTRAL_MANTENIMIENTO = chkmant(3).value
            .setBIMESTRAL_MANTENIMIENTO = chkmant(4).value
            .setSEMESTRAL_MANTENIMIENTO = chkmant(5).value
            .setANUAL_MANTENIMIENTO = chkmant(6).value
            
            ' Diario
            .setDIARIO_PROCEDIMIENTO = txtDatos(0)
            .setDIARIO_MODALIDAD_ID = IIf(cmbDiarioModalidad.BoundText <> "", cmbDiarioModalidad.BoundText, 0)
            If UCase(cmbDiarioModalidad.Text) = "INTERNA" Then
               .setDIARIO_RESPONSABLE_INTERNO_ID = cmbDiarioResp_interno.getPK_SALIDA
               .setDIARIO_RESPONSABLE_EXTERNO_ID = 0
            ElseIf UCase(cmbDiarioModalidad.Text) = "EXTERNA" Then
               .setDIARIO_RESPONSABLE_EXTERNO_ID = cmbDiarioResp_externo.getPK_SALIDA
               .setDIARIO_RESPONSABLE_INTERNO_ID = 0
            Else
               .setDIARIO_RESPONSABLE_EXTERNO_ID = 0
               .setDIARIO_RESPONSABLE_INTERNO_ID = 0
            End If
            
            ' Semanal
            .setSEMANAL_REGISTRO = txtDatos(4)
            .setSEMANAL_PROCEDIMIENTO = txtDatos(3)
            .setSEMANAL_MODALIDAD_ID = IIf(cmbSemanalModalidad.BoundText <> "", cmbSemanalModalidad.BoundText, 0)
            If UCase(cmbSemanalModalidad.Text) = "INTERNA" Then
               .setSEMANAL_RESPONSABLE_INTERNO_ID = cmbSemanalResp_interno.getPK_SALIDA
               .setSEMANAL_RESPONSABLE_EXTERNO_ID = 0
            ElseIf UCase(cmbSemanalModalidad.Text) = "EXTERNA" Then
               .setSEMANAL_RESPONSABLE_EXTERNO_ID = cmbSemanalResp_externo.getPK_SALIDA
               .setSEMANAL_RESPONSABLE_INTERNO_ID = 0
            Else
               .setSEMANAL_RESPONSABLE_EXTERNO_ID = 0
               .setSEMANAL_RESPONSABLE_INTERNO_ID = 0
            End If
            
            If Check1(0).value = Checked Then
               .setSEMANAL_FECHA = Format(fecha(0), "yyyy-mm-dd")
            Else
               .setSEMANAL_FECHA = "1900-01-01"
            End If
            
            ' Mensual
            .setMENSUAL_REGISTRO = txtDatos(5)
            .setMENSUAL_PROCEDIMIENTO = txtDatos(6)
            .setMENSUAL_MODALIDAD_ID = IIf(cmbMensualModalidad.BoundText <> "", cmbMensualModalidad.BoundText, 0)
            If UCase(cmbMensualModalidad.Text) = "INTERNA" Then
               .setMENSUAL_RESPONSABLE_INTERNO_ID = cmbMensualResp_interno.getPK_SALIDA
               .setMENSUAL_RESPONSABLE_EXTERNO_ID = 0
            ElseIf UCase(cmbMensualModalidad.Text) = "EXTERNA" Then
               .setMENSUAL_RESPONSABLE_EXTERNO_ID = cmbMensualResp_externo.getPK_SALIDA
               .setMENSUAL_RESPONSABLE_INTERNO_ID = 0
            Else
               .setMENSUAL_RESPONSABLE_EXTERNO_ID = 0
               .setMENSUAL_RESPONSABLE_INTERNO_ID = 0
            End If
            
            If Check1(1).value = Checked Then
               .setMENSUAL_FECHA = Format(fecha(1), "yyyy-mm-dd")
            Else
               .setMENSUAL_FECHA = "1900-01-01"
            End If
             
            ' Plan de mantenimiento
            Dim i As Integer
            Dim sfecha As String
            For i = 1 To 12
                If chkop(i).value = Checked Then
                   sfecha = Format(f1(i), "yyyy-mm-dd")
                Else
                   sfecha = "1900-01-01"
                End If
                Select Case i
                    Case 1
                        .setMENSUAL_ENERO = sfecha
                    Case 2
                         .setMENSUAL_FEBRERO = sfecha
                    Case 3
                         .setMENSUAL_MARZO = sfecha
                    Case 4
                         .setMENSUAL_ABRIL = sfecha
                    Case 5
                         .setMENSUAL_MAYO = sfecha
                    Case 6
                         .setMENSUAL_JUNIO = sfecha
                    Case 7
                         .setMENSUAL_JULIO = sfecha
                    Case 8
                         .setMENSUAL_AGOSTO = sfecha
                    Case 9
                         .setMENSUAL_SEPTIEMBRE = sfecha
                    Case 10
                         .setMENSUAL_OCTUBRE = sfecha
                    Case 11
                         .setMENSUAL_NOVIEMBRE = sfecha
                    Case 12
                         .setMENSUAL_DICIEMBRE = sfecha
                    End Select
                    If Format(sfecha, "yyyy-mm-dd") = "1900-01-01" Or Not IsDate(sfecha) Then
                        chkop(i).value = Unchecked
                        f1(i).Enabled = False
                    Else
                        chkop(i).value = Checked
                        f1(i) = sfecha
                    End If
                Next
                .Insertar
                
                'frmEquipos_Detalle.datos_mantenimiento (PK)
                
            End With
            
            MsgBox "El mantenimiento del equipo se ha actualizado correctamente.", vbOKOnly + vbInformation, App.Title
            Unload Me
    End If
End Sub

' función que comprueba los datos del mantenimiento
Private Function datos_mantenimiento_correctos() As Boolean
    ' diario
    If chkmant(0).value = Checked Then
        If cmbDiarioModalidad.BoundText = "0" Or cmbDiarioModalidad.BoundText = "" Then
            MsgBox "Debe indicar la modalidad del mantenimiento diario.", vbInformation, App.Title
            cmbDiarioModalidad.SetFocus
            datos_mantenimiento_correctos = False
            Exit Function
        Else
            If cmbDiarioModalidad.BoundText = "1" Then ' interna
                If cmbDiarioResp_interno.getPK_SALIDA = 0 Then
                    MsgBox "Debe indicar responsable del mantenimiento diario.", vbInformation, App.Title
                    cmbDiarioResp_interno.SetFocus
                    datos_mantenimiento_correctos = False
                    Exit Function
                End If
            ElseIf cmbDiarioModalidad.BoundText = "2" Then ' externa
                If cmbDiarioResp_externo.getPK_SALIDA = 0 Then
                    MsgBox "Debe indicar responsable del mantenimiento diario.", vbInformation, App.Title
                    cmbDiarioResp_externo.SetFocus
                    datos_mantenimiento_correctos = False
                    Exit Function
                End If
            End If
        End If
    End If
    ' ----------------------
    
    ' semanal
    If chkmant(1).value = Checked Then
        If cmbSemanalModalidad.BoundText = "0" Or cmbSemanalModalidad.BoundText = "" Then
            MsgBox "Debe indicar la modalidad del mantenimiento semanal.", vbInformation, App.Title
            cmbSemanalModalidad.SetFocus
            datos_mantenimiento_correctos = False
            Exit Function
        Else
            If cmbSemanalModalidad.BoundText = "1" Then ' interna
                If cmbSemanalResp_interno.getPK_SALIDA = 0 Then
                    MsgBox "Debe indicar responsable del mantenimiento semanal.", vbInformation, App.Title
                    cmbSemanalResp_interno.SetFocus
                    datos_mantenimiento_correctos = False
                    Exit Function
                End If
            ElseIf cmbSemanalModalidad.BoundText = "2" Then ' externa
                If cmbSemanalResp_externo.getPK_SALIDA = 0 Then
                    MsgBox "Debe indicar responsable del mantenimiento semanal.", vbInformation, App.Title
                    cmbSemanalResp_externo.SetFocus
                    datos_mantenimiento_correctos = False
                    Exit Function
                End If
            End If
        End If
        
        If Format(fecha(0), "yyyy-mm-dd") = "1900-01-01" Then
            MsgBox "Debe indicar la próxima fecha del mantenimiento semanal.", vbInformation, App.Title
            If fecha(0).Enabled = True Then
                fecha(0).SetFocus
            End If
            datos_mantenimiento_correctos = False
            Exit Function
        End If
        
    End If
    ' ----------------------
    
    ' mensual
    If (chkmant(2).value = Checked Or chkmant(3).value = Checked Or chkmant(4).value = Checked _
       Or chkmant(5).value = Checked Or chkmant(6).value = Checked) Then
       
        If cmbMensualModalidad.BoundText = "0" Or cmbMensualModalidad.BoundText = "" Then
            MsgBox "Debe indicar la modalidad del mantenimiento.", vbInformation, App.Title
            cmbMensualModalidad.SetFocus
            datos_mantenimiento_correctos = False
            Exit Function
        Else
            If cmbMensualModalidad.BoundText = "1" Then ' interna
                If cmbMensualResp_interno.getPK_SALIDA = 0 Then
                    MsgBox "Debe indicar responsable del mantenimiento.", vbInformation, App.Title
                    cmbMensualResp_interno.SetFocus
                    datos_mantenimiento_correctos = False
                    Exit Function
                End If
            ElseIf cmbMensualModalidad.BoundText = "2" Then ' externa
                If cmbMensualResp_externo.getPK_SALIDA = 0 Then
                    MsgBox "Debe indicar responsable del mantenimiento.", vbInformation, App.Title
                    cmbMensualResp_externo.SetFocus
                    datos_mantenimiento_correctos = False
                    Exit Function
                End If
            End If
        End If
        
        If Format(fecha(1), "yyyy-mm-dd") = "1900-01-01" Then
            MsgBox "Debe indicar la fecha próxima del mantenimiento.", vbInformation, App.Title
            If fecha(1).Enabled = True Then
                fecha(1).SetFocus
            End If
            datos_mantenimiento_correctos = False
            Exit Function
        End If
        
    End If
    ' ----------------------
    
    datos_mantenimiento_correctos = True
End Function

Private Sub estado_marco_diario(booActivo As Boolean)
    txtDatos(0).Enabled = booActivo
    cmbDiarioModalidad.Enabled = booActivo
    If Not booActivo Then
        cmbDiarioResp_interno.Limpiar
        cmbDiarioResp_interno.desactivar
        cmbDiarioResp_externo.Limpiar
        cmbDiarioResp_externo.desactivar
        cmbDiarioModalidad.BoundText = 0
        txtDatos(0) = ""
    Else
        cmbDiarioResp_interno.activar
        cmbDiarioResp_externo.activar
    End If
    
End Sub

Private Sub estado_marco_semanal(booActivo As Boolean)
    txtDatos(3).Enabled = booActivo
    txtDatos(4).Enabled = booActivo
    cmbSemanalModalidad.Enabled = booActivo
    If Not booActivo Then
        cmbSemanalResp_interno.Limpiar
        cmbSemanalResp_interno.desactivar
        cmbSemanalResp_externo.Limpiar
        cmbSemanalResp_externo.desactivar
    Else
        cmbSemanalResp_interno.activar
        cmbSemanalResp_externo.activar
    End If
    fecha(0).Enabled = booActivo
    Check1(0).Enabled = booActivo
End Sub

Private Sub estado_marco_mensual(booActivo As Boolean)
            txtDatos(5).Enabled = booActivo
            txtDatos(6).Enabled = booActivo
            cmbMensualModalidad.Enabled = booActivo
            If Not booActivo Then
                cmbMensualResp_interno.Limpiar
                cmbMensualResp_interno.desactivar
                cmbMensualResp_externo.Limpiar
                cmbMensualResp_externo.desactivar
                
                Check1(1).Enabled = False
                Check1(1).value = Unchecked
                fecha(1).value = Format("1900-01-01", "yyyy-mm-dd")
                fecha(1).Enabled = False
                cmbMensualModalidad.BoundText = 0
                cmbMensualResp_interno.Limpiar
                cmbMensualResp_externo.Limpiar
                txtDatos(5) = ""
                txtDatos(6) = ""
            Else
                cmbMensualResp_interno.activar
                cmbMensualResp_externo.activar
            End If
            fecha(1).Enabled = booActivo
            Check1(1).Enabled = booActivo
End Sub
