VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmPlasma_Dureza_Shore 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Registro de Resultados Muestra de Plasma"
   ClientHeight    =   8910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13725
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlasma_Dureza_Shore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   13725
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDurezaEspesor 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   6345
      MaxLength       =   255
      TabIndex        =   62
      Top             =   8415
      Visible         =   0   'False
      Width           =   4830
   End
   Begin VB.TextBox txtUnidades 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2520
      TabIndex        =   58
      Top             =   8190
      Visible         =   0   'False
      Width           =   1815
   End
   Begin Geslab.ControlPanelXP cpReactivos 
      Height          =   3975
      Left            =   6885
      TabIndex        =   21
      Top             =   2745
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   7011
      Caption         =   "Reactivos Utilizados en la Muestra"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   3975
      Begin VB.Frame frmReactivos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Reactivos"
         ForeColor       =   &H80000008&
         Height          =   3480
         Left            =   45
         TabIndex        =   22
         Top             =   450
         Width           =   6630
         Begin VB.CommandButton cmdAnadirReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   750
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   24
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   1395
            Width           =   915
         End
         Begin VB.CommandButton cmdEliminarReactivo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   23
            Tag             =   "Elimina el campo seleccionado"
            Top             =   450
            Width           =   915
         End
         Begin MSComctlLib.ListView listaReactivos 
            Height          =   2460
            Left            =   45
            TabIndex        =   25
            Top             =   135
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   4339
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
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
         Begin pryCombo.miCombo cmbReactivos 
            Height          =   330
            Left            =   765
            TabIndex        =   26
            Top             =   2700
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbReactivosInternos 
            Height          =   330
            Left            =   765
            TabIndex        =   27
            Top             =   3060
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   582
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Externo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   5
            Left            =   90
            TabIndex        =   29
            Top             =   2745
            Width           =   540
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Interno"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
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
            TabIndex        =   28
            Top             =   3105
            Width           =   495
         End
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP2 
      Height          =   3975
      Left            =   45
      TabIndex        =   14
      Top             =   2745
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   7011
      Caption         =   "Equipos Utilizados en la Muestra"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      Object.Height          =   3975
      Begin VB.Frame frmEquipos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   90
         TabIndex        =   15
         Top             =   405
         Width           =   6585
         Begin VB.CommandButton cmdVerificacion 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Verificación"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   31
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   1920
            Width           =   915
         End
         Begin VB.CommandButton cmdAnadirEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Añadir"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   765
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   17
            Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
            Top             =   1110
            Width           =   915
         End
         Begin VB.CommandButton cmdEliminarEquipo 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Eliminar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Left            =   5670
            Style           =   1  'Graphical
            TabIndex        =   16
            Tag             =   "Elimina el campo seleccionado"
            Top             =   270
            Width           =   915
         End
         Begin MSComctlLib.ListView listaEquipos 
            Height          =   2730
            Left            =   0
            TabIndex        =   18
            Top             =   270
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   4815
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
         Begin pryCombo.miCombo cmbEquipos 
            Height          =   330
            Left            =   0
            TabIndex        =   19
            Top             =   3060
            Width           =   5595
            _ExtentX        =   9869
            _ExtentY        =   582
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Marque los equipos que deben salir en el informe"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Index           =   0
            Left            =   0
            TabIndex        =   20
            Top             =   45
            Width           =   4335
         End
      End
   End
   Begin VB.Frame frmRockwell 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1005
      Left            =   45
      TabIndex        =   45
      Top             =   6750
      Width           =   13650
      Begin VB.TextBox txtDurezaAverageR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7740
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   59
         Top             =   360
         Visible         =   0   'False
         Width           =   1815
      End
      Begin VB.TextBox txtPOR 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   12465
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   56
         Top             =   540
         Width           =   1140
      End
      Begin VB.TextBox txtSD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10935
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   54
         Top             =   540
         Width           =   1185
      End
      Begin VB.TextBox txtDurezaDimension 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1710
         MaxLength       =   255
         TabIndex        =   47
         Top             =   495
         Width           =   5325
      End
      Begin VB.TextBox txtDurezaReq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9315
         MaxLength       =   255
         TabIndex        =   51
         Top             =   180
         Width           =   4290
      End
      Begin VB.TextBox txtDurezaAverage 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9315
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   48
         Top             =   540
         Width           =   1050
      End
      Begin VB.TextBox txtDurezaResults 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1710
         MaxLength       =   255
         TabIndex        =   46
         Top             =   180
         Width           =   5325
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "%"
         Height          =   195
         Index           =   21
         Left            =   12240
         TabIndex        =   57
         Top             =   585
         Width           =   150
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "S.D."
         Height          =   195
         Index           =   20
         Left            =   10440
         TabIndex        =   55
         Top             =   585
         Width           =   390
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.SPECIFICATION"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   53
         Top             =   540
         Width           =   1590
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "REQUIREMENT"
         Height          =   195
         Index           =   0
         Left            =   7875
         TabIndex        =   52
         Top             =   225
         Width           =   1380
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "AVERAGE"
         Height          =   195
         Index           =   6
         Left            =   8280
         TabIndex        =   50
         Top             =   585
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "RESULTS"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   49
         Top             =   225
         Width           =   870
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "RESULT"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4545
      TabIndex        =   42
      Top             =   7920
      Width           =   4425
      Begin VB.CheckBox chkResult 
         BackColor       =   &H00C0C0C0&
         Caption         =   "chkResult"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   225
         TabIndex        =   43
         Top             =   270
         Value           =   1  'Checked
         Width           =   240
      End
      Begin VB.Label lblResult 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "PASS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   540
         TabIndex        =   44
         Top             =   225
         Width           =   3390
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "SPECIMEN ID AND DESCRIPTION"
      ForeColor       =   &H80000008&
      Height          =   2370
      Left            =   45
      TabIndex        =   33
      Top             =   360
      Width           =   13650
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   50
         Left            =   1350
         MaxLength       =   255
         TabIndex        =   2
         Top             =   990
         Width           =   11085
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   51
         Left            =   1350
         MaxLength       =   255
         TabIndex        =   3
         Top             =   1665
         Width           =   2715
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   54
         Left            =   1350
         MaxLength       =   255
         TabIndex        =   6
         Top             =   1980
         Width           =   2715
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   52
         Left            =   5580
         MaxLength       =   255
         TabIndex        =   4
         Top             =   1665
         Width           =   2760
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   55
         Left            =   5580
         MaxLength       =   255
         TabIndex        =   7
         Top             =   1980
         Width           =   2760
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   53
         Left            =   9630
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1665
         Width           =   2805
      End
      Begin pryCombo.miCombo cmbProcess 
         Height          =   345
         Left            =   1350
         TabIndex        =   0
         Top             =   270
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbCustomer 
         Height          =   345
         Left            =   1350
         TabIndex        =   1
         Top             =   630
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbnatype 
         Height          =   345
         Left            =   1350
         TabIndex        =   60
         Top             =   1305
         Width           =   12150
         _ExtentX        =   21431
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº AND TYPE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   54
         Left            =   135
         TabIndex        =   61
         Top             =   1350
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "S/N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   41
         Top             =   2025
         Width           =   345
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "SPECIMEN ID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   40
         Top             =   1035
         Width           =   1035
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "P/N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   39
         Top             =   1725
         Width           =   345
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "PROCESS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   38
         Top             =   345
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "CUSTOMER"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   14
         Left            =   135
         TabIndex        =   37
         Top             =   705
         Width           =   915
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRODUCT S/N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   15
         Left            =   4185
         TabIndex        =   36
         Top             =   2025
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRODUCT TYPE:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   4185
         TabIndex        =   35
         Top             =   1725
         Width           =   1305
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "MODULE S/N:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   8505
         TabIndex        =   34
         Top             =   1710
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdImagen 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imagenes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7920
      Width           =   1140
   End
   Begin VB.CommandButton cmdObservador 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Observador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   9
      Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
      Top             =   7920
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.CheckBox chkModificar 
      Caption         =   "Permiso Modificar Cerrada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9315
      TabIndex        =   32
      Top             =   7875
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CheckBox chkDuplicada 
      Caption         =   "Duplicada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   9315
      TabIndex        =   30
      Top             =   8100
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   12510
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7920
      Width           =   1140
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   11340
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7920
      Width           =   1140
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "THICKNESS"
      Height          =   195
      Index           =   19
      Left            =   5220
      TabIndex        =   63
      Top             =   8550
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label lblCerrada 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   345
      Left            =   11925
      TabIndex        =   13
      Top             =   0
      Width           =   1770
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resultados de Muestra de Plasma"
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
      Height          =   330
      Left            =   45
      TabIndex        =   12
      Top             =   0
      Width           =   13725
   End
End
Attribute VB_Name = "frmPlasma_Dureza_Shore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub chkAlabe_Click()
    If chkAlabe.Value = Checked Then
        frmVickers.Enabled = True
    Else
        frmVickers.Enabled = False
    End If
End Sub

Private Sub chkPreparation_Click()
    If chkPreparation.Value = Checked Then
        frmPreparation.Enabled = True
        cmbUsuarioPreparation.MostrarElemento USUARIO.getID_EMPLEADO
    Else
        frmPreparation.Enabled = False
        cmbUsuarioPreparation.limpiar
    End If
End Sub
Private Sub chkResult_Click()
    If chkResult.Value = Checked Then
        lblResult.Caption = "PASS"
        lblResult.ForeColor = &H8000&
    Else
        lblResult.Caption = "FAIL"
        lblResult.ForeColor = vbRed
    End If
End Sub

Private Sub cmbProcess_change()
    Dim oPP As New clsPlasma_procesos
    Dim oPF As New clsPlasma_ficha
    Dim oPE As New clsPlasma_ensayos
    
    If cmbProcess.getTEXTO = "" Then
        txtDurezaReq = ""
    Else
'        oPP.Carga cmbProcess.getPK_SALIDA
'        oPF.Carga oPP.getBOND_COAT_FICHA_ID
'        If opTipo(0).Value = True Then ' Rockwell
'            txtDurezaReq = oPF.getMACRO_DUREZA_REQ
'            oPE.Carga oPF.getMACRO_DUREZA
'        Else ' Vicker
'            txtDurezaReq = oPF.getMICRO_DUREZA_REQ
'            oPE.Carga oPF.getMICRO_DUREZA
'        End If
'        Dim ounidad As New clsUnidades
'        ounidad.CARGAR oPE.getUNIDAD_ID
'        txtUnidades = ounidad.getNOMBRE
    End If
    Set oPP = Nothing
    Set oPF = Nothing
End Sub

Private Sub cmdImagen_Click()
    With frmCE_Imagenes
        .PK = PK
        .Show 1
    End With
End Sub

Private Sub cmdObservador_Click()

    Dim objfrm As New frmObservadorEnsayo

    objfrm.FORMULARIO_ORIGEN = 2 'Sellantes asociado al número 2
    objfrm.ES_CONTROL_EFICACIA = False
    objfrm.MUESTRA_ID = PK ' Id de la muestra
    objfrm.DETERMINACION_ENSAYO_ID = 0
    objfrm.SELLANTE_ID = txtID_SELLANTE
    objfrm.ENSAYO = lista.ListItems(lista.selectedItem.Index)
    
    If (UCase(lblCerrada) <> "CERRADA") Then
        objfrm.MUESTRA_CERRADA = False
    Else
        objfrm.MUESTRA_CERRADA = True
    End If

    objfrm.Show vbModal
    
    Set objfrm = Nothing

End Sub
'MANTIS-807-F'

Private Sub cmdok_Click()
    Dim oPRE As New clsPlasma_recepcion
   On Error GoTo cmdok_Click_Error

   On Error GoTo cmdok_Click_Error
    ' Validar equipos pendientes de verificacion
    Dim cont As Integer
    Dim oEV As New clsEquipoVerificacion
    Dim salidaVerificacion As String
    Dim salidaVerificacionAux As String
    For cont = 1 To listaEquipos.ListItems.Count
        salidaVerificacionAux = oEV.pendienteVerificacion(listaEquipos.ListItems(cont).Text, Date)
        If salidaVerificacionAux <> "" Then
            salidaVerificacion = salidaVerificacion & " - " & salidaVerificacionAux & vbNewLine
        End If
    Next
    If salidaVerificacion <> "" Then
        If MsgBox("ATENCIÓN : " & vbNewLine & salidaVerificacion & vbNewLine & " ¿Desea continuar?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    ' Validaciones de campos
    Dim i As Integer
    Dim listaResultados() As String
    If txtDurezaResults <> "" Then
        listaResultados = Split(txtDurezaResults, "-")
        If UBound(listaResultados) <> 3 Then
            If MsgBox("ATENCIÓN : " & vbNewLine & " NO HA INTRODUCIDO 4 RESULTADOS " & vbNewLine & " ¿Desea continuar?", vbExclamation + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    ' Validar rangos DUREZA
    If txtDurezaAverage <> "" Then
        If IsNumeric(txtDurezaAverage) Then
            If CInt(txtDurezaAverage) >= 55 And CInt(txtDurezaAverage) <= 85 Then
                chkResult.Value = Checked
            Else
                If CInt(txtDurezaAverage) < 55 Then
                    If MsgBox("El porcentaje de DUREZA es menor de 55. Se marcará el ensayo como FAIL. ¿Desea continuar?", vbCritical + vbYesNo, App.Title) = vbNo Then
                        Exit Sub
                    Else
                        chkResult.Value = Unchecked
                    End If
                End If
                If CInt(txtDurezaAverage) > 85 Then
                    If MsgBox("El porcentaje de DUREZA es mayor de 85. Se marcará el ensayo como FAIL. ¿Desea continuar?", vbCritical + vbYesNo, App.Title) = vbNo Then
                        Exit Sub
                    Else
                        chkResult.Value = Unchecked
                    End If
                End If
            End If
        End If
    End If
    ' Grabación de datos
    Me.MousePointer = 11
    With oPRE
        .setPROCESO_ID = cmbProcess.getPK_SALIDA
        .setCUSTOMER_ID = cmbCustomer.getPK_SALIDA
        .setSPECIMEN_ID = txtDatos(50)
        .setNTYPE = cmbnatype.getPK_SALIDA
        .setPN = txtDatos(51)
        .setPRODUCT_TYPE = txtDatos(52)
        .setMODULE_SN = txtDatos(53)
        .setSN = txtDatos(54)
        .setPRODUCT_SN = txtDatos(55)
        .setMP = 0
        .setMP_FECHA = "NULL"
        .setMP_USUARIO_ID = 0
        .setMP_PASS = 0
        .Modificar PK
        .ModificarResultado PK, chkResult.Value
        .informarControlSpecification PK
    End With
    Set oPRE = Nothing
    ' RESULTADOS
    Dim opd As New clsPlasma_dureza
    Dim res As String
    With opd
        .setMUESTRA_ID = PK
        .setIDENTIFICATION = "HARDNESS TEST (PNT IB 208)"
        .setDIMENSION = txtDurezaDimension
        .setESPESOR = txtDurezaEspesor
        .setREQUIREMENT = txtDurezaReq
        .setRESULT = txtDurezaResults
        .setAVERAGE = txtDurezaAverage
        If txtSD = "" Then
            .setSD = 0
        Else
            .setSD = Replace(txtSD, ",", ".")
        End If
        If txtPOR = "" Then
            .setPOR = 0
        Else
            .setPOR = Replace(txtPOR, ",", ".")
        End If
        .setPASS = chkResult.Value
        .Insertar
    End With
    
    Dim oPRH As New clsPlasma_resultados_historico
    oPRH.generar_dureza PK
    Set oPRH = Nothing
    
    Me.MousePointer = 0
    MsgBox "Datos almacenados correctamente.", vbInformation, App.Title
    If USUARIO.getPER_CIERRE = True Then
        Dim oMuestra As New clsMuestra
        oMuestra.CargaMuestra PK
        If oMuestra.getCERRADA = 0 Then
            If MsgBox("¿Desea cerrar la muestra?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                oMuestra.Cerrar PK
            End If
        End If
    End If
    Unload Me
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmPlasma_Dureza_Shore"
End Sub

Private Sub cmdTraccion_Click(Index As Integer)
    With frmPlasma_Traccion
        .MUESTRA_ID = PK
        .tipo = Index
        .Show 1
    End With
End Sub

Private Sub cmdVerificacion_Click()
    If listaEquipos.ListItems.Count > 0 Then
        Dim objfrm  As New frmEquipoEdicionVerificacion
        Dim oEquipo As New clsEquipos
        oEquipo.Carga listaEquipos.ListItems(listaEquipos.selectedItem.Index).Text
        Set objfrm.EQUIPO = oEquipo
        
        If listaEquipos.ListItems(listaEquipos.selectedItem.Index).SubItems(3) = 0 Then
            
            objfrm.TipoEdicion = Alta
            objfrm.idVerificadorInternoInicial = USUARIO.getID_EMPLEADO
            objfrm.FechaProximaInicial = Now
            'MANTIS-810-I
            'objfrm.IdPeriodoInicial = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO
            objfrm.IdPeriodoInicial = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO
            'MANTIS-810-F
            objfrm.IdTipoVerificacionIncial = 1
            
            'MANTIS-810-I
            'objfrm.copiarUltimaVerificacionPeriodo = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO
            objfrm.copiarUltimaVerificacionPeriodo = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO
            'MANTIS-810-F
              
            objfrm.Show vbModal
            If objfrm.ID_VERIFICACION <> 0 Then
                listaEquipos.ListItems(listaEquipos.selectedItem.Index).SubItems(3) = objfrm.ID_VERIFICACION
            End If
            grabar_equipos
        Else
            objfrm.ID = listaEquipos.ListItems(listaEquipos.selectedItem.Index).SubItems(3)
            objfrm.TipoEdicion = visualizar
            objfrm.copiarUltimaVerificacionPeriodo = 0
            objfrm.Show vbModal
        End If
        
        Unload objfrm
        Set objfrm = Nothing
        Exit Sub
    End If
End Sub
Private Sub cmdAnadirEquipo_Click()
    If cmbEquipos.getPK_SALIDA <> 0 Then
        For i = 1 To listaEquipos.ListItems.Count
            If listaEquipos.ListItems(i) = cmbEquipos.getPK_SALIDA Then
                MsgBox "El equipo ya se encuentra en la lista.", vbExclamation, App.Title
                Exit Sub
            End If
        Next
        Dim oEquipo As New clsEquipos
        oEquipo.Carga_Datos_Basicos cmbEquipos.getPK_SALIDA
        With listaEquipos.ListItems.Add(, , oEquipo.getID_EQUIPO)
            .SubItems(1) = oEquipo.getNOMBRE
            .SubItems(2) = oEquipo.getSERIE
            .SubItems(3) = "0"
        End With
        listaEquipos.ListItems(listaEquipos.ListItems.Count).EnsureVisible
        cmbEquipos.limpiar
        grabar_equipos
    End If

End Sub

Private Sub cmdAnadirReactivo_Click()
    ' Interno (I)
    If cmbReactivos.getTEXTO <> "" Then
        Dim oBote As New clsBotes_ex
        Dim oTb As New clsTipos_bote_ex
        Dim oTR As New clsTipos_reactivo_ex
        oBote.CARGAR cmbReactivos.getPK_SALIDA
        oTb.CARGAR oBote.getTIPO_BOTE_EX_ID
        oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
        With listaReactivos.ListItems.Add(, , oBote.getID_BOTE_EX)
            .SubItems(1) = oTR.getNOMBRE
            .SubItems(2) = Format(oBote.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "E"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Externo (E)
    If cmbReactivosInternos.getTEXTO <> "" Then
        Dim oRPR As New clsRpr_botes
        Dim oTRPR As New clsRPR_Tipos
        oRPR.Carga cmbReactivosInternos.getPK_SALIDA
        oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
        With listaReactivos.ListItems.Add(, , oRPR.getID_BOTE_PR)
            .SubItems(1) = oTRPR.getNOMBRE
            .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "dd-mm-yyyy")
            .SubItems(3) = "I"
        End With
        listaReactivos.ListItems(listaReactivos.ListItems.Count).EnsureVisible
    End If
    ' Limpiar Combos
    cmbReactivos.limpiar
    cmbReactivosInternos.limpiar
    grabar_reactivos
End Sub
Private Sub cmdEliminarEquipo_Click()
    If listaEquipos.ListItems.Count > 0 Then
        listaEquipos.ListItems.Remove listaEquipos.selectedItem.Index
        grabar_equipos
    End If
End Sub

Private Sub cmdEliminarReactivo_Click()
    If listaReactivos.ListItems.Count > 0 Then
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
        cmbReactivos.limpiar
        cmbReactivosInternos.limpiar
    End If
    grabar_reactivos
End Sub
Private Sub cmdSalir_Click()
'    grabar_equipos
    Dim oMuestra As New clsMuestra
    oMuestra.comprobar_cierre (PK)
    Unload Me
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    permisos
    If PK > 0 Then
        cargar_muestra
    End If
End Sub
Private Sub permisos()
    ' Permiso para modificar la vida
    Dim op As New clsParametros
    Dim s() As String
    Dim i As Integer
    op.Carga parametros.PARAM_USUARIOS_MODIFICAN_EQUIPOS_MUESTRA_CERRADA, ""
    If op.getVALOR <> "" Then
        s = Split(op.getVALOR, ",")
        For i = LBound(s) To UBound(s)
            If USUARIO.getID_EMPLEADO = CInt(s(i)) Then
                chkModificar.Value = Checked
                Exit For
            End If
        Next
    End If
    Set op = Nothing
End Sub

Private Sub cabecera()
    With listaEquipos.ColumnHeaders
        .Add , , "NºEquipo", 800, lvwColumnLeft
        .Add , , "Nombre", 3200, lvwColumnLeft
        .Add , , "NºSerie", 1200, lvwColumnCenter
        .Add , , "Verificación", 1, lvwColumnCenter
    End With
    With listaReactivos.ColumnHeaders
        .Add , , "Número", 800, lvwColumnLeft
        .Add , , "Reactivo", 3200, lvwColumnLeft
        .Add , , "Caducidad", 1200, lvwColumnCenter
        .Add , , "TIPO", 0, lvwColumnCenter ' (I-E) Interno o externo
    End With
End Sub
Private Sub cargar_muestra()
    'Titulo
    Dim oMuestra As New clsMuestra
   On Error GoTo cargar_muestra_Error

    oMuestra.CargaMuestra (PK)
    ' Duplicada
    If oMuestra.getANALISIS_DUPLICADO = 1 Then
        chkDuplicada.Value = Checked
    End If
    lbltitulo = "Registro resultados DUREZA SHORE A : " & Trim(str(oMuestra.getID_GENERAL)) & " (" & oMuestra.CodigoParticular(PK) & ")"
    Me.Caption = lbltitulo
    'Equipos
    cargar_equipos PK
    cargar_reactivos PK
    ' Cargar datos de recepción
    Dim oPlasmaRecepcion As New clsPlasma_recepcion
    With oPlasmaRecepcion
        If .Carga(PK) Then
            cmbProcess.MostrarElemento .getPROCESO_ID
            cmbCustomer.MostrarElemento .getCUSTOMER_ID
            cmbnatype.MostrarElemento .getNTYPE
            txtDatos(50) = .getSPECIMEN_ID
            txtDatos(51) = .getPN
            txtDatos(52) = .getPRODUCT_TYPE
            txtDatos(53) = .getMODULE_SN
            txtDatos(54) = .getSN
            txtDatos(55) = .getPRODUCT_SN
            chkResult = .getRESULT
        End If
    End With
    ' Resultados
    Dim opd As New clsPlasma_dureza
    If opd.Carga(PK) = True Then
        txtDurezaResults = opd.getRESULT
        txtDurezaReq = opd.getREQUIREMENT
        txtDurezaAverage = opd.getAVERAGE
        txtDurezaDimension = opd.getDIMENSION
        txtDurezaEspesor = opd.getESPESOR
        txtSD = opd.getSD
        txtPOR = opd.getPOR
    End If
    If txtDurezaReq = "" Then
        Dim oPP As New clsPlasma_procesos
        Dim oPF As New clsPlasma_ficha
        oPP.Carga oPlasmaRecepcion.getPROCESO_ID
        oPF.Carga oPP.getBOND_COAT_FICHA_ID
        
        txtDurezaReq = oPF.getSHOREA_REQ
    End If
    If txtDurezaDimension = "" Then
        txtDurezaDimension = "ATTACHED REPAIR ORDER"
    End If
    
    Set oPlasmaRecepcion = Nothing
    Set opd = Nothing
    proteger_campos oMuestra.getCERRADA

   On Error GoTo 0
   Exit Sub

cargar_muestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestra of Formulario frmPlasma_Dureza_Shore"
End Sub

Private Sub listaEquipos_DblClick()
    If listaEquipos.ListItems.Count > 0 Then
        frmEquipoEdicion.PK = CLng(listaEquipos.ListItems(listaEquipos.selectedItem.Index).Text)
        frmEquipoEdicion.Show 1
    End If
End Sub

Private Sub listaEquipos_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    grabar_equipos
End Sub

Private Sub calcularDesviacion()
    Dim total As Single
    Dim CANTIDAD As Integer
    Dim sumatorio As Single
    Dim medida As Single
    Dim numero_medidas As Integer
    Dim resultado As Single

    media = 0
    sumatorio = 0
    numero_medidas = 0
    
    lista = Split(txtDurezaResults, "-")
    If UBound(lista) < 2 Then
'        txtSD(CAMPO + 1) = ""
        Exit Sub
    End If
    ' MEDIA
    For i = LBound(lista) To UBound(lista)
        If IsNumeric(lista(i)) Then
            total = total + lista(i)
            CANTIDAD = CANTIDAD + 1
        End If
    Next
    If CANTIDAD > 0 Then
        media = CInt(total / CANTIDAD)
    End If
    ' DESVIACION
    For i = LBound(lista) To UBound(lista)
        If IsNumeric(lista(i)) Then
            medida = lista(i)
            sumatorio = sumatorio + ((medida - media) * (medida - media))
            numero_medidas = numero_medidas + 1
        End If
    Next
    txtSD = formatear(Sqr(sumatorio / (numero_medidas - 1)), 5, 1)

   On Error GoTo 0
   Exit Sub

calcularDesviacion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularDesviacion of Formulario frmPlasma_Dureza_Shore"
End Sub

Private Sub calcularPorcentaje()
    Dim total As Single
    Dim CANTIDAD As Integer
    Dim sumatorio As Single
    Dim medida As Single
    Dim numero_medidas As Integer
    Dim resultado As Single

   On Error GoTo calcularDesviacion_Error

    media = 0
    sumatorio = 0
    numero_medidas = 0
    
    lista = Split(txtDurezaResults, "-")
    If UBound(lista) < 2 Then
'        txtSD(CAMPO + 1) = ""
        Exit Sub
    End If
    ' MEDIA
    For i = LBound(lista) To UBound(lista)
        If IsNumeric(lista(i)) Then
            total = total + lista(i)
            CANTIDAD = CANTIDAD + 1
        End If
    Next
    If CANTIDAD > 0 Then
        media = total / CANTIDAD
    End If
    Dim mayor As Single
    Dim menor As Single
    mayor = 0
    menor = 9999999
    For i = LBound(lista) To UBound(lista)
        If IsNumeric(lista(i)) Then
            If lista(i) > mayor Then
                mayor = lista(i)
            End If
            If lista(i) < menor Then
                menor = lista(i)
            End If
        End If
    Next
    ' PORCENTAJE
    Dim porcentaje As Single
    porcentaje = ((mayor - menor) / media) * 100
    txtPOR = formatear(CStr(porcentaje), 3, 2)

   On Error GoTo 0
   Exit Sub

calcularDesviacion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularDesviacion of Formulario frmPlasma_Dureza_Shore"
End Sub
Private Function calcularMedia(resultados As String) As String
   On Error GoTo calcularMedia_Error

    If resultados <> "" And resultados <> "N/A" Then
        Dim lista() As String
        Dim resultado As String
        Dim total As Single
        Dim CANTIDAD As Integer
        resultado = ""
        CANTIDAD = 0
        lista = Split(resultados, "-")
        For i = LBound(lista) To UBound(lista)
            If IsNumeric(lista(i)) Then
                total = total + lista(i)
                CANTIDAD = CANTIDAD + 1
            End If
        Next
        calcularMedia = CInt(total / CANTIDAD)
    Else
        calcularMedia = ""
    End If

   On Error GoTo 0
   Exit Function

calcularMedia_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularMedia of Formulario frmPlasma_Dureza_Shore"
End Function

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = vbYellow
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 8 Or Index = 9 Then
        If Trim(txtDatos(Index)) <> "" Then
            If Right(txtDatos(Index), 2) <> "ºC" Then
                txtDatos(Index) = txtDatos(Index) & " ºC"
            End If
        End If
    End If
End Sub

Private Sub txtvalor_GotFocus()
    txtValor.BackColor = vbYellow
    txtValor.SelStart = 0
    txtValor.SelLength = Len(txtValor)
End Sub
Private Sub txtvalor_LostFocus()
    txtValor.BackColor = vbWhite
End Sub
Private Sub proteger_campos(CERRADA As Integer)
    If (CERRADA = 1 Or CERRADA = 3) And chkModificar.Value = Unchecked Then
        cmdEliminarReactivo.Enabled = False
        cmdAnadirReactivo.Enabled = False
        cmdEliminarEquipo.Enabled = False
        cmdAnadirEquipo.Enabled = False
        cmbEquipos.desactivar
        cmbReactivos.desactivar
        cmbReactivosInternos.desactivar
        cmbProcess.desactivar
        cmbCustomer.desactivar
        txtDatos(50).Enabled = False
        txtDatos(51).Enabled = False
        txtDatos(52).Enabled = False
        txtDatos(53).Enabled = False
        txtDatos(54).Enabled = False
        txtDatos(55).Enabled = False
        txtDurezaResults.Enabled = False
        chkResult.Enabled = False
        cmdok.visible = False
    Else
        cmdEliminarReactivo.Enabled = True
        cmdAnadirReactivo.Enabled = True
        cmdEliminarEquipo.Enabled = True
        cmdAnadirEquipo.Enabled = True
        cmbEquipos.activar
        cmbReactivos.activar
        cmbReactivosInternos.activar
        cmbProcess.activar
        cmbCustomer.activar
        txtDatos(50).Enabled = True
        txtDatos(51).Enabled = True
        txtDatos(52).Enabled = True
        txtDatos(53).Enabled = True
        txtDatos(54).Enabled = True
        txtDatos(55).Enabled = True
        txtDurezaResults.Enabled = True
        chkResult.Enabled = True
        cmdok.visible = True
    End If
    Select Case CERRADA
        Case 0
            lblCerrada = "ABIERTA"
        Case 1
            lblCerrada = "CERRADA"
            lblCerrada.BackColor = vbRed
        Case 2
            lblCerrada = "PTE. CIERRE"
        Case 3
            lblCerrada = "C.SIN INFORME"
    End Select
End Sub

Private Sub cargar_equipos(muestra As Long)
    Dim oPE As New clsPlasma_equipos
    Dim rs As ADODB.Recordset
    Set rs = oPE.Listado(muestra)
    listaEquipos.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With listaEquipos.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(5) ' VERIFICACION
            End With
            If rs("EN_INFORME") = 1 Then
                listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
            Else
                listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = False
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oPE = Nothing
    
End Sub

Private Sub cargar_reactivos(muestra As Long)
    Dim oPR As New clsPlasma_Reactivos
    Dim oReactivo As New clsBotes_ex
    Dim oTb As New clsTipos_bote_ex
    Dim oTR As New clsTipos_reactivo_ex
    
    Dim oRPR As New clsRpr_botes
    Dim oTRPR As New clsRPR_Tipos
    Dim rs As ADODB.Recordset
    Set rs = oPR.Listado(muestra)
    If rs.RecordCount > 0 Then
        Do
            If rs(1) = "E" Then
               oReactivo.CARGAR CLng(rs(0))
               oTb.CARGAR oReactivo.getTIPO_BOTE_EX_ID
               oTR.CARGAR oTb.getTIPO_REACTIVO_EX_ID
               With listaReactivos.ListItems.Add(, , rs(0))
                  .SubItems(1) = oTR.getNOMBRE
                  .SubItems(2) = Format(oReactivo.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                  .SubItems(3) = "E"
               End With
            Else
                oRPR.Carga CLng(rs(0))
                oTRPR.CARGAR oRPR.getTIPO_REACTIVO_PR_ID
                With listaReactivos.ListItems.Add(, , rs(0))
                    .SubItems(1) = oTRPR.getCODIGO & "-" & Format(oRPR.getNUMERO, "000") & " " & oTRPR.getNOMBRE
                    .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                    .SubItems(3) = "I"
                End With
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub

Private Sub cargar_combos()
    llenar_combo cmbProcess, New clsPlasma_procesos, 0, frmPlasma_Procesos_Detalle, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbCustomer, DECODIFICADORA.DECODIFICADORA_PLASMA_CLIENTES_INTERNOS
    oDeco.cargar_mi_combo cmbnatype, DECODIFICADORA.DECODIFICADORA_PLASMA_NUMBER_AND_TYPE
    
    llenar_combo cmbEquipos, New clsEquipos, 0, frmEquipoEdicion, ""
    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, " AND ABIERTO = 1 AND FINALIZADO = 0 "
    llenar_combo cmbReactivosInternos, New clsRpr_botes, 0, Me, " AND ISNULL(fecha_fin)"
    
End Sub
Private Sub grabar_equipos()
    Dim Equipos As String
    Dim oPE As New clsPlasma_equipos
    oPE.Eliminar PK
    Dim i As Integer
    For i = 1 To listaEquipos.ListItems.Count
        Equipos = Equipos & listaEquipos.ListItems(i).Text & ";"
        With oPE
            .setMUESTRA_ID = PK
            .setORDEN = i
            .setEQUIPO_ID = listaEquipos.ListItems(i).Text
            .setVERIFICACION_ID = listaEquipos.ListItems(i).SubItems(3)
            .setEN_INFORME = Abs(listaEquipos.ListItems(i).Checked)
            .Insertar
        End With
    Next
    ' Usos de los equipos
    Dim oEU As New clsEq_usos
    oEU.Eliminar PK, 0
    For i = 1 To listaEquipos.ListItems.Count
      With oEU
          .setMUESTRA_ID = PK
          .setEQUIPO_ID = listaEquipos.ListItems(i).Text
          .setDETERMINACION_ID = 0
          .setUSOS = 1
          .Insertar
      End With
    Next
    Set oEU = Nothing
End Sub
Private Sub grabar_reactivos()
    Dim oPR As New clsPlasma_Reactivos
    oPR.Eliminar PK
    Dim i As Integer
    For i = 1 To listaReactivos.ListItems.Count
        With oPR
            .setMUESTRA_ID = PK
            .setBOTE_EX_ID = listaReactivos.ListItems(i).Text
            .setTIPO = listaReactivos.ListItems(i).SubItems(3)
            .setORDEN = i
            .Insertar
        End With
    Next
    Set oPR = Nothing
End Sub


Private Sub txtDurezaResults_Change()
    txtDurezaAverageR = calcularMedia(txtDurezaResults)
    txtDurezaAverage = txtDurezaAverageR & " " & txtUnidades
    calcularDesviacion
    calcularPorcentaje
End Sub
Private Sub txtVickersA_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 13
       SendKeys "{Tab}", True
    End Select
End Sub

Private Sub txtVickersB_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 13
       SendKeys "{Tab}", True
    End Select

End Sub
