VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{77DDF307-D82B-4757-8B3A-106EC9D75D4B}#8.0#0"; "tdbg8.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmPlasma_Resultados 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Registro de Resultados Muestra de Plasma"
   ClientHeight    =   11370
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14040
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPlasma_Resultados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11370
   ScaleWidth      =   14040
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkAbierta 
      Caption         =   "Muestra Abierta"
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
      Left            =   9270
      TabIndex        =   203
      Top             =   10530
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.TextBox fichaTop 
      Height          =   330
      Left            =   8100
      TabIndex        =   185
      Top             =   10935
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.TextBox fichaBond 
      Height          =   330
      Left            =   6840
      TabIndex        =   184
      Top             =   10935
      Visible         =   0   'False
      Width           =   915
   End
   Begin Geslab.ControlPanelXP cpReactivos 
      Height          =   3975
      Left            =   45
      TabIndex        =   22
      Top             =   2700
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   7011
      Caption         =   "Reactivos Utilizados en la Muestra"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   3975
      Begin VB.Frame frmReactivos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Reactivos"
         ForeColor       =   &H80000008&
         Height          =   3480
         Left            =   45
         TabIndex        =   23
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
            TabIndex        =   25
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
            TabIndex        =   24
            Tag             =   "Elimina el campo seleccionado"
            Top             =   450
            Width           =   915
         End
         Begin MSComctlLib.ListView listaReactivos 
            Height          =   2460
            Left            =   45
            TabIndex        =   26
            Top             =   90
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
            TabIndex        =   27
            Top             =   2700
            Width           =   5280
            _ExtentX        =   9313
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbReactivosInternos 
            Height          =   330
            Left            =   765
            TabIndex        =   28
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
            TabIndex        =   30
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
            TabIndex        =   29
            Top             =   3105
            Width           =   495
         End
      End
   End
   Begin Geslab.ControlPanelXP ControlPanelXP2 
      Height          =   3840
      Left            =   6885
      TabIndex        =   15
      Top             =   2700
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   6773
      Caption         =   "Equipos Utilizados en la Muestra"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   3840
      Begin VB.Frame frmEquipos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3390
         Left            =   90
         TabIndex        =   16
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
            TabIndex        =   32
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
            TabIndex        =   18
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
            TabIndex        =   17
            Tag             =   "Elimina el campo seleccionado"
            Top             =   270
            Width           =   915
         End
         Begin MSComctlLib.ListView listaEquipos 
            Height          =   2325
            Left            =   0
            TabIndex        =   19
            Top             =   270
            Width           =   5580
            _ExtentX        =   9843
            _ExtentY        =   4101
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
            TabIndex        =   20
            Top             =   2700
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
            TabIndex        =   21
            Top             =   45
            Width           =   4335
         End
      End
   End
   Begin XtremeSuiteControls.Resizer Resizer1 
      Height          =   7290
      Left            =   45
      TabIndex        =   47
      Top             =   3195
      Width           =   13920
      _Version        =   851970
      _ExtentX        =   24553
      _ExtentY        =   12859
      _StockProps     =   1
      VScrollLargeChange=   500
      VScrollSmallChange=   100
      VScrollMaximum  =   12700
      ClientMinHeight =   8000
      Begin VB.Frame frmMicroTemp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "MICRO HARDNESS TEMPERATURE"
         ForeColor       =   &H000000FF&
         Height          =   960
         Left            =   6840
         TabIndex        =   181
         Top             =   11534
         Width           =   3840
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
            Index           =   10
            Left            =   1935
            MaxLength       =   255
            TabIndex        =   62
            Top             =   270
            Width           =   1770
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
            Index           =   11
            Left            =   1935
            MaxLength       =   255
            TabIndex        =   63
            Top             =   585
            Width           =   1770
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "30 MIN. BEFORE TEST"
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
            Index           =   51
            Left            =   90
            TabIndex        =   183
            Top             =   315
            Width           =   1725
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "TEMPERATURE TEST:"
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
            Index           =   50
            Left            =   90
            TabIndex        =   182
            Top             =   630
            Width           =   1740
         End
      End
      Begin VB.Frame frmMacroTemp 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "MACRO HARDNESS TEMPERATURE"
         ForeColor       =   &H000000FF&
         Height          =   960
         Left            =   2970
         TabIndex        =   178
         Top             =   11534
         Width           =   3840
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
            Index           =   9
            Left            =   1935
            MaxLength       =   255
            TabIndex        =   61
            Top             =   585
            Width           =   1770
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
            Index           =   8
            Left            =   1935
            MaxLength       =   255
            TabIndex        =   60
            Top             =   270
            Width           =   1770
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "TEMPERATURE TEST:"
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
            Index           =   53
            Left            =   90
            TabIndex        =   180
            Top             =   630
            Width           =   1740
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "30 MIN. BEFORE TEST"
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
            Index           =   52
            Left            =   90
            TabIndex        =   179
            Top             =   315
            Width           =   1725
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   10650
         Left            =   6840
         TabIndex        =   98
         Top             =   765
         Width           =   6765
         Begin VB.Frame frmTopMicro 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "MICRO HARDNESS"
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   90
            TabIndex        =   127
            Top             =   7560
            Width           =   6540
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   19
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   199
               Top             =   585
               Width           =   5280
            End
            Begin VB.TextBox txtSD 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   47
               Left            =   3555
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   167
               Top             =   900
               Width           =   1185
            End
            Begin VB.TextBox txtPOR 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   47
               Left            =   5220
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   147
               Top             =   900
               Width           =   1185
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
               Index           =   46
               Left            =   1125
               MaxLength       =   255
               TabIndex        =   58
               Top             =   270
               Width           =   5280
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   47
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   130
               Top             =   900
               Width           =   1905
            End
            Begin VB.CheckBox chkTopMicro 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PASS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   240
               Left            =   5355
               TabIndex        =   129
               Top             =   0
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox chkTopMicroNR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "N.R."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   4275
               TabIndex        =   128
               Top             =   0
               Width           =   825
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "REQMNT:"
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
               Index           =   61
               Left            =   90
               TabIndex        =   200
               Top             =   630
               Width           =   750
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "S.D.:"
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
               Index           =   44
               Left            =   3105
               TabIndex        =   168
               Top             =   945
               Width           =   360
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "% :"
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
               Index           =   33
               Left            =   4905
               TabIndex        =   148
               Top             =   945
               Width           =   210
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "RESULTS:"
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
               Index           =   18
               Left            =   90
               TabIndex        =   132
               Top             =   315
               Width           =   795
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "AVERAGE:"
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
               Index           =   20
               Left            =   90
               TabIndex        =   131
               Top             =   945
               Width           =   810
            End
         End
         Begin VB.Frame frmTopMacro 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "MACRO HARDNESS"
            ForeColor       =   &H80000008&
            Height          =   1905
            Left            =   90
            TabIndex        =   121
            Top             =   5580
            Width           =   6540
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   18
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   197
               Top             =   585
               Width           =   5280
            End
            Begin VB.TextBox txtPOR 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   45
               Left            =   5220
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   165
               Top             =   1530
               Width           =   1185
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
               Index           =   7
               Left            =   1125
               MaxLength       =   255
               TabIndex        =   56
               Top             =   900
               Width           =   5280
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
               Index           =   6
               Left            =   1125
               MaxLength       =   255
               TabIndex        =   57
               Top             =   1215
               Width           =   5280
            End
            Begin VB.TextBox txtSD 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   45
               Left            =   3555
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   145
               Top             =   1530
               Width           =   1185
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
               Index           =   44
               Left            =   1125
               MaxLength       =   255
               TabIndex        =   55
               Top             =   270
               Width           =   5280
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   45
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   124
               Top             =   1530
               Width           =   1905
            End
            Begin VB.CheckBox chkTopMacro 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PASS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   240
               Left            =   5355
               TabIndex        =   123
               Top             =   0
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox chkTopMacroNR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "N.R."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   4275
               TabIndex        =   122
               Top             =   0
               Width           =   825
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "REQMNT:"
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
               Index           =   60
               Left            =   90
               TabIndex        =   198
               Top             =   630
               Width           =   750
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "% :"
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
               Index           =   43
               Left            =   4905
               TabIndex        =   166
               Top             =   1575
               Width           =   210
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "DIMENSION:"
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
               Index           =   42
               Left            =   90
               TabIndex        =   164
               Top             =   945
               Width           =   960
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "THICKNESS:"
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
               Index           =   41
               Left            =   90
               TabIndex        =   163
               Top             =   1260
               Width           =   960
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "S.D.:"
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
               Index           =   32
               Left            =   3105
               TabIndex        =   146
               Top             =   1575
               Width           =   360
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "RESULTS:"
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
               Index           =   21
               Left            =   90
               TabIndex        =   126
               Top             =   315
               Width           =   795
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "AVERAGE:"
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
               Index           =   22
               Left            =   90
               TabIndex        =   125
               Top             =   1575
               Width           =   810
            End
         End
         Begin VB.Frame frmTopTraccion 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "TENSILE STRENGTH"
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   90
            TabIndex        =   114
            Top             =   4050
            Width           =   6540
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   17
               Left            =   945
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   195
               Top             =   675
               Width           =   4020
            End
            Begin VB.CommandButton cmdTraccion 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Results"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1005
               Index           =   2
               Left            =   5040
               Style           =   1  'Graphical
               TabIndex        =   150
               Top             =   315
               Width           =   1365
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   43
               Left            =   945
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   118
               Top             =   990
               Width           =   4020
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   42
               Left            =   945
               MaxLength       =   255
               TabIndex        =   117
               Top             =   360
               Width           =   4020
            End
            Begin VB.CheckBox chkTopTraccion 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PASS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   240
               Left            =   5355
               TabIndex        =   116
               Top             =   0
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox chkTopTraccionNR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "N.R."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   4275
               TabIndex        =   115
               Top             =   0
               Width           =   825
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "REQMNT:"
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
               Index           =   59
               Left            =   90
               TabIndex        =   196
               Top             =   720
               Width           =   750
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "AVERAGE:"
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
               Index           =   23
               Left            =   90
               TabIndex        =   120
               Top             =   1035
               Width           =   810
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "RESULTS:"
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
               Index           =   24
               Left            =   90
               TabIndex        =   119
               Top             =   405
               Width           =   795
            End
         End
         Begin VB.Frame frmTopMicroEstructura 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "METALLOGRAPHIC EXAMINATION"
            ForeColor       =   &H80000008&
            Height          =   2580
            Left            =   90
            TabIndex        =   111
            Top             =   1440
            Width           =   6540
            Begin VB.CheckBox chkTopMetalografiaNR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "N.R."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   5310
               TabIndex        =   112
               Top             =   0
               Visible         =   0   'False
               Width           =   825
            End
            Begin TrueDBGrid80.TDBGrid gridTOP 
               Height          =   2250
               Left            =   90
               TabIndex        =   113
               Top             =   225
               Width           =   6315
               _ExtentX        =   11139
               _ExtentY        =   3969
               _LayoutType     =   4
               _RowHeight      =   16
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "ID_ENSAYO"
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "TEST"
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "RESULT (%)"
               Columns(2).DataField=   ""
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   4
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "PASS"
               Columns(3).DataField=   ""
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   4
               Splits(0)._UserFlags=   0
               Splits(0).ExtendRightColumn=   -1  'True
               Splits(0).MarqueeStyle=   1
               Splits(0).AllowRowSizing=   0   'False
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0).AllowColSelect=   0   'False
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=4"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=6800"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6720"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
               Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8196"
               Splits(0)._ColumnProps(13)=   "Column(1).AllowFocus=0"
               Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(15)=   "Column(2).Width=2699"
               Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2619"
               Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=1"
               Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(21)=   "Column(3).Width=212"
               Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=132"
               Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
               Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=1"
               Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               DataMode        =   4
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               TabAction       =   2
               WrapCellPointer =   -1  'True
               MultipleLines   =   0
               EmptyRows       =   -1  'True
               CellTipsWidth   =   0
               MultiSelect     =   2
               DeadAreaBackColor=   12632256
               RowDividerColor =   12632256
               RowSubDividerColor=   12632256
               DirectionAfterEnter=   2
               DirectionAfterTab=   1
               MaxRows         =   250000
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HDEEDFA&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=41"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=38,.bgcolor=&H8080FF&,.fgcolor=&H0&"
               _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=39,.bgcolor=&H8000000A&,.bold=0"
               _StyleDefs(14)  =   ":id=3,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
               _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFFFFF&,.fgcolor=&HFFFFFF&"
               _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=42"
               _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=43"
               _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=44"
               _StyleDefs(22)  =   "RecordSelectorStyle:id=45,.parent=2,.namedParent=47"
               _StyleDefs(23)  =   "FilterBarStyle:id=48,.parent=1,.namedParent=50"
               _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
               _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
               _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.namedParent=38"
               _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
               _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
               _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6,.namedParent=40"
               _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7,.namedParent=40"
               _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
               _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9,.namedParent=43"
               _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10,.namedParent=44"
               _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=46,.parent=45"
               _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=49,.parent=48"
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=11,.bgcolor=&HD7D7D7&,.locked=-1"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=11,.alignment=2"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
               _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=11,.alignment=2"
               _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=12"
               _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=13"
               _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=15"
               _StyleDefs(52)  =   "Named:id=37:Normal"
               _StyleDefs(53)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
               _StyleDefs(54)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(55)  =   ":id=37,.fontname=MS Sans Serif"
               _StyleDefs(56)  =   "Named:id=38:Heading"
               _StyleDefs(57)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
               _StyleDefs(58)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
               _StyleDefs(59)  =   ":id=38,.strikethrough=0,.charset=0"
               _StyleDefs(60)  =   ":id=38,.fontname=MS Sans Serif"
               _StyleDefs(61)  =   "Named:id=39:Footing"
               _StyleDefs(62)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(63)  =   "Named:id=40:Selected"
               _StyleDefs(64)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
               _StyleDefs(65)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(66)  =   ":id=40,.fontname=MS Sans Serif"
               _StyleDefs(67)  =   "Named:id=41:Caption"
               _StyleDefs(68)  =   ":id=41,.parent=38,.alignment=2"
               _StyleDefs(69)  =   "Named:id=42:HighlightRow"
               _StyleDefs(70)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
               _StyleDefs(71)  =   "Named:id=43:EvenRow"
               _StyleDefs(72)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
               _StyleDefs(73)  =   "Named:id=44:OddRow"
               _StyleDefs(74)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
               _StyleDefs(75)  =   "Named:id=47:RecordSelector"
               _StyleDefs(76)  =   ":id=47,.parent=38"
               _StyleDefs(77)  =   "Named:id=50:FilterBar"
               _StyleDefs(78)  =   ":id=50,.parent=37"
            End
         End
         Begin VB.Frame Frame14 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "TESTING SPECIMEN"
            ForeColor       =   &H80000008&
            Height          =   915
            Left            =   90
            TabIndex        =   106
            Top             =   495
            Width           =   6540
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
               Index           =   41
               Left            =   1710
               MaxLength       =   255
               TabIndex        =   108
               Top             =   540
               Width           =   4695
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
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
               Index           =   40
               Left            =   1710
               MaxLength       =   255
               TabIndex        =   107
               Top             =   225
               Width           =   4695
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "BATCH TOP COAT:"
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
               Index           =   25
               Left            =   90
               TabIndex        =   110
               Top             =   585
               Width           =   1440
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "TOP COAT:"
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
               Index           =   26
               Left            =   90
               TabIndex        =   109
               Top             =   270
               Width           =   855
            End
         End
         Begin VB.Frame frmTopEspesor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "THICKNESS"
            ForeColor       =   &H80000008&
            Height          =   1680
            Left            =   90
            TabIndex        =   99
            Top             =   8865
            Width           =   6540
            Begin VB.TextBox txtUnidadEspesorTop 
               Height          =   285
               Left            =   810
               TabIndex        =   206
               Top             =   1305
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   20
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   201
               Top             =   630
               Width           =   5280
            End
            Begin VB.TextBox txtPOR 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   49
               Left            =   3555
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   176
               Top             =   1260
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   49
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   170
               Top             =   945
               Width           =   1905
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
               Index           =   48
               Left            =   1125
               MaxLength       =   255
               TabIndex        =   59
               Top             =   315
               Width           =   5280
            End
            Begin VB.TextBox txtSD 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   49
               Left            =   3555
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   169
               Top             =   945
               Width           =   1185
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   3
               Left            =   5265
               MaxLength       =   255
               TabIndex        =   103
               Top             =   1260
               Width           =   1140
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   2
               Left            =   5265
               MaxLength       =   255
               TabIndex        =   102
               Top             =   945
               Width           =   1140
            End
            Begin VB.CheckBox chkTopEspesor 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PASS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   240
               Left            =   5355
               TabIndex        =   101
               Top             =   0
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox chkTopEspesorNR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "N.R."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   4275
               TabIndex        =   100
               Top             =   0
               Width           =   825
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "REQMNT:"
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
               Index           =   62
               Left            =   90
               TabIndex        =   202
               Top             =   675
               Width           =   750
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "% :"
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
               Index           =   49
               Left            =   3240
               TabIndex        =   177
               Top             =   1305
               Visible         =   0   'False
               Width           =   210
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "AVERAGE:"
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
               Index           =   47
               Left            =   90
               TabIndex        =   173
               Top             =   990
               Width           =   810
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "RESULTS:"
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
               Index           =   46
               Left            =   90
               TabIndex        =   172
               Top             =   360
               Width           =   795
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "S.D.:"
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
               Index           =   45
               Left            =   3105
               TabIndex        =   171
               Top             =   990
               Width           =   360
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Máx:"
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
               Index           =   28
               Left            =   4860
               TabIndex        =   105
               Top             =   1305
               Width           =   345
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Min:"
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
               Index           =   29
               Left            =   4860
               TabIndex        =   104
               Top             =   990
               Width           =   300
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H0080C0FF&
            Caption         =   "TOP COAT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   45
            TabIndex        =   133
            Top             =   135
            Width           =   6675
         End
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   10650
         Left            =   45
         TabIndex        =   48
         Top             =   765
         Width           =   6765
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "TESTING SPECIMEN"
            ForeColor       =   &H80000008&
            Height          =   915
            Left            =   90
            TabIndex        =   92
            Top             =   495
            Width           =   6540
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
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
               Index           =   30
               Left            =   1710
               MaxLength       =   255
               TabIndex        =   94
               Top             =   225
               Width           =   4695
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
               Index           =   31
               Left            =   1710
               MaxLength       =   255
               TabIndex        =   93
               Top             =   540
               Width           =   4695
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "BOND COAT:"
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
               Index           =   0
               Left            =   90
               TabIndex        =   96
               Top             =   270
               Width           =   990
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "BATCH BOND COAT:"
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
               Index           =   4
               Left            =   90
               TabIndex        =   95
               Top             =   585
               Width           =   1575
            End
         End
         Begin VB.Frame frmBondMicroEstructura 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "METALLOGRAPHIC EXAMINATION"
            ForeColor       =   &H80000008&
            Height          =   2580
            Left            =   90
            TabIndex        =   89
            Top             =   1440
            Width           =   6540
            Begin VB.CheckBox chkBondMetalografiaNR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "N.R."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   5310
               TabIndex        =   90
               Top             =   0
               Visible         =   0   'False
               Width           =   825
            End
            Begin TrueDBGrid80.TDBGrid gridBOND 
               Height          =   2250
               Left            =   90
               TabIndex        =   91
               Top             =   225
               Width           =   6315
               _ExtentX        =   11139
               _ExtentY        =   3969
               _LayoutType     =   4
               _RowHeight      =   16
               _WasPersistedAsPixels=   0
               Columns(0)._VlistStyle=   0
               Columns(0)._MaxComboItems=   5
               Columns(0).Caption=   "ID_ENSAYO"
               Columns(0).DataField=   ""
               Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(1)._VlistStyle=   0
               Columns(1)._MaxComboItems=   5
               Columns(1).Caption=   "TEST"
               Columns(1).DataField=   ""
               Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(2)._VlistStyle=   0
               Columns(2)._MaxComboItems=   5
               Columns(2).Caption=   "RESULT (%)"
               Columns(2).DataField=   ""
               Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns(3)._VlistStyle=   4
               Columns(3)._MaxComboItems=   5
               Columns(3).Caption=   "PASS"
               Columns(3).DataField=   ""
               Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
               Columns.Count   =   4
               Splits(0)._UserFlags=   0
               Splits(0).ExtendRightColumn=   -1  'True
               Splits(0).MarqueeStyle=   1
               Splits(0).AllowRowSizing=   0   'False
               Splits(0).RecordSelectors=   0   'False
               Splits(0).RecordSelectorWidth=   503
               Splits(0).AllowColSelect=   0   'False
               Splits(0).DividerColor=   12632256
               Splits(0).SpringMode=   0   'False
               Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
               Splits(0)._ColumnProps(0)=   "Columns.Count=4"
               Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
               Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
               Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
               Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
               Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
               Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
               Splits(0)._ColumnProps(7)=   "Column(1).Width=6800"
               Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
               Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=6720"
               Splits(0)._ColumnProps(10)=   "Column(1)._EditAlways=0"
               Splits(0)._ColumnProps(11)=   "Column(1).AllowSizing=0"
               Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=8196"
               Splits(0)._ColumnProps(13)=   "Column(1).AllowFocus=0"
               Splits(0)._ColumnProps(14)=   "Column(1).Order=2"
               Splits(0)._ColumnProps(15)=   "Column(2).Width=2699"
               Splits(0)._ColumnProps(16)=   "Column(2).DividerColor=0"
               Splits(0)._ColumnProps(17)=   "Column(2)._WidthInPix=2619"
               Splits(0)._ColumnProps(18)=   "Column(2)._EditAlways=0"
               Splits(0)._ColumnProps(19)=   "Column(2)._ColStyle=1"
               Splits(0)._ColumnProps(20)=   "Column(2).Order=3"
               Splits(0)._ColumnProps(21)=   "Column(3).Width=212"
               Splits(0)._ColumnProps(22)=   "Column(3).DividerColor=0"
               Splits(0)._ColumnProps(23)=   "Column(3)._WidthInPix=132"
               Splits(0)._ColumnProps(24)=   "Column(3)._EditAlways=0"
               Splits(0)._ColumnProps(25)=   "Column(3)._ColStyle=1"
               Splits(0)._ColumnProps(26)=   "Column(3).Order=4"
               Splits.Count    =   1
               PrintInfos(0)._StateFlags=   3
               PrintInfos(0).Name=   "piInternal 0"
               PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
               PrintInfos(0).PageHeaderHeight=   0
               PrintInfos(0).PageFooterHeight=   0
               PrintInfos.Count=   1
               DataMode        =   4
               DefColWidth     =   0
               HeadLines       =   1
               FootLines       =   1
               TabAction       =   2
               WrapCellPointer =   -1  'True
               MultipleLines   =   0
               EmptyRows       =   -1  'True
               CellTipsWidth   =   0
               MultiSelect     =   2
               DeadAreaBackColor=   12632256
               RowDividerColor =   12632256
               RowSubDividerColor=   12632256
               DirectionAfterEnter=   2
               DirectionAfterTab=   1
               MaxRows         =   250000
               _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
               _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=0,.valignment=0,.bgcolor=&H80000005&"
               _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
               _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
               _StyleDefs(3)   =   ":id=0,.borderColor=&H0&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
               _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
               _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=37,.bgcolor=&HDEEDFA&,.bold=0,.fontsize=825"
               _StyleDefs(7)   =   ":id=1,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
               _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=41"
               _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=38,.bgcolor=&H8080FF&,.fgcolor=&H0&"
               _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
               _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=39,.bgcolor=&H8000000A&,.bold=0"
               _StyleDefs(14)  =   ":id=3,.fontsize=975,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
               _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=40,.bgcolor=&HFFFFFF&"
               _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1,.bgcolor=&HFFFFFF&,.fgcolor=&HFFFFFF&"
               _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=42"
               _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=43"
               _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=44"
               _StyleDefs(22)  =   "RecordSelectorStyle:id=45,.parent=2,.namedParent=47"
               _StyleDefs(23)  =   "FilterBarStyle:id=48,.parent=1,.namedParent=50"
               _StyleDefs(24)  =   "Splits(0).Style:id=11,.parent=1"
               _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=20,.parent=4"
               _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=12,.parent=2,.namedParent=38"
               _StyleDefs(27)  =   "Splits(0).FooterStyle:id=13,.parent=3"
               _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=14,.parent=5"
               _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=16,.parent=6,.namedParent=40"
               _StyleDefs(30)  =   "Splits(0).EditorStyle:id=15,.parent=7,.namedParent=40"
               _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=17,.parent=8"
               _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=18,.parent=9,.namedParent=43"
               _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=19,.parent=10,.namedParent=44"
               _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=46,.parent=45"
               _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=49,.parent=48"
               _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=24,.parent=11"
               _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=21,.parent=12"
               _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=22,.parent=13"
               _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=23,.parent=15"
               _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=11,.bgcolor=&HD7D7D7&,.locked=-1"
               _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=12"
               _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=13"
               _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=15"
               _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=11,.alignment=2"
               _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=12"
               _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=13"
               _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=15"
               _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=54,.parent=11,.alignment=2"
               _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=51,.parent=12"
               _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=52,.parent=13"
               _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=53,.parent=15"
               _StyleDefs(52)  =   "Named:id=37:Normal"
               _StyleDefs(53)  =   ":id=37,.parent=0,.alignment=3,.bgcolor=&H80000018&,.bgpicMode=2,.bold=0"
               _StyleDefs(54)  =   ":id=37,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(55)  =   ":id=37,.fontname=MS Sans Serif"
               _StyleDefs(56)  =   "Named:id=38:Heading"
               _StyleDefs(57)  =   ":id=38,.parent=37,.valignment=2,.bgcolor=&H80000016&,.fgcolor=&H80000012&"
               _StyleDefs(58)  =   ":id=38,.wraptext=-1,.bold=0,.fontsize=825,.italic=0,.underline=0"
               _StyleDefs(59)  =   ":id=38,.strikethrough=0,.charset=0"
               _StyleDefs(60)  =   ":id=38,.fontname=MS Sans Serif"
               _StyleDefs(61)  =   "Named:id=39:Footing"
               _StyleDefs(62)  =   ":id=39,.parent=37,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
               _StyleDefs(63)  =   "Named:id=40:Selected"
               _StyleDefs(64)  =   ":id=40,.parent=37,.bgcolor=&H8080FF&,.fgcolor=&H0&,.bold=0,.fontsize=825"
               _StyleDefs(65)  =   ":id=40,.italic=0,.underline=0,.strikethrough=0,.charset=0"
               _StyleDefs(66)  =   ":id=40,.fontname=MS Sans Serif"
               _StyleDefs(67)  =   "Named:id=41:Caption"
               _StyleDefs(68)  =   ":id=41,.parent=38,.alignment=2"
               _StyleDefs(69)  =   "Named:id=42:HighlightRow"
               _StyleDefs(70)  =   ":id=42,.parent=37,.bgcolor=&H80000008&,.fgcolor=&H80000005&"
               _StyleDefs(71)  =   "Named:id=43:EvenRow"
               _StyleDefs(72)  =   ":id=43,.parent=37,.bgcolor=&HFFFFFF&,.wraptext=-1"
               _StyleDefs(73)  =   "Named:id=44:OddRow"
               _StyleDefs(74)  =   ":id=44,.parent=37,.bgcolor=&HD5ECF9&"
               _StyleDefs(75)  =   "Named:id=47:RecordSelector"
               _StyleDefs(76)  =   ":id=47,.parent=38"
               _StyleDefs(77)  =   "Named:id=50:FilterBar"
               _StyleDefs(78)  =   ":id=50,.parent=37"
            End
         End
         Begin VB.Frame frmBondTraccion 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "TENSILE STRENGTH"
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   90
            TabIndex        =   82
            Top             =   4050
            Width           =   6540
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   13
               Left            =   945
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   187
               Top             =   675
               Width           =   4020
            End
            Begin VB.CommandButton cmdTraccion 
               BackColor       =   &H00E0E0E0&
               Caption         =   "Results"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1005
               Index           =   1
               Left            =   5040
               Style           =   1  'Graphical
               TabIndex        =   149
               Top             =   315
               Width           =   1365
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   32
               Left            =   945
               MaxLength       =   255
               TabIndex        =   86
               Top             =   360
               Width           =   4020
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   33
               Left            =   945
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   85
               Top             =   990
               Width           =   4020
            End
            Begin VB.CheckBox chkBondTraccion 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PASS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   240
               Left            =   5355
               TabIndex        =   84
               Top             =   0
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox chkBondTraccionNR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "N.R."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   4185
               TabIndex        =   83
               Top             =   0
               Width           =   825
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "REQMNT:"
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
               Index           =   55
               Left            =   90
               TabIndex        =   188
               Top             =   720
               Width           =   750
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "RESULTS:"
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
               Index           =   7
               Left            =   90
               TabIndex        =   88
               Top             =   405
               Width           =   795
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "AVERAGE:"
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
               Index           =   6
               Left            =   90
               TabIndex        =   87
               Top             =   1035
               Width           =   810
            End
         End
         Begin VB.Frame frmBondMacro 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "MACRO HARDNESS"
            ForeColor       =   &H80000008&
            Height          =   1905
            Left            =   90
            TabIndex        =   76
            Top             =   5580
            Width           =   6540
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   14
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   189
               Top             =   585
               Width           =   5280
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
               Index           =   5
               Left            =   1125
               MaxLength       =   255
               TabIndex        =   52
               Top             =   1215
               Width           =   5280
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
               Index           =   4
               Left            =   1125
               MaxLength       =   255
               TabIndex        =   51
               Top             =   900
               Width           =   5280
            End
            Begin VB.TextBox txtPOR 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   35
               Left            =   5220
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   152
               Top             =   1530
               Width           =   1185
            End
            Begin VB.TextBox txtSD 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   35
               Left            =   3555
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   141
               Top             =   1530
               Width           =   1185
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   35
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   79
               Top             =   1530
               Width           =   1905
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
               Index           =   34
               Left            =   1125
               MaxLength       =   255
               TabIndex        =   50
               Top             =   270
               Width           =   5280
            End
            Begin VB.CheckBox chkBondMacro 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PASS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   240
               Left            =   5355
               TabIndex        =   78
               Top             =   0
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox chkBondMacroNR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "N.R."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   4185
               TabIndex        =   77
               Top             =   0
               Width           =   825
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "REQMNT:"
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
               Index           =   56
               Left            =   90
               TabIndex        =   190
               Top             =   630
               Width           =   750
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "THICKNESS:"
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
               Index           =   36
               Left            =   90
               TabIndex        =   155
               Top             =   1260
               Width           =   960
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "DIMENSION:"
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
               Index           =   35
               Left            =   90
               TabIndex        =   154
               Top             =   945
               Width           =   960
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "% :"
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
               Index           =   34
               Left            =   4905
               TabIndex        =   153
               Top             =   1575
               Width           =   210
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "S.D.:"
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
               Index           =   30
               Left            =   3105
               TabIndex        =   142
               Top             =   1575
               Width           =   360
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "AVERAGE:"
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
               Index           =   9
               Left            =   90
               TabIndex        =   81
               Top             =   1575
               Width           =   810
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "RESULTS:"
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
               Index           =   10
               Left            =   90
               TabIndex        =   80
               Top             =   315
               Width           =   795
            End
         End
         Begin VB.Frame frmBondMicro 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "MICRO HARDNESS"
            ForeColor       =   &H80000008&
            Height          =   1275
            Left            =   90
            TabIndex        =   70
            Top             =   7560
            Width           =   6540
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   15
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   191
               Top             =   585
               Width           =   5280
            End
            Begin VB.TextBox txtSD 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   37
               Left            =   3555
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   156
               Top             =   900
               Width           =   1185
            End
            Begin VB.TextBox txtPOR 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   37
               Left            =   5265
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   143
               Top             =   900
               Width           =   1140
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   37
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   73
               Top             =   900
               Width           =   1905
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
               Index           =   36
               Left            =   1125
               MaxLength       =   255
               TabIndex        =   53
               Top             =   270
               Width           =   5280
            End
            Begin VB.CheckBox chkBondMicro 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PASS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   240
               Left            =   5355
               TabIndex        =   72
               Top             =   0
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox chkBondMicroNR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "N.R."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   4185
               TabIndex        =   71
               Top             =   0
               Width           =   825
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "REQMNT:"
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
               Index           =   57
               Left            =   90
               TabIndex        =   192
               Top             =   630
               Width           =   750
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "S.D.:"
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
               Index           =   37
               Left            =   3105
               TabIndex        =   157
               Top             =   945
               Width           =   360
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "% :"
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
               Index           =   31
               Left            =   4905
               TabIndex        =   144
               Top             =   945
               Width           =   210
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "AVERAGE:"
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
               Index           =   12
               Left            =   90
               TabIndex        =   75
               Top             =   945
               Width           =   810
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "RESULTS:"
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
               Index           =   13
               Left            =   90
               TabIndex        =   74
               Top             =   315
               Width           =   795
            End
         End
         Begin VB.Frame frmBondEspesor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "THICKNESS"
            ForeColor       =   &H80000008&
            Height          =   1680
            Left            =   90
            TabIndex        =   49
            Top             =   8865
            Width           =   6540
            Begin VB.TextBox txtUnidadEspesorBond 
               Height          =   285
               Left            =   945
               TabIndex        =   205
               Top             =   1305
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   16
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   193
               Top             =   630
               Width           =   5280
            End
            Begin VB.TextBox txtPOR 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   39
               Left            =   3555
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   174
               Top             =   1260
               Visible         =   0   'False
               Width           =   1185
            End
            Begin VB.TextBox txtSD 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   39
               Left            =   3555
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   161
               Top             =   945
               Width           =   1185
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
               Index           =   38
               Left            =   1125
               MaxLength       =   255
               TabIndex        =   54
               Top             =   315
               Width           =   5280
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   39
               Left            =   1125
               Locked          =   -1  'True
               MaxLength       =   255
               TabIndex        =   158
               Top             =   945
               Width           =   1905
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   0
               Left            =   5265
               MaxLength       =   255
               TabIndex        =   67
               Top             =   945
               Width           =   1140
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Index           =   1
               Left            =   5265
               MaxLength       =   255
               TabIndex        =   66
               Top             =   1260
               Width           =   1140
            End
            Begin VB.CheckBox chkBondEspesor 
               BackColor       =   &H00C0C0C0&
               Caption         =   "PASS"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00008000&
               Height          =   240
               Left            =   5355
               TabIndex        =   65
               Top             =   0
               Value           =   1  'Checked
               Width           =   1050
            End
            Begin VB.CheckBox chkBondEspesorNR 
               BackColor       =   &H00C0C0C0&
               Caption         =   "N.R."
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FF0000&
               Height          =   240
               Left            =   4185
               TabIndex        =   64
               Top             =   0
               Width           =   825
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "REQMNT:"
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
               Index           =   58
               Left            =   90
               TabIndex        =   194
               Top             =   675
               Width           =   750
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "% :"
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
               Index           =   48
               Left            =   3240
               TabIndex        =   175
               Top             =   1305
               Visible         =   0   'False
               Width           =   210
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "S.D.:"
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
               Index           =   40
               Left            =   3105
               TabIndex        =   162
               Top             =   990
               Width           =   360
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "RESULTS:"
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
               Index           =   39
               Left            =   90
               TabIndex        =   160
               Top             =   360
               Width           =   795
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "AVERAGE:"
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
               Index           =   38
               Left            =   90
               TabIndex        =   159
               Top             =   990
               Width           =   810
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Min:"
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
               Index           =   19
               Left            =   4815
               TabIndex        =   69
               Top             =   990
               Width           =   300
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Máx:"
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
               Index           =   27
               Left            =   4815
               TabIndex        =   68
               Top             =   1305
               Width           =   345
            End
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackColor       =   &H008080FF&
            Caption         =   "BOND COAT"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   45
            TabIndex        =   97
            Top             =   135
            Width           =   6675
         End
      End
      Begin VB.Frame frmPreparation 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "METALLOGRAPHIC PREPARATION"
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   45
         TabIndex        =   134
         Top             =   0
         Width           =   13560
         Begin VB.OptionButton opPreparation 
            BackColor       =   &H00C0C0C0&
            Caption         =   "N.R."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   2
            Left            =   12195
            TabIndex        =   151
            Top             =   315
            Width           =   1095
         End
         Begin VB.OptionButton opPreparation 
            BackColor       =   &H00C0C0C0&
            Caption         =   "FAIL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   195
            Index           =   1
            Left            =   10890
            TabIndex        =   140
            Top             =   315
            Width           =   1095
         End
         Begin VB.OptionButton opPreparation 
            BackColor       =   &H00C0C0C0&
            Caption         =   "PASS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   195
            Index           =   0
            Left            =   9540
            TabIndex        =   139
            Top             =   315
            Width           =   1095
         End
         Begin MSComCtl2.DTPicker fechaPreparation 
            Height          =   330
            Left            =   990
            TabIndex        =   135
            Top             =   270
            Width           =   1290
            _ExtentX        =   2275
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
            Format          =   51642369
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin pryCombo.miCombo cmbUsuarioPreparation 
            Height          =   330
            Left            =   3195
            TabIndex        =   136
            Top             =   270
            Width           =   6090
            _ExtentX        =   10742
            _ExtentY        =   582
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Usuario"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   2520
            TabIndex        =   138
            Top             =   315
            Width           =   600
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   360
            TabIndex        =   137
            Top             =   315
            Width           =   735
         End
      End
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
      Left            =   5580
      TabIndex        =   46
      Top             =   10530
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "RESULT"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4635
      TabIndex        =   43
      Top             =   10530
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
         TabIndex        =   44
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
         TabIndex        =   45
         Top             =   225
         Width           =   3390
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "SPECIMEN ID AND DESCRIPTION"
      ForeColor       =   &H80000008&
      Height          =   2325
      Left            =   45
      TabIndex        =   34
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
         TabIndex        =   4
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
         TabIndex        =   7
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
         TabIndex        =   5
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
         TabIndex        =   8
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
         TabIndex        =   6
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
         TabIndex        =   3
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
         TabIndex        =   186
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
         TabIndex        =   42
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   1710
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
         TabIndex        =   39
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
      TabIndex        =   9
      Top             =   10485
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
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   10
      Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
      Top             =   10485
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
      Left            =   9270
      TabIndex        =   33
      Top             =   10755
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
      Left            =   9270
      TabIndex        =   31
      Top             =   10980
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
      Left            =   12645
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   10485
      Width           =   1140
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10620
      Top             =   10665
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
            Picture         =   "frmPlasma_Resultados.frx":08CA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdCopiarResultados 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copiar Resultados"
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
      Left            =   1260
      Picture         =   "frmPlasma_Resultados.frx":11A4
      Style           =   1  'Graphical
      TabIndex        =   204
      Top             =   10485
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
      Left            =   11475
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   10485
      Width           =   1140
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
      Left            =   12240
      TabIndex        =   14
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
      TabIndex        =   13
      Top             =   0
      Width           =   13995
   End
End
Attribute VB_Name = "frmPlasma_Resultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Dim xBOND As New XArrayDB
Dim xTOP As New XArrayDB
Dim tooltipBond As New cTooltip
Dim tooltipTop As New cTooltip
Const filasGrid As Integer = 20
Const ColGrid As Integer = 3
Private Enum ColsGrid
    ENSAYO_ID = 0
    ENSAYO = 1
    VALOR = 2
    RESULT = 3
End Enum
Private Sub chkBondEspesor_Click()
    If chkBondEspesor.Value = Checked Then
        chkBondEspesor.Caption = "PASS"
        chkBondEspesor.ForeColor = &H8000&
    Else
        chkBondEspesor.Caption = "FAIL"
        chkBondEspesor.ForeColor = vbRed
    End If
End Sub

Private Sub chkBondEspesorNR_Click()
    If chkBondEspesorNR.Value = Unchecked Then
        chkBondEspesor.Enabled = True
    Else
        chkBondEspesor.Enabled = False
    End If
End Sub

Private Sub chkBondMacro_Click()
    If chkBondMacro.Value = Checked Then
        chkBondMacro.Caption = "PASS"
        chkBondMacro.ForeColor = &H8000&
    Else
        chkBondMacro.Caption = "FAIL"
        chkBondMacro.ForeColor = vbRed
    End If
End Sub

Private Sub chkBondMacroNR_Click()
    If chkBondMacroNR.Value = Unchecked Then
        chkBondMacro.Enabled = True
    Else
        chkBondMacro.Enabled = False
        If txtDatos(34) = "" Then
            txtDatos(34) = "--"
        End If
    End If
End Sub

Private Sub chkBondMetalografiaNR_Click()
    If chkBondMetalografiaNR.Value = Unchecked Then
        gridBOND.Enabled = True
    Else
        gridBOND.Enabled = False
        gridBOND.Col = ColsGrid.RESULT
        On Error Resume Next
        For i = 0 To 5
            gridBOND.Row = i
            gridBOND.Text = "0"
        Next
    End If
End Sub

Private Sub chkBondMicro_Click()
    If chkBondMicro.Value = Checked Then
        chkBondMicro.Caption = "PASS"
        chkBondMicro.ForeColor = &H8000&
    Else
        chkBondMicro.Caption = "FAIL"
        chkBondMicro.ForeColor = vbRed
    End If
End Sub

Private Sub chkBondMicroNR_Click()
    If chkBondMicroNR.Value = Unchecked Then
        chkBondMicro.Enabled = True
    Else
        chkBondMicro.Enabled = False
        If txtDatos(36) = "" Then
            txtDatos(36) = "--"
        End If
    End If
End Sub

Private Sub chkBondTraccion_Click()
    If chkBondTraccion.Value = Checked Then
        chkBondTraccion.Caption = "PASS"
        chkBondTraccion.ForeColor = &H8000&
    Else
        chkBondTraccion.Caption = "FAIL"
        chkBondTraccion.ForeColor = vbRed
    End If
End Sub

Private Sub chkBondTraccionNR_Click()
    If chkBondTraccionNR.Value = Unchecked Then
        chkBondTraccion.Enabled = True
    Else
        chkBondTraccion.Enabled = False
        If txtDatos(32) = "" Then
            txtDatos(32) = "--"
        End If
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

Private Sub chkTopEspesor_Click()
    If chkTopEspesor.Value = Checked Then
        chkTopEspesor.Caption = "PASS"
        chkTopEspesor.ForeColor = &H8000&
    Else
        chkTopEspesor.Caption = "FAIL"
        chkTopEspesor.ForeColor = vbRed
    End If
End Sub

Private Sub chkTopEspesorNR_Click()
    If chkTopEspesorNR.Value = Unchecked Then
        chkTopEspesor.Enabled = True
    Else
        chkTopEspesor.Enabled = False
    End If

End Sub

Private Sub chkTopMacro_Click()
    If chkTopMacro.Value = Checked Then
        chkTopMacro.Caption = "PASS"
        chkTopMacro.ForeColor = &H8000&
    Else
        chkTopMacro.Caption = "FAIL"
        chkTopMacro.ForeColor = vbRed
    End If

End Sub

Private Sub chkTopMacroNR_Click()
    If chkTopMacroNR.Value = Unchecked Then
        chkTopMacro.Enabled = True
    Else
        chkTopMacro.Enabled = False
        If txtDatos(44) = "" Then
            txtDatos(44) = "--"
        End If
    End If

End Sub

Private Sub chkTopMetalografiaNR_Click()
    If chkTopMetalografiaNR.Value = Unchecked Then
        gridTOP.Enabled = True
    Else
        gridTOP.Enabled = False
        gridTOP.Col = ColsGrid.RESULT
        On Error Resume Next
        For i = 0 To 5
            gridTOP.Row = i
            gridTOP.Text = "0"
        Next
    End If
End Sub

Private Sub chkTopMicro_Click()
    If chkTopMicro.Value = Checked Then
        chkTopMicro.Caption = "PASS"
        chkTopMicro.ForeColor = &H8000&
    Else
        chkTopMicro.Caption = "FAIL"
        chkTopMicro.ForeColor = vbRed
    End If

End Sub

Private Sub chkTopMicroNR_Click()
    If chkTopMicroNR.Value = Unchecked Then
        chkTopMicro.Enabled = True
    Else
        chkTopMicro.Enabled = False
        If txtDatos(46) = "" Then
            txtDatos(46) = "--"
        End If
    End If

End Sub

Private Sub chkTopTraccion_Click()
    If chkTopTraccion.Value = Checked Then
        chkTopTraccion.Caption = "PASS"
        chkTopTraccion.ForeColor = &H8000&
    Else
        chkTopTraccion.Caption = "FAIL"
        chkTopTraccion.ForeColor = vbRed
    End If
End Sub

Private Sub chkTopTraccionNR_Click()
    If chkTopTraccionNR.Value = Unchecked Then
        chkTopTraccion.Enabled = True
    Else
        chkTopTraccion.Enabled = False
        If txtDatos(42) = "" Then
            txtDatos(42) = "--"
        End If
    End If
End Sub

Private Sub cmdCopiarResultados_Click()
    frmPlasma_CopiaResultados.PK = PK
    frmPlasma_CopiaResultados.Show 1
End Sub

Private Sub cmdImagen_Click()
    With frmCE_Imagenes
        .PK = PK
        .Show 1
    End With
End Sub

Private Sub cmdObservador_Click()

    Dim objfrm As New frmObservadorEnsayo

   On Error GoTo cmdObservador_Click_Error

    objfrm.FORMULARIO_ORIGEN = 5 'PLASMA
    objfrm.ES_CONTROL_EFICACIA = False
    objfrm.MUESTRA_ID = PK ' Id de la muestra
    Dim oM As New clsMuestra
    oM.CargaMuestra PK
    objfrm.DETERMINACION_ENSAYO_ID = o
    objfrm.ENSAYO = oM.getTIPO_ANALISIS_ID
    
    If (UCase(lblCerrada) <> "CERRADA") Then
        objfrm.MUESTRA_CERRADA = False
    Else
        objfrm.MUESTRA_CERRADA = True
    End If

    objfrm.Show vbModal
    
    Set objfrm = Nothing

   On Error GoTo 0
   Exit Sub

cmdObservador_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdObservador_Click of Formulario frmPlasma_Resultados"

End Sub

Private Sub cmdok_Click()
    Dim oPR As New clsPlasma_resultados
    Dim oPRE As New clsPlasma_recepcion
            Dim h As Integer
            Dim ensayoid As String
            Dim VALOR As String
            Dim RESULT As String
   On Error GoTo cmdok_Click_Error
    ' Validar reactivos caducados (1090)
    Dim cont As Integer
    Dim existen As Boolean
    existen = False
    For cont = 1 To listaReactivos.ListItems.Count
        If Trim(listaReactivos.ListItems(cont).SubItems(2)) <> "" Then
            If Format(listaReactivos.ListItems(cont).SubItems(2), "yyyy-mm-dd") < Format(Date, "yyyy-mm-dd") Then
                existen = True
            End If
        End If
    Next
    If existen Then
        If MsgBox("Existen reactivos CADUCADOS. ¿ESTA SEGURO DE ALMACENAR LOS DATOS DE LA MUESTRA?", vbExclamation + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    ' Validar equipos pendientes de verificacion
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
    ' Validar preparacion metalografica
    If opPreparation(0).Value = False And opPreparation(1).Value = False And opPreparation(2).Value = False Then
        MsgBox "Debe indicar el resultado de la preparación metalográfica (PASS, FAIL, N.R.).", vbCritical, App.Title
        Exit Sub
    End If
    ' Validar rangos MICRO
    Dim res() As String
    Dim microPor As Integer
    If txtPOR(37) <> "" Then
        res = Split(txtDatos(37), " ")
        If IsNumeric(Trim(res(0))) Then
            If Trim(res(0)) < 100 Then
                microPor = 999
            ElseIf Trim(res(0)) <= 240 Then
                microPor = 13
            ElseIf Trim(res(0)) <= 600 Then
                microPor = 10
            Else
                microPor = 8
            End If
            If CInt(txtPOR(37)) > microPor Then
                If MsgBox("El porcentaje de la MICRO en la BOND COAT es mayor de " & microPor & ", no cumple repetibilidad. Se marcará el ensayo como FAIL. ¿Desea continuar?", vbCritical + vbYesNo, App.Title) = vbNo Then
                    Exit Sub
                Else
                    chkBondMicro.Value = Unchecked
                    chkResult.Value = Unchecked
                End If
            End If
        End If
    End If
    If txtPOR(47) <> "" Then
        res = Split(txtDatos(47), " ")
        If IsNumeric(Trim(res(0))) Then
            If Trim(res(0)) < 100 Then
                microPor = 999
            ElseIf Trim(res(0)) <= 240 Then
                microPor = 13
            ElseIf Trim(res(0)) <= 600 Then
                microPor = 10
            Else
                microPor = 8
            End If
            If CInt(txtPOR(47)) > microPor Then
                If MsgBox("El porcentaje de la MICRO en la TOP COAT es mayor de " & microPor & ", no cumple repetibilidad. Se marcará el ensayo como FAIL. ¿Desea continuar?", vbCritical + vbYesNo, App.Title) = vbNo Then
                    Exit Sub
                Else
                    chkTopMicro.Value = Unchecked
                    chkResult.Value = Unchecked
                End If
            End If
        End If
    End If
    ' Validar rangos MACRO
    Dim unidades() As String
    Dim rango As Integer
    unidades = Split(txtUnidades, ";")
    If txtPOR(35) <> "" Then
        rango = 0
        If unidades(1) = "HR15N" Or unidades(1) = "HRBW" Or unidades(1) = "HRC" Then
            rango = 10
        End If
        If unidades(1) = "HR15Y" Then
            rango = 15
        End If
        If CInt(txtPOR(35)) > rango And rango <> 0 Then
            If MsgBox("El porcentaje de la MACRO en la TOP COAT es mayor de " & rango & ", no cumple repetibilidad. Se marcará el ensayo como FAIL. ¿Desea continuar?", vbCritical + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            Else
                chkTopMacro.Value = Unchecked
                chkResult.Value = Unchecked
            End If
        End If
    End If
    If txtPOR(45) <> "" Then
        rango = 0
        If unidades(4) = "HR15N" Or unidades(4) = "HRBW" Or unidades(4) = "HRC" Then
            rango = 10
        End If
        If unidades(4) = "HR15Y" Then
            rango = 15
        End If
        If CInt(txtPOR(45)) > rango And rango <> 0 Then
            If MsgBox("El porcentaje de la MACRO en la TOP COAT es mayor de " & rango & ", no cumple repetibilidad. Se marcará el ensayo como FAIL. ¿Desea continuar?", vbCritical + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            Else
                chkTopMacro.Value = Unchecked
                chkResult.Value = Unchecked
            End If
        End If
    End If
    ' Validar temperaturas
'    If frmMacroTemp.Enabled = True Then
'        If Trim(txtDatos(8)) = "" Then
'            MsgBox "Debe indicar la temperatura 30 minutos antes de Macro Dureza.", vbCritical, App.Title
'            txtDatos(8).SetFocus
'            Exit Sub
'        End If
'        If Trim(txtDatos(9)) = "" Then
'            MsgBox "Debe indicar la temperatura de ensayo de Macro Dureza.", vbCritical, App.Title
'            txtDatos(9).SetFocus
'            Exit Sub
'        End If
'    End If
'    If frmMicroTemp.Enabled = True Then
'        If Trim(txtDatos(10)) = "" Then
'            MsgBox "Debe indicar la temperatura 30 minutos antes de Micro Dureza.", vbCritical, App.Title
'            txtDatos(10).SetFocus
'            Exit Sub
'        End If
'        If Trim(txtDatos(11)) = "" Then
'            MsgBox "Debe indicar la temperatura de ensayo de Micro Dureza.", vbCritical, App.Title
'            txtDatos(11).SetFocus
'            Exit Sub
'        End If
'    End If
    ' Almacenar Datos
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
        
        .setMP = 1
        .setMP_FECHA = "'" & Format(fechaPreparation, "yyyy-mm-dd") & "'"
        .setMP_USUARIO_ID = cmbUsuarioPreparation.getPK_SALIDA
        If opPreparation(0).Value = True Then
            .setMP_PASS = 0
        ElseIf opPreparation(1).Value = True Then
            .setMP_PASS = 1
        ElseIf opPreparation(2).Value = True Then
            .setMP_PASS = 2
        End If
        
        .setMACRO_DUREZA_T1 = txtDatos(8)
        .setMACRO_DUREZA_T2 = txtDatos(9)
        .setMICRO_DUREZA_T1 = txtDatos(10)
        .setMICRO_DUREZA_T2 = txtDatos(11)
        
        .Modificar PK
    End With
    With oPR
        ' BOND
        .setMUESTRA_ID = PK
        .setTIPO = 1
        .setBATCH = txtDatos(31)
        .setMICROESTRUCTURA1 = 0
        .setMICROESTRUCTURA2 = 0
        .setMICROESTRUCTURA3 = 0
        .setMICROESTRUCTURA4 = 0
        .setMICROESTRUCTURA5 = 0
        .setMICROESTRUCTURA6 = 0
        .setMICROESTRUCTURA1_VALOR = ""
        .setMICROESTRUCTURA2_VALOR = ""
        .setMICROESTRUCTURA3_VALOR = ""
        .setMICROESTRUCTURA4_VALOR = ""
        .setMICROESTRUCTURA5_VALOR = ""
        .setMICROESTRUCTURA6_VALOR = ""
        If chkBondMetalografiaNR.Value = Checked Then
            .setMICROESTRUCTURA1 = 2
            .setMICROESTRUCTURA2 = 2
            .setMICROESTRUCTURA3 = 2
            .setMICROESTRUCTURA4 = 2
            .setMICROESTRUCTURA5 = 2
            .setMICROESTRUCTURA6 = 2
        Else
            For h = 0 To 5
                gridBOND.Row = h
                gridBOND.Col = ColsGrid.ENSAYO_ID
                ensayoid = gridBOND.Text
                gridBOND.Col = ColsGrid.VALOR
                VALOR = gridBOND.Text
                gridBOND.Col = ColsGrid.RESULT
                RESULT = ""
                If gridBOND.Text <> "" Then
                    RESULT = Abs(gridBOND.Text)
                End If
                If ensayoid <> "" Then
                    Select Case CInt(ensayoid)
                    Case 1
                        .setMICROESTRUCTURA1_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA1 = 1
                        Else
                            .setMICROESTRUCTURA1 = 0
                        End If
                    Case 2
                        .setMICROESTRUCTURA2_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA2 = 1
                        Else
                            .setMICROESTRUCTURA2 = 0
                        End If
                    Case 3
                        .setMICROESTRUCTURA3_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA3 = 1
                        Else
                            .setMICROESTRUCTURA3 = 0
                        End If
                    Case 4
                        .setMICROESTRUCTURA4_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA4 = 1
                        Else
                            .setMICROESTRUCTURA4 = 0
                        End If
                    Case 5
                        .setMICROESTRUCTURA5_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA5 = 1
                        Else
                            .setMICROESTRUCTURA5 = 0
                        End If
                    Case 6
                        .setMICROESTRUCTURA6_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA6 = 1
                        Else
                            .setMICROESTRUCTURA6 = 0
                        End If
                    End Select
                End If
            Next
        End If
        .setTRACCION = txtDatos(32)
        .setTRACCION_RES = txtDatos(33)
        
        .setMACRO_DUREZA = txtDatos(34)
        .setMACRO_DUREZA_DIMENSION = txtDatos(4)
        .setMACRO_DUREZA_ESPESOR = txtDatos(5)
        .setMACRO_DUREZA_RES = txtDatos(35)
        If txtSD(35) = "" Then
            .setMACRO_DUREZA_SD = 0
        Else
            .setMACRO_DUREZA_SD = Replace(txtSD(35), ",", ".")
        End If
        If txtSD(35) = "" Then
            .setMACRO_DUREZA_POR = 0
        Else
            .setMACRO_DUREZA_POR = Replace(txtPOR(35), ",", ".")
        End If
        
        .setMICRO_DUREZA = txtDatos(36)
        .setMICRO_DUREZA_RES = txtDatos(37)
        
        If txtSD(37) = "" Then
            .setMICRO_DUREZA_SD = 0
        Else
            .setMICRO_DUREZA_SD = Replace(txtSD(37), ",", ".")
        End If
        If txtPOR(37) = "" Then
            .setMICRO_DUREZA_POR = 0
        Else
            .setMICRO_DUREZA_POR = Replace(txtPOR(37), ",", ".")
        End If
        .setESPESOR = txtDatos(38)
        .setESPESOR_MIN = txtDatos(0)
        .setESPESOR_MAX = txtDatos(1)
        .setESPESOR_RES = txtDatos(39)
        If txtSD(39) = "" Then
            .setESPESOR_SD = 0
        Else
            .setESPESOR_SD = Replace(txtSD(39), ",", ".")
        End If
        If txtPOR(39) = "" Then
            .setESPESOR_POR = 0
        Else
            .setESPESOR_POR = Replace(txtPOR(39), ",", ".")
        End If
        
        If chkBondTraccionNR.Value = Checked Then
            .setTRACCION_PASS = 2
        Else
            .setTRACCION_PASS = chkBondTraccion.Value
        End If
        If chkBondMacroNR.Value = Checked Then
            .setMACRO_DUREZA_PASS = 2
        Else
            .setMACRO_DUREZA_PASS = chkBondMacro.Value
        End If
        If chkBondMicroNR.Value = Checked Then
            .setMICRO_DUREZA_PASS = 2
        Else
            .setMICRO_DUREZA_PASS = chkBondMicro.Value
        End If
        If chkBondEspesorNR.Value = Checked Then
            .setESPEROR_PASS = 2
        Else
            .setESPEROR_PASS = chkBondEspesor.Value
        End If
        
        .Insertar
        ' TOP
        .setTIPO = 2
        .setBATCH = txtDatos(41)
        .setMICROESTRUCTURA1 = 0
        .setMICROESTRUCTURA2 = 0
        .setMICROESTRUCTURA3 = 0
        .setMICROESTRUCTURA4 = 0
        .setMICROESTRUCTURA5 = 0
        .setMICROESTRUCTURA6 = 0
        .setMICROESTRUCTURA1_VALOR = ""
        .setMICROESTRUCTURA2_VALOR = ""
        .setMICROESTRUCTURA3_VALOR = ""
        .setMICROESTRUCTURA4_VALOR = ""
        .setMICROESTRUCTURA5_VALOR = ""
        .setMICROESTRUCTURA6_VALOR = ""
        If chkTopMetalografiaNR.Value = Checked Then
            .setMICROESTRUCTURA1 = 2
            .setMICROESTRUCTURA2 = 2
            .setMICROESTRUCTURA3 = 2
            .setMICROESTRUCTURA4 = 2
            .setMICROESTRUCTURA5 = 2
            .setMICROESTRUCTURA6 = 2
        Else
            For h = 0 To 5
                gridTOP.Row = h
                gridTOP.Col = ColsGrid.ENSAYO_ID
                ensayoid = gridTOP.Text
                gridTOP.Col = ColsGrid.VALOR
                VALOR = gridTOP.Text
                gridTOP.Col = ColsGrid.RESULT
                RESULT = ""
                If gridTOP.Text <> "" Then
                    RESULT = Abs(gridTOP.Text)
                End If
                If ensayoid <> "" Then
                    Select Case CInt(ensayoid)
                    Case 1
                        .setMICROESTRUCTURA1_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA1 = 1
                        Else
                            .setMICROESTRUCTURA1 = 0
                        End If
                    Case 2
                        .setMICROESTRUCTURA2_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA2 = 1
                        Else
                            .setMICROESTRUCTURA2 = 0
                        End If
                    Case 3
                        .setMICROESTRUCTURA3_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA3 = 1
                        Else
                            .setMICROESTRUCTURA3 = 0
                        End If
                    Case 4
                        .setMICROESTRUCTURA4_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA4 = 1
                        Else
                            .setMICROESTRUCTURA4 = 0
                        End If
                    Case 5
                        .setMICROESTRUCTURA5_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA5 = 1
                        Else
                            .setMICROESTRUCTURA5 = 0
                        End If
                    Case 6
                        .setMICROESTRUCTURA6_VALOR = VALOR
                        If RESULT = "1" Then
                            .setMICROESTRUCTURA6 = 1
                        Else
                            .setMICROESTRUCTURA6 = 0
                        End If
                    End Select
                End If
            Next
        End If
        .setTRACCION = txtDatos(42)
        .setTRACCION_RES = txtDatos(43)
        
        .setMACRO_DUREZA = txtDatos(44)
        .setMACRO_DUREZA_DIMENSION = txtDatos(7)
        .setMACRO_DUREZA_ESPESOR = txtDatos(6)
        .setMACRO_DUREZA_RES = txtDatos(45)
        If txtSD(45) = "" Then
            .setMACRO_DUREZA_SD = 0
        Else
            .setMACRO_DUREZA_SD = Replace(txtSD(45), ",", ".")
        End If
        If txtSD(45) = "" Then
            .setMACRO_DUREZA_POR = 0
        Else
            .setMACRO_DUREZA_POR = Replace(txtPOR(45), ",", ".")
        End If
        
        .setMICRO_DUREZA = txtDatos(46)
        .setMICRO_DUREZA_RES = txtDatos(47)
        If txtSD(47) = "" Then
            .setMICRO_DUREZA_SD = 0
        Else
            .setMICRO_DUREZA_SD = Replace(txtSD(47), ",", ".")
        End If
        If txtPOR(47) = "" Then
            .setMICRO_DUREZA_POR = 0
        Else
            .setMICRO_DUREZA_POR = Replace(txtPOR(47), ",", ".")
        End If
        .setESPESOR = txtDatos(48)
        .setESPESOR_MIN = txtDatos(2)
        .setESPESOR_MAX = txtDatos(3)
        .setESPESOR_RES = txtDatos(49)
        If txtSD(49) = "" Then
            .setESPESOR_SD = 0
        Else
            .setESPESOR_SD = Replace(txtSD(49), ",", ".")
        End If
        If txtPOR(49) = "" Then
            .setESPESOR_POR = 0
        Else
            .setESPESOR_POR = Replace(txtPOR(49), ",", ".")
        End If
        If chkTopTraccionNR.Value = Checked Then
            .setTRACCION_PASS = 2
        Else
            .setTRACCION_PASS = chkTopTraccion.Value
        End If
        If chkTopMacroNR.Value = Checked Then
            .setMACRO_DUREZA_PASS = 2
        Else
            .setMACRO_DUREZA_PASS = chkTopMacro.Value
        End If
        If chkTopMicroNR.Value = Checked Then
            .setMICRO_DUREZA_PASS = 2
        Else
            .setMICRO_DUREZA_PASS = chkTopMicro.Value
        End If
        If chkTopEspesorNR.Value = Checked Then
            .setESPEROR_PASS = 2
        Else
            .setESPEROR_PASS = chkTopEspesor.Value
        End If
        .Insertar
    End With
    Set oPR = Nothing
    oPRE.informarRequirement PK
    oPRE.informarControlSpecification PK
    oPRE.ModificarResultado PK, chkResult.Value
    Set oPRE = Nothing
    Dim oPRH As New clsPlasma_resultados_historico
    oPRH.generar PK
    Set oPRH = Nothing
    Me.MousePointer = 0
    MsgBox "Datos almacenados correctamente.", vbInformation, App.Title
    If USUARIO.getPER_CIERRE = True Then
        Dim oMuestra As New clsMuestra
        oMuestra.CargaMuestra PK
        If oMuestra.getCERRADA = 0 Then
            If MsgBox("¿Desea cerrar la muestra?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                If Trim(txtDatos(31)) = "" Or Trim(txtDatos(41)) = "" Then
                    MsgBox "Debe indicar los campos BATCH BOND COAT y TOP BOND COAT para poder cerrar la muestra.", vbExclamation, App.Title
                Else
                    oMuestra.Cerrar PK
                End If
            End If
        End If
    End If
    Unload Me
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmPlasma_Resultados"
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
    If chkAbierta.Value = Checked Then
        If MsgBox("Va a salir sin guardar los cambios.¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
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
    inicializar_grid 1
    inicializar_grid 2
    toolBond
    toolTop
    fechaPreparation = Date
'SOLICITUD LAURA M2600 cmbUsuarioPreparation.MostrarElemento USUARIO.getID_EMPLEADO
    permisos
    If PK > 0 Then
        cargar_muestra
    End If
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
    
    lbltitulo = "Registro resultados PLASMA : " & Trim(str(oMuestra.getID_GENERAL)) & " (" & oMuestra.CodigoParticular(PK) & ")"
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
            txtDatos(50) = .getSPECIMEN_ID
            cmbnatype.MostrarElemento .getNTYPE
            txtDatos(51) = .getPN
            txtDatos(52) = .getPRODUCT_TYPE
            txtDatos(53) = .getMODULE_SN
            txtDatos(54) = .getSN
            txtDatos(55) = .getPRODUCT_SN
            If .getMP_FECHA <> "" Then
                fechaPreparation = .getMP_FECHA
                cmbUsuarioPreparation.MostrarElemento .getMP_USUARIO_ID
                opPreparation(.getMP_PASS).Value = True
            Else
                opPreparation(2).Value = True
            End If
            txtDatos(8) = .getMACRO_DUREZA_T1
            txtDatos(9) = .getMACRO_DUREZA_T2
            txtDatos(10) = .getMICRO_DUREZA_T1
            txtDatos(11) = .getMICRO_DUREZA_T2
        End If
    End With
    Set oPlasmaRecepcion = Nothing
    cargar_ficha
    cargar_resultados
    proteger_campos oMuestra.getCERRADA
    If oMuestra.getCERRADA = 0 Then
        chkAbierta.Value = Checked
    End If
   On Error GoTo 0
   Exit Sub

cargar_muestra_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_muestra of Formulario frmPlasma_Resultados"
End Sub
Private Sub cargar_ficha()
    Dim oPR As New clsPlasma_recepcion
    If oPR.Carga(PK) Then
        Dim oPP As New clsPlasma_procesos
        If oPP.Carga(oPR.getPROCESO_ID) Then
            Dim oPF As New clsPlasma_ficha
            Dim oPE As New clsPlasma_ensayos
            Dim ounidades As New clsUnidades
            Dim texto As String
            ' BOND
            fichaBond = oPP.getBOND_COAT_FICHA_ID
            If oPF.Carga(oPP.getBOND_COAT_FICHA_ID) Then
                txtDatos(30) = oPF.getMETCO
                cargarMicroEstructura oPP.getBOND_COAT_FICHA_ID, 1
                frmBondMicroEstructura.Enabled = oPP.getBOND_MICROESTRUCTURA
                frmBondTraccion.Enabled = oPP.getBOND_TRACCION
                frmBondMacro.Enabled = oPP.getBOND_MACRO_DUREZA
                frmBondMicro.Enabled = oPP.getBOND_MICRO_DUREZA
                frmBondEspesor.Enabled = oPP.getBOND_ESPESOR
                
                chkBondTraccion.visible = oPP.getBOND_TRACCION
                chkBondMacro.visible = oPP.getBOND_MACRO_DUREZA
                chkBondMicro.visible = oPP.getBOND_MICRO_DUREZA
                chkBondEspesor.visible = oPP.getBOND_ESPESOR
                
                chkBondMetalografiaNR.visible = oPP.getBOND_MICROESTRUCTURA
                chkBondTraccionNR.visible = oPP.getBOND_TRACCION
                chkBondMacroNR.visible = oPP.getBOND_MACRO_DUREZA
                chkBondMicroNR.visible = oPP.getBOND_MICRO_DUREZA
                chkBondEspesorNR.visible = oPP.getBOND_ESPESOR
                
                If oPP.getBOND_TRACCION = 1 Then
                    txtDatos(13) = oPF.getTRACCION_REQ
                End If
                If oPP.getBOND_MACRO_DUREZA = 1 Then
                    txtDatos(14) = oPF.getMACRO_DUREZA_REQ
                End If
                If oPP.getBOND_MICRO_DUREZA = 1 Then
                    txtDatos(15) = oPF.getMICRO_DUREZA_REQ
                End If
                If oPP.getBOND_ESPESOR = 1 Then
                    txtDatos(16) = oPF.getESPESOR_REQ
                End If
                ' Unidades->traccion
                texto = ""
                If oPF.getTRACCION <> 0 Then
                    If oPE.Carga(oPF.getTRACCION) Then
                        If ounidades.CARGAR(oPE.getUNIDAD_ID) Then
                            texto = ounidades.getNOMBRE
                        End If
                    End If
                End If
                txtUnidades = txtUnidades & texto & ";"
                ' Unidades->macro
                texto = ""
                If oPF.getMACRO_DUREZA <> 0 Then
                    If oPE.Carga(oPF.getMACRO_DUREZA) Then
                        If ounidades.CARGAR(oPE.getUNIDAD_ID) Then
                            texto = ounidades.getNOMBRE
                        End If
                    End If
                End If
                txtUnidades = txtUnidades & texto & ";"
                ' Unidades->micro
                texto = ""
                If oPF.getMICRO_DUREZA <> 0 Then
                    If oPE.Carga(oPF.getMICRO_DUREZA) Then
                        If ounidades.CARGAR(oPE.getUNIDAD_ID) Then
                            texto = ounidades.getNOMBRE
                        End If
                    End If
                End If
                txtUnidades = txtUnidades & texto & ";"
                ' Unidad->espesor
                txtUnidadEspesorBond = ""
                If oPF.getESPESOR <> 0 Then
                    If oPE.Carga(oPF.getESPESOR) Then
                        If ounidades.CARGAR(oPE.getUNIDAD_ID) Then
                            txtUnidadEspesorBond = ounidades.getNOMBRE
                        End If
                    End If
                End If
            End If
            ' TOP
            fichaTop = oPP.getTOP_COAT_FICHA_ID
            If oPF.Carga(oPP.getTOP_COAT_FICHA_ID) Then
                txtDatos(40) = oPF.getMETCO
                cargarMicroEstructura oPP.getTOP_COAT_FICHA_ID, 2
                frmTopMicroEstructura.Enabled = oPP.getTOP_MICROESTRUCTURA
                frmTopTraccion.Enabled = oPP.getTOP_TRACCION
                frmTopMacro.Enabled = oPP.getTOP_MACRO_DUREZA
                frmTopMicro.Enabled = oPP.getTOP_MICRO_DUREZA
                frmTopEspesor.Enabled = oPP.getTOP_ESPESOR
                
                chkTopTraccion.visible = oPP.getTOP_TRACCION
                chkTopMacro.visible = oPP.getTOP_MACRO_DUREZA
                chkTopMicro.visible = oPP.getTOP_MICRO_DUREZA
                chkTopEspesor.visible = oPP.getTOP_ESPESOR
                
                chkTopMetalografiaNR.visible = oPP.getTOP_MICROESTRUCTURA
                chkTopTraccionNR.visible = oPP.getTOP_TRACCION
                chkTopMacroNR.visible = oPP.getTOP_MACRO_DUREZA
                chkTopMicroNR.visible = oPP.getTOP_MICRO_DUREZA
                chkTopEspesorNR.visible = oPP.getTOP_ESPESOR
                
                If oPP.getTOP_TRACCION = 1 Then
                    txtDatos(17) = oPF.getTRACCION_REQ
                End If
                If oPP.getTOP_MACRO_DUREZA = 1 Then
                    txtDatos(18) = oPF.getMACRO_DUREZA_REQ
                End If
                If oPP.getTOP_MICRO_DUREZA = 1 Then
                    txtDatos(19) = oPF.getMICRO_DUREZA_REQ
                End If
                If oPP.getTOP_ESPESOR = 1 Then
                    txtDatos(20) = oPF.getESPESOR_REQ
                End If
                
                ' Unidades->traccion
                texto = ""
                If oPF.getTRACCION <> 0 Then
                    If oPE.Carga(oPF.getTRACCION) Then
                        If ounidades.CARGAR(oPE.getUNIDAD_ID) Then
                            texto = ounidades.getNOMBRE
                        End If
                    End If
                End If
                txtUnidades = txtUnidades & texto & ";"
                ' Unidades->macro
                texto = ""
                If oPF.getMACRO_DUREZA <> 0 Then
                    If oPE.Carga(oPF.getMACRO_DUREZA) Then
                        If ounidades.CARGAR(oPE.getUNIDAD_ID) Then
                            texto = ounidades.getNOMBRE
                        End If
                    End If
                End If
                txtUnidades = txtUnidades & texto & ";"
                ' Unidades->micro
                texto = ""
                If oPF.getMICRO_DUREZA <> 0 Then
                    If oPE.Carga(oPF.getMICRO_DUREZA) Then
                        If ounidades.CARGAR(oPE.getUNIDAD_ID) Then
                            texto = ounidades.getNOMBRE
                        End If
                    End If
                End If
                txtUnidades = txtUnidades & texto & ";"
                ' Unidad->espesor
                txtUnidadEspesorTop = ""
                If oPF.getESPESOR <> 0 Then
                    If oPE.Carga(oPF.getESPESOR) Then
                        If ounidades.CARGAR(oPE.getUNIDAD_ID) Then
                            txtUnidadEspesorTop = ounidades.getNOMBRE
                        End If
                    End If
                End If
            End If
            frmMacroTemp.Enabled = False
            frmMicroTemp.Enabled = False
            If frmBondMacro.Enabled = True Or frmTopMacro.Enabled = True Then
                frmMacroTemp.Enabled = True
            End If
            If frmBondMicro.Enabled = True Or frmTopMicro.Enabled = True Then
                frmMicroTemp.Enabled = True
            End If
        End If
    End If
End Sub
Private Sub cargar_resultados()
    Dim oPR As New clsPlasma_resultados
   On Error GoTo cargar_resultados_Error
    Dim h As Integer
    If oPR.Carga(PK, 1) Then
        txtDatos(31) = oPR.getBATCH
        txtDatos(32) = oPR.getTRACCION
        txtDatos(33) = oPR.getTRACCION_RES
        txtDatos(34) = oPR.getMACRO_DUREZA
        txtDatos(4) = oPR.getMACRO_DUREZA_DIMENSION
        txtDatos(5) = oPR.getMACRO_DUREZA_ESPESOR
        txtDatos(36) = oPR.getMICRO_DUREZA
        txtDatos(38) = oPR.getESPESOR
        txtDatos(0) = oPR.getESPESOR_MIN
        txtDatos(1) = oPR.getESPESOR_MAX
        If oPR.getMICROESTRUCTURA1 = 2 Then
            If Not IsEmpty(xBOND(0, ColsGrid.ENSAYO)) Then
                xBOND(0, ColsGrid.VALOR) = ""
                xBOND(0, ColsGrid.RESULT) = "0"
            End If
            If Not IsEmpty(xBOND(1, ColsGrid.ENSAYO)) Then
                xBOND(1, ColsGrid.VALOR) = ""
                xBOND(1, ColsGrid.RESULT) = "0"
            End If
            If Not IsEmpty(xBOND(2, ColsGrid.ENSAYO)) Then
                xBOND(2, ColsGrid.VALOR) = ""
                xBOND(2, ColsGrid.RESULT) = "0"
            End If
            If Not IsEmpty(xBOND(3, ColsGrid.ENSAYO)) Then
                xBOND(3, ColsGrid.VALOR) = ""
                xBOND(3, ColsGrid.RESULT) = "0"
            End If
            If Not IsEmpty(xBOND(4, ColsGrid.ENSAYO)) Then
                xBOND(4, ColsGrid.VALOR) = ""
                xBOND(4, ColsGrid.RESULT) = "0"
            End If
            If Not IsEmpty(xBOND(5, ColsGrid.ENSAYO)) Then
                xBOND(5, ColsGrid.VALOR) = ""
                xBOND(5, ColsGrid.RESULT) = "0"
            End If
            gridBOND.Refresh
            chkBondMetalografiaNR.Value = Checked
            gridBOND.Enabled = False
        Else
            For h = 0 To 5
                If Not IsEmpty(xBOND(h, ColsGrid.ENSAYO)) Then
                    Select Case CInt(xBOND(h, ColsGrid.ENSAYO_ID))
                    Case 1
                        xBOND(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA1_VALOR
                        xBOND(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA1
                    Case 2
                        xBOND(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA2_VALOR
                        xBOND(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA2
                    Case 3
                        xBOND(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA3_VALOR
                        xBOND(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA3
                    Case 4
                        xBOND(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA4_VALOR
                        xBOND(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA4
                    Case 5
                        xBOND(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA5_VALOR
                        xBOND(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA5
                    Case 6
                        xBOND(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA6_VALOR
                        xBOND(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA6
                    End Select
                End If
            Next
'            If Not IsEmpty(xBOND(1, ColsGrid.ENSAYO)) Then
'                xBOND(1, ColsGrid.valor) = oPR.getMICROESTRUCTURA2_VALOR
'                xBOND(1, ColsGrid.result) = oPR.getMICROESTRUCTURA2
'            End If
'            If Not IsEmpty(xBOND(2, ColsGrid.ENSAYO)) Then
'                xBOND(2, ColsGrid.valor) = oPR.getMICROESTRUCTURA3_VALOR
'                xBOND(2, ColsGrid.result) = oPR.getMICROESTRUCTURA3
'            End If
'            If Not IsEmpty(xBOND(3, ColsGrid.ENSAYO)) Then
'                xBOND(3, ColsGrid.valor) = oPR.getMICROESTRUCTURA4_VALOR
'                xBOND(3, ColsGrid.result) = oPR.getMICROESTRUCTURA4
'            End If
'            If Not IsEmpty(xBOND(4, ColsGrid.ENSAYO)) Then
'                xBOND(4, ColsGrid.valor) = oPR.getMICROESTRUCTURA5_VALOR
'                xBOND(4, ColsGrid.result) = oPR.getMICROESTRUCTURA5
'            End If
'            If Not IsEmpty(xBOND(5, ColsGrid.ENSAYO)) Then
'                xBOND(5, ColsGrid.valor) = oPR.getMICROESTRUCTURA6_VALOR
'                xBOND(5, ColsGrid.result) = oPR.getMICROESTRUCTURA6
'            End If
            gridBOND.Refresh
        End If
        If oPR.getTRACCION_PASS = 2 Then
            chkBondTraccionNR.Value = Checked
            chkBondTraccion.Enabled = False
        Else
            chkBondTraccion.Value = oPR.getTRACCION_PASS
        End If
        If oPR.getMACRO_DUREZA_PASS = 2 Then
            chkBondMacroNR.Value = Checked
            chkBondMacro.Enabled = False
        Else
            chkBondMacro.Value = oPR.getMACRO_DUREZA_PASS
        End If
        If oPR.getMICRO_DUREZA_PASS = 2 Then
            chkBondMicroNR.Value = Checked
            chkBondMicro.Enabled = False
        Else
            chkBondMicro.Value = oPR.getMICRO_DUREZA_PASS
        End If
        If oPR.getESPESOR_PASS = 2 Then
            chkBondEspesorNR.Value = Checked
            chkBondEspesor.Enabled = False
        Else
            chkBondEspesor.Value = oPR.getESPESOR_PASS
        End If
    End If
    If oPR.Carga(PK, 2) Then
        txtDatos(41) = oPR.getBATCH
        txtDatos(42) = oPR.getTRACCION
        txtDatos(43) = oPR.getTRACCION_RES
        txtDatos(44) = oPR.getMACRO_DUREZA
        txtDatos(7) = oPR.getMACRO_DUREZA_DIMENSION
        txtDatos(6) = oPR.getMACRO_DUREZA_ESPESOR
        txtDatos(46) = oPR.getMICRO_DUREZA
        txtDatos(48) = oPR.getESPESOR
        txtDatos(2) = oPR.getESPESOR_MIN
        txtDatos(3) = oPR.getESPESOR_MAX
        
        If oPR.getMICROESTRUCTURA1 = 2 Then
            If Not IsEmpty(xTOP(0, ColsGrid.ENSAYO)) Then
                xTOP(0, ColsGrid.VALOR) = ""
                xTOP(0, ColsGrid.RESULT) = "0"
            End If
            If Not IsEmpty(xTOP(1, ColsGrid.ENSAYO)) Then
                xTOP(1, ColsGrid.VALOR) = ""
                xTOP(1, ColsGrid.RESULT) = "0"
            End If
            If Not IsEmpty(xTOP(2, ColsGrid.ENSAYO)) Then
                xTOP(2, ColsGrid.VALOR) = ""
                xTOP(2, ColsGrid.RESULT) = "0"
            End If
            If Not IsEmpty(xTOP(3, ColsGrid.ENSAYO)) Then
                xTOP(3, ColsGrid.VALOR) = ""
                xTOP(3, ColsGrid.RESULT) = "0"
            End If
            If Not IsEmpty(xTOP(4, ColsGrid.ENSAYO)) Then
                xTOP(4, ColsGrid.VALOR) = ""
                xTOP(4, ColsGrid.RESULT) = "0"
            End If
            If Not IsEmpty(xTOP(5, ColsGrid.ENSAYO)) Then
                xTOP(5, ColsGrid.VALOR) = ""
                xTOP(5, ColsGrid.RESULT) = "0"
            End If
            gridTOP.Refresh
            chkTopMetalografiaNR.Value = Checked
            gridTOP.Enabled = False
        Else
            For h = 0 To 5
                If Not IsEmpty(xTOP(h, ColsGrid.ENSAYO)) Then
                    Select Case CInt(xTOP(h, ColsGrid.ENSAYO_ID))
                    Case 1
                        xTOP(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA1_VALOR
                        xTOP(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA1
                    Case 2
                        xTOP(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA2_VALOR
                        xTOP(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA2
                    Case 3
                        xTOP(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA3_VALOR
                        xTOP(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA3
                    Case 4
                        xTOP(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA4_VALOR
                        xTOP(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA4
                    Case 5
                        xTOP(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA5_VALOR
                        xTOP(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA5
                    Case 6
                        xTOP(h, ColsGrid.VALOR) = oPR.getMICROESTRUCTURA6_VALOR
                        xTOP(h, ColsGrid.RESULT) = oPR.getMICROESTRUCTURA6
                    End Select
                End If
            Next
'            If Not IsEmpty(xTOP(0, ColsGrid.ENSAYO)) Then
'                xTOP(0, ColsGrid.valor) = oPR.getMICROESTRUCTURA1_VALOR
'                xTOP(0, ColsGrid.result) = oPR.getMICROESTRUCTURA1
'            End If
'            If Not IsEmpty(xTOP(1, ColsGrid.ENSAYO)) Then
'                xTOP(1, ColsGrid.valor) = oPR.getMICROESTRUCTURA2_VALOR
'                xTOP(1, ColsGrid.result) = oPR.getMICROESTRUCTURA2
'            End If
'            If Not IsEmpty(xTOP(2, ColsGrid.ENSAYO)) Then
'                xTOP(2, ColsGrid.valor) = oPR.getMICROESTRUCTURA3_VALOR
'                xTOP(2, ColsGrid.result) = oPR.getMICROESTRUCTURA3
'            End If
'            If Not IsEmpty(xTOP(3, ColsGrid.ENSAYO)) Then
'                xTOP(3, ColsGrid.valor) = oPR.getMICROESTRUCTURA4_VALOR
'                xTOP(3, ColsGrid.result) = oPR.getMICROESTRUCTURA4
'            End If
'            If Not IsEmpty(xTOP(4, ColsGrid.ENSAYO)) Then
'                xTOP(4, ColsGrid.valor) = oPR.getMICROESTRUCTURA5_VALOR
'                xTOP(4, ColsGrid.result) = oPR.getMICROESTRUCTURA5
'            End If
'            If Not IsEmpty(xTOP(5, ColsGrid.ENSAYO)) Then
'                xTOP(5, ColsGrid.valor) = oPR.getMICROESTRUCTURA6_VALOR
'                xTOP(5, ColsGrid.result) = oPR.getMICROESTRUCTURA6
'            End If
            gridTOP.Refresh
        End If
        If oPR.getTRACCION_PASS = 2 Then
            chkTopTraccionNR.Value = Checked
            chkTopTraccion.Enabled = False
        Else
            chkTopTraccion.Value = oPR.getTRACCION_PASS
        End If
        If oPR.getMACRO_DUREZA_PASS = 2 Then
            chkTopMacroNR.Value = Checked
            chkTopMacro.Enabled = False
        Else
            chkTopMacro.Value = oPR.getMACRO_DUREZA_PASS
        End If
        If oPR.getMICRO_DUREZA_PASS = 2 Then
            chkTopMicroNR.Value = Checked
            chkTopMicro.Enabled = False
        Else
            chkTopMicro.Value = oPR.getMICRO_DUREZA_PASS
        End If
        If oPR.getESPESOR_PASS = 2 Then
            chkTopEspesorNR.Value = Checked
            chkTopEspesor.Enabled = False
        Else
            chkTopEspesor.Value = oPR.getESPESOR_PASS
        End If
    End If
    
    Dim oPRE As New clsPlasma_recepcion
    oPRE.Carga PK
    chkResult.Value = oPRE.getRESULT
    Set oPRE = Nothing
    
   On Error GoTo 0
   Exit Sub

cargar_resultados_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_resultados of Formulario frmPlasma_Resultados"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tooltipBond = Nothing
    Set tooltipTop = Nothing
    
End Sub

Private Sub gridBOND_SelChange(Cancel As Integer)
    If fichaBond <> "" Then
        Dim oPFE As New clsPlasma_ficha_estructura
        If oPFE.Carga(fichaBond, xBOND(gridBOND.Row, ColsGrid.ENSAYO_ID)) Then
            tooltipBond.ToolText(gridBOND) = oPFE.getREQUIREMENT
        End If
    End If
End Sub

Private Sub gridTOP_SelChange(Cancel As Integer)
    If fichaTop <> "" Then
        Dim oPFE As New clsPlasma_ficha_estructura
        If oPFE.Carga(fichaTop, xTOP(gridTOP.Row, ColsGrid.ENSAYO_ID)) Then
            tooltipTop.ToolText(gridTOP) = oPFE.getREQUIREMENT
        End If
    End If
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

Private Sub txtDatos_Change(Index As Integer)
    
    Select Case Index
        Case 34, 36, 38, 44, 46, 48
            calcularMedia Index
'            If Index = 34 Or Index = 44 Then
                calcularDesviacion Index
'            End If
'            If Index = 36 Or Index = 46 Then
                calcularPorcentaje Index
'            End If
            If Index = 38 Or Index = 48 Then
                maximos Index
            End If
    End Select
End Sub
Private Sub maximos(campo)

    Dim lista() As String
   On Error GoTo maximos_Error

    lista = Split(txtDatos(campo), "-")
    Dim min As Single
    Dim Max As Single
    If UBound(lista) > 0 Then
        min = lista(0)
        Max = lista(0)
    Else
        If campo = 38 Then
            txtDatos(0) = ""
            txtDatos(1) = ""
        End If
        If campo = 48 Then
            txtDatos(2) = ""
            txtDatos(3) = ""
        End If
        Exit Sub
    End If
    ' MEDIA
    For i = LBound(lista) To UBound(lista)
        If IsNumeric(lista(i)) And lista(i) <> "" Then
            If lista(i) > Max Then
                Max = lista(i)
            End If
            If lista(i) < min Then
                min = lista(i)
            End If
        End If
    Next
    If campo = 38 Then
'        txtDatos(0) = min & """"
'        txtDatos(1) = Max & """"
        txtDatos(0) = min & txtUnidadEspesorBond
        txtDatos(1) = Max & txtUnidadEspesorBond
    End If
    If campo = 48 Then
'        txtDatos(2) = min & """"
'        txtDatos(3) = Max & """"
        txtDatos(2) = min & txtUnidadEspesorTop
        txtDatos(3) = Max & txtUnidadEspesorTop
    End If

   On Error GoTo 0
   Exit Sub

maximos_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure maximos of Formulario frmPlasma_Resultados"
End Sub
Private Sub calcularDesviacion(campo)
    Dim total As Single
    Dim CANTIDAD As Integer
    Dim sumatorio As Single
    Dim medida As Single
    Dim numero_medidas As Integer
    Dim resultado As Single

    media = 0
    sumatorio = 0
    numero_medidas = 0
    
    lista = Split(txtDatos(campo), "-")
    If UBound(lista) < 2 Then
        txtSD(campo + 1) = ""
        Exit Sub
    End If
    ' MEDIA
    For i = LBound(lista) To UBound(lista)
        If IsNumeric(lista(i)) And lista(i) <> "" Then
            total = total + lista(i)
            CANTIDAD = CANTIDAD + 1
        End If
    Next
    If CANTIDAD > 0 Then
        If campo = 38 Or campo = 48 Then
            media = total / CANTIDAD
        Else
            media = CInt(total / CANTIDAD)
        End If
    Else
        txtSD(campo + 1) = ""
        Exit Sub
    End If
    ' DESVIACION
    For i = LBound(lista) To UBound(lista)
        If IsNumeric(lista(i)) Then
            medida = lista(i)
            sumatorio = sumatorio + ((medida - media) * (medida - media))
            numero_medidas = numero_medidas + 1
        End If
    Next
    If campo = 38 Or campo = 48 Then
        txtSD(campo + 1) = formatear(Sqr(sumatorio / (numero_medidas - 1)), 5, 3)
    Else
        txtSD(campo + 1) = formatear(Sqr(sumatorio / (numero_medidas - 1)), 5, 1)
    End If
On Error GoTo 0
   Exit Sub

calcularDesviacion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularDesviacion of Formulario frmPlasma_Resultados"
End Sub

Private Sub calcularPorcentaje(campo)
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
    
    lista = Split(txtDatos(campo), "-")
    If UBound(lista) < 2 Then
        txtPOR(campo + 1) = ""
        Exit Sub
    End If
    ' MEDIA
    For i = LBound(lista) To UBound(lista)
        If IsNumeric(lista(i)) And lista(i) <> "" Then
            total = total + lista(i)
            CANTIDAD = CANTIDAD + 1
        End If
    Next
    If CANTIDAD > 0 Then
        media = total / CANTIDAD
    Else
        txtPOR(campo + 1) = ""
        Exit Sub
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
    txtPOR(campo + 1) = formatear(CStr(porcentaje), 3, 2)

   On Error GoTo 0
   Exit Sub

calcularDesviacion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularDesviacion of Formulario frmPlasma_Resultados"
End Sub
Private Sub calcularMedia(campo)
   On Error GoTo calcularMedia_Error

    If txtDatos(campo) <> "" Then
        Dim lista() As String
        Dim resultado As String
        Dim total As Single
        Dim CANTIDAD As Integer
        resultado = ""
        CANTIDAD = 0
        lista = Split(txtDatos(campo), "-")
        For i = LBound(lista) To UBound(lista)
            If IsNumeric(lista(i)) Then
                total = total + lista(i)
                CANTIDAD = CANTIDAD + 1
            End If
        Next
        If CANTIDAD > 0 Then
            If campo = 38 Or campo = 48 Then
                txtDatos(campo + 1) = formatear(CStr(total / CANTIDAD), 3, 3)
            Else
                txtDatos(campo + 1) = CInt(total / CANTIDAD)
            End If
        End If
        ' Unidad
        If Trim(txtUnidades) <> "" Then
            Dim unidades() As String
            unidades = Split(txtUnidades, ";")
            Select Case campo
            Case 32
                txtDatos(campo + 1) = Trim(txtDatos(campo + 1) & " " & unidades(0))
            Case 34
                txtDatos(campo + 1) = Trim(txtDatos(campo + 1) & " " & unidades(1))
            Case 36
                txtDatos(campo + 1) = Trim(txtDatos(campo + 1) & " " & unidades(2))
            Case 42
                txtDatos(campo + 1) = Trim(txtDatos(campo + 1) & " " & unidades(3))
            Case 44
                txtDatos(campo + 1) = Trim(txtDatos(campo + 1) & " " & unidades(4))
            Case 46
                txtDatos(campo + 1) = Trim(txtDatos(campo + 1) & " " & unidades(5))
            End Select
        End If
    Else
        txtDatos(campo + 1) = ""
    End If

   On Error GoTo 0
   Exit Sub

calcularMedia_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure calcularMedia of Formulario frmPlasma_Resultados"
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = vbYellow
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 38 Or Index = 48 Or Index = 8 Or Index = 9 Or Index = 10 Or Index = 11 Then
        If KeyAscii = 46 Then
             KeyAscii = 44
        End If
    End If
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 8 Or Index = 9 Or Index = 10 Or Index = 11 Then
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
        cmbnatype.desactivar
        txtDatos(51).Enabled = False
        txtDatos(52).Enabled = False
        txtDatos(53).Enabled = False
        txtDatos(54).Enabled = False
        txtDatos(55).Enabled = False
        txtDatos(31).Enabled = False
        txtDatos(41).Enabled = False
        gridBOND.Enabled = False
        gridTOP.Enabled = False
        txtDatos(32).Enabled = False
        txtDatos(34).Enabled = False
        txtDatos(35).Enabled = False
        txtDatos(42).Enabled = False
        txtDatos(44).Enabled = False
        txtDatos(46).Enabled = False
'        frmBondEspesor.Enabled = False
'        frmTopEspesor.Enabled = False
'        txtDatos(0).Enabled = False
'        txtDatos(1).Enabled = False
'        txtDatos(2).Enabled = False
'        txtDatos(3).Enabled = False
        txtDatos(38).Enabled = False
        txtDatos(48).Enabled = False
        chkResult.Enabled = False
        cmdok.visible = False
        frmPreparation.Enabled = False
        
        txtDatos(8).Enabled = False
        txtDatos(9).Enabled = False
        txtDatos(10).Enabled = False
        txtDatos(11).Enabled = False
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
        cmbnatype.activar
        txtDatos(51).Enabled = True
        txtDatos(52).Enabled = True
        txtDatos(53).Enabled = True
        txtDatos(54).Enabled = True
        txtDatos(55).Enabled = True
        txtDatos(31).Enabled = True
        txtDatos(41).Enabled = True
        If chkBondMetalografiaNR.Value = Unchecked Then
            gridBOND.Enabled = True
        End If
        If chkTopMetalografiaNR.Value = Unchecked Then
            gridTOP.Enabled = True
        End If
        txtDatos(32).Enabled = True
        txtDatos(34).Enabled = True
        txtDatos(35).Enabled = True
        txtDatos(42).Enabled = True
        txtDatos(44).Enabled = True
        txtDatos(46).Enabled = True
'        frmBondEspesor.Enabled = True
'        frmTopEspesor.Enabled = True
'        txtDatos(0).Enabled = True
'        txtDatos(1).Enabled = True
'        txtDatos(2).Enabled = True
'        txtDatos(3).Enabled = True
        txtDatos(38).Enabled = True
        txtDatos(48).Enabled = True
        chkResult.Enabled = True
        cmdok.visible = True
        frmPreparation.Enabled = True
    
        txtDatos(8).Enabled = True
        txtDatos(9).Enabled = True
        txtDatos(10).Enabled = True
        txtDatos(11).Enabled = True
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
    
    llenar_combo cmbUsuarioPreparation, New clsUsuarios, 0, frmUsuarios, ""
    
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
Private Sub cargarMicroEstructura(FICHA As Long, tipo As Integer)
    If FICHA = 0 Then
        inicializar_grid tipo
    Else
        Dim oPFE As New clsPlasma_ficha_estructura
        Dim i As Integer
        i = 0
        Dim rs As ADODB.Recordset
        Set rs = oPFE.Listado(FICHA)
        If rs.RecordCount > 0 Then
            Do
                If rs(2) <> "N/A" Then
                    If tipo = 1 Then
                        xBOND(i, ColsGrid.ENSAYO_ID) = CStr(rs(0))
                        xBOND(i, ColsGrid.ENSAYO) = CStr(rs(1))
                        xBOND(i, ColsGrid.VALOR) = ""
                        xBOND(i, ColsGrid.RESULT) = "0"
                    Else
                        xTOP(i, ColsGrid.ENSAYO_ID) = CStr(rs(0))
                        xTOP(i, ColsGrid.ENSAYO) = CStr(rs(1))
                        xTOP(i, ColsGrid.VALOR) = ""
                        xTOP(i, ColsGrid.RESULT) = "0"
                    End If
                    i = i + 1
                End If
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oPFE = Nothing
    End If
End Sub

Private Sub inicializar_grid(tipo)
   On Error GoTo inicializar_grid_Error

    If tipo = 1 Then
        xBOND.Clear
        xBOND.ReDim 0, filasGrid, 0, ColGrid
        xBOND.Clear
        Set gridBOND.Array = xBOND
        gridBOND.Refresh
    End If
    If tipo = 2 Then
        xTOP.Clear
        xTOP.ReDim 0, filasGrid, 0, ColGrid
        xTOP.Clear
        Set gridTOP.Array = xTOP
        gridTOP.Refresh
    End If
    
   On Error GoTo 0
   Exit Sub

inicializar_grid_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure inicializar_grid of Formulario frmCE_Recepcion_Nuevo2"
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
Private Sub toolBond()
   With tooltipBond
    Call .Create(Me)
    .MaxTipWidth = 600
    .Margin(ttMarginBottom) = 7
    .Margin(ttMarginTop) = 7
    .Margin(ttMarginLeft) = 5
    .Margin(ttMarginRight) = 5
    .DelayTime(ttDelayShow) = 10000
    .AddTool gridBOND
   End With
End Sub
Private Sub toolTop()
   With tooltipTop
    Call .Create(Me)
    .MaxTipWidth = 600
    .Margin(ttMarginBottom) = 7
    .Margin(ttMarginTop) = 7
    .Margin(ttMarginLeft) = 5
    .Margin(ttMarginRight) = 5
    .DelayTime(ttDelayShow) = 10000
    .AddTool gridTOP
   End With
End Sub

