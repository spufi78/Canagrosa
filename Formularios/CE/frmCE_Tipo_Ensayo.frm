VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmCE_Tipo_Ensayo 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Nuevo Tipo de Ensayo de Eficacia"
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13245
   Icon            =   "frmCE_Tipo_Ensayo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   13245
   StartUpPosition =   2  'CenterScreen
   Begin XtremeSuiteControls.TabControl TabControl1 
      Height          =   6180
      Left            =   45
      TabIndex        =   12
      Top             =   1845
      Width           =   13155
      _Version        =   851970
      _ExtentX        =   23204
      _ExtentY        =   10901
      _StockProps     =   68
      Appearance      =   9
      Color           =   4
      PaintManager.BoldSelected=   -1  'True
      PaintManager.ShowIcons=   -1  'True
      PaintManager.FixedTabWidth=   120
      ItemCount       =   5
      SelectedItem    =   1
      Item(0).Caption =   "Detalle del Ensayo"
      Item(0).ControlCount=   1
      Item(0).Control(0)=   "Frame1"
      Item(1).Caption =   "Criterio Aceptación"
      Item(1).ControlCount=   13
      Item(1).Control(0)=   "txtDatos(4)"
      Item(1).Control(1)=   "Label2"
      Item(1).Control(2)=   "Label3"
      Item(1).Control(3)=   "Label4"
      Item(1).Control(4)=   "txtDatos(7)"
      Item(1).Control(5)=   "cmbNormaCA"
      Item(1).Control(6)=   "lblCampos(7)"
      Item(1).Control(7)=   "Label5"
      Item(1).Control(8)=   "txtCaMetodo"
      Item(1).Control(9)=   "listaNormasCA"
      Item(1).Control(10)=   "cmbNormaCAS"
      Item(1).Control(11)=   "Command1"
      Item(1).Control(12)=   "Command2"
      Item(2).Caption =   "Normas"
      Item(2).ControlCount=   10
      Item(2).Control(0)=   "txtDatos(10)"
      Item(2).Control(1)=   "cmbNormaEnsayo"
      Item(2).Control(2)=   "lblCampos(12)"
      Item(2).Control(3)=   "Command3"
      Item(2).Control(4)=   "Command4"
      Item(2).Control(5)=   "Label6"
      Item(2).Control(6)=   "Label7"
      Item(2).Control(7)=   "listaNormas"
      Item(2).Control(8)=   "txtMetodoNorma"
      Item(2).Control(9)=   "cmbNormas"
      Item(3).Caption =   "Equipos"
      Item(3).ControlCount=   5
      Item(3).Control(0)=   "cmdAnadirEquipo"
      Item(3).Control(1)=   "cmdEliminarEquipo"
      Item(3).Control(2)=   "listaEquipos"
      Item(3).Control(3)=   "cmbEquipos"
      Item(3).Control(4)=   "Label1"
      Item(4).Caption =   "Reactivos"
      Item(4).ControlCount=   7
      Item(4).Control(0)=   "cmdAnadirReactivo"
      Item(4).Control(1)=   "cmdEliminarReactivo"
      Item(4).Control(2)=   "listaReactivos"
      Item(4).Control(3)=   "cmbReactivos"
      Item(4).Control(4)=   "cmbReactivosInternos"
      Item(4).Control(5)=   "lblCampos(26)"
      Item(4).Control(6)=   "lblCampos(25)"
      Begin VB.CommandButton Command4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   750
         Left            =   -57895
         Picture         =   "frmCE_Tipo_Ensayo.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   74
         Tag             =   "Elimina el campo seleccionado"
         Top             =   495
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   765
         Left            =   -57895
         Picture         =   "frmCE_Tipo_Ensayo.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   73
         Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
         Top             =   4095
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   765
         Left            =   12105
         Picture         =   "frmCE_Tipo_Ensayo.frx":1A5E
         Style           =   1  'Graphical
         TabIndex        =   72
         Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
         Top             =   4095
         Width           =   915
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   750
         Left            =   12105
         Picture         =   "frmCE_Tipo_Ensayo.frx":2328
         Style           =   1  'Graphical
         TabIndex        =   71
         Tag             =   "Elimina el campo seleccionado"
         Top             =   2340
         Width           =   915
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   10
         Left            =   -68335
         TabIndex        =   61
         Top             =   5760
         Visible         =   0   'False
         Width           =   5625
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   7
         Left            =   1215
         TabIndex        =   58
         Top             =   5805
         Width           =   5670
      End
      Begin XtremeSuiteControls.FlatEdit txtCaMetodo 
         Height          =   285
         Left            =   990
         TabIndex        =   57
         Top             =   5310
         Width           =   9915
         _Version        =   851970
         _ExtentX        =   17489
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   5
         UseVisualStyle  =   0   'False
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1065
         Index           =   4
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   855
         Width           =   11820
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   5640
         Left            =   -69910
         TabIndex        =   24
         Top             =   405
         Visible         =   0   'False
         Width           =   13005
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            Height          =   285
            Index           =   5
            Left            =   9945
            TabIndex        =   84
            Top             =   3105
            Visible         =   0   'False
            Width           =   2550
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Acreditaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1320
            Left            =   6525
            TabIndex        =   65
            Top             =   1755
            Width           =   5370
            Begin VB.CheckBox chkNADCAP 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "NADCAP"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00404040&
               Height          =   240
               Left            =   2925
               TabIndex        =   69
               Top             =   315
               Width           =   1140
            End
            Begin VB.OptionButton opENAC 
               BackColor       =   &H00C0C0C0&
               Caption         =   "NO ENAC"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   68
               Top             =   315
               Value           =   -1  'True
               Width           =   1410
            End
            Begin VB.OptionButton opENAC 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ENAC COMPLETA"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   67
               Top             =   630
               Width           =   1950
            End
            Begin VB.OptionButton opENAC 
               BackColor       =   &H00C0C0C0&
               Caption         =   "ENAC PARCIAL"
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   66
               Top             =   945
               Width           =   1590
            End
         End
         Begin VB.CheckBox chkambientales 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Condiciones Ambientales"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   6660
            TabIndex        =   29
            Top             =   3150
            Width           =   2130
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   17
            Left            =   1395
            TabIndex        =   36
            Top             =   2340
            Width           =   1515
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   18
            Left            =   1395
            TabIndex        =   35
            Top             =   1710
            Width           =   1515
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   19
            Left            =   1395
            TabIndex        =   34
            Top             =   2025
            Width           =   1515
         End
         Begin VB.CheckBox chkEspesor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Incluye Espesor"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   33
            Top             =   2700
            Width           =   1545
         End
         Begin VB.CheckBox chkLote 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Lote Probetas"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   32
            Top             =   2970
            Width           =   1365
         End
         Begin VB.TextBox txtDatos 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   645
            Index           =   6
            Left            =   1395
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   31
            Top             =   990
            Width           =   10470
         End
         Begin VB.Frame frmCondicionesAmbientales 
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            Height          =   1185
            Left            =   6525
            TabIndex        =   30
            Top             =   3105
            Width           =   5370
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   8
               Left            =   3960
               TabIndex        =   40
               Top             =   675
               Width           =   1335
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   3
               Left            =   2295
               TabIndex        =   39
               Top             =   675
               Width           =   1245
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   2
               Left            =   3960
               TabIndex        =   38
               Top             =   360
               Width           =   1335
            End
            Begin VB.TextBox txtDatos 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Index           =   1
               Left            =   2295
               TabIndex        =   37
               Top             =   360
               Width           =   1245
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   6
               Left            =   3735
               TabIndex        =   83
               Top             =   720
               Width           =   75
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "-"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   4
               Left            =   3735
               TabIndex        =   82
               Top             =   405
               Width           =   75
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Rango Humedad (% Hr)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   135
               TabIndex        =   81
               Top             =   720
               Width           =   1995
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Rango Temperatura (°C)"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   80
               Top             =   405
               Width           =   2070
            End
         End
         Begin VB.CheckBox chkDuplicado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Realizar por duplicado"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   28
            Top             =   3240
            Width           =   2220
         End
         Begin VB.CheckBox chkEsPintura 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ensayo de Pintura"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   27
            Top             =   3510
            Width           =   2085
         End
         Begin VB.CheckBox CHKESSIMPLIFICADO 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ensayo Simplificado"
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   90
            TabIndex        =   26
            Top             =   3780
            Width           =   2085
         End
         Begin VB.CheckBox CHKESSUBCONTRATABLE 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Subcontratable"
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
            Height          =   240
            Left            =   90
            TabIndex        =   25
            Top             =   4050
            Width           =   2085
         End
         Begin pryCombo.miCombo cmbPNT 
            Height          =   330
            Left            =   1395
            TabIndex        =   41
            Top             =   270
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbFormula 
            Height          =   330
            Left            =   1395
            TabIndex        =   42
            Top             =   630
            Width           =   11535
            _ExtentX        =   20346
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbunidades 
            Height          =   330
            Left            =   3285
            TabIndex        =   70
            Top             =   4590
            Width           =   9600
            _ExtentX        =   16933
            _ExtentY        =   582
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Duración (h:m)"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   13
            Left            =   90
            TabIndex        =   49
            Top             =   2385
            Width           =   1260
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Rango Min."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   14
            Left            =   90
            TabIndex        =   48
            Top             =   1755
            Width           =   990
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Rango Max."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   15
            Left            =   90
            TabIndex        =   47
            Top             =   2070
            Width           =   1035
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Unidad en ensayos con resultado numérico"
            Height          =   195
            Index           =   18
            Left            =   90
            TabIndex        =   46
            Top             =   4635
            Width           =   3060
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "PNT Vinculado"
            Height          =   195
            Index           =   20
            Left            =   90
            TabIndex        =   45
            Top             =   315
            Width           =   1080
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Observacion"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Index           =   21
            Left            =   90
            TabIndex        =   44
            Top             =   1170
            Width           =   1080
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Formula"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   43
            Top             =   675
            Width           =   555
         End
      End
      Begin VB.CommandButton cmdEliminarReactivo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   750
         Left            =   -57895
         Picture         =   "frmCE_Tipo_Ensayo.frx":2BF2
         Style           =   1  'Graphical
         TabIndex        =   18
         Tag             =   "Elimina el campo seleccionado"
         Top             =   585
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdAnadirReactivo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   765
         Left            =   -57895
         Picture         =   "frmCE_Tipo_Ensayo.frx":34BC
         Style           =   1  'Graphical
         TabIndex        =   17
         Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
         Top             =   4365
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdEliminarEquipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   810
         Left            =   -57895
         Style           =   1  'Graphical
         TabIndex        =   14
         Tag             =   "Elimina el campo seleccionado"
         Top             =   810
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.CommandButton cmdAnadirEquipo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Añadir"
         Height          =   810
         Left            =   -57850
         Style           =   1  'Graphical
         TabIndex        =   13
         Tag             =   "Añade campo o modifica el campo existente con el mismo nombre"
         Top             =   4500
         Visible         =   0   'False
         Width           =   915
      End
      Begin MSComctlLib.ListView listaEquipos 
         Height          =   4530
         Left            =   -69865
         TabIndex        =   15
         Top             =   810
         Visible         =   0   'False
         Width           =   11880
         _ExtentX        =   20955
         _ExtentY        =   7990
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
         BackColor       =   16777215
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
         Left            =   -69865
         TabIndex        =   16
         Top             =   5490
         Visible         =   0   'False
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   582
      End
      Begin MSComctlLib.ListView listaReactivos 
         Height          =   4650
         Left            =   -69865
         TabIndex        =   19
         Top             =   495
         Visible         =   0   'False
         Width           =   11880
         _ExtentX        =   20955
         _ExtentY        =   8202
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
         Left            =   -69130
         TabIndex        =   20
         Top             =   5250
         Visible         =   0   'False
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbReactivosInternos 
         Height          =   330
         Left            =   -69130
         TabIndex        =   21
         Top             =   5580
         Visible         =   0   'False
         Width           =   10920
         _ExtentX        =   19262
         _ExtentY        =   582
      End
      Begin MSComctlLib.ListView listaNormasCA 
         Height          =   2580
         Left            =   90
         TabIndex        =   53
         Top             =   2340
         Width           =   11880
         _ExtentX        =   20955
         _ExtentY        =   4551
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
      Begin pryCombo.miCombo cmbNormaCAS 
         Height          =   330
         Left            =   990
         TabIndex        =   54
         Top             =   4950
         Width           =   10950
         _ExtentX        =   19315
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbNormaCA 
         Height          =   330
         Left            =   6930
         TabIndex        =   59
         Top             =   5805
         Width           =   6120
         _ExtentX        =   10795
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbNormaEnsayo 
         Height          =   330
         Left            =   -62665
         TabIndex        =   62
         Top             =   5760
         Visible         =   0   'False
         Width           =   5760
         _ExtentX        =   10160
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.FlatEdit txtMetodoNorma 
         Height          =   285
         Left            =   -69010
         TabIndex        =   75
         Top             =   5310
         Visible         =   0   'False
         Width           =   9915
         _Version        =   851970
         _ExtentX        =   17489
         _ExtentY        =   503
         _StockProps     =   77
         BackColor       =   -2147483643
         Appearance      =   5
         UseVisualStyle  =   0   'False
      End
      Begin MSComctlLib.ListView listaNormas 
         Height          =   4380
         Left            =   -69910
         TabIndex        =   76
         Top             =   495
         Visible         =   0   'False
         Width           =   11880
         _ExtentX        =   20955
         _ExtentY        =   7726
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
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
      Begin pryCombo.miCombo cmbNormas 
         Height          =   330
         Left            =   -69010
         TabIndex        =   77
         Top             =   4950
         Visible         =   0   'False
         Width           =   10950
         _ExtentX        =   19315
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.Label Label7 
         Height          =   195
         Left            =   -69910
         TabIndex        =   79
         Top             =   4995
         Visible         =   0   'False
         Width           =   465
         _Version        =   851970
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Norma"
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label6 
         Height          =   195
         Left            =   -69910
         TabIndex        =   78
         Top             =   5355
         Visible         =   0   'False
         Width           =   540
         _Version        =   851970
         _ExtentX        =   953
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Método"
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label5 
         Height          =   195
         Left            =   90
         TabIndex        =   64
         Top             =   2070
         Width           =   2400
         _Version        =   851970
         _ExtentX        =   4233
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Normas del Criterio de Aceptación"
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Norma de ensayo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   -69910
         TabIndex        =   63
         Top             =   5850
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Norma (C.A)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   60
         Top             =   5850
         Width           =   1035
      End
      Begin XtremeSuiteControls.Label Label4 
         Height          =   195
         Left            =   90
         TabIndex        =   56
         Top             =   5355
         Width           =   540
         _Version        =   851970
         _ExtentX        =   953
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Método"
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label3 
         Height          =   195
         Left            =   90
         TabIndex        =   55
         Top             =   4995
         Width           =   465
         _Version        =   851970
         _ExtentX        =   820
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Norma"
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label2 
         Height          =   195
         Left            =   90
         TabIndex        =   52
         Top             =   585
         Width           =   1560
         _Version        =   851970
         _ExtentX        =   2752
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Criterio de Aceptación"
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin XtremeSuiteControls.Label Label1 
         Height          =   195
         Left            =   -61450
         TabIndex        =   50
         Top             =   585
         Visible         =   0   'False
         Width           =   3450
         _Version        =   851970
         _ExtentX        =   6085
         _ExtentY        =   344
         _StockProps     =   79
         Caption         =   "Marque los equipos que deben salir en el informe"
         ForeColor       =   255
         Transparent     =   -1  'True
         AutoSize        =   -1  'True
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Interno"
         Height          =   195
         Index           =   25
         Left            =   -69805
         TabIndex        =   23
         Top             =   5625
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Externos"
         Height          =   195
         Index           =   26
         Left            =   -69805
         TabIndex        =   22
         Top             =   5280
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   9585
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8055
      Width           =   1365
   End
   Begin VB.CheckBox chkActivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "ACTIVO"
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
      Height          =   240
      Left            =   11610
      TabIndex        =   10
      Top             =   90
      Width           =   1545
   End
   Begin VB.Frame Frame2 
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
      ForeColor       =   &H00FF0000&
      Height          =   1365
      Left            =   45
      TabIndex        =   5
      Top             =   450
      Width           =   13155
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1530
         TabIndex        =   0
         Top             =   270
         Width           =   11505
      End
      Begin pryCombo.miCombo cmbTA 
         Height          =   330
         Left            =   1530
         TabIndex        =   1
         Top             =   585
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbPB 
         Height          =   330
         Left            =   1530
         TabIndex        =   2
         Top             =   945
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   8
         Top             =   315
         Width           =   660
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   7
         Top             =   990
         Width           =   705
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Análisis"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   17
         Left            =   135
         TabIndex        =   6
         Top             =   630
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   11025
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8055
      Width           =   1050
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo Tipo de Ensayo de  Eficacia"
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
      Height          =   285
      Left            =   0
      TabIndex        =   9
      Top             =   45
      Width           =   13170
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   420
      Left            =   0
      Top             =   0
      Width           =   13275
   End
End
Attribute VB_Name = "frmCE_Tipo_Ensayo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmdAnadirReactivo_Click()
    ' Externo (E)
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
    ' Interno (I)
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
End Sub

Private Sub cmdEliminarReactivo_Click()
    If listaReactivos.ListItems.Count > 0 Then
        listaReactivos.ListItems.Remove listaReactivos.selectedItem.Index
        cmbReactivos.limpiar
        cmbReactivosInternos.limpiar
    End If
End Sub



Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_CE_TIPO_ENSAYO_EFICACIA
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Tipo Ensayo " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub

Private Sub chkambientales_Click()
    If chkambientales.Value = Checked Then
        frmCondicionesAmbientales.Enabled = True
    Else
        frmCondicionesAmbientales.Enabled = False
    End If
End Sub

'Private Sub chkEquipo_Click()
'    If chkEquipo.Value = Checked Then
'        frmEquipos.Enabled = True
'    Else
'        frmEquipos.Enabled = False
'        listaEquipos.ListItems.Clear
'    End If
'End Sub

Private Sub cmdAnadirEquipo_Click()
    If cmbEquipos.getPK_SALIDA <> 0 Then
        Dim oEquipo As New clsEquipos
        oEquipo.Carga cmbEquipos.getPK_SALIDA
        With listaEquipos.ListItems.Add(, , oEquipo.getID_EQUIPO)
            .SubItems(1) = oEquipo.getNOMBRE
            .SubItems(2) = oEquipo.getSERIE
        End With
        listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
        listaEquipos.ListItems(listaEquipos.ListItems.Count).EnsureVisible
        cmbEquipos.limpiar
    End If
End Sub

Private Sub cmdEliminarEquipo_Click()
    If listaEquipos.ListItems.Count > 0 Then
        listaEquipos.ListItems.Remove listaEquipos.selectedItem.Index
    End If

End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    If validar = True Then
      Dim tipo_ensayo As Long
      Dim oce_tipo_ensayo As New clsCe_tipos_ensayos
      With oce_tipo_ensayo
            .setNOMBRE = txtDatos(0)
'            .setNOMBRE_INGLES = txtDatos(1)
            .setPROCESO_BASE_ID = cmbPB.getPK_SALIDA
            .setCRITERIO = txtDatos(4)
            .setINCLUYE_ESPESOR = chkEspesor.Value
            .setDUPLICADO = chkDuplicado.Value
            .setACTIVO = chkActivo.Value
            .setENAC = 0
            If opENAC(1).Value = True Then
                .setENAC = 1
            ElseIf opENAC(2).Value = True Then
                .setENAC = 2
            End If
            .setNADCAP = chkNADCAP.Value
            .setLOTE_PROBETAS = chkLote.Value
            .setSEGUN_NORMA = txtDatos(7)
            .setRANGO_MIN = txtDatos(18)
            .setRANGO_MAX = txtDatos(19)
            .setHORAS = txtDatos(17)
            .setNORMA = txtDatos(10)
            
            .setNORMA_ID_CA = cmbNormaCA.getPK_SALIDA
            .setNORMA_ID_ENSAYO = cmbNormaEnsayo.getPK_SALIDA
            
            .setTIPO_ANALISIS_ID = cmbTA.getPK_SALIDA
            .setPNT_VINCULADO = cmbPNT.getPK_SALIDA
            .setUNIDAD_ID = cmbUnidades.getPK_SALIDA
            
            .setCONDICIONES_AMBIENTALES = chkambientales.Value
            If chkambientales.Value = Checked Then
                .setRANGO_MIN_TA = numerico_bd(txtDatos(1))
                .setRANGO_MAX_TA = numerico_bd(txtDatos(2))
                .setRANGO_MIN_HUMEDAD = numerico_bd(txtDatos(3))
                .setRANGO_MAX_HUMEDAD = numerico_bd(txtDatos(8))
            Else
                .setRANGO_MIN_TA = numerico_bd("")
                .setRANGO_MAX_TA = numerico_bd("")
                .setRANGO_MIN_HUMEDAD = numerico_bd("")
                .setRANGO_MAX_HUMEDAD = numerico_bd("")
            End If
            .setOBSERVACIONES = Trim(txtDatos(6))
            .setFORMULA_ID = cmbFormula.getPK_SALIDA
            Dim g As Integer
            .setINCLUYE_EQUIPO = 0
            For g = 1 To listaEquipos.ListItems.Count
                If listaEquipos.ListItems(g).Checked = True Then
                    .setINCLUYE_EQUIPO = 1
                End If
            Next
            .setESPINTURA = chkEsPintura.Value
            .setESSIMPLIFICADO = CHKESSIMPLIFICADO.Value
            'MXXXX-I
            .setES_SUBCONTRATABLE = CHKESSUBCONTRATABLE.Value
            'MXXXX-F
      End With
      Dim ohc As New clsHistorial_cambios
      If PK = 0 Then
        If MsgBox("Va a introducir el nuevo tipo de ensayo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            tipo_ensayo = oce_tipo_ensayo.Insertar
            If tipo_ensayo <> 0 Then
                With ohc
                    .setTIPO = HC_TIPOS.HC_CE_TIPO_ENSAYO_EFICACIA
                    .setIDENTIFICADOR = tipo_ensayo
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = HC_CREACION
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el tipo de ensayo. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            tipo_ensayo = PK
            frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación."
            frmMotivo.Show 1
            If Trim(MOTIVO) = "" Then
                MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                Exit Sub
            End If
            If oce_tipo_ensayo.Modificar(PK) = False Then
                Exit Sub
            End If
            If PK <> 0 Then
                With ohc
                    .setTIPO = HC_TIPOS.HC_CE_TIPO_ENSAYO_EFICACIA
                    .setIDENTIFICADOR = PK
                    .setIDENTIFICADOR_TEXTO = txtDatos(0)
                    .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                    .setMOTIVO = Trim(MOTIVO)
                    .Insertar
                End With
            End If
        Else
            Exit Sub
        End If
      End If
      Set ohc = Nothing
      Dim i As Integer
      ' NORMAS_CA
      Dim oNCA As New clsCe_tipos_ensayos_normas_ca
      oNCA.Eliminar tipo_ensayo
      For i = 1 To listaNormasCA.ListItems.Count
        With oNCA
            .setTIPO_ENSAYO_ID = tipo_ensayo
            .setNORMA_ID = listaNormasCA.ListItems(i).Text
            .setMETODO = listaNormasCA.ListItems(i).SubItems(4)
            .setORDEN = i
            .Insertar
        End With
      Next
      ' NORMAS
      Dim oNC As New clsCe_tipos_ensayos_normas
      oNC.Eliminar tipo_ensayo
      For i = 1 To listaNormas.ListItems.Count
        With oNC
            .setTIPO_ENSAYO_ID = tipo_ensayo
            .setNORMA_ID = listaNormas.ListItems(i).Text
            .setMETODO = listaNormas.ListItems(i).SubItems(4)
            .setORDEN = i
            .Insertar
        End With
      Next
      ' EQUIPOS
      oce_tipo_ensayo.Equipos_Eliminar tipo_ensayo
'      If chkEquipo.Value = Checked Then
        For i = 1 To listaEquipos.ListItems.Count
          oce_tipo_ensayo.Equipos_Insertar tipo_ensayo, listaEquipos.ListItems(i), i, listaEquipos.ListItems(i).Checked
        Next
'      End If
      ' Reactivos
      Dim oCER As New clsCe_tipos_ensayos_botes_ex
      oCER.Eliminar tipo_ensayo
      For i = 1 To listaReactivos.ListItems.Count
        With oCER
            .setTIPO_ENSAYO_ID = tipo_ensayo
            .setBOTE_EX_ID = listaReactivos.ListItems(i).Text
            .setTIPO = listaReactivos.ListItems(i).SubItems(3)
            .setORDEN = i
            .Insertar
        End With
      Next
      ' MENSAJE DE SALIDA
      If PK = 0 Then
          MsgBox "El tipo de ensayo se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
      Else
          MsgBox "El tipo de ensayo se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
      End If
      Unload Me
    End If
    Exit Sub
fallo:
    error_grave ("Error al insertar el tipo de ensayo : " & Err.Description)
End Sub

Private Sub Command1_Click()
    If listaNormasCA.ListItems.Count > 0 Then
        listaNormasCA.ListItems.Remove listaNormasCA.selectedItem.Index
    End If
End Sub

Private Sub Command2_Click()
   On Error GoTo Command2_Click_Error

    If cmbNormaCAS.getPK_SALIDA <> 0 Then
        Dim oNorma As New clsCa_normas
        oNorma.Carga cmbNormaCAS.getPK_SALIDA
        With listaNormasCA.ListItems.Add(, , oNorma.getID_NORMA)
            .SubItems(1) = oNorma.getNOMBRE
            .SubItems(2) = oNorma.getCODIGO
            .SubItems(3) = oNorma.getEDICION
            .SubItems(4) = txtCaMetodo
            .SubItems(5) = IIf(oNorma.getENAC = 0, "*", "")
        End With
        cmbNormaCAS.limpiar
        txtCaMetodo = ""
    End If

   On Error GoTo 0
   Exit Sub

Command2_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Command2_Click of Formulario frmCE_Tipo_Ensayo"
End Sub

Private Sub Command3_Click()
    If cmbNormas.getPK_SALIDA <> 0 Then
        Dim oNorma As New clsCa_normas
        oNorma.Carga cmbNormas.getPK_SALIDA
        With listaNormas.ListItems.Add(, , oNorma.getID_NORMA)
            .SubItems(1) = oNorma.getNOMBRE
            .SubItems(2) = oNorma.getCODIGO
            .SubItems(3) = oNorma.getEDICION
            .SubItems(4) = txtMetodoNorma
            .SubItems(5) = IIf(oNorma.getENAC = 0, "*", "")
        End With
        cmbNormas.limpiar
        txtMetodoNorma = ""
    End If
End Sub

Private Sub Command4_Click()
    If listaNormas.ListItems.Count > 0 Then
        listaNormas.ListItems.Remove listaNormas.selectedItem.Index
    End If
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_combos
    If PK <> 0 Then
        cargar_datos
    Else
        chkActivo.Value = Checked
    End If
    TabControl1.Item(0).Selected = True

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmCE_Tipo_Ensayo"
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub listaEquipos_DblClick()
    If listaEquipos.ListItems.Count > 0 Then
        frmEquipoEdicion.PK = listaEquipos.ListItems(listaEquipos.selectedItem.Index)
        frmEquipoEdicion.Show 1
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 4 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
    If Index = 1 Or Index = 2 Or Index = 3 Or Index = 8 Then
        KeyAscii = KeyAscii_SoloDecimal(txtDatos(Index), KeyAscii, True)
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
'Private Function validar_ensayo() As Boolean
'    validar_ensayo = True
'    If Trim(txtDatos(14)) = "" Then
'        MsgBox "Debe darle un nombre al ensayo.", vbInformation, App.Title
'        validar_ensayo = False
'        Exit Function
'    End If
'End Function

Private Sub cargar_datos()
    Dim oce_tipo_ensayo As New clsCe_tipos_ensayos
    With oce_tipo_ensayo
        If .Carga(PK) = True Then
            lbltitulo.Caption = "Modificación control de eficacia : " & .getNOMBRE
            Me.Caption = lbltitulo.Caption
            txtDatos(0) = .getNOMBRE
'            txtDatos(1) = .getNOMBRE_INGLES
            cmbPB.MostrarElemento .getPROCESO_BASE_ID
            txtDatos(4) = .getCRITERIO
            chkEspesor.Value = .getINCLUYE_ESPESOR
            chkDuplicado.Value = .getDUPLICADO
            chkActivo.Value = .getACTIVO
            opENAC(.getENAC).Value = True
            chkNADCAP.Value = .getNADCAP
            chkLote.Value = .getLOTE_PROBETAS
            txtDatos(7) = .getSEGUN_NORMA
            txtDatos(18) = .getRANGO_MIN
            txtDatos(19) = .getRANGO_MAX
            txtDatos(17) = .getHORAS
            txtDatos(10) = .getNORMA
            If .getNORMA_ID_CA <> 0 Then
                cmbNormaCA.MostrarElemento .getNORMA_ID_CA
            End If
            If .getNORMA_ID_ENSAYO <> 0 Then
                cmbNormaEnsayo.MostrarElemento .getNORMA_ID_ENSAYO
            End If
'            If .getEQUIPO_ID <> 0 Then
'                cmbEquipo.MostrarElemento .getEQUIPO_ID
'            End If
            If .getPNT_VINCULADO <> 0 Then
                cmbPNT.MostrarElemento .getPNT_VINCULADO
            End If
'            txtDatos(12) = .getEQUIPO
'            txtDatos(13) = .getEQUIPO_INGLES
            cmbTA.MostrarElemento .getTIPO_ANALISIS_ID
            cmbUnidades.MostrarElemento .getUNIDAD_ID
            ' Condiciones ambientales
            If .getCONDICIONES_AMBIENTALES = 1 Then
                chkambientales.Value = Checked
                frmCondicionesAmbientales.Enabled = True
                txtDatos(1) = .getRANGO_MIN_TA
                txtDatos(2) = .getRANGO_MAX_TA
                txtDatos(3) = .getRANGO_MIN_HUMEDAD
                txtDatos(8) = .getRANGO_MAX_HUMEDAD
            Else
                chkambientales.Value = Unchecked
                frmCondicionesAmbientales.Enabled = False
            End If
            txtDatos(6) = .getOBSERVACIONES
            cmbFormula.MostrarElemento .getFORMULA_ID
            
            cargar_normas_ca PK
            cargar_normas PK
            ' EQUIPOS
'            If .getINCLUYE_EQUIPO = 0 Then
'                chkEquipo.Value = Unchecked
'            Else
                cargar_equipos PK
'            End If
            ' REACTIVOS
            cargar_reactivos PK
            chkEsPintura.Value = .getESPINTURA
            CHKESSIMPLIFICADO.Value = .getESSIMPLIFICADO
            'MXXXX-I
            CHKESSUBCONTRATABLE.Value = .getES_SUBCONTRATABLE
            'MXXXX-F
        End If
    End With
    Set oce_tipo_ensayo = Nothing
End Sub

Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe introducir una descripción del ensayo.", vbCritical, App.Title
        validar = False
        txtDatos(0).SetFocus
        Exit Function
    End If
    If cmbTA.getTEXTO = "" Then
        MsgBox "Debe introducir un tipo de análisis.", vbCritical, App.Title
        validar = False
        cmbTA.SetFocus
        Exit Function
    End If
    If cmbPB.getTEXTO = "" Then
        MsgBox "Debe introducir un proceso.", vbCritical, App.Title
        validar = False
        cmbPB.SetFocus
        Exit Function
    End If
    If chkambientales.Value = Checked Then
        If Trim(txtDatos(1)) = "" Or Trim(txtDatos(2)) = "" Or Trim(txtDatos(3)) = "" Or Trim(txtDatos(8)) = "" Then
            MsgBox "Debe introducir todos los campos de condiciones ambientales.", vbCritical, App.Title
            validar = False
            txtDatos(1).SetFocus
            Exit Function
        End If
    End If
End Function
Private Sub cargar_combos()
    llenar_combo cmbTA, New clsTipos_analisis, 0, frmTA_Detalle, "ANULADO=0"
    llenar_combo cmbPB, New clsProceso_base, 0, Me, ""
    llenar_combo cmbNormaCA, New clsCa_normas, 0, frmCA_Normas, ""
    llenar_combo cmbNormaCAS, New clsCa_normas, 0, frmCA_Normas, ""
    llenar_combo cmbNormas, New clsCa_normas, 0, frmCA_Normas, ""
    llenar_combo cmbNormaEnsayo, New clsCa_normas, 0, frmCA_Normas, ""
    llenar_combo cmbEquipos, New clsEquipos, 0, frmEquipoEdicion, ""
    llenar_combo cmbPNT, New clsCa_documentos, 0, frmCA_Documento, " NOMBRE LIKE '%PNT%' "
    llenar_combo cmbFormula, New clsFormulas, 0, frmFORMULA_Detalle, ""
    llenar_combo cmbReactivos, New clsBotes_ex, 0, Me, "AND ABIERTO = 1 AND FINALIZADO = 0 "
    llenar_combo cmbReactivosInternos, New clsRpr_botes, 0, Me, " AND ISNULL(fecha_fin)"
    llenar_combo cmbUnidades, New clsUnidades, 0, Me, ""
End Sub
Private Sub cabecera()
    With listaNormasCA.ColumnHeaders
        .Add , , "NORMA_ID", 1, lvwColumnLeft
        .Add , , "Norma", 5000, lvwColumnLeft
        .Add , , "Código", 2000, lvwColumnCenter
        .Add , , "Edición", 1000, lvwColumnCenter
        .Add , , "Método", 3000, lvwColumnCenter
        .Add , , "Enac", 600, lvwColumnCenter
    End With
    With listaNormas.ColumnHeaders
        .Add , , "NORMA_ID", 1, lvwColumnLeft
        .Add , , "Norma", 5000, lvwColumnLeft
        .Add , , "Código", 2000, lvwColumnCenter
        .Add , , "Edición", 1000, lvwColumnCenter
        .Add , , "Método", 3000, lvwColumnCenter
        .Add , , "Enac", 600, lvwColumnCenter
    End With
    With listaEquipos.ColumnHeaders
        .Add , , "NºEquipo", 1400, lvwColumnLeft
        .Add , , "Nombre", 7000, lvwColumnLeft
        .Add , , "NºSerie", 3000, lvwColumnCenter
    End With
    With listaReactivos.ColumnHeaders
        .Add , , "Número", 1500, lvwColumnLeft
        .Add , , "Reactivo", 7000, lvwColumnLeft
        .Add , , "Caducidad", 1500, lvwColumnCenter
        .Add , , "Tipo", 1000, lvwColumnCenter
    End With
End Sub
Private Sub cargar_normas_ca(CE As Long)
    Dim oTE_N As New clsCe_tipos_ensayos_normas_ca
    Set rs = oTE_N.Listado(CE)
    If rs.RecordCount > 0 Then
        Do
               With listaNormasCA.ListItems.Add(, , rs(0))
                  .SubItems(1) = rs(1)
                  .SubItems(2) = rs(2)
                  .SubItems(3) = rs(3)
                  .SubItems(4) = rs(4)
                  .SubItems(5) = rs(5)
               End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub
Private Sub cargar_normas(CE As Long)
    Dim oTEN As New clsCe_tipos_ensayos_normas
    Set rs = oTEN.Listado(CE)
    If rs.RecordCount > 0 Then
        Do
               With listaNormas.ListItems.Add(, , rs(0))
                  .SubItems(1) = rs(1)
                  .SubItems(2) = rs(2)
                  .SubItems(3) = rs(3)
                  .SubItems(4) = rs(4)
                  .SubItems(5) = rs(5)
               End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub

Private Sub cargar_equipos(CE As Long)
    Dim oCE As New clsCe_tipos_ensayos
    Dim rs As ADODB.Recordset
'    chkEquipo.Value = Checked
    Set rs = oCE.Equipos_Listado(CE)
    If rs.RecordCount > 0 Then
        Do
            With listaEquipos.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
            End With
            If rs("EN_INFORME") = 1 Then
                listaEquipos.ListItems(listaEquipos.ListItems.Count).Checked = True
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oCE = Nothing
End Sub

Private Sub cargar_reactivos(CE As Long)
    ' Reactivos
    Dim oCER As New clsCe_tipos_ensayos_botes_ex
    Dim oReactivo As New clsBotes_ex
    Dim oTb As New clsTipos_bote_ex
    Dim oTR As New clsTipos_reactivo_ex
    
    Dim oRPR As New clsRpr_botes
    Dim oTRPR As New clsRPR_Tipos
    Set rs = oCER.Listado(CE)
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
                    .SubItems(1) = oTRPR.getNOMBRE
                    .SubItems(2) = Format(oRPR.getFECHA_CADUCIDAD, "DD-MM-YYYY")
                    .SubItems(3) = "I"
                End With
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub

