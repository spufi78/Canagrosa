VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmCA_Documento 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documento de Calidad"
   ClientHeight    =   10500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15075
   Icon            =   "frmCA_Documento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10500
   ScaleWidth      =   15075
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdValoracion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cuestionarios de Valoración"
      Height          =   825
      Left            =   8910
      Picture         =   "frmCA_Documento.frx":2AFA
      Style           =   1  'Graphical
      TabIndex        =   69
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   9585
      Width           =   2100
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Requerimientos para la creación/modificación del PNT"
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
      Height          =   735
      Index           =   3
      Left            =   45
      TabIndex        =   64
      Top             =   7965
      Width           =   9105
      Begin MSDataListLib.DataCombo cmbFamilia_Req 
         Height          =   315
         Left            =   1755
         TabIndex        =   65
         Top             =   270
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin XtremeSuiteControls.PushButton cmbRequisitos 
         Height          =   480
         Left            =   7560
         TabIndex        =   67
         Top             =   180
         Width           =   1470
         _Version        =   851970
         _ExtentX        =   2593
         _ExtentY        =   847
         _StockProps     =   79
         Caption         =   "Requisitos"
         Appearance      =   5
         Picture         =   "frmCA_Documento.frx":33C4
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Requerimientos"
         Height          =   195
         Index           =   13
         Left            =   150
         TabIndex        =   66
         Top             =   315
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   825
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   56
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   9585
      Width           =   1875
   End
   Begin VB.CommandButton cmdEquipos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Equipos Referenciados"
      Height          =   825
      Left            =   5130
      Picture         =   "frmCA_Documento.frx":9C26
      Style           =   1  'Graphical
      TabIndex        =   50
      Top             =   9585
      Width           =   1875
   End
   Begin VB.CommandButton cmdAnotaciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anotaciones"
      Height          =   825
      Left            =   3420
      MaskColor       =   &H000000FF&
      Picture         =   "frmCA_Documento.frx":A4F0
      Style           =   1  'Graphical
      TabIndex        =   49
      Top             =   9585
      Width           =   1680
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incluir en la lista de formación "
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
      Height          =   735
      Index           =   0
      Left            =   45
      TabIndex        =   45
      Top             =   8775
      Width           =   9105
      Begin VB.CheckBox chkFormacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2790
         TabIndex        =   46
         Top             =   0
         Width           =   285
      End
      Begin MSDataListLib.DataCombo cmbPlanes 
         Height          =   315
         Left            =   1755
         TabIndex        =   48
         Top             =   270
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Plan de Formación"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   47
         Top             =   315
         Width           =   1335
      End
   End
   Begin VB.CheckBox chkVinculo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documento Vínculado"
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
      Left            =   5850
      TabIndex        =   42
      Top             =   6570
      Width           =   2220
   End
   Begin VB.CommandButton cmdPNT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Versiones"
      Height          =   825
      Left            =   1710
      Picture         =   "frmCA_Documento.frx":ADBA
      Style           =   1  'Graphical
      TabIndex        =   39
      Top             =   9585
      Width           =   1680
   End
   Begin VB.CommandButton cmdMostrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar Documento"
      Height          =   825
      Index           =   0
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   9585
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Copia Controlada"
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
      Height          =   1320
      Index           =   2
      Left            =   45
      TabIndex        =   34
      Top             =   6570
      Width           =   5685
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   4
         Left            =   1125
         MaxLength       =   255
         TabIndex        =   16
         Top             =   810
         Width           =   4485
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   2
         Left            =   1125
         MaxLength       =   255
         TabIndex        =   15
         Top             =   360
         Width           =   4485
      End
      Begin VB.CheckBox chkcopia 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1710
         TabIndex        =   14
         Top             =   0
         Width           =   285
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Número"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   36
         Top             =   855
         Width           =   585
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Laboratorio"
         Height          =   195
         Index           =   10
         Left            =   135
         TabIndex        =   35
         Top             =   405
         Width           =   825
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   9000
      Top             =   9585
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   14010
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   9585
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del documento"
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
      Height          =   5820
      Index           =   1
      Left            =   45
      TabIndex        =   22
      Top             =   720
      Width           =   9135
      Begin VB.CheckBox chkMTL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "MTL"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4050
         TabIndex        =   68
         Top             =   2385
         Width           =   870
      End
      Begin VB.CheckBox chkNoEdicion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "No aplica edición (NA)"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4950
         TabIndex        =   41
         Top             =   2385
         Width           =   1995
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   5
         Left            =   7515
         MaxLength       =   255
         TabIndex        =   6
         Top             =   1890
         Width           =   1515
      End
      Begin VB.CheckBox chkEQA 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "EQA"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   3150
         TabIndex        =   9
         Top             =   2385
         Width           =   750
      End
      Begin VB.CheckBox chkuso 
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Caption         =   "Documento en USO"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   5445
         Width           =   4200
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H000000FF&
         Height          =   285
         Index           =   1
         Left            =   5130
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1890
         Width           =   1605
      End
      Begin VB.CheckBox chkNADCAP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "NADCAP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   2025
         TabIndex        =   8
         Top             =   2385
         Width           =   960
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1905
         Index           =   10
         Left            =   1080
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   2700
         Width           =   7965
      End
      Begin VB.CheckBox chkENAC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENAC"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   1080
         TabIndex        =   7
         Top             =   2385
         Width           =   810
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   3
         Left            =   1080
         MaxLength       =   255
         TabIndex        =   4
         Top             =   1890
         Width           =   3000
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1080
         TabIndex        =   0
         Top             =   405
         Width           =   7965
      End
      Begin MSDataListLib.DataCombo cmbfamilias 
         Height          =   315
         Left            =   1080
         TabIndex        =   1
         Top             =   765
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   4680
         Width           =   3405
         _ExtentX        =   6006
         _ExtentY        =   556
         _Version        =   393216
         Locked          =   -1  'True
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo cmbSubfamilia 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   1125
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo cmbresponsables 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   1485
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
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
      Begin MSDataListLib.DataCombo cmbPlantilla 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   5040
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
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Plantilla"
         Height          =   195
         Index           =   12
         Left            =   90
         TabIndex        =   40
         Top             =   5130
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   9
         Left            =   90
         TabIndex        =   33
         Top             =   1530
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "SubFamilia"
         Height          =   195
         Index           =   8
         Left            =   90
         TabIndex        =   32
         Top             =   1170
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   29
         Top             =   4770
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   3
         Left            =   6975
         TabIndex        =   28
         Top             =   1935
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edición"
         Height          =   195
         Index           =   1
         Left            =   4500
         TabIndex        =   27
         Top             =   1935
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   26
         Top             =   1935
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalle"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   25
         Top             =   2925
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   23
         Top             =   810
         Width           =   480
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   480
         Width           =   555
      End
   End
   Begin VB.Frame frmVinculo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Enabled         =   0   'False
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
      Height          =   1320
      Left            =   5760
      TabIndex        =   30
      Top             =   6570
      Width           =   3390
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar Vínculo"
         Height          =   825
         Index           =   1
         Left            =   1845
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   360
         Width           =   1500
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   315
         TabIndex        =   31
         Top             =   270
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   465
         Index           =   0
         Left            =   2565
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   90
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Insertar Vínculo"
         Height          =   825
         Index           =   0
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   360
         Width           =   1635
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   510
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   315
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   17
         Top             =   675
         Visible         =   0   'False
         Width           =   1665
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   825
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   9585
      Width           =   1005
   End
   Begin Geslab.ControlPanelXP cpNormas 
      Height          =   4155
      Left            =   9225
      TabIndex        =   51
      Top             =   735
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   7329
      Caption         =   "Normas de Referencia"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   4155
      Begin XtremeSuiteControls.PushButton cmdEliminarNorma 
         Height          =   435
         Left            =   3555
         TabIndex        =   55
         Top             =   3645
         Width           =   2115
         _Version        =   851970
         _ExtentX        =   3731
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar Norma"
         Appearance      =   5
         Picture         =   "frmCA_Documento.frx":B684
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirNorma 
         Height          =   435
         Left            =   90
         TabIndex        =   54
         Top             =   3645
         Width           =   2145
         _Version        =   851970
         _ExtentX        =   3784
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir Norma"
         Appearance      =   5
         Picture         =   "frmCA_Documento.frx":11EE6
      End
      Begin pryCombo.miCombo cmbNormas 
         Height          =   330
         Left            =   60
         TabIndex        =   53
         Top             =   3285
         Width           =   5640
         _ExtentX        =   9948
         _ExtentY        =   582
      End
      Begin MSComctlLib.ListView listaNormas 
         Height          =   2805
         Left            =   45
         TabIndex        =   52
         Top             =   450
         Width           =   5670
         _ExtentX        =   10001
         _ExtentY        =   4948
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
   End
   Begin Geslab.ControlPanelXP cpDocumentos 
      Height          =   4650
      Left            =   9225
      TabIndex        =   57
      Top             =   4905
      Width           =   5790
      _ExtentX        =   10213
      _ExtentY        =   8202
      Caption         =   "Documentos asociados (PNT)"
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   4650
      Begin XtremeSuiteControls.PushButton cmdModificarPNT 
         Height          =   435
         Left            =   1935
         TabIndex        =   63
         Top             =   4095
         Width           =   1830
         _Version        =   851970
         _ExtentX        =   3228
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Modificar PNT"
         Appearance      =   5
         Picture         =   "frmCA_Documento.frx":18748
      End
      Begin VB.CheckBox chkMarcarPNT 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Marcar si el PNT debe estar en vigor para generar edición"
         Height          =   240
         Left            =   90
         TabIndex        =   62
         Top             =   3780
         Width           =   4425
      End
      Begin MSComctlLib.ListView listaDocumentos 
         Height          =   2940
         Left            =   90
         TabIndex        =   61
         Top             =   450
         Width           =   5580
         _ExtentX        =   9843
         _ExtentY        =   5186
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
      Begin pryCombo.miCombo cmbDocumentos 
         Height          =   330
         Left            =   105
         TabIndex        =   60
         Top             =   3420
         Width           =   5550
         _ExtentX        =   9790
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirPNT 
         Height          =   435
         Left            =   90
         TabIndex        =   59
         Top             =   4095
         Width           =   1830
         _Version        =   851970
         _ExtentX        =   3228
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir PNT"
         Appearance      =   5
         Picture         =   "frmCA_Documento.frx":1EFAA
      End
      Begin XtremeSuiteControls.PushButton cmdEliminarPNT 
         Height          =   435
         Left            =   3780
         TabIndex        =   58
         Top             =   4095
         Width           =   1890
         _Version        =   851970
         _ExtentX        =   3334
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar PNT"
         Appearance      =   5
         Picture         =   "frmCA_Documento.frx":2580C
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Generación de nuevo PNT"
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
      Top             =   30
      Width           =   2760
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rellene todos los campos para la creación/modificación  de un nuevo documento de Calidad"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   43
      Top             =   330
      Width           =   6585
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   15030
   End
End
Attribute VB_Name = "frmCA_Documento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long


Private Sub cmbRequisitos_Click()
    If cmbFamilia_Req.BoundText <> "" Then
        If cmbFamilia_Req.BoundText > 0 Then
            frmCA_Req_Detalle.PK = cmbFamilia_Req.BoundText
            frmCA_Req_Detalle.Show 1
        End If
    End If
End Sub

Private Sub cmdAdjuntos_Click()
'M0499-I
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_CA_DOCUMENTO
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
'M0499-F
End Sub


Private Sub cmdAnadirNorma_Click()
    If cmbNormas.getPK_SALIDA = 0 Then
        MsgBox "Debe seleccionar una de entre las existentes", vbOK + vbExclamation, "Añadir Norma"
        Exit Sub
    End If
    Dim i As Integer
    For i = 1 To listaNormas.ListItems.Count
        If CLng(listaNormas.ListItems(i).Text) = CLng(cmbNormas.getPK_SALIDA) Then
            MsgBox "La norma ya se encuentra en el documento.", vbExclamation, App.Title
            Exit Sub
        End If
    Next
    Dim oDC As New clsCa_documentos_normas
    With oDC
        .setDOCUMENTO_ID = PK
        .setNORMA_ID = cmbNormas.getPK_SALIDA
        .setORDEN = listaNormas.ListItems.Count + 1
        .Insertar
    End With
    cmbNormas.limpiar
    Call cargar_normas
End Sub

'BUG-XXXX-I

Private Sub cmdAnadirPNT_Click()

    If cmbDocumentos.getPK_SALIDA = 0 Then
        MsgBox "Debe seleccionar un PNT de entre los existentes", vbOK + vbExclamation, "Añadir PNT"
        Exit Sub
    End If
    
    Dim i As Integer
    
    For i = 1 To listaDocumentos.ListItems.Count
        If CLng(listaDocumentos.ListItems(i).Text) = CLng(cmbDocumentos.getPK_SALIDA) Then
            MsgBox "Este PNT ya se encuentra asociado al documento.", vbExclamation, App.Title
            Exit Sub
        End If
    Next
    
    Dim oDC As New clsCa_documentos_PNT
    With oDC
        .setDOCUMENTO_ID = PK
        .setPNT_ID = cmbDocumentos.getPK_SALIDA
        .setVINCULADO = chkMarcarPNT.Value
        .setORDEN = listaDocumentos.ListItems.Count + 1
        .Insertar
    End With
    cmbDocumentos.limpiar
    chkMarcarPNT.Value = Unchecked
    Call Cargar_Documentos
End Sub

Private Sub cmdEliminarPNT_Click()
    If listaDocumentos.ListItems.Count = 0 Then Exit Sub
    
    If MsgBox("¿Esta seguro de eliminar el PNT?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oDP As New clsCa_documentos_PNT
        oDP.Eliminar PK, listaDocumentos.ListItems(listaDocumentos.selectedItem.Index).Text
        Cargar_Documentos
    End If
End Sub

'BUG-XXXX-F


Private Sub cmdEquipos_Click()
    If PK > 0 Then
        frmDonde.lbltitulo = "Listado de Equipos con el Documento : " & txtDatos(0)
        With frmDonde.lista.ColumnHeaders
            .Add , , "NºEQUIPO", 1000, lvwColumnLeft
            .Add , , "NOMBRE", 4500, lvwColumnLeft
            .Add , , "NºSERIE", 2000, lvwColumnCenter
            .Add , , "MODELO", 2000, lvwColumnCenter
        End With
        Dim rs As ADODB.Recordset
        Dim c As String
        c = "SELECT A.ID_EQUIPO, A.NOMBRE, A.SERIE,A.MODELO " & _
            " FROM EQUIPOS A, EQ_NORMAS_EQUIPOS B " & _
            " WHERE A.ID_EQUIPO = B.EQUIPO_ID " & _
            "   AND B.DOCUMENTO_ID = " & PK & _
            "   AND TIPO = 0 " & _
            " ORDER BY A.NOMBRE "
        Set rs = datos_bd(c)
        frmDonde.lblsubtitulo = "Equipos encontrados : " & rs.RecordCount
        If rs.RecordCount > 0 Then
            Do
                With frmDonde.lista.ListItems.Add(, , Format(rs(0), "0000"))
                 .SubItems(1) = rs(1)
                 .SubItems(2) = rs(2)
                 .SubItems(3) = rs(3)
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = Nothing
        frmDonde.tipo = 0
        frmDonde.Show 1
    End If
End Sub

Private Sub chkVinculo_Click()
    If chkVinculo.Value = Checked Then
        frmVinculo.Enabled = True
    Else
        frmVinculo.Enabled = False
    End If
End Sub

Private Sub cmdAnotaciones_Click()
    If PK > 0 Then
        frmCA_Documento_Anotaciones.PK_ID = PK
        frmCA_Documento_Anotaciones.Show 1
    End If
End Sub

Private Sub cmdModificarPNT_Click()
   On Error GoTo cmdModificarPNT_Click_Error

    If cmbDocumentos.getPK_SALIDA = 0 Then
        MsgBox "Debe seleccionar un PNT de entre los existentes.", vbOK + vbExclamation, "Añadir PNT"
        Exit Sub
    End If
    If listaDocumentos.ListItems.Count = 0 Then
        MsgBox "Debe seleccionar un PNT de la lista.", vbOK + vbExclamation, "Añadir PNT"
        Exit Sub
    End If
    Dim oDC As New clsCa_documentos_PNT
    With oDC
        .setVINCULADO = chkMarcarPNT.Value
        .Modificar PK, cmbDocumentos.getPK_SALIDA
    End With
    cmbDocumentos.limpiar
    chkMarcarPNT.Value = Unchecked
    Call Cargar_Documentos

   On Error GoTo 0
   Exit Sub

cmdModificarPNT_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificarPNT_Click of Formulario frmCA_Documento"

End Sub

Private Sub cmdPNT_Click()
    If chkVinculo.Value = Checked Then
        MsgBox "No se pueden generar ediciones sobre un documento vinculado.", vbExclamation, App.Title
        Exit Sub
    End If
    If Not IsNumeric(txtDatos(1)) Then
        MsgBox "La edición del documento debe ser numérica.", vbExclamation, App.Title
        txtDatos(1).SetFocus
        Exit Sub
    End If
'    frmCA_PNT.PK = gCA_documento
    frmCA_PNT.PK = PK
    
    If cmbestados.BoundText = C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_CREACION Then
        frmCA_PNT.txtDatos(0) = "1"
    Else
        If txtDatos(1) = "" Then
            frmCA_PNT.txtDatos(0) = "1"
        Else
            If IsNumeric(txtDatos(1)) Then
                frmCA_PNT.txtDatos(0) = CInt(txtDatos(1)) + 1 ' Edición
            Else
                frmCA_PNT.txtDatos(0) = "1"
            End If
        End If
    End If
    frmCA_PNT.fecha(4) = Date
    frmCA_PNT.txtDatos(1) = txtDatos(3) ' Código
    frmCA_PNT.txtDatos(2) = txtDatos(0) ' Descripción
    frmCA_PNT.Show 1
    cargar_documento
End Sub

Private Sub chkcopia_Click()
    If chkcopia.Value = Checked Then
        txtDatos(2).Enabled = True
        txtDatos(4).Enabled = True
    Else
        txtDatos(2).Enabled = False
        txtDatos(4).Enabled = False
    End If
End Sub

Private Sub cmdMostrar_Click(Index As Integer)
    Dim oCA_Documento As New clsCa_documentos
    oCA_Documento.mostrar PK, True
    Set oCA_Documento = Nothing
End Sub

Private Sub cmdAdjuntar_Click(Index As Integer)
    
    On Error GoTo fallo
    Dim oCA_Documento As New clsCa_documentos
    Dim oDoc As New clsDocumentacion
    Select Case Index
    Case 0
        On Error GoTo fallo
        On Error Resume Next
        cd.DialogTitle = "Abrir fichero"
        cd.InitDir = "c:\"
        cd.ShowOpen
        If cd.FileName <> "" Then
            datos(0).Text = cd.FileName  ' cd.FileTitle
            datos(1).Text = cd.FileTitle
        Else
            Exit Sub
        End If
        On Error GoTo fallo
        Me.MousePointer = 11
        ' Validar ruta seleccionada
        If validar_ruta = False Then
            Me.MousePointer = 0
            Exit Sub
        End If
        Dim MOTIVO As String
        MOTIVO = ""
        oDoc.SubirDocumento TOBJETO.TOBJETO_CA_DOCUMENTO, PK, 0, datos(0), datos(1), MOTIVO, 1, 0
        ' Copiar documento a la nueva ruta
        ' Informamos las rutas del documento
'        Dim RUTA As String
'        Dim oDeco As New clsDecodificadora
'        RUTA = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Documentos\PDF\"
        ' Cargamos la descripción de la familia
'        oDeco.Carga_valor DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS, CLng(cmbfamilias.BoundText)
        ' Creamos la carpeta de la familia por si no existe
'        On Error Resume Next
'        MkDir RUTA_TRABAJO
'        RUTA = RUTA & oDeco.getDESCRIPCION & "\" & datos(1)
'        On Error GoTo fallo
'        FileCopy datos(0), RUTA
        ' Informar la nueva ruta
'        oCA_Documento.Informar_ruta PK, Replace(RUTA, "\", "/")
        Me.MousePointer = 0
        MsgBox "Se ha adjuntado el vínculo correctamente.", vbInformation, App.Title
    Case 1
        If MsgBox("¿Desea realmente eliminar el vínculo?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            oDoc.PasoHistoricoDocumento PK
            oCA_Documento.Informar_ruta PK, ""
            datos(0) = ""
            Me.MousePointer = 0
            MsgBox "Se ha eliminado el vínculo correctamente.", vbInformation, App.Title
        End If
    End Select
    Exit Sub
fallo:
    Me.MousePointer = 0
    error_grave "Error al adjuntar el archivo. " & Err.Description
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
    cd.InitDir = "c:\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(0).Text = cd.FileName  ' cd.FileTitle
        datos(1).Text = cd.FileTitle
    End If
End Sub
Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
      If listaNormas.ListItems.Count = 0 Then
        MsgBox "El documento no tiene normas vínculadas. Se va a guardar el documento, pero compruebe que esto es correcto.", vbInformation, App.Title
      End If
      Dim copiar_documento As Boolean
      Dim documento As Long
      Dim oCA_Documento As New clsCa_documentos
      Dim origen_documento_trabajo As String
      Dim origen_documento_pdf As String
      ' Si el documento cambia de Familia/Código, copiar
      ' el documento de trabajo, el pdf y la ruta
      copiar_documento = False
      If PK <> 0 Then
        oCA_Documento.Carga PK
        If oCA_Documento.getCODIGO <> txtDatos(3) Or _
           CInt(oCA_Documento.getFAMILIA_ID) <> CInt(cmbfamilias.BoundText) Then
            MsgBox "El documento ha cambiado de código/Familia. Se procederá a realizar una copia.", vbInformation, App.Title
            copiar_documento = True
            ' origen documento de trabajo
            origen_documento_trabajo = calidad_ruta_documento_trabajo(PK)
            ' origen documento pdf
            origen_documento_pdf = oCA_Documento.getRUTA
        End If
      End If
      With oCA_Documento
            .setNOMBRE = txtDatos(0)
            .setFAMILIA_ID = cmbfamilias.BoundText
            .setSUBFAMILIA_ID = cmbSubfamilia.BoundText
            .setRESPONSABLE_ID = cmbresponsables.BoundText
            .setCODIGO = txtDatos(3)
            .setEDICION = txtDatos(1)
            .setENAC = chkENAC.Value
            .setNADCAP = chkNADCAP.Value
            .setMTL = chkMTL.Value
            .setEQA = chkEQA.Value
            .setFECHA = txtDatos(5)
            .setOBSERVACIONES = txtDatos(10)
            .setESTADO_ID = cmbestados.BoundText
            .setUSO = chkuso.Value
            .setCOPIA_CONTROLADA = chkcopia.Value
            .setCOPIA_LABORATORIO = txtDatos(2)
            .setCOPIA_NUMERO = txtDatos(4)
            .setPLANTILLA_ID = 0
            If cmbPlantilla.Text <> "" Then
                .setPLANTILLA_ID = cmbPlantilla.BoundText
            End If
            .setNO_EDICION = chkNoEdicion.Value
            .setDOCUMENTO_VINCULADO = chkVinculo.Value
            .setFORMACION = chkFormacion.Value
            If chkFormacion.Value = Checked Then
                If cmbPlanes.BoundText = "" Then
                    .setPLAN_ID = 0
                Else
                    .setPLAN_ID = cmbPlanes.BoundText
                    'M1106-I
                    Dim oPlanDocs As New clsFormacion_pf_docs
                    oPlanDocs.setDOCUMENTO_ID = PK
                    oPlanDocs.setPLAN_FORMACION_ID = cmbPlanes.BoundText
                    oPlanDocs.Insertar_Ignorar
                    Set oPlanDocs = Nothing
                    'M1106-F
                End If
            Else
                .setPLAN_ID = 0
            End If
            If cmbFamilia_Req.BoundText = "" Then
                .setFAMILIA_REQ_ID = 0
            Else
                .setFAMILIA_REQ_ID = cmbFamilia_Req.BoundText
            End If
      End With
'      If gCA_documento = 0 Then
      If PK = 0 Then
        If MsgBox("Va a introducir un nuevo documento. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            documento = oCA_Documento.Insertar
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el documento. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            If oCA_Documento.Modificar(PK) Then
                If copiar_documento Then
                    ' Copiar el documento origen al destino
                    FileCopy origen_documento_trabajo, calidad_ruta_documento_trabajo(PK)
                    ' Copiar el pdf origen al destino
                    If origen_documento_pdf <> "" Then
                        If Dir(origen_documento_pdf) <> "" Then
                            FileCopy origen_documento_pdf, calidad_ruta_pdf(PK) & "\" & calidad_nombre_documento_pdf(PK)
                        End If
                    End If
                    ' actualizar la ruta del pdf
                    oCA_Documento.Informar_ruta PK, calidad_ruta_pdf(PK) & "\" & calidad_nombre_documento_pdf(PK)
                End If
                documento = PK
            End If
        Else
            Exit Sub
        End If
      End If
'      If gCA_documento = 0 Then
      If PK = 0 Then
          MsgBox "El documento se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
'          gCA_documento = DOCUMENTO
          PK = documento
          Form_Load
      Else
          MsgBox "El documento se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
          Unload Me
      End If
      
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmCA_Documento"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdValoracion_Click()
    If PK > 0 Then
        frmCA_Valoracion_Listado.PK = PK
        frmCA_Valoracion_Listado.Show 1
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    fecha = Date
    Call cargar_combos
    If PK <> 0 Then
        lbltitulo = "Modificación de Documento de calidad"
        cargar_documento
        
    Else
        lbltitulo = "Alta de Documento de calidad"
        txtDatos(1).Locked = True
        txtDatos(5).Locked = True
        cmbestados.Locked = True
        
        txtDatos(1) = "1"
        txtDatos(5) = Format(Date, "DD-MM-YYYY")
        datos(0).Locked = True
        cmdMostrar(0).Enabled = False
        cmdPNT.Enabled = False
        cmdAdjuntos.Enabled = False
        chkVinculo.Enabled = False
        frmVinculo.Enabled = False
        cmbestados.BoundText = C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_CREACION
        chkuso.Value = Checked
        
        cpNormas.PanelOpen = False
        cpNormas.CanExpand = False
    End If
    permisos
End Sub



Private Sub listaDocumentos_Click()
   On Error GoTo listaDocumentos_Click_Error

    If listaDocumentos.ListItems.Count > 0 Then
        cmbDocumentos.MostrarElemento listaDocumentos.ListItems(listaDocumentos.selectedItem.Index).Text
        If listaDocumentos.ListItems(listaDocumentos.selectedItem.Index).SubItems(2) = "X" Then
            chkMarcarPNT.Value = Checked
        Else
            chkMarcarPNT.Value = Unchecked
        End If
    End If

   On Error GoTo 0
   Exit Sub

listaDocumentos_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure listaDocumentos_Click of Formulario frmCA_Documento"
End Sub

Private Sub listaNormas_DblClick()
'    Dim strDest As String
'    Dim r As Double

    If listaNormas.ListItems.Count > 0 Then
        Dim oNorma As New clsCa_normas
        oNorma.mostrar listaNormas.ListItems(listaNormas.selectedItem.Index).Text, True
        Set oNorma = Nothing
'        strDest = Replace(listaNormas.ListItems(listaNormas.selectedItem.Index).SubItems(2), "/", "\")
'        If Dir(strDest, vbArchive) <> "" Then
'            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & strDest, vbNormalFocus)
'        End If
    End If
End Sub
Private Sub cmdEliminarNorma_Click()
    'BUG-XXXX-I
    ' If listaNormas.ListItems.Count = 0 Then Exit Sub
      If listaNormas.ListItems.Count = 0 Then
         cpNormas.PanelOpen = False
         Exit Sub
      End If
    'BUG-XXXX-F
    If MsgBox("¿Esta seguro de eliminar la norma?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oDN As New clsCa_documentos_normas
        oDN.Eliminar PK, listaNormas.ListItems(listaNormas.selectedItem.Index).Text
        cargar_normas
    End If
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 10 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
    If Index = 3 Then ' Código
        If txtDatos(Index) <> "" Then
            Dim oCA As New clsCa_documentos
'            If oCA.Verificar_Codigo(txtDatos(Index), gCA_documento) Then
            If oCA.Verificar_Codigo(txtDatos(Index), PK) Then
                MsgBox "Ya existe este código de documento.", vbExclamation, App.Title
                txtDatos(Index) = ""
                txtDatos(Index).SetFocus
            End If
        End If
    End If
End Sub
Private Sub cargar_documento()
    Dim oCA_Documento As New clsCa_documentos
   On Error GoTo cargar_documento_Error

    With oCA_Documento
'        If .Carga(gCA_documento) = True Then
        If .Carga(PK) = True Then
            txtDatos(0) = .getNOMBRE
            cmbfamilias.BoundText = .getFAMILIA_ID
            cmbSubfamilia.BoundText = .getSUBFAMILIA_ID
            cmbresponsables.BoundText = .getRESPONSABLE_ID
            txtDatos(1) = .getEDICION
            txtDatos(3) = .getCODIGO
            chkENAC.Value = .getENAC
            chkNADCAP.Value = .getNADCAP
            chkMTL.Value = .getMTL
            chkEQA.Value = .getEQA

'            If IsDate(.getFECHA) Then
'                fecha = .getFECHA
'            Else
'                fecha = Date
'            End If
            If IsDate(.getFECHA) Then
                txtDatos(5) = Format(.getFECHA, "dd-mm-yyyy")
            Else
                txtDatos(5) = .getFECHA
            End If
            txtDatos(10) = .getOBSERVACIONES
            cmbestados.BoundText = .getESTADO_ID
            cmbPlantilla.BoundText = .getPLANTILLA_ID
            If .getUSO = 1 Then
                chkuso.Value = Checked
                chkuso.Caption = "Documento EN USO"
                chkuso.BackColor = vbGreen
            Else
                chkuso.Value = Unchecked
                chkuso.BackColor = vbRed
                chkuso.Caption = "Documento NO SE USA"
            End If
            If .getCOPIA_CONTROLADA = 1 Then
                chkcopia.Value = Checked
                txtDatos(2).Enabled = True
                txtDatos(4).Enabled = True
            Else
                chkcopia.Value = Unchecked
                txtDatos(2).Enabled = False
                txtDatos(4).Enabled = False
            End If
            txtDatos(2) = .getCOPIA_LABORATORIO
            txtDatos(4) = .getCOPIA_NUMERO
            datos(0) = Replace(.getRUTA, "/", "\")
            If .getNO_EDICION = 1 Then
                chkNoEdicion.Value = Checked
            Else
                chkNoEdicion.Value = Unchecked
            End If
            chkVinculo.Value = .getDOCUMENTO_VINCULADO
            ' Formacion
            chkFormacion.Value = .getFORMACION
            cmbPlanes.BoundText = .getPLAN_ID
            cmbFamilia_Req.BoundText = .getFAMILIA_REQ_ID
        End If
    End With
    
    cargar_normas
    If listaNormas.ListItems.Count > 0 Then
        cpNormas.PanelOpen = True
    End If
'BUG-XXXX-I
    Cargar_Documentos
'BUG-XXXX-F
'1018 (Numero de anotaciones)
    Dim ocada As New clsCa_documentos_anotaciones
    Dim cont As Integer
    cont = ocada.anotaciones(PK)
    If cont <> 0 Then
        cmdAnotaciones.Caption = "Anotaciones (" & cont & ")"
        cmdAnotaciones.Font.bold = True
        
    End If
   On Error GoTo 0
   Exit Sub

cargar_documento_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_documento of Formulario frmCA_Documento"
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle un nombre al documento.", vbExclamation, App.Title
        validar = False
        txtDatos(0).SetFocus
        Exit Function
    End If
    If Trim(txtDatos(1)) = "" Then
        MsgBox "Debe indicar la edición del documento.", vbExclamation, App.Title
        txtDatos(1).SetFocus
        validar = False
        Exit Function
'    Else
'        If Not IsNumeric(txtDatos(1)) Then
'            MsgBox "La edición del documento debe ser numérica.", vbExclamation, App.Title
'            txtDatos(1).SetFocus
'            validar = False
'            Exit Function
'        End If
    End If
    If Trim(txtDatos(3)) = "" Then
        MsgBox "Debe darle un código al documento.", vbExclamation, App.Title
        validar = False
        txtDatos(3).SetFocus
        Exit Function
    End If
    If cmbfamilias.BoundText = "" Then
        MsgBox "Debe asignar una familia al documento.", vbExclamation, App.Title
        validar = False
        cmbfamilias.SetFocus
        Exit Function
    End If
    If cmbSubfamilia.BoundText = "" Then
        MsgBox "Debe asignar una SubFamilia al documento.", vbExclamation, App.Title
        validar = False
        cmbSubfamilia.SetFocus
        Exit Function
    End If
    If cmbresponsables.BoundText = "" Then
        MsgBox "Debe asignar un Responsable al documento.", vbExclamation, App.Title
        validar = False
        cmbresponsables.SetFocus
        Exit Function
    End If
    If cmbestados.BoundText = "" Then
        MsgBox "Debe asignar un estado al documento.", vbExclamation, App.Title
        validar = False
        cmbestados.SetFocus
        Exit Function
    End If
End Function
Private Sub cabecera()
    With listaNormas.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Nombre", listaNormas.Width * 0.99, lvwColumnLeft
'        .Add , , "Ruta", 0, lvwColumnLeft
    End With
'BUG-XXXX-I
    With listaDocumentos.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Nombre", 4500, lvwColumnLeft
        .Add , , "Relación", 1000, lvwColumnCenter
'        .Add , , "Ruta", 0, lvwColumnLeft
    End With
'BUG-XXXX-F
    
End Sub
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbfamilias, DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS
    oDeco.cargar_combo cmbSubfamilia, DECODIFICADORA.CA_DOCUMENTOS_SUBFAMILIAS
    oDeco.cargar_combo cmbestados, DECODIFICADORA.CA_DOCUMENTOS_ESTADOS
    oDeco.cargar_combo cmbresponsables, DECODIFICADORA.CA_DOCUMENTOS_RESPONSABLES
    oDeco.cargar_combo cmbPlantilla, DECODIFICADORA.CALIDAD_PLANTILLAS_DOCUMENTOS
'M1110-I
'    cargar_combo cmbPlanes, New clsEmpleados_plan_formacion
    cargar_combo cmbPlanes, New clsFormacion_pf
'M1110-F
    llenar_combo cmbNormas, New clsCa_normas, 0, frmCA_Normas, ""
'BUG-XXXX-I
    llenar_combo cmbDocumentos, New clsCa_documentos, 0, Me, ""
'BUG-XXXX-F
    cargar_combo cmbFamilia_Req, New clsCa_req_familias
End Sub
Private Function validar_ruta() As Boolean
    validar_ruta = False
    If datos(0) = "" Then
        MsgBox "Escriba una ruta.", vbExclamation, App.Title
        Exit Function
    End If
    If Dir(datos(0)) = "" Then
        MsgBox "La ruta introducida no existe.", vbExclamation, App.Title
        Exit Function
    End If
    If cmbfamilias.Text = "" Then
        MsgBox "El documento debe pertenecer a una familia.", vbExclamation, App.Title
        Exit Function
    End If
    validar_ruta = True
End Function

Private Sub permisos()
    ' Permiso gestión documentación
    If Not USUARIO.getPER_DOCUMENTACION_CALIDAD Then
        cmdok.Enabled = False
        cmdPNT.Enabled = False
        cmdAdjuntos.Enabled = False
        chkVinculo.Enabled = False
        frmVinculo.Enabled = False
    Else
        ' Permiso creación ediciones
        If Not USUARIO.getPER_PNT Then
            cmdPNT.Enabled = False
        End If
    End If
    ' Superusuario pnt, puede tocar edicion, fecha y estado
    If USUARIO.getPER_ADMIN_PNT Then
        txtDatos(1).Locked = False
        txtDatos(5).Locked = False
        cmbestados.Locked = False
        
        chkVinculo.Enabled = True
        frmVinculo.Enabled = True
    End If
End Sub
Private Sub cargar_normas()
    Dim rs As ADODB.Recordset
    Dim oCN As New clsCa_documentos_normas
    Set rs = oCN.Listado(PK)

    listaNormas.ListItems.Clear
    If rs.RecordCount > 0 Then
'BUG-XXXX-I
        cpNormas.PanelOpen = True
'BUG-XXXX-F
        Do
            With listaNormas.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
'                .SubItems(2) = rs(2)
            End With
            rs.MoveNext
        Loop Until rs.EOF
'BUG-XXXX-I
    Else
        cpNormas.PanelOpen = False
'BUG-XXXX-F
    End If
    Set rs = Nothing
End Sub

'BUG-XXXX-I
Private Sub Cargar_Documentos()
    Dim rs As ADODB.Recordset
    Dim oCP As New clsCa_documentos_PNT
    Dim oDoc As New clsCa_documentos
    
    Set rs = oCP.Listado(PK)
    listaDocumentos.ListItems.Clear
    If rs.RecordCount > 0 Then
        cpDocumentos.PanelOpen = True
        
        If rs.RecordCount = 1 Then
           oDoc.VINCULAR (PK)
        End If
        
        Do
            With listaDocumentos.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = IIf(rs(3) = 1, "X", "")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    Else
        cpDocumentos.PanelOpen = False
        oDoc.Desvincular (PK)
    End If
    
    
    Set rs = Nothing
End Sub
'BUG-XXXX-F
