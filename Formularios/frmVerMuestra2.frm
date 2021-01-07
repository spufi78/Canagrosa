VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmVerMuestra 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Otros"
   ClientHeight    =   11175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   15780
   Icon            =   "frmVerMuestra2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   15780
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDatosEspeciales 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Especiales"
      ClipControls    =   0   'False
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
      Height          =   4620
      Left            =   4725
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   6570
      Begin VB.CommandButton cmdwww 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quitar Revisado WEB Procesos"
         Height          =   870
         Left            =   90
         Picture         =   "frmVerMuestra2.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   208
         Top             =   3645
         Width           =   2310
      End
      Begin VB.CheckBox chkInformeManual 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Informe Manual"
         Enabled         =   0   'False
         Height          =   240
         Left            =   135
         TabIndex        =   207
         Top             =   3240
         Width           =   2175
      End
      Begin VB.CheckBox chkAnulada 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Muestra ANULADA"
         Enabled         =   0   'False
         Height          =   240
         Left            =   135
         TabIndex        =   201
         Top             =   2925
         Width           =   2175
      End
      Begin VB.CheckBox chkFechaEnvio 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   195
         Left            =   135
         TabIndex        =   45
         Top             =   2160
         Width           =   240
      End
      Begin VB.CheckBox chkFechaCierre 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   195
         Left            =   135
         TabIndex        =   44
         Top             =   1800
         Width           =   240
      End
      Begin VB.CommandButton cmdRegenerarInforme 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Regenerar informe sin tocar edición"
         Height          =   1095
         Left            =   4185
         Picture         =   "frmVerMuestra2.frx":711C
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   540
         Width           =   2310
      End
      Begin VB.CommandButton cmdModificarDatosEspeciales 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   870
         Left            =   4095
         Picture         =   "frmVerMuestra2.frx":79E6
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   3645
         Width           =   2310
      End
      Begin VB.TextBox txtedicion 
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
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   1530
         TabIndex        =   24
         Top             =   675
         Width           =   1455
      End
      Begin MSComCtl2.DTPicker fcomienzo 
         Height          =   330
         Left            =   1530
         TabIndex        =   25
         Top             =   1035
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   60358657
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker ffin 
         Height          =   330
         Left            =   1530
         TabIndex        =   28
         Top             =   1755
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   60358657
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbUsuarioCierre 
         Height          =   330
         Left            =   1530
         TabIndex        =   30
         Top             =   2520
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fFinalizacion 
         Height          =   330
         Left            =   1530
         TabIndex        =   34
         Top             =   1395
         Width           =   1470
         _ExtentX        =   2593
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
         Format          =   60358657
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker hfin 
         Height          =   330
         Left            =   3060
         TabIndex        =   40
         Top             =   1755
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   60358658
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fEnvio 
         Height          =   330
         Left            =   1530
         TabIndex        =   41
         Top             =   2115
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   60358657
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker hEnvio 
         Height          =   330
         Left            =   3060
         TabIndex        =   43
         Top             =   2115
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   60358658
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00BD814F&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "DATOS ESPECIALES"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   12
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   45
         TabIndex        =   199
         Top             =   0
         Width           =   6435
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Envío"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   36
         Left            =   405
         TabIndex        =   42
         Top             =   2160
         Width           =   1230
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Finalización"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   30
         Left            =   135
         TabIndex        =   35
         Top             =   1440
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cerrada por"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   24
         Left            =   135
         TabIndex        =   31
         Top             =   2565
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Cierre"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   25
         Left            =   405
         TabIndex        =   29
         Top             =   1800
         Width           =   1230
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Comienzo"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   28
         Left            =   135
         TabIndex        =   27
         Top             =   1080
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edición"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   26
         Left            =   135
         TabIndex        =   26
         Top             =   720
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdContra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Contradictorio"
      Height          =   915
      Left            =   13275
      Picture         =   "frmVerMuestra2.frx":82B0
      Style           =   1  'Graphical
      TabIndex        =   48
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   10170
      Visible         =   0   'False
      Width           =   1050
   End
   Begin XtremeSuiteControls.PushButton cmdCambiar 
      Height          =   300
      Index           =   1
      Left            =   1890
      TabIndex        =   38
      Top             =   765
      Width           =   1905
      _Version        =   851970
      _ExtentX        =   3360
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Siguiente"
      Appearance      =   5
      Picture         =   "frmVerMuestra2.frx":8B7A
      TextImageRelation=   4
   End
   Begin XtremeSuiteControls.PushButton cmdCambiar 
      Height          =   300
      Index           =   0
      Left            =   45
      TabIndex        =   39
      Top             =   765
      Width           =   1860
      _Version        =   851970
      _ExtentX        =   3281
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Anterior"
      Appearance      =   5
      Picture         =   "frmVerMuestra2.frx":F3DC
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   45
      TabIndex        =   1
      Top             =   9990
      Width           =   13125
      Begin VB.CommandButton cmdEtiquetaSoluciones 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Soluciones"
         Height          =   915
         Left            =   6030
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Muestra el informe de registrro de la muestra"
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton cmdEdiciones 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ediciones"
         Height          =   915
         Left            =   10980
         Picture         =   "frmVerMuestra2.frx":15C3E
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton cmdFirmaCliente 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Firma Cliente"
         Height          =   915
         Left            =   7020
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton cmdFluido 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fluido"
         Height          =   915
         Left            =   11970
         Picture         =   "frmVerMuestra2.frx":1C490
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "Consultar Detalle del Fluido"
         Top             =   180
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdRecarga 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recarga"
         Height          =   915
         Left            =   9990
         Style           =   1  'Graphical
         TabIndex        =   21
         ToolTipText     =   "Enviar informe por E-mail"
         Top             =   180
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.CommandButton cmdAdjuntos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntos"
         Height          =   915
         Left            =   8010
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Enviar informe por E-mail"
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton cmdVida 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vida "
         Height          =   915
         Left            =   9000
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Enviar informe por E-mail"
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton cmdEtiqueta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Etiqueta"
         Height          =   915
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   18
         ToolTipText     =   "Muestra el informe de registrro de la muestra"
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton cmdInfRegistro 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Doc.Registro"
         Height          =   915
         Left            =   4050
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Muestra el informe de registrro de la muestra"
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton cmdInforme 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informe"
         Height          =   915
         Left            =   3060
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Previsualizar informe de ensayo"
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton cmdDeter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Registro"
         Height          =   915
         Left            =   2070
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton cmdAnular 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Anular"
         Height          =   915
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   180
         Width           =   990
      End
      Begin VB.CommandButton cmdModificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   915
         Left            =   75
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Códigos de Registro"
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
      Height          =   645
      Left            =   3810
      TabIndex        =   9
      Top             =   360
      Width           =   8760
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   345
         Index           =   5
         Left            =   720
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   225
         Width           =   1230
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   345
         Index           =   0
         Left            =   4410
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   1230
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   345
         Index           =   1
         Left            =   2610
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   555
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         Height          =   345
         Index           =   3
         Left            =   3150
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   225
         Width           =   975
      End
      Begin MSDataListLib.DataCombo cmbCentroMuestra 
         Bindings        =   "frmVerMuestra2.frx":1C79A
         Height          =   315
         Left            =   6435
         TabIndex        =   216
         Top             =   225
         Width           =   2265
         _ExtentX        =   3995
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
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   38
         Left            =   5850
         TabIndex        =   215
         Top             =   270
         Width           =   600
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ID"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   22
         Left            =   4185
         TabIndex        =   15
         Top             =   270
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   14
         Top             =   270
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   2025
         TabIndex        =   13
         Top             =   270
         Width           =   600
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   960
      Left            =   14625
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   10170
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   960
      Left            =   13545
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   10170
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdAbrirMuestra 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Abrir Muestra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   12660
      Style           =   1  'Graphical
      TabIndex        =   22
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   780
      Visible         =   0   'False
      Width           =   3015
   End
   Begin TabDlg.SSTab tabPrincipal 
      Height          =   8835
      Left            =   45
      TabIndex        =   51
      Top             =   1125
      Width           =   15675
      _ExtentX        =   27649
      _ExtentY        =   15584
      _Version        =   393216
      Style           =   1
      TabHeight       =   670
      BackColor       =   8421504
      MouseIcon       =   "frmVerMuestra2.frx":1C7E0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " Datos Generales"
      TabPicture(0)   =   "frmVerMuestra2.frx":1C7FC
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "frmPlasma"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "frameDeterminaciones"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Frame5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Frame4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "frmFacturacion"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Frame9"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).ControlCount=   8
      TabCaption(1)   =   "Otros Datos"
      TabPicture(1)   =   "frmVerMuestra2.frx":2305E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame10"
      Tab(1).Control(1)=   "Frame8"
      Tab(1).Control(2)=   "frmADS"
      Tab(1).Control(3)=   "frmIndicadores"
      Tab(1).Control(4)=   "frmENAC"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "Fluidos"
      TabPicture(2)   =   "frmVerMuestra2.frx":298C0
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "frmFluidos"
      Tab(2).Control(1)=   "frmAIM"
      Tab(2).ControlCount=   2
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Muestra en Consulta"
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
         Height          =   1965
         Left            =   -74820
         TabIndex        =   210
         Top             =   6435
         Width           =   7875
         Begin VB.TextBox txtCONSULTA_OBSERVACIONES 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   960
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   212
            Top             =   900
            Width           =   7710
         End
         Begin VB.CheckBox chkConsulta 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   90
            TabIndex        =   211
            Top             =   315
            Width           =   240
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Observación"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   50
            Left            =   90
            TabIndex        =   214
            Top             =   630
            Width           =   1590
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Muestra en Consulta o Falta documentación"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   46
            Left            =   405
            TabIndex        =   213
            Top             =   315
            Width           =   4695
         End
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Normas En Estudio"
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
         Height          =   3390
         Left            =   -66900
         TabIndex        =   202
         Top             =   1845
         Width           =   7260
         Begin XtremeSuiteControls.PushButton cmdEliminarNorma 
            Height          =   435
            Left            =   5040
            TabIndex        =   203
            Top             =   2835
            Width           =   2115
            _Version        =   851970
            _ExtentX        =   3731
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Eliminar Norma"
            Appearance      =   5
            Picture         =   "frmVerMuestra2.frx":30122
         End
         Begin XtremeSuiteControls.PushButton cmdAnadirNorma 
            Height          =   435
            Left            =   90
            TabIndex        =   204
            Top             =   2880
            Width           =   2145
            _Version        =   851970
            _ExtentX        =   3784
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Añadir Norma"
            Appearance      =   5
            Picture         =   "frmVerMuestra2.frx":36984
         End
         Begin pryCombo.miCombo cmbNormas 
            Height          =   330
            Left            =   90
            TabIndex        =   205
            Top             =   2430
            Width           =   7080
            _ExtentX        =   12488
            _ExtentY        =   582
         End
         Begin MSComctlLib.ListView listaNormas 
            Height          =   2130
            Left            =   90
            TabIndex        =   206
            Top             =   270
            Width           =   7110
            _ExtentX        =   12541
            _ExtentY        =   3757
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
      Begin VB.Frame frmADS 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos ADS"
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
         Height          =   2550
         Left            =   -74820
         TabIndex        =   188
         Top             =   3825
         Width           =   7875
         Begin pryCombo.miCombo cmbProgramaADS 
            Height          =   330
            Left            =   1080
            TabIndex        =   189
            Top             =   765
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbEnsayo 
            Height          =   330
            Left            =   1080
            TabIndex        =   190
            Top             =   360
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbSection 
            Height          =   330
            Left            =   1080
            TabIndex        =   191
            Top             =   1170
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbFluid 
            Height          =   330
            Left            =   1080
            TabIndex        =   192
            Top             =   1575
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbFacility 
            Height          =   330
            Left            =   1080
            TabIndex        =   193
            Top             =   1980
            Width           =   6630
            _ExtentX        =   11695
            _ExtentY        =   582
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Programa"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   45
            Left            =   90
            TabIndex        =   198
            Top             =   810
            Width           =   870
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ensayo"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   44
            Left            =   90
            TabIndex        =   197
            Top             =   405
            Width           =   690
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Section"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   43
            Left            =   90
            TabIndex        =   196
            Top             =   1215
            Width           =   870
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fluid"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   42
            Left            =   90
            TabIndex        =   195
            Top             =   1620
            Width           =   870
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Facility"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   41
            Left            =   90
            TabIndex        =   194
            Top             =   2025
            Width           =   870
         End
      End
      Begin VB.Frame frmFluidos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "FLUIDO"
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
         Left            =   -74865
         TabIndex        =   185
         Top             =   3015
         Width           =   7650
         Begin VB.TextBox txtFluidoNormativa 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Left            =   1260
            Locked          =   -1  'True
            TabIndex        =   186
            Top             =   270
            Width           =   6180
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Norm. Aplicable"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   39
            Left            =   90
            TabIndex        =   187
            Top             =   315
            Width           =   1140
         End
      End
      Begin VB.Frame frmAIM 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clasificación Fluido AIM (Aplicación control de procesos)"
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
         Height          =   2280
         Left            =   -74865
         TabIndex        =   174
         Top             =   630
         Width           =   7665
         Begin pryCombo.miCombo cmbCentro 
            Height          =   375
            Left            =   1305
            TabIndex        =   175
            Top             =   630
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   661
         End
         Begin pryCombo.miCombo cmbTipoEnsayo 
            Height          =   375
            Left            =   1305
            TabIndex        =   176
            Top             =   1035
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   661
         End
         Begin pryCombo.miCombo cmbSeccion 
            Height          =   375
            Left            =   1305
            TabIndex        =   177
            Top             =   1440
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   661
         End
         Begin pryCombo.miCombo cmbEstacion 
            Height          =   375
            Left            =   1305
            TabIndex        =   178
            Top             =   1845
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   661
         End
         Begin pryCombo.miCombo cmbPrograma 
            Height          =   375
            Left            =   1305
            TabIndex        =   179
            Top             =   270
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   661
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Estación"
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   184
            Top             =   1890
            Width           =   615
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sección"
            Height          =   195
            Index           =   7
            Left            =   90
            TabIndex        =   183
            Top             =   1485
            Width           =   585
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Programa"
            Height          =   195
            Index           =   20
            Left            =   90
            TabIndex        =   182
            Top             =   270
            Width           =   675
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Centro"
            Height          =   195
            Index           =   18
            Left            =   90
            TabIndex        =   181
            Top             =   675
            Width           =   465
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo de Ensayo"
            Height          =   195
            Index           =   19
            Left            =   90
            TabIndex        =   180
            Top             =   1080
            Width           =   1110
         End
      End
      Begin VB.Frame frmIndicadores 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Indicadores"
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
         Height          =   3180
         Left            =   -74820
         TabIndex        =   157
         Top             =   630
         Width           =   7875
         Begin VB.CheckBox chkFechaSolicitud 
            Caption         =   "Check1"
            Enabled         =   0   'False
            Height          =   195
            Left            =   90
            TabIndex        =   161
            Top             =   315
            Width           =   240
         End
         Begin VB.CheckBox chkIPA 
            BackColor       =   &H00C0C0C0&
            Caption         =   "I.P.A."
            Enabled         =   0   'False
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   90
            TabIndex        =   160
            Top             =   1440
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   960
            Index           =   10
            Left            =   1575
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   159
            Top             =   2115
            Width           =   4785
         End
         Begin VB.CheckBox chkFechaSolicitudNA 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Aplica"
            Enabled         =   0   'False
            Height          =   195
            Left            =   4185
            TabIndex        =   158
            Top             =   315
            Width           =   1950
         End
         Begin MSComCtl2.DTPicker fechaSolicitud 
            Height          =   330
            Left            =   1575
            TabIndex        =   162
            Top             =   270
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   60358657
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker horaSolicitud 
            Height          =   330
            Left            =   3060
            TabIndex        =   163
            Top             =   270
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   60358658
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker fechaCierre 
            Height          =   330
            Left            =   1575
            TabIndex        =   164
            Top             =   630
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   60358657
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker horaCierre 
            Height          =   330
            Left            =   3060
            TabIndex        =   165
            Top             =   630
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   60358658
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker fechaEnvio 
            Height          =   330
            Left            =   1575
            TabIndex        =   166
            Top             =   990
            Width           =   1470
            _ExtentX        =   2593
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   60358657
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSComCtl2.DTPicker horaEnvio 
            Height          =   330
            Left            =   3060
            TabIndex        =   167
            Top             =   990
            Width           =   1065
            _ExtentX        =   1879
            _ExtentY        =   582
            _Version        =   393216
            Enabled         =   0   'False
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
            Format          =   60358658
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin pryCombo.miCombo cmbMotivoRetraso 
            Height          =   330
            Left            =   1575
            TabIndex        =   168
            Top             =   1710
            Width           =   5820
            _ExtentX        =   10266
            _ExtentY        =   582
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha solicitud"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   32
            Left            =   405
            TabIndex        =   173
            Top             =   315
            Width           =   1140
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Motivo Restraso"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   33
            Left            =   90
            TabIndex        =   172
            Top             =   1755
            Width           =   1590
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha de cierre"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   34
            Left            =   90
            TabIndex        =   171
            Top             =   675
            Width           =   1140
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha de envío"
            ForeColor       =   &H00000000&
            Height          =   240
            Index           =   35
            Left            =   90
            TabIndex        =   170
            Top             =   1035
            Width           =   1140
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Observación"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   37
            Left            =   90
            TabIndex        =   169
            Top             =   2430
            Width           =   1590
         End
      End
      Begin VB.Frame frmENAC 
         Appearance      =   0  'Flat
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
         ForeColor       =   &H80000008&
         Height          =   1185
         Left            =   -66900
         TabIndex        =   152
         Top             =   630
         Width           =   7260
         Begin VB.OptionButton opENAC 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ENAC PARCIAL (Algún Ensayo no esta certificado por ENAC)"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   156
            Top             =   765
            Width           =   4965
         End
         Begin VB.OptionButton opENAC 
            BackColor       =   &H00C0C0C0&
            Caption         =   "ENAC COMPLETA (Todos los ensayos estan certificados por ENAC)"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   155
            Top             =   540
            Width           =   5280
         End
         Begin VB.OptionButton opENAC 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO ENAC"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   154
            Top             =   315
            Value           =   -1  'True
            Width           =   3480
         End
         Begin VB.CheckBox chkNadcap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "NADCAP"
            ForeColor       =   &H00000000&
            Height          =   240
            Left            =   5850
            TabIndex        =   153
            Top             =   225
            Width           =   1140
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos Específicos"
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
         Height          =   2580
         Left            =   7875
         TabIndex        =   148
         Top             =   6075
         Width           =   7620
         Begin VB.CommandButton cmdespecificas 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Dat. Especificos"
            Height          =   915
            Left            =   6255
            Picture         =   "frmVerMuestra2.frx":3D1E6
            Style           =   1  'Graphical
            TabIndex        =   150
            Top             =   1575
            Width           =   1230
         End
         Begin VB.CommandButton cmdSC 
            BackColor       =   &H000080FF&
            Caption         =   "Pedido S.C."
            Height          =   915
            Left            =   6255
            Picture         =   "frmVerMuestra2.frx":3DAB0
            Style           =   1  'Graphical
            TabIndex        =   149
            Top             =   270
            Width           =   1230
         End
         Begin MSComctlLib.ListView datos 
            Height          =   2235
            Left            =   90
            TabIndex        =   151
            Top             =   270
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   3942
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
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
      End
      Begin VB.Frame frmFacturacion 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturación"
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
         Height          =   1005
         Left            =   7875
         TabIndex        =   140
         Top             =   4995
         Width           =   7650
         Begin VB.CheckBox chkPRECIO_FIJADO 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Precio Fijado"
            Enabled         =   0   'False
            Height          =   240
            Left            =   3015
            TabIndex        =   209
            Top             =   270
            Width           =   1410
         End
         Begin VB.CheckBox chkAjuste 
            BackColor       =   &H00C0C0C0&
            Caption         =   "AJUSTE"
            Enabled         =   0   'False
            Height          =   240
            Left            =   4590
            TabIndex        =   200
            Top             =   630
            Width           =   1410
         End
         Begin VB.TextBox Text1 
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
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   2
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   144
            Top             =   585
            Width           =   1455
         End
         Begin VB.CommandButton cmdfactura 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Ver factura"
            Height          =   285
            Left            =   3015
            Style           =   1  'Graphical
            TabIndex        =   143
            Top             =   585
            Width           =   1275
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   19
            Left            =   1485
            Locked          =   -1  'True
            TabIndex        =   142
            Top             =   225
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Alignment       =   1  'Right Justify
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
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   6
            Left            =   6165
            Locked          =   -1  'True
            TabIndex        =   141
            Top             =   225
            Width           =   1320
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Factura número"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   6
            Left            =   180
            TabIndex        =   147
            Top             =   630
            Width           =   1185
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Precio"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   19
            Left            =   180
            TabIndex        =   146
            Top             =   270
            Width           =   645
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Por Determinaciones"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   23
            Left            =   4590
            TabIndex        =   145
            Top             =   270
            Width           =   1470
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Otros Datos"
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
         Height          =   1200
         Left            =   7875
         TabIndex        =   133
         Top             =   3735
         Width           =   7650
         Begin VB.TextBox Text1 
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
            ForeColor       =   &H000000FF&
            Height          =   300
            Index           =   7
            Left            =   90
            Locked          =   -1  'True
            TabIndex        =   135
            Top             =   810
            Visible         =   0   'False
            Width           =   1320
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   555
            Index           =   21
            Left            =   1215
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   134
            Top             =   555
            Width           =   6270
         End
         Begin MSComCtl2.DTPicker FechaEntrega 
            Height          =   330
            Left            =   1215
            TabIndex        =   136
            Top             =   180
            Width           =   1470
            _ExtentX        =   2593
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
            Format          =   60358657
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin XtremeSuiteControls.PushButton cmdPaquete 
            Height          =   300
            Left            =   3015
            TabIndex        =   137
            Top             =   180
            Width           =   4470
            _Version        =   851970
            _ExtentX        =   7885
            _ExtentY        =   529
            _StockProps     =   79
            Caption         =   "Material Devuelto en Paquete : "
            Appearance      =   5
            Picture         =   "frmVerMuestra2.frx":3E37A
            TextImageRelation=   4
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Incidencias"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   20
            Left            =   45
            TabIndex        =   139
            Top             =   675
            Width           =   1140
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha entrega"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   21
            Left            =   45
            TabIndex        =   138
            Top             =   210
            Width           =   1140
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos de Recepción"
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
         Height          =   3285
         Left            =   7875
         TabIndex        =   116
         Top             =   405
         Width           =   7650
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   4
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   119
            Top             =   1305
            Width           =   6180
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   13
            Left            =   1305
            Locked          =   -1  'True
            TabIndex        =   118
            Top             =   2025
            Width           =   6180
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   525
            Index           =   14
            Left            =   1305
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   117
            Top             =   2385
            Width           =   6180
         End
         Begin MSComCtl2.DTPicker fechaRecepcion 
            Height          =   330
            Left            =   1305
            TabIndex        =   120
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   60358657
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo cmbDatos 
            Bindings        =   "frmVerMuestra2.frx":44BDC
            Height          =   315
            Index           =   4
            Left            =   1305
            TabIndex        =   121
            Top             =   960
            Width           =   6225
            _ExtentX        =   10980
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
         Begin MSDataListLib.DataCombo cmbDatos 
            Height          =   315
            Index           =   5
            Left            =   1305
            TabIndex        =   122
            Top             =   1665
            Width           =   6180
            _ExtentX        =   10901
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
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
         Begin pryCombo.miCombo cmbUsuario 
            Height          =   330
            Left            =   1305
            TabIndex        =   123
            Top             =   630
            Width           =   6225
            _ExtentX        =   10980
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbResponsable 
            Height          =   330
            Left            =   1305
            TabIndex        =   124
            Top             =   2925
            Width           =   6180
            _ExtentX        =   10901
            _ExtentY        =   582
         End
         Begin MSComCtl2.DTPicker horaRecepcion 
            Height          =   330
            Left            =   3420
            TabIndex        =   218
            Top             =   270
            Width           =   1155
            _ExtentX        =   2037
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
            CustomFormat    =   "00:00:00"
            Format          =   60358658
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Hora"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   47
            Left            =   2880
            TabIndex        =   217
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Precinto"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   29
            Left            =   90
            TabIndex        =   132
            Top             =   1350
            Width           =   1995
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Responsable"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   27
            Left            =   90
            TabIndex        =   131
            Top             =   2925
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Recepcion"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   9
            Left            =   90
            TabIndex        =   130
            Top             =   630
            Width           =   915
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   10
            Left            =   90
            TabIndex        =   129
            Top             =   315
            Width           =   735
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Entregada por"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   11
            Left            =   90
            TabIndex        =   128
            Top             =   1710
            Width           =   1995
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Envase"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   12
            Left            =   90
            TabIndex        =   127
            Top             =   990
            Width           =   1995
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Observaciones"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   13
            Left            =   90
            TabIndex        =   126
            Top             =   2520
            Width           =   1365
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Detalles"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   14
            Left            =   90
            TabIndex        =   125
            Top             =   2025
            Width           =   1995
         End
      End
      Begin VB.Frame frameDeterminaciones 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Determinaciones"
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
         Height          =   2580
         Left            =   180
         TabIndex        =   113
         Top             =   6075
         Width           =   7530
         Begin VB.CommandButton cmdListadoDeter 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Determinaciones"
            Height          =   915
            Left            =   6210
            Picture         =   "frmVerMuestra2.frx":44C22
            Style           =   1  'Graphical
            TabIndex        =   114
            Top             =   1575
            Width           =   1275
         End
         Begin MSComctlLib.ListView deter 
            Height          =   2235
            Left            =   90
            TabIndex        =   115
            Top             =   270
            Width           =   6045
            _ExtentX        =   10663
            _ExtentY        =   3942
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
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
      End
      Begin VB.Frame frmPlasma 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "SPECIMEN ID AND DESCRIPTION"
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
         Height          =   2595
         Left            =   180
         TabIndex        =   96
         Top             =   6075
         Width           =   7530
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   1440
            MaxLength       =   255
            TabIndex        =   102
            Top             =   1350
            Width           =   5895
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   1440
            MaxLength       =   255
            TabIndex        =   101
            Top             =   1035
            Width           =   5895
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   4950
            MaxLength       =   255
            TabIndex        =   100
            Top             =   1980
            Width           =   2390
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   1440
            MaxLength       =   255
            TabIndex        =   99
            Top             =   1665
            Width           =   2130
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   4950
            MaxLength       =   255
            TabIndex        =   98
            Top             =   1665
            Width           =   2390
         End
         Begin VB.TextBox txtDatos 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   1440
            MaxLength       =   255
            TabIndex        =   97
            Top             =   1980
            Width           =   2130
         End
         Begin pryCombo.miCombo cmbProcess 
            Height          =   345
            Left            =   1440
            TabIndex        =   103
            Top             =   315
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   609
         End
         Begin pryCombo.miCombo cmbCustomer 
            Height          =   345
            Left            =   1440
            TabIndex        =   104
            Top             =   675
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   609
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "S/N:"
            Height          =   195
            Index           =   3
            Left            =   90
            TabIndex        =   112
            Top             =   1710
            Width           =   345
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "SPECIMEN ID"
            Height          =   195
            Index           =   2
            Left            =   90
            TabIndex        =   111
            Top             =   1080
            Width           =   1035
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "P/N:"
            Height          =   195
            Index           =   8
            Left            =   3690
            TabIndex        =   110
            Top             =   1995
            Width           =   345
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "PROCESS"
            Height          =   195
            Index           =   11
            Left            =   90
            TabIndex        =   109
            Top             =   390
            Width           =   765
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "CUSTOMER"
            Height          =   195
            Index           =   14
            Left            =   90
            TabIndex        =   108
            Top             =   720
            Width           =   915
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "PRODUCT S/N:"
            Height          =   195
            Index           =   15
            Left            =   3690
            TabIndex        =   107
            Top             =   1710
            Width           =   1185
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "PRODUCT TYPE:"
            Height          =   195
            Index           =   16
            Left            =   90
            TabIndex        =   106
            Top             =   1410
            Width           =   1305
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "MODULE S/N:"
            Height          =   195
            Index           =   17
            Left            =   90
            TabIndex        =   105
            Top             =   2025
            Width           =   1080
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos del Muestreo"
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
         Height          =   1770
         Left            =   180
         TabIndex        =   79
         Top             =   4230
         Width           =   7545
         Begin VB.CheckBox chkFMSinEspecificar 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sin especificar"
            Enabled         =   0   'False
            Height          =   240
            Left            =   2790
            TabIndex        =   89
            Top             =   225
            Width           =   1410
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00004080&
            Height          =   315
            Index           =   17
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   88
            Top             =   900
            Width           =   4560
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   450
            Index           =   18
            Left            =   1350
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   87
            Top             =   1245
            Width           =   4560
         End
         Begin VB.CheckBox chkOpcion 
            BackColor       =   &H00C0C0C0&
            Caption         =   "No Rutinaria"
            Enabled         =   0   'False
            Height          =   240
            Index           =   3
            Left            =   4230
            TabIndex        =   86
            Top             =   225
            Width           =   1365
         End
         Begin VB.Frame frmTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo"
            ForeColor       =   &H80000008&
            Height          =   1230
            Left            =   5985
            TabIndex        =   81
            Top             =   450
            Width           =   1500
            Begin VB.CheckBox chkOpcion 
               BackColor       =   &H00C0C0C0&
               Caption         =   "A/C-Vuelo"
               Enabled         =   0   'False
               Height          =   240
               Index           =   0
               Left            =   135
               TabIndex        =   85
               Top             =   450
               Width           =   1095
            End
            Begin VB.CheckBox chkOpcion 
               BackColor       =   &H00C0C0C0&
               Caption         =   "In Situ"
               Enabled         =   0   'False
               Height          =   240
               Index           =   1
               Left            =   135
               TabIndex        =   84
               Top             =   675
               Width           =   1095
            End
            Begin VB.CheckBox chkOpcion 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Lab.Movil"
               Enabled         =   0   'False
               Height          =   240
               Index           =   2
               Left            =   135
               TabIndex        =   83
               Top             =   900
               Width           =   1095
            End
            Begin VB.CheckBox chkOpcion 
               BackColor       =   &H00C0C0C0&
               Caption         =   "No Aplica"
               Enabled         =   0   'False
               Height          =   240
               Index           =   4
               Left            =   135
               TabIndex        =   82
               Top             =   225
               Width           =   1095
            End
         End
         Begin VB.CheckBox chkOpcion 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Repetición"
            Enabled         =   0   'False
            Height          =   240
            Index           =   5
            Left            =   5670
            TabIndex        =   80
            Top             =   225
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker fechaMuestreo 
            Height          =   330
            Left            =   1350
            TabIndex        =   90
            Top             =   195
            Width           =   1380
            _ExtentX        =   2434
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
            Format          =   60358657
            CurrentDate     =   38000
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo cmbDatos 
            Height          =   315
            Index           =   6
            Left            =   1350
            TabIndex        =   91
            Top             =   555
            Width           =   4590
            _ExtentX        =   8096
            _ExtentY        =   556
            _Version        =   393216
            Locked          =   -1  'True
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
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Realizada por"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   15
            Left            =   90
            TabIndex        =   95
            Top             =   585
            Width           =   1995
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   16
            Left            =   90
            TabIndex        =   94
            Top             =   240
            Width           =   1995
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Observaciones"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   17
            Left            =   90
            TabIndex        =   93
            Top             =   1380
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Detalles"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   18
            Left            =   90
            TabIndex        =   92
            Top             =   930
            Width           =   1995
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Datos del Registro"
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
         Height          =   3765
         Left            =   180
         TabIndex        =   52
         Top             =   405
         Width           =   7545
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   9
            Left            =   1350
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   60
            Top             =   3015
            Width           =   6090
         End
         Begin VB.TextBox Text1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   315
            Index           =   8
            Left            =   1350
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            TabIndex        =   59
            Top             =   1665
            Width           =   6090
         End
         Begin VB.OptionButton opDuplicado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   2190
            TabIndex        =   58
            Top             =   2025
            Width           =   615
         End
         Begin VB.OptionButton opDuplicado 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "SI"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   1590
            TabIndex        =   57
            Top             =   2025
            Width           =   615
         End
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   4050
            Picture         =   "frmVerMuestra2.frx":454EC
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   56
            Top             =   2025
            Width           =   240
         End
         Begin VB.Frame Frame7 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5175
            TabIndex        =   53
            Top             =   1980
            Width           =   1455
            Begin VB.OptionButton opUrgente 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "SI"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   1
               Left            =   90
               TabIndex        =   55
               Top             =   45
               Width           =   615
            End
            Begin VB.OptionButton opUrgente 
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               Caption         =   "NO"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   225
               Index           =   0
               Left            =   720
               TabIndex        =   54
               Top             =   45
               Value           =   -1  'True
               Width           =   615
            End
         End
         Begin pryCombo.miCombo cmbClientes 
            Height          =   330
            Left            =   1350
            TabIndex        =   61
            Top             =   225
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   582
         End
         Begin MSDataListLib.DataCombo cmbPedidos 
            Bindings        =   "frmVerMuestra2.frx":4BD3E
            Height          =   315
            Left            =   1350
            TabIndex        =   62
            Top             =   2295
            Width           =   5085
            _ExtentX        =   8969
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
         Begin pryCombo.miCombo cmbTM 
            Height          =   330
            Left            =   1350
            TabIndex        =   63
            Top             =   585
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbTA 
            Height          =   330
            Left            =   1350
            TabIndex        =   64
            Top             =   945
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   582
         End
         Begin pryCombo.miCombo cmbBano 
            Height          =   375
            Left            =   1350
            TabIndex        =   65
            Top             =   1305
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   661
         End
         Begin pryCombo.miCombo cmbOferta 
            Height          =   375
            Left            =   1350
            TabIndex        =   66
            Top             =   2655
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   661
         End
         Begin pryCombo.miCombo cmbReplacement 
            Height          =   345
            Left            =   1350
            TabIndex        =   67
            Top             =   3375
            Width           =   6135
            _ExtentX        =   10821
            _ExtentY        =   609
         End
         Begin VB.Image imgVerPedido 
            Height          =   360
            Left            =   6480
            Picture         =   "frmVerMuestra2.frx":4BD84
            Stretch         =   -1  'True
            ToolTipText     =   "Ver Pedido"
            Top             =   2295
            Width           =   360
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Producto"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   31
            Left            =   90
            TabIndex        =   78
            Top             =   3045
            Width           =   1065
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pedido"
            Height          =   195
            Index           =   4
            Left            =   90
            TabIndex        =   77
            Top             =   2340
            Width           =   525
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cliente"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   2
            Left            =   90
            TabIndex        =   76
            Top             =   300
            Width           =   1995
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo de muestra"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   3
            Left            =   90
            TabIndex        =   75
            Top             =   630
            Width           =   1995
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Nombre Baño"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   4
            Left            =   90
            TabIndex        =   74
            Top             =   1335
            Width           =   1995
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo de analisis"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   5
            Left            =   90
            TabIndex        =   73
            Top             =   1005
            Width           =   1995
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Analisis duplicado"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   7
            Left            =   90
            TabIndex        =   72
            Top             =   2025
            Width           =   1455
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ref. muestra"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   8
            Left            =   90
            TabIndex        =   71
            Top             =   1695
            Width           =   1065
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Oferta"
            Height          =   195
            Index           =   0
            Left            =   90
            TabIndex        =   70
            Top             =   2700
            Width           =   465
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Replacement"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   69
            Top             =   3450
            Width           =   945
         End
         Begin VB.Label Label2 
            BackColor       =   &H00C0C0C0&
            Caption         =   "URGENTE"
            ForeColor       =   &H00000000&
            Height          =   285
            Index           =   40
            Left            =   4320
            TabIndex        =   68
            Top             =   2025
            Width           =   915
         End
      End
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "  Consulta de Muestras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Index           =   5
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15705
   End
   Begin VB.Label lblTipo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ABIERTA"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   30
      TabIndex        =   37
      Top             =   360
      Width           =   3765
   End
   Begin VB.Label lblf5 
      BackColor       =   &H00C0C0FF&
      Caption         =   "F5 - DATOS ESPECIALES"
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
      Left            =   11550
      TabIndex        =   33
      Top             =   60
      Visible         =   0   'False
      Width           =   2805
   End
   Begin VB.Label lblestado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ABIERTA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   12630
      TabIndex        =   6
      Top             =   405
      Width           =   3045
   End
End
Attribute VB_Name = "frmVerMuestra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Enum ColsTAB
    COL_GENERAL = 0
    COL_OTROS = 1
    COL_FLUIDOS = 2
End Enum
Dim frecepcion As Date
Dim fmuestreo As Date
Dim fentrega As Date

Private Sub cmdAnadirNorma_Click()
   On Error GoTo cmdAnadirNorma_Click_Error

    If cmbNormas.getPK_SALIDA = 0 Then
        MsgBox "Debe seleccionar una de entre las existentes", vbOK + vbExclamation, "Añadir Norma"
        Exit Sub
    End If
    Dim oMN As New clsMuestras_normas
    With oMN
        .setMUESTRA_ID = gmuestra
        .setNORMA_ID = cmbNormas.getPK_SALIDA
        .Insertar
    End With
    cmbNormas.limpiar
    Call cargar_normas

   On Error GoTo 0
   Exit Sub

cmdAnadirNorma_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadirNorma_Click of Formulario frmVerMuestra"
End Sub
Private Sub cargar_normas()
    Dim rs As ADODB.Recordset
    Dim oMN As New clsMuestras_normas
    listaNormas.ListItems.Clear
    Set rs = oMN.Listado(gmuestra)
    If rs.RecordCount > 0 Then
        Do
            With listaNormas.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(1)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub
Private Sub cmdEdiciones_Click()
    With frmMuestras_Ediciones
        .PK = gmuestra
        .Show 1
    End With
End Sub
Private Sub chkFechaCierre_Click()
    If chkFechaCierre.Value = Checked Then
        ffin.Enabled = True
        hfin.Enabled = True
    Else
        ffin.Enabled = False
        hfin.Enabled = False
    End If
End Sub
Private Sub chkFechaEnvio_Click()
    If chkFechaEnvio.Value = Checked Then
        fEnvio.Enabled = True
        hEnvio.Enabled = True
    Else
        fEnvio.Enabled = False
        hEnvio.Enabled = False
    End If
End Sub
Private Sub chkFechaSolicitud_Click()
    If chkFechaSolicitud.Value = Checked Then
        fechaSolicitud = Date
        horaSolicitud = Date & " " & Time
        fechaSolicitud.Enabled = True
        horaSolicitud.Enabled = True
    Else
        fechaSolicitud.Enabled = False
        horaSolicitud.Enabled = False
    End If
End Sub

Private Sub cmdCambiar_Click(Index As Integer)
    Dim omue As New clsMuestra
    If omue.CargaMuestra(gmuestra) = True Then
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        If Index = 1 Then
            consulta = "select id_muestra from muestras where tipo_muestra_id = " & omue.getTIPO_MUESTRA_ID & " and id_muestra > " & gmuestra & " order by id_muestra asc"
        Else
            consulta = "select id_muestra from muestras where tipo_muestra_id = " & omue.getTIPO_MUESTRA_ID & " and id_muestra < " & gmuestra & " order by id_muestra desc"
        End If
        Set rs = datos_bd(consulta)
        If rs.RecordCount <> 0 Then
            gmuestra = rs.Fields(0)
            consulta_muestra
         Else
            If Index = 1 Then
                MsgBox "No existen muestras con código superior.", vbInformation, App.Title
            Else
                MsgBox "No existen muestras con código inferior.", vbInformation, App.Title
            End If
        End If
    End If
    Set omue = Nothing
End Sub
Private Sub chkFMSinEspecificar_Click()
    If chkFMSinEspecificar.Value = Checked Then
        fechaMuestreo.Enabled = False
    Else
        fechaMuestreo.Enabled = True
    End If
End Sub

Private Sub cmbBano_Change()
    If cmbBano.getTEXTO <> "" Then
        Dim oBANO As New clsBanos
        Dim oMuestra As New clsMuestra
        If Not oMuestra.esControlEficacia(gmuestra) Then
            oBANO.cargar_bano (cmbBano.getPK_SALIDA)
            llenar_combo cmbTA, New clsTipos_analisis, 0, frmTA_Detalle, ""
            cmbTA.cargar_datos
            cmbTA.MostrarElemento oBANO.getID_SOLUCION
            If Text1(8) = "" Then
                Text1(8) = cmbBano.getTEXTO
            End If
        End If
    End If
End Sub
Private Sub cmbClientes_change()
    cmbPedidos.Text = ""
    If cmbClientes.getTEXTO <> "" Then
         pedidos (cmbClientes.getPK_SALIDA)
    End If
End Sub

Private Sub cmbTM_Change()
    On Error Resume Next
    Dim tipo As New clsMuestra
    cmbBano.limpiar
    cmbTA.limpiar
    If cmbTM.getTEXTO <> "" Then
     If Not tipo.esBano(cmbTM.getPK_SALIDA) Then   'es un baño Id_Espacial
        ' Es una determinacion de muestra
        Dim oAnalisis As New clsTipos_analisis
        cmbBano.desactivar
        Label2(5).Caption = "Tipo de análisis"
        cmbTA.activar
        llenar_combo cmbTA, New clsTipos_analisis, cmbTM.getPK_SALIDA, frmTA_Detalle, " ANULADO = 0 "
        cmbTA.cargar_datos
        If tipo.esControlEficacia(gmuestra) Then
            cmbBano.activar
        End If
        CARGAR_COMBO_BANOS cmbClientes.getPK_SALIDA, ""
     Else
      ' Es un baño especial
         Label2(5).Caption = "Solución"
         If cmbClientes.getTEXTO = "" Then
             MsgBox "Seleccione primero un cliente.", vbInformation, App.Title
             cmbTM.limpiar
             cmbClientes.SetFocus
             Exit Sub
         End If
         cmbTA.desactivar
         Dim oBANO As New clsBanos
         Dim rsbano As ADODB.Recordset
         Set rsbano = oBANO.banos_cliente(cmbClientes.getPK_SALIDA, cmbTM.getPK_SALIDA)
         If rsbano.RecordCount = 0 Then
             MsgBox "No hay baños para el cliente y tipo de muestra seleccionado.", vbInformation, App.Title
             cmbTM.limpiar
             Exit Sub
         Else
'             CARGAR_COMBO_BANOS cmbClientes.getPK_SALIDA, " TIPO_MUESTRA_ID = " & cmbTM.getPK_SALIDA
             CARGAR_COMBO_BANOS cmbClientes.getPK_SALIDA, ""
         End If
         cmbBano.activar
         'JGM-I
         If tipo.esControlEficacia(gmuestra) Then
            cmbTA.activar
            llenar_combo cmbTA, New clsTipos_analisis, cmbTM.getPK_SALIDA, frmTA_Detalle, " ANULADO = 0 "
            cmbTA.cargar_datos
         End If
         'JGM-F
         Set oBANO = Nothing
     End If 'fin del esBanno
     Set tipo = Nothing
     Set oAnalisis = Nothing
    End If
End Sub

Private Sub cmdAbrirMuestra_Click()
    On Error GoTo fallo
    ' Verificar si la edición anterior se genero
    Dim oMuestra As New clsMuestra
    If oMuestra.CargaMuestra(gmuestra) Then
        If oMuestra.getULT_EDICION_IMP = 0 Then
            If oMuestra.Abrir(gmuestra, False) Then
                consulta_muestra
            End If
            Exit Sub
        End If
'        Dim destino As String
'        destino = NOMBRE_DOCUMENTO(gmuestra, True) & ".pdf"
'        If Dir(destino) = "" Then
'            MsgBox "La edición anterior falló al generarse, por lo que no pueden generarse nuevas ediciones. Contacte con mantenimiento.", vbExclamation, App.Title
'            Exit Sub
'        End If

        ' Verificar si es Agua o Baño si tiene alguna determinación pendiente y si es así, reabrir la muestra y disminuir Edición
        If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.TM_AGUA Or oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.TM_BANO Then
            Dim oDET As New clsDeterminaciones
            If oDET.existePendiente(gmuestra) = True Then
                If oMuestra.Abrir(gmuestra, True) Then
                    oMuestra.disminuir_edicion_impresa gmuestra
                    MsgBox "Muestra REABIERTA sin nueva edición, ya que existen determinaciones PENDIENTES.", vbExclamation, App.Title
                    consulta_muestra
                    Exit Sub
                End If
            End If
        End If

        If MsgBox("¿Desea abrir la muestra? Generará una nueva edición del informe.", vbQuestion + vbYesNo, App.Title) = vbYes Then
             ' Motivo de nueva edición
             frmMotivo.Show 1
             If Trim(MOTIVO) = "" Then
                 MsgBox "Para generar una nueva edición es necesario introducir un motivo.", vbInformation, App.Title
                 Exit Sub
             End If
             Dim oMe As New clsMuestras_ediciones
'             If oMuestra.Nueva_Edicion(gmuestra, oMuestra.getULT_EDICION_IMP + 1, Trim(MOTIVO), Date) = False Then
'                 Exit Sub
'             End If
             With oMe
                .setMUESTRA_ID = gmuestra
                .setEDICION = oMuestra.getULT_EDICION_IMP + 1
                .setOBSERVACIONES = Trim(MOTIVO)
                .setFECHA = Date
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                If .Insertar(False) = 0 Then
                    MsgBox "Error al insertar la edición.", vbCritical, App.Title
                Else
                    If oMuestra.Abrir(gmuestra, False) Then
                       With frmMuestras_Ediciones
                           .PK = gmuestra
                           .Show 1
                       End With
                       consulta_muestra
                    End If
                End If
             End With
        End If
    End If
    Me.MousePointer = 0
    Me.SetFocus
    Exit Sub
fallo:
    MsgBox "Se ha producido un error al abrir la muestra.", vbCritical, App.Title
End Sub

Private Sub cmdAdjuntos_Click()
'M0499-I
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_MUESTRAS
        .COBJETO = gmuestra
        .Show 1
    End With
    Set frmAdjuntos = Nothing
'M0499-F
    consulta_muestra
End Sub

Private Sub cmdAnular_Click()
    If MsgBox("Va a anular la muestra. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        frmMotivo.Show 1
        If Trim(MOTIVO) = "" Then
            MsgBox "Para anular la muestra es necesario introducir un motivo.", vbInformation, App.Title
            Exit Sub
        End If
        Dim oMuestra As New clsMuestra
        If oMuestra.Anular(CLng(Text1(0)), Trim(MOTIVO)) Then
            consulta_muestra
        End If
        Set oMuestra = Nothing
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdContra_Click()
    If MsgBox("Va a generar el contradictorio de este análisis. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
      On Error GoTo fallo
      Dim omuestra_origen As New clsMuestra
      Dim omuestra_destino As New clsMuestra
      omuestra_origen.CargaMuestra (CLng(Text1(0)))
      With omuestra_destino
        .setTIPO_MUESTRA_ID = omuestra_origen.getTIPO_MUESTRA_ID
        .setTIPO_ANALISIS_ID = omuestra_origen.getTIPO_ANALISIS_ID
        .setANALISIS_MODIFICADO = omuestra_origen.getANALISIS_MODIFICADO
        ' Incidencia 291
        '.setFECHA_MUESTREO = omuestra_origen.getFECHA_MUESTREO
        .setFECHA_MUESTREO = Format(CDate(omuestra_origen.getFECHA_MUESTREO), "yyyy-mm-dd")
        .setENTIDAD_MUESTREO_ID = omuestra_origen.getENTIDAD_MUESTREO_ID
        .setDETALLE_MUESTREO = omuestra_origen.getDETALLE_MUESTREO
        .setOBSERVACIONES_MUESTREO = omuestra_origen.getOBSERVACIONES_MUESTREO
        ' Incidencia 291
        '.setFECHA_RECEPCION = Format(Date, "yyyy-mm-dd")
        .setFECHA_RECEPCION = Format(CDate(omuestra_origen.getFECHA_RECEPCION), "yyyy-mm-dd")
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setFORMATO_ID = omuestra_origen.getFORMATO_ID
        .setENTIDAD_ENTREGA_ID = omuestra_origen.getENTIDAD_ENTREGA_ID
        .setDETALLE_ENTREGA = omuestra_origen.getDETALLE_ENTREGA
        .setOBSERVACIONES_ENTREGA = omuestra_origen.getOBSERVACIONES_ENTREGA
        .setCLIENTE_ID = omuestra_origen.getCLIENTE_ID
        .setREFERENCIA_CLIENTE = omuestra_origen.getREFERENCIA_CLIENTE & " (2ª.MUESTRA)"
        .setPRECIO = omuestra_origen.getPRECIO
        ' Incidencia 291
        '.setFECHA_PREV_FIN = Format(Date + 10, "yyyy-mm-dd")
        .setFECHA_PREV_FIN = Format(CDate(omuestra_origen.getFECHA_RECEPCION) + 15, "yyyy-mm-dd")
        .setOBSERVACIONES = omuestra_origen.getOBSERVACIONES
        .setANULADA = 0
        .setPRECINTO = omuestra_origen.getPRECINTO
        .setBANO_ID = omuestra_origen.getBANO_ID
        ' J51
        .setFECHA_COMIENZO = "0000-00-00"
        .setFECHA_FINALIZACION = "0000-00-00"
        .setFECHA_CIERRE = "0000-00-00"
        .setCERRADA = 0
        .setDOCUMENTO_PAGO = 0
        .setULT_EDICION_IMP = 0
        .setREPLACEMENT_ID = 0
        .guardarMuestra
     End With
     ' Insertar determinaciones
     Dim rs_deter As ADODB.Recordset
     Dim oDeter As New clsDeterminaciones
     Dim oDatosDet As New clsDatos_determinaciones
     Dim ocampos As New clsFormulas_campos
     Dim rscampos As ADODB.Recordset
     Dim DETERMINACION As Long
     Set rs_deter = oDeter.lista_contradictorio(CLng(Text1(0)))
     If rs_deter.RecordCount <> 0 Then
        Do
            oDeter.setMUESTRA_ID = omuestra_destino.getID_MUESTRA
            oDeter.setTIPO_DETERMINACION_ID = rs_deter("tipo_determinacion_id")
            oDeter.setORDEN = rs_deter("orden")
            oDeter.setFORMULA_ID = rs_deter("formula_id")
            oDeter.setES_DUPLICADO = rs_deter("es_duplicado")
            oDeter.setSITUACION = rs_deter("situacion")
            DETERMINACION = oDeter.InsertarDeterminacion
            ' Recuperar formulas_camposs (CAMPO_ID)
            Set rscampos = ocampos.ListaFormulas(rs_deter("formula_id"))
            ' Insertar Datos_Determinaciones
            If rscampos.RecordCount <> 0 Then
              Do
               oDatosDet.setDETERMINACION_ID = DETERMINACION
               oDatosDet.setCAMPO_ID = rscampos("id_campo")
'               oDatosDet.setVALOR_1 = "I-1"
'               oDatosDet.setVALOR_2 = "I-2"
               oDatosDet.setVALOR_1 = ""
               oDatosDet.setVALOR_2 = ""
               oDatosDet.Insertar
               rscampos.MoveNext
              Loop Until rscampos.EOF
            End If
            rs_deter.MoveNext
        Loop Until rs_deter.EOF
     End If
     ' Datos_valores
     Dim ovalmuestra As New clsDatos_valores
     ' Insertar por defecto el 100 (Fecha y hora del contradictorio)
     With ovalmuestra
        .setMUESTRA_ID = omuestra_destino.getID_MUESTRA
        .setTIPO_DATO_ID = 100
        .setVALOR = ""
        .setORDEN = 1
        .Insertar
     End With
     Dim indice As Integer
     Dim rs_oval As ADODB.Recordset
     indice = 2
     Set rs_oval = ovalmuestra.datos_muestra(CLng(Text1(0)))
     If rs_oval.RecordCount <> 0 Then
        Do
            ovalmuestra.setMUESTRA_ID = omuestra_destino.getID_MUESTRA
            ovalmuestra.setBANO_ID = rs_oval("bano_id")
            ovalmuestra.setTIPO_DATO_ID = rs_oval("tipo_dato_id")
            If rs_oval("tipo_dato_id") = 15 Then ' Contradictorio
                ovalmuestra.setVALOR = omuestra_origen.getID_GENERAL & "/" & omuestra_origen.getANNO
            Else
                ovalmuestra.setVALOR = rs_oval("valor")
            End If
            ovalmuestra.setORDEN = indice
            ovalmuestra.Insertar
            indice = indice + 1
            rs_oval.MoveNext
        Loop Until rs_oval.EOF
     End If
     ' Informe de recepcion
'     imprimir omuestra_destino.getID_MUESTRA, 10, False
     MsgBox "El contradictorio ha sido registrado con el Nº: " & omuestra_destino.getID_GENERAL & " y código: " & omuestra_destino.CodigoParticular(omuestra_destino.getID_MUESTRA), vbInformation, App.Title
    End If
    Exit Sub
fallo:
    MsgBox "Error al generar el contradictorio. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdDeter_Click()
        gmuestra = CLng(Text1(0))
        abrirRegistroMuestra gmuestra

'        Dim oMuestra As New clsMuestra
'        oMuestra.CargaMuestra (gmuestra)
'        Select Case oMuestra.getANALISIS_MODIFICADO
'            Case 2 ' Control de eficacia
'                With frmCE_Resultados
'                    .PK_ID_MUESTRA = gmuestra
'                    .Show 1
'                End With
'            Case 3 ' Sellante
'                frmSE_Resultados.Show 1
'            Case 5 ' Plasma
'                If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_ROCKWELL Or _
'                   oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_BRINELL Or _
'                   oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_VICKERS Then
'                    With frmPlasma_Dureza
'                        .PK = gmuestra
'                        .Show 1
'                    End With
'                Else
'                    With frmPlasma_Resultados
'                        .PK = gmuestra
'                        .Show 1
'                    End With
'                End If
'            Case Else
'                frmDeterminaciones.Show 1
'        End Select
        consulta_muestra
End Sub

Private Sub cmdEliminarNorma_Click()
   On Error GoTo cmdEliminarNorma_Click_Error

      If listaNormas.ListItems.Count = 0 Then
         Exit Sub
      End If
      Dim oMN As New clsMuestras_normas
      oMN.Eliminar gmuestra, listaNormas.ListItems(listaNormas.selectedItem.Index).Text
      cargar_normas

   On Error GoTo 0
   Exit Sub

cmdEliminarNorma_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminarNorma_Click of Formulario frmVerMuestra"
End Sub

Private Sub cmdespecificas_Click()
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra (CLng(Text1(0)))
    frmDatosEspecificos.PK_MUESTRA = CLng(Text1(0))
    frmDatosEspecificos.PK_BANO = oMuestra.getBANO_ID
    frmDatosEspecificos.Show 1
    cargar_datos_especificos
End Sub

Private Sub cmdetiqueta_Click()
    ReDim ETIQUETAS(1)
    ETIQUETAS(1) = gmuestra
    frmEtiquetas.Show 1
End Sub

Private Sub cmdEtiquetaSoluciones_Click()
    frmSoluciones_Etiqueta.PK = gmuestra
    frmSoluciones_Etiqueta.Show 1
End Sub

Private Sub cmdFactura_Click()
   On Error GoTo cmdfactura_Click_Error
    Dim oDoc_pago As New clsDocs_pago
    oDoc_pago.generar_factura CLng(Text1(7)), False, "", "rptFactura"
   On Error GoTo 0
   Exit Sub

cmdfactura_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdfactura_Click of Formulario frmVerMuestra"
End Sub

Private Sub cmdFirmaCliente_Click()
    frmFirma.Show 1
End Sub

Private Sub cmdFluido_Click()
    If cmbBano.getTEXTO = "" Then
        MsgBox "No tiene baño asignado.", vbExclamation, App.Title
    Else
        Dim oFluido As New clsFluidos_ficha
        If oFluido.Carga_por_BANO(cmbBano.getPK_SALIDA) = True Then
            frmFluidos_Detalle.PK = oFluido.getID_FLUIDO
            frmFluidos_Detalle.Show 1
        Else
            MsgBox "No tiene fluido asociado.", vbExclamation, App.Title
        End If
    End If
End Sub

Private Sub cmdInforme_Click()
    Me.MousePointer = 11
    MostrarInforme CLng(Text1(0))
    Me.MousePointer = 0
End Sub

Private Sub cmdInfRegistro_Click()
    Dim oMuestra As New clsMuestra
    oMuestra.Informe_Recepcion gmuestra, False
    Set oMuestra = Nothing
End Sub

Private Sub cmdListadoDeter_Click()
    gmuestra = CLng(Text1(0))
    frmVerDeterminaciones.Show 1
    consulta_muestra
'    gmuestra = 0
    cargar_determinaciones
End Sub

Private Sub cmdModificar_Click()
    Dim color As Single
    color = &H80C0FF
    'Titulo
'    Label1(5).BackColor = color
'    Label1(5).Caption = "  MODIFICACION de la Muestra : " & Text1(5) & " (" & Text1(1) & "-" & Text1(3) & ")"
    Label1(5).Caption = Replace(Label1(5).Caption, "Consulta", "MODIFICACION")
    Me.Caption = Label1(5).Caption
    'Registro
    Text1(8).Locked = False
    Text1(9).Locked = False
    'Recepcion
    cmbClientes.activar
    'M1009-I
    cmbOferta.activar
    'M1009-F
    cmbTM.activar
    cmbTA.activar
    cmbUsuario.activar
    cmbBano.activar
    cmbDatos(4).Locked = False
    cmbCentroMuestra.Locked = False
    cmbReplacement.activar
    cmbDatos(5).Locked = False
    Text1(4).Locked = False
    Text1(13).Locked = False
    Text1(14).Locked = False
    opDuplicado(0).Enabled = True
    opDuplicado(1).Enabled = True
    opUrgente(0).Enabled = True
    opUrgente(1).Enabled = True
    chkOpcion(0).Enabled = True
    chkOpcion(1).Enabled = True
    chkOpcion(2).Enabled = True
    chkOpcion(3).Enabled = True
    chkOpcion(5).Enabled = True
    chkAjuste.Enabled = True
    chkConsulta.Enabled = True
    txtCONSULTA_OBSERVACIONES.Locked = False
    chkPRECIO_FIJADO.Enabled = True
    ' Muestreo
    cmbDatos(6).Locked = False
    Text1(17).Locked = False
    Text1(18).Locked = False
    Text1(10).Locked = False
    txtFluidoNormativa.Locked = False
    ' Otros datos
    If USUARIO.getPER_FACTURACION = True Then
        Text1(19).Locked = False
    End If
    Text1(21).Locked = False
    ' Botones
    cmdAnular.Enabled = False
    cmdDeter.Enabled = False
    cmdInfRegistro.Enabled = False
    cmdContra.Enabled = False
    cmdok.visible = True
    cmdModificar.Enabled = False
    cmdVida.Enabled = False
    cmdListadoDeter.Enabled = False
    cmdInforme.Enabled = False
    cmdespecificas.Enabled = False
    cmbPedidos.Locked = False
    cmbResponsable.activar
    chkFMSinEspecificar.Enabled = True
    
    chkIPA.Enabled = True
    'M1105-I
    frmAIM.Enabled = True
    frmADS.Enabled = True
    'M1105-F
    frmENAC.Enabled = True
    chkFechaSolicitud.Enabled = True
    ' No modificar ref si cerrada
'    Dim oMuestra As New clsMuestra
'    oMuestra.CargaMuestra (CLng(Text1(0)))
'    If oMuestra.getCERRADA = 1 Then
'        MsgBox "ATENCION : Al modificar la muestra cerrada, se regenerará el informe.", vbInformation, App.Title
'        Text1(8).Locked = True
'    End If
    cmdListadoDeter.Enabled = False
    cmdespecificas.Enabled = False
    Text1(8).SetFocus
End Sub
Private Function validar() As Boolean
    validar = True
    If IsNumeric(Text1(19)) = False Then
        MsgBox "El precio debe ser numérico.", vbCritical, "Error"
        Text1(19).SetFocus
        validar = False
        Exit Function
    End If
    If cmbClientes.getTEXTO = "" Then
        MsgBox "El cliente debe estar informado.", vbInformation, App.Title
        validar = False
        cmbClientes.SetFocus
        Exit Function
    End If
    If cmbTM.getTEXTO = "" Then
        MsgBox "El tipo de muestra debe estar informado.", vbInformation, App.Title
        validar = False
        cmbTM.SetFocus
        Exit Function
    End If
    If cmbTA.getTEXTO = "" Then
        MsgBox "El tipo de análisis o el baño deben estar informados.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If cmbUsuario.getTEXTO = "" Then
        MsgBox "El usuario que recepciona debe estar informado.", vbInformation, App.Title
        validar = False
        cmbUsuario.SetFocus
        Exit Function
    End If
    If cmbCentroMuestra.Text = "" Then
        MsgBox "El CENTRO de la muestra debe estar informado.", vbInformation, App.Title
        validar = False
        cmbCentroMuestra.SetFocus
        Exit Function
    End If
    If chkFMSinEspecificar.Value = Unchecked Then
        If Format(fechaMuestreo.Value, "yyyy-mm-dd") = "1900-01-01" Then
            MsgBox "La fecha de muestreo debe estar informada.", vbInformation, App.Title
            validar = False
            fechaMuestreo.SetFocus
            Exit Function
        End If
    End If
End Function

Private Sub cmdModificarDatosEspeciales_Click()
    If MsgBox("¿Esta seguro de modificar los datos?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim oMuestra As New clsMuestra
        With oMuestra
            .setULT_EDICION_IMP = txtedicion
            .setFECHA_CIERRE = ffin.Value
            .setHORA_CIERRE = Format(hfin.Value, "hh:mm:ss")
            .setFECHA_COMIENZO = fcomienzo.Value
            .setFECHA_FINALIZACION = fFinalizacion.Value
            .setCERRADA_USUARIO = cmbUsuarioCierre.getPK_SALIDA
            .setANULADA = chkAnulada.Value
            .setINFORME_MANUAL = chkInformeManual.Value
            If chkFechaEnvio.Value = Checked Then
                .setFECHA_ENVIO = Format(fEnvio, "yyyy-mm-dd") & " " & Format(hEnvio, "hh:mm:ss")
            Else
                .setFECHA_ENVIO = ""
            End If
            .Modificar_Otros_datos CLng(Text1(0))
        End With
        MsgBox "Datos modificados correctamente.", vbInformation + vbOKOnly, App.Title
        frmDatosEspeciales.visible = False
        consulta_muestra
        Set oMuestra = Nothing
    End If
End Sub

Private Sub cmdok_Click()
    On Error GoTo fallo
    Dim s As String
    Dim consulta As String
    If validar = False Then
        Exit Sub
    End If
    ' Modificamos la muestra
    Me.MousePointer = 11
    Dim oMuestra As New clsMuestra
    ' Validar el cambio de tipo de muestra
    oMuestra.CargaMuestra (CLng(Text1(0)))
    'M1053-I
    Dim auxIPA As Integer
    auxIPA = oMuestra.getIPA
    'M1053-F
    If oMuestra.getTIPO_MUESTRA_ID <> cmbTM.getPK_SALIDA Then
        If MsgBox("Al cambiar el tipo de muestra, cambia la numeracion. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    If oMuestra.getBANO_ID <> 0 And cmbBano.getTEXTO = "" Then
        If MsgBox("Se esta modificando la muestra SIN EL BAÑO INFORMADO. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    With oMuestra
        .setREFERENCIA_CLIENTE = Trim(Text1(8))
        .setFECHA_RECEPCION = Format(fechaRecepcion.Value, "yyyy-mm-dd")
        .setHORA_RECEPCION = Format(horaRecepcion.Value, "hh:mm:ss")
    
        If cmbDatos(4).BoundText <> "" Then
            .setFORMATO_ID = cmbDatos(4).BoundText
        Else
            .setFORMATO_ID = 0
        End If
        If cmbCentroMuestra.BoundText <> "" Then
            .setCENTRO_ID = cmbCentroMuestra.BoundText
        Else
            .setCENTRO_ID = 0
        End If
        If cmbReplacement.getTEXTO = "" Then
            .setREPLACEMENT_ID = 0
        Else
            .setREPLACEMENT_ID = cmbReplacement.getPK_SALIDA
        End If
        .setENTIDAD_ENTREGA_ID = cmbDatos(5).BoundText
        .setDETALLE_ENTREGA = Trim(Text1(13))
        .setOBSERVACIONES_ENTREGA = Trim(Text1(14))
        If chkFMSinEspecificar.Value = Checked Then
            .setFECHA_MUESTREO = "1900-01-01"
        Else
            .setFECHA_MUESTREO = Format(fechaMuestreo.Value, "yyyy-mm-dd")
        End If
        .setENTIDAD_MUESTREO_ID = cmbDatos(6).BoundText
        .setCLIENTE_ID = cmbClientes.getPK_SALIDA
        .setDETALLE_MUESTREO = Trim(Text1(17))
        .setOBSERVACIONES_MUESTREO = Trim(Text1(18))
        .setRETRASO_OBSERVACION = Text1(10)
        .setFECHA_PREV_FIN = Format(FechaEntrega.Value, "yyyy-mm-dd")
        .setPRECIO = Replace(Format(Text1(19), "####0.00"), ",", ".")
        .setOBSERVACIONES = Trim(Text1(21))
        .setPRECINTO = Text1(4)
        If opDuplicado(0).Value = True Then
           .setANALISIS_DUPLICADO = 1
        Else
           .setANALISIS_DUPLICADO = 0
        End If
        If opUrgente(0).Value = True Then
            .setURGENTE = 0
        Else
            .setURGENTE = 1
        End If
        .setAJUSTE = chkAjuste.Value
        .setCONSULTA = chkConsulta.Value
        .setCONSULTA_OBSERVACIONES = txtCONSULTA_OBSERVACIONES
        .setPRECIO_FIJADO = chkPRECIO_FIJADO.Value
        If cmbPedidos.Text = "" Then
            .setPEDIDO_ID = 0
        Else
            .setPEDIDO_ID = cmbPedidos.BoundText
        End If
        'M1009-I
        If cmbOferta.getTEXTO = "" Then
            .setOFERTA_ID = 0
        Else
            .setOFERTA_ID = cmbOferta.getPK_SALIDA
        End If
        'M1009-F
        .setEMPLEADO_ID = cmbUsuario.getPK_SALIDA
        .setRESPONSABLE_ID = cmbResponsable.getPK_SALIDA
        .setPRODUCTO = Text1(9)
        .setOP_VUELO = chkOpcion(0).Value
        .setOP_INSITU = chkOpcion(1).Value
        .setOP_LABMOVIL = chkOpcion(2).Value
        .setOP_NORUTINARIA = chkOpcion(3).Value
        .setOP_REPETICION = chkOpcion(5).Value
        ' INDICADORES
        If chkFechaSolicitud.Value = Checked Then
            .setFECHA_RECOGIDA = Format(fechaSolicitud, "yyyy-mm-dd") & " " & Format(horaSolicitud, "hh:mm:ss")
        Else
            .setFECHA_RECOGIDA = ""
        End If
        .setIPA = chkIPA.Value
        If cmbMotivoRetraso.getTEXTO = "" Then
            .setMOTIVO_RETRASO_ID = 0
        Else
            .setMOTIVO_RETRASO_ID = cmbMotivoRetraso.getPK_SALIDA
        End If
        ' ENAC
        .setENAC = 0
        If opENAC(1).Value = True Then
            .setENAC = 1
        ElseIf opENAC(2).Value = True Then
            .setENAC = 2
        End If
        .setNADCAP = chkNadcap.Value
        .Modificar CLng(Text1(0))
        'M1105-I
'        If frmAIM.visible = True Then
        If tabPrincipal.TabEnabled(ColsTAB.COL_FLUIDOS) = True Then
            Dim oAIM As New clsMuestras_aim
            With oAIM
                .setAIM_PROGRAMA_ID = cmbPrograma.getPK_SALIDA
                .setAIM_CENTRO_ID = cmbCentro.getPK_SALIDA
                .setAIM_TIPO_ENSAYO_ID = cmbTipoEnsayo.getPK_SALIDA
                .setAIM_SECCION_ID = cmbSeccion.getPK_SALIDA
                .setAIM_ESTACION_ID = cmbEstacion.getPK_SALIDA
                .setMUESTRA_ID = CLng(Text1(0))
                .Insertar
'                .Modificar CLng(Text1(0))
            End With
            Set oAIM = Nothing
        End If
        'M1105-F
        'ADS
        If frmADS.visible = True Then
            Dim oADS As New clsMuestras_airbus
            With oADS
                .setMUESTRA_ID = CLng(Text1(0))
                .setENSAYO_ID = IIf(cmbEnsayo.getTEXTO = "", 0, cmbEnsayo.getPK_SALIDA)
                .setPROGRAMA_ID = IIf(cmbProgramaADS.getTEXTO = "", 0, cmbProgramaADS.getPK_SALIDA)
                .setSECTION_ID = IIf(cmbSection.getTEXTO = "", 0, cmbSection.getPK_SALIDA)
                .setFLUID_ID = IIf(cmbFluid.getTEXTO = "", 0, cmbFluid.getPK_SALIDA)
                .setFACILITY_ID = IIf(cmbFacility.getTEXTO = "", 0, cmbFacility.getPK_SALIDA)
                .Insertar True, True, True, True, True
            End With
        End If
    End With
    'M1053-I Mandar correo si se marca IPA
    If chkIPA.Value = Checked And auxIPA = 0 Then
        Dim oParametro As New clsParametros
        Dim destinatario As String
        Dim ASUNTO As String
        Dim mensaje As String
        oParametro.Carga parametros.PARAM_MUESTRA_IPA, ""
        destinatario = oParametro.getVALOR
        If destinatario <> "" Then
                ASUNTO = "Muestra marcada como IPA, Nº General : " & Text1(5) & " Código : " & Text1(1) & "-" & Text1(3)
                mensaje = "Se ha marcado como IPA la muestra del asunto. " & vbNewLine & vbNewLine
                mensaje = mensaje & vbNewLine
                mensaje = mensaje & " Motivo Retraso : " & cmbMotivoRetraso.getTEXTO & vbNewLine
                mensaje = mensaje & " Observaciones : " & Text1(10) & vbNewLine
                mensaje = mensaje & vbNewLine
                mensaje = mensaje & " Realizado por : " & USUARIO.getUSUARIO & vbNewLine
                mensaje = mensaje & " Fecha : " & Format(Date, "dd-mm-yyyy") & vbNewLine
                mensaje = mensaje & " Hora : " & Format(Time, "hh:mm:ss") & vbNewLine
                mensaje = mensaje & vbNewLine
                
                mensaje = mensaje & " Mensaje enviado automáticamente desde Geslab, el " & Format(Date, "dd-mm-yyyy") & " a las " & Format(Time, "hh:mm:ss")
                ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
        End If
        Set oParametro = Nothing
    End If
    'M1053-F
    ' Verificamos si se modifica el tipo de la muestra
    Dim TIPO_MUESTRA_ID As Long
    Dim ID_PARTICULAR As Long
    Dim TIPO_ANALISIS_ID As Integer
    Dim BANO_ID As Integer
    ' Tipo analisis
    If cmbTA.getTEXTO <> "" Then
        TIPO_ANALISIS_ID = cmbTA.getPK_SALIDA
    Else
        TIPO_ANALISIS_ID = 0
    End If
    ' Baño
    If cmbBano.getTEXTO = "" Then
        BANO_ID = 0
    Else
        BANO_ID = cmbBano.getPK_SALIDA
        If oMuestra.esControlEficacia(CLng(Text1(0))) Then
            TIPO_ANALISIS_ID = cmbTA.getPK_SALIDA
        Else
            Dim oBANO As New clsBanos
            oBANO.cargar_bano (BANO_ID)
            TIPO_ANALISIS_ID = oBANO.getID_SOLUCION
        End If
    End If
    If oMuestra.getTIPO_MUESTRA_ID <> cmbTM.getPK_SALIDA Or _
       oMuestra.getTIPO_ANALISIS_ID <> TIPO_ANALISIS_ID Or _
       oMuestra.getBANO_ID <> BANO_ID Then
        TIPO_MUESTRA_ID = cmbTM.getPK_SALIDA
        If oMuestra.getTIPO_MUESTRA_ID <> cmbTM.getPK_SALIDA Then
            oMuestra.CrearIdCodigoParticular (cmbTM.getPK_SALIDA)
            ID_PARTICULAR = oMuestra.getID_PARTICULAR
        Else
            ID_PARTICULAR = oMuestra.getID_PARTICULAR
        End If
        ' Modificamos la muestra
        consulta = "update muestras set " & _
                    " tipo_muestra_id = " & TIPO_MUESTRA_ID & "," & _
                    " id_particular = " & ID_PARTICULAR & "," & _
                    " tipo_analisis_id = " & TIPO_ANALISIS_ID & "," & _
                    " bano_id = " & BANO_ID & _
                    " where id_muestra = " & CLng(Text1(0))
        execute_bd consulta
        
'        If Not oMuestra.esControlEficacia(CLng(Text1(0))) Then
'            ' Borramos las determinaciones
'            Dim oDeter As New clsDeterminaciones
'            oDeter.Eliminar_Por_Muestra (CLng(Text1(0)))
'            ' Insertamos las determinaciones por defecto
'            oDeter.Insertar_determinaciones_por_defecto (CLng(Text1(0)))
'            ' Datos especificos
'            Dim oDE As New clsDatos_valores
'            oDE.Eliminar_datos_especificos_vacios CLng(Text1(0))
'            oDE.Insertar_datos_especificos_por_defecto CLng(Text1(0))
'        End If
    End If
    If frmFluidos.visible = True Then
        Dim oFR As New clsFluidos_recepcion
        With oFR
            .setMUESTRA_ID = gmuestra
            .setNORMATIVA_APLICABLE = txtFluidoNormativa
            .Insertar
        End With
    End If
'    imprimir_recepcion
'    If omuestra.getCERRADA = 1 And Not omuestra.esControlEficacia(CLng(Text1(0))) Then
'        omuestra.disminuir_edicion_impresa CLng(Text1(0))
'        imprimir CLng(Text1(0)), 1, False
'    End If
    Set oMuestra = Nothing
    consulta_muestra
    proteger_campos
    Me.MousePointer = 0
    MsgBox "Datos modificados correctamente.", vbInformation, App.Title
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub cmdPaquete_Click()
    Dim oMuestra As New clsMuestra
    If (oMuestra.CargaMuestra((Text1(0)))) Then
        If oMuestra.getPAQUETE_ID <> 0 Then
            frmEP_Paquete_Detalle.PK = oMuestra.getPAQUETE_ID
            frmEP_Paquete_Detalle.Show 1
        End If
    End If
    Set oMuestra = Nothing
End Sub

Private Sub cmdRecarga_Click()
    gmuestra = CLng(Text1(0))
    frmEads_Recarga.Show 1
'    gmuestra = 0
End Sub

Private Sub cmdRegenerarInforme_Click()
   On Error GoTo cmdRegenerarInforme_Click_Error

    If CInt(txtedicion) > 0 Then
        'M1054-I
        Dim oMuestra As New clsMuestra
        oMuestra.informar_situacion CLng(Text1(0))
        oMuestra.CargaMuestra CLng(Text1(0))
        If oMuestra.getINFORME_MANUAL = 0 Then
            imprimir CLng(Text1(0)), 2, False
        Else
            imprimir CLng(Text1(0)), 70, False
        End If
        'M1054-F
'        imprimir CLng(Text1(0)), 2, False
        MsgBox "Se ha enviado a reimprimir la edición. Espere unos instantes antes de consultarla.", vbInformation, App.Title
        frmDatosEspeciales.visible = False
    End If
    
   On Error GoTo 0
   Exit Sub

cmdRegenerarInforme_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdRegenerarInforme_Click of Formulario frmVerMuestra"
End Sub

Private Sub cmdSC_Click()
        gmuestra = CLng(Text1(0))
        Dim oPaquete As New clsSC_Paquetes
        Dim idSc As Long
        Dim salida As String
        Dim EDICION As Long
        salida = oPaquete.paqueteAsociadoMuestra(gmuestra)
        If salida <> "" Then
        
            Dim oMuestra As New clsMuestra
            oMuestra.CargaMuestra (gmuestra)
            
            Dim lista() As String
            lista = Split(salida, "/")
            idSc = lista(0)
            EDICION = lista(1)
            Select Case oMuestra.getANALISIS_MODIFICADO
                Case 2 ' Control de eficacia
                    If IsLoadForm("frmSC_Paquete_Detalle_CE") Then
                        MsgBox "La ventana de detalle de S.C. ya esta abierta. No esta permitido abrirla dos veces.", vbCritical, App.Title
                    Else
                        frmSC_Paquete_Detalle_CE.PK = idSc
                        frmSC_Paquete_Detalle_CE.EDICION = EDICION
                        frmSC_Paquete_Detalle_CE.Show 1
                    End If
                Case Else
                    If IsLoadForm("frmSC_Paquete_Detalle") Then
                        MsgBox "La ventana de detalle de S.C. ya esta abierta. No esta permitido abrirla dos veces.", vbCritical, App.Title
                    Else
                        frmSC_Paquete_Detalle.PK = idSc
                        frmSC_Paquete_Detalle.EDICION = EDICION
                        frmSC_Paquete_Detalle.Show 1
                    End If
            End Select
        End If
End Sub

Private Sub cmdVida_Click()
    frmVidaMuestra.PK = CLng(Text1(0))
    frmVidaMuestra.Show 1
End Sub

Private Sub cmdwww_Click()
    If MsgBox("¿Esta seguro de eliminar la revisión de la web?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim c As String
        c = "delete from web_muestras_revision where muestra_id = " & CLng(Text1(0))
        execute_bd c
        MsgBox "Datos modificados correctamente.", vbInformation + vbOKOnly, App.Title
        frmDatosEspeciales.visible = False
    End If
End Sub

Private Sub deter_DblClick()
    If deter.ListItems.Count > 0 Then
        If lblTipo <> "ENSAYO DE EFICACIA" And lblTipo <> "SELLANTE" Then
            frmTD_Detalle.PK = deter.ListItems(deter.selectedItem.Index).SubItems(3)
            frmTD_Detalle.Show 1
        End If
    End If
End Sub

Private Sub FechaEntrega_Change()
    If cmdok.visible = False Or USUARIO.getPER_PLAZO_ENTREGA_CAMBIO = 0 Then
        FechaEntrega = fentrega
    End If
End Sub
Private Sub fechaMuestreo_Change()
    If cmdok.visible = False Then
        fechaMuestreo = fmuestreo
    End If
End Sub
Private Sub fechaRecepcion_Change()
    If cmdok.visible = False Then
        fechaRecepcion = frecepcion
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.MousePointer = 0
    Select Case KeyCode
        Case 27
            cmdcancel_Click
        Case 116 ' F5 Datos especiales
            If USUARIO.getPER_DATOS_ESPECIALES Then
                frmDatosEspeciales.visible = Not frmDatosEspeciales.visible
            End If
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    cabecera
    permisos
    consulta_muestra
    tabPrincipal.Tab = ColsTAB.COL_GENERAL
'    Resizer1.VScrollPosition = 0
End Sub
Private Sub consulta_muestra()
    Dim oMuestra As New clsMuestra
    Dim CODIGO As String
    Dim pos As Integer
    With oMuestra
    If .CargaMuestra(gmuestra) Then
        Label1(5).Caption = "Consulta de la muestra : " & .getTITULO_MUESTRA
        Me.Caption = Label1(5).Caption
        
        Text1(0) = .getID_MUESTRA
        Text1(5) = .getID_GENERAL
        CODIGO = .CodigoParticular(.getID_MUESTRA)
        pos = InStr(1, CODIGO, "-", vbTextCompare)
        Text1(1) = Mid(.CodigoParticular(.getID_MUESTRA), 1, pos - 1)
        fechaRecepcion = Format(.getFECHA_RECEPCION, "dd/mm/yyyy")
        horaRecepcion = Format(.getFECHA_RECEPCION & " " & .getHORA_RECEPCION, "dd/mm/yyyy hh:mm:ss")
        Text1(3) = .getID_PARTICULAR
        ' Cliente
        cmbClientes.MostrarElemento .getCLIENTE_ID
        ' Pedidos
'        pedidos (.getCLIENTE_ID)
        cmbPedidos.BoundText = .getPEDIDO_ID
        cmbTM.MostrarElemento .getTIPO_MUESTRA_ID
        ' Tipo de Análisis
        cmbBano.limpiar
        If .getBANO_ID = 0 And oMuestra.getANALISIS_MODIFICADO <> 2 Then
            Label2(5).Caption = "Tipo de análisis"
        Else
            Label2(5).Caption = "Solución"
'            If omuestra.getANALISIS_MODIFICADO = 2 Then ' CE
                CARGAR_COMBO_BANOS .getCLIENTE_ID, ""
'            Else
'                CARGAR_COMBO_BANOS .getCLIENTE_ID, "TIPO_MUESTRA_ID = " & cmbTM.getPK_SALIDA
'            End If
            cmbBano.MostrarElemento .getBANO_ID
            cmdRecarga.visible = True
        End If
        If .getTIPO_ANALISIS_ID <> 0 Then
            cmbTA.MostrarElemento .getTIPO_ANALISIS_ID
        End If
        If .getANALISIS_DUPLICADO = 0 Then
            opDuplicado(1).Value = True
        Else
            opDuplicado(0).Value = True
        End If
        opUrgente(.getURGENTE).Value = True
        chkAjuste.Value = .getAJUSTE
        chkConsulta.Value = .getCONSULTA
        txtCONSULTA_OBSERVACIONES = .getCONSULTA_OBSERVACIONES
        chkPRECIO_FIJADO.Value = .getPRECIO_FIJADO
        If .getURGENTE = 1 Then
            opUrgente(1).BackColor = &HFF&
        End If
        Text1(4) = .getPRECINTO
        Text1(8) = .getREFERENCIA_CLIENTE
        Text1(9) = .getPRODUCTO
        chkOpcion(0).Value = .getOP_VUELO
        chkOpcion(1).Value = .getOP_INSITU
        chkOpcion(2).Value = .getOP_LABMOVIL
        chkOpcion(3).Value = .getOP_NORUTINARIA
        chkOpcion(5).Value = .getOP_REPETICION
        'M1105-I
        If chkOpcion(0).Value = Unchecked And chkOpcion(1).Value = Unchecked And chkOpcion(2).Value = Unchecked Then
            chkOpcion(4).Value = Checked
        End If
        'M1105-F
        frecepcion = fechaRecepcion.Value
        cmbUsuario.MostrarElemento .getEMPLEADO_ID
        cmbDatos(4).BoundText = .getFORMATO_ID
        cmbCentroMuestra.BoundText = .getCENTRO_ID
        cmbReplacement.MostrarElemento .getREPLACEMENT_ID
        cmbDatos(5).BoundText = .getENTIDAD_ENTREGA_ID
        Text1(13) = .getDETALLE_ENTREGA
        Text1(14) = .getOBSERVACIONES_ENTREGA
        If IsNull(.getFECHA_MUESTREO) = False And Trim(.getFECHA_MUESTREO) <> "" Then
            If Format(.getFECHA_MUESTREO, "yyyy-mm-dd") = "1900-01-01" Then
                chkFMSinEspecificar.Value = Checked
                fechaMuestreo.Enabled = False
            Else
                chkFMSinEspecificar.Value = Unchecked
                fechaMuestreo.Enabled = True
            End If
            fechaMuestreo = Format(.getFECHA_MUESTREO, "dd/mm/yyyy")
            fmuestreo = fechaMuestreo.Value
        Else
            fechaMuestreo = Format(Date, "dd/mm/yyyy")
            fmuestreo = fechaMuestreo.Value
        End If
        cmbDatos(6).BoundText = .getENTIDAD_MUESTREO_ID
        Text1(17) = .getDETALLE_MUESTREO
        Text1(18) = .getOBSERVACIONES_MUESTREO
        Text1(10) = .getRETRASO_OBSERVACION
        Text1(19) = Format(.getPRECIO, "currency")
        If .getANALISIS_MODIFICADO = 0 Then
            Text1(6) = Format(oMuestra.ImporteMuestraPorDeterminaciones(gmuestra, .getCLIENTE_ID), "currency")
        End If
        If IsNull(.getFECHA_PREV_FIN) = False And .getFECHA_PREV_FIN <> "" Then
            FechaEntrega = Format(.getFECHA_PREV_FIN, "dd/mm/yyyy")
        Else
            FechaEntrega = Format(Date, "dd/mm/yyyy")
        End If
        fentrega = FechaEntrega.Value
        Text1(21) = .getOBSERVACIONES
        cmdAbrirMuestra.visible = False
        Select Case .getCERRADA
        Case 1
            lblestado = "CERRADA"
            lblestado.BackColor = &HFF&
            If USUARIO.getPER_CIERRE = True Then
                cmdAbrirMuestra.visible = True
            End If
        Case 2
            lblestado = "PTE.CIERRE"
            lblestado.BackColor = &HC0C0FF
        Case 3
            lblestado = "C.SIN INFORME"
            lblestado.BackColor = &HC0C0FF
        Case Else
            lblestado = "ABIERTA"
            lblestado.BackColor = &HC000&
            cerrar_muestra
        End Select
        If .getANULADA <> 0 Then
            lblestado = "ANULADA"
            lblestado.BackColor = &HFFFF&
            cmdAnular.Enabled = False
            cmdModificar.Enabled = False
            cmdDeter.Enabled = False
        End If
        cmdFluido.visible = False
        'M1105-I
'        frmAIM.visible = False
        tabPrincipal.TabVisible(ColsTAB.COL_FLUIDOS) = False
        'M1105-F
        ' TIPO
        Dim oDeco As New clsDecodificadora
        frmPlasma.visible = False
        If .getANALISIS_MODIFICADO = 2 Then
            lblTipo = "ENSAYO DE EFICACIA"
        ElseIf .getANALISIS_MODIFICADO = 3 Then
            lblTipo = "SELLANTE"
        ElseIf .getANALISIS_MODIFICADO = 5 Then
            lblTipo = "PLASMA"
            frameDeterminaciones.visible = False
            frmPlasma.visible = True
            frmPlasma.Enabled = False
            llenar_combo cmbProcess, New clsPlasma_procesos, 0, frmPlasma_Procesos_Detalle, ""
            oDeco.cargar_mi_combo cmbCustomer, DECODIFICADORA.DECODIFICADORA_PLASMA_CLIENTES_INTERNOS
            Dim oPR As New clsPlasma_recepcion
            oPR.Carga gmuestra
            cmbProcess.MostrarElemento oPR.getPROCESO_ID
            cmbCustomer.MostrarElemento oPR.getCUSTOMER_ID
            txtDatos(0) = oPR.getSPECIMEN_ID
            txtDatos(2) = oPR.getPRODUCT_TYPE
            txtDatos(3) = oPR.getSN
            txtDatos(4) = oPR.getPRODUCT_SN
            txtDatos(5) = oPR.getMODULE_SN
            txtDatos(1) = oPR.getPN
        Else
            Dim otm As New clsTipos_muestra
            otm.CARGAR .getTIPO_MUESTRA_ID
            Select Case otm.getTIPO_ESPECIAL_ID
            Case tipo_especial.agua
                lblTipo = "AGUA"
            Case tipo_especial.BANO
                lblTipo = "BAÑO"
            Case tipo_especial.control_eficacia
                lblTipo = "ENSAYO DE EFICACIA"
            Case tipo_especial.CONTROLES_PROCESOS
                lblTipo = "CONTROL DE PROCESOS"
            Case tipo_especial.FLUIDO
                lblTipo = "FLUIDO"
                cmdFluido.visible = True
                frmFluidos.visible = True
                Dim oFR As New clsFluidos_recepcion
                oFR.Carga gmuestra
                txtFluidoNormativa = oFR.getNORMATIVA_APLICABLE
                'M1105-I
                oDeco.cargar_mi_combo cmbPrograma, DECODIFICADORA.FLUIDOS_PROGRAMAS
                oDeco.cargar_mi_combo cmbCentro, DECODIFICADORA.FLUIDOS_CENTROS
                oDeco.cargar_mi_combo cmbTipoEnsayo, DECODIFICADORA.FLUIDOS_TIPOS_ENSAYOS
                oDeco.cargar_mi_combo cmbSeccion, DECODIFICADORA.FLUIDOS_SECCIONES
                oDeco.cargar_mi_combo cmbEstacion, DECODIFICADORA.FLUIDOS_ESTACIONES
                Set oDeco = Nothing
                Dim oMuestraAIM As New clsMuestras_aim
                If oMuestraAIM.Carga(gmuestra) Then
                    cmbPrograma.MostrarElemento oMuestraAIM.getAIM_PROGRAMA_ID
                    cmbCentro.MostrarElemento oMuestraAIM.getAIM_CENTRO_ID
                    cmbTipoEnsayo.MostrarElemento oMuestraAIM.getAIM_TIPO_ENSAYO_ID
                    cmbSeccion.MostrarElemento oMuestraAIM.getAIM_SECCION_ID
                    cmbEstacion.MostrarElemento oMuestraAIM.getAIM_ESTACION_ID
                End If
'                frmAIM.visible = True
                tabPrincipal.TabVisible(ColsTAB.COL_FLUIDOS) = True
                'M1105-F
            Case Else
                lblTipo = "ENSAYO NORMALIZADO"
            End Select
        End If
        ' Numero de factura
        cmdfactura.Enabled = False
        If .getDOCUMENTO_PAGO = 2 Then
            Dim oDoc_pago_muestra As New clsDocs_pago_muestras
            Text1(7) = oDoc_pago_muestra.EstaEnLaFacturaNumeroID(.getID_MUESTRA)
            If Text1(7) <> 0 Then
                Dim oDoc_pago As New clsDocs_pago
                If oDoc_pago.CargarDocumento(Text1(7)) Then
                     Text1(2) = oDoc_pago.getNUMERO & "/" & Format(oDoc_pago.getFECHA_FACTURA, "yyyy")
                     cmdfactura.Enabled = True
                End If
            End If
        End If
        ' Datos especiales
        chkAnulada.Value = .getANULADA
        If .getANULADA = 1 Then
            chkAnulada.Enabled = True
        End If
        chkInformeManual.Value = .getINFORME_MANUAL
        If .getINFORME_MANUAL = 1 Then
            chkInformeManual.Enabled = True
        End If
        txtedicion = .getULT_EDICION_IMP
        If IsDate(.getFECHA_COMIENZO) Then
            fcomienzo = .getFECHA_COMIENZO
        Else
            fcomienzo = Date
        End If
        If IsDate(.getFECHA_FINALIZACION) Then
            fFinalizacion = .getFECHA_FINALIZACION
        Else
            fFinalizacion = Date
        End If
        If IsDate(.getFECHA_CIERRE) Then
            ffin = .getFECHA_CIERRE
            fechaCierre = .getFECHA_CIERRE
        Else
            ffin = Date
            fechaCierre = Date
        End If
        If .getCERRADA = 1 Then
            cmbUsuarioCierre.MostrarElemento .getCERRADA_USUARIO
        End If
        If .getRESPONSABLE_ID <> 0 Then
            cmbResponsable.MostrarElemento .getRESPONSABLE_ID
        End If
        If .getANALISIS_MODIFICADO = 2 Then
            cargar_ce
        ElseIf .getANALISIS_MODIFICADO = 3 Then
            cargar_sellante
        ElseIf .getANALISIS_MODIFICADO = tipo_especial.PLASMA Then
            
            
        Else
            cargar_determinaciones
        End If
        cargar_datos_especificos
        ' INDICADORES
        If .getFECHA_RECOGIDA <> "" Then
            chkFechaSolicitud.Value = Checked
            'M1105-I
            chkFechaSolicitudNA.Value = Unchecked
            'M1105-F
'            fechaSolicitud = Left(.getFECHA_RECOGIDA, 10)
'            horaSolicitud = Right(.getFECHA_RECOGIDA, 8)
            fechaSolicitud = .getFECHA_RECOGIDA
            horaSolicitud = .getFECHA_RECOGIDA
        Else
            chkFechaSolicitud.Value = Unchecked
            'M1105-I
            chkFechaSolicitudNA.Value = Checked
            'M1105-F
        End If
        chkIPA.Value = .getIPA
        If .getCERRADA = 0 Then
            chkFechaCierre.Value = Unchecked
        Else
            chkFechaCierre.Value = Checked
        End If
        If .getHORA_CIERRE <> "" Then
            hfin = .getFECHA_CIERRE & " " & .getHORA_CIERRE
        End If
        If .getENVIADO_CORREO = 0 Then
            chkFechaEnvio.Value = Unchecked
        Else
            chkFechaEnvio.Value = Checked
        End If
        If .getFECHA_ENVIO <> "" Then
            fEnvio = .getFECHA_ENVIO
            hEnvio = .getFECHA_ENVIO
        End If
        cmbMotivoRetraso.MostrarElemento .getMOTIVO_RETRASO_ID
        'M1009-I
        llenar_combo cmbOferta, New clsOfertas, 0, frmOferta_Nueva2, " AND O.CLIENTE_ID = " & .getCLIENTE_ID
        cmbOferta.MostrarElemento .getOFERTA_ID
        'M1009-F
        ' PAQUETE
        If .getPAQUETE_ID = 0 Then
            cmdPaquete.visible = False
        Else
            cmdPaquete.visible = True
            cmdPaquete.Caption = cmdPaquete.Caption & " " & .getPAQUETE_ID
        End If
        ' ENAC
        opENAC(.getENAC).Value = True
        chkNadcap.Value = .getNADCAP
        ' BOTON S.C.
        Dim oDeter As New clsDeterminaciones
        cmdSC.visible = oDeter.tieneSC(.getID_MUESTRA)
        Set oDeter = Nothing
        ' ADS
        Dim oCliente As New clsCliente
        oCliente.CargaCliente cmbClientes.getPK_SALIDA
        Dim ID_PLANTA As String
        ID_PLANTA = CStr(oCliente.getPLANT_ID)
        frmADS.Enabled = False
        If oCliente.getAIRBUS = 1 Then
            frmADS.Enabled = True
            If ID_PLANTA = "0" Then
                MsgBox "El cliente ADS no tiene informada la planta. Es necesario informarla en la ficha de cliente.", vbCritical, App.Title
                Exit Sub
            Else
                'Cargar combos
                oDeco.cargar_mi_combo_parametro cmbEnsayo, DECODIFICADORA.AIRBUS_TIPOS_ENSAYOS, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbProgramaADS, DECODIFICADORA.AIRBUS_PROGRAMAS, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbSection, DECODIFICADORA.AIRBUS_SECTION, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbFluid, DECODIFICADORA.AIRBUS_FLUID, ID_PLANTA
                oDeco.cargar_mi_combo_parametro cmbFacility, DECODIFICADORA.AIRBUS_FACILITY, ID_PLANTA
                Dim oADS As New clsMuestras_airbus
                If oADS.Carga(.getID_MUESTRA) Then
                    cmbEnsayo.MostrarElemento oADS.getENSAYO_ID
                    cmbProgramaADS.MostrarElemento oADS.getPROGRAMA_ID
                    cmbSection.MostrarElemento oADS.getSECTION_ID
                    cmbFluid.MostrarElemento oADS.getFLUID_ID
                    cmbFacility.MostrarElemento oADS.getFACILITY_ID
                End If
            End If
        End If
        
        cargar_normas
    End If
    End With
'    Label1(5).BackColor = &HC0FFFF
'    Label1(5).Caption = "  Consulta de Muestra : " & Text1(5) & " (" & Text1(1) & "-" & Text1(3) & ")"
    proteger_campos
    Set oMuestra = Nothing
End Sub

Private Sub imgVerPedido_Click()
    If cmbPedidos.Text = "" Then
        MsgBox "Seleccione algún pedido de la lista.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim oAdjunto As New clsAdjuntos
'    oAdjunto.CargarDocumento TOBJETO.TOBJETO_CLIENTES_PEDIDOS, cmbPedidos.BoundText, 0, 1, True
    oAdjunto.CargarDocumento TOBJETO.TOBJETO_CLIENTES_PEDIDOS, cmbPedidos.BoundText, 0, 0, True
    Set oAdjunto = Nothing
    
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title

End Sub

Private Sub Text1_GotFocus(Index As Integer)
    Text1(Index).BackColor = &H80C0FF
    Text1(Index).SelStart = 0
    Text1(Index).SelLength = Len(Text1(Index))
End Sub

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
    ' Escribir ',' al pulsar '.'
    If Index = 19 And KeyAscii = 46 Then
         KeyAscii = 44
    End If
End Sub

Private Sub permisos()
    If USUARIO.getPER_FACTURACION = False Then
        Text1(19).Locked = True
        frmFacturacion.visible = False
    End If
    If USUARIO.getPER_MODIFICACION = False Then
        cmdModificar.Enabled = False
    End If
    If USUARIO.getPER_DATOS_ESPECIALES Then
        lblf5.visible = True
    End If
'    cmdVida.Visible = False
End Sub
Private Sub proteger_campos()
    Dim color As Single
    color = &HC0FFFF
    'Titulo
'    Label1(5).BackColor = color
'    Label1(5).Caption = "  Consulta de Muestra : " & Text1(5) & " (" & Text1(1) & "-" & Text1(3) & ")"
'    Me.Caption = Label1(5).Caption
    'Registro
    Text1(8).Locked = True
    Text1(9).Locked = True
    'Recepcion
    cmbResponsable.desactivar
    cmbClientes.desactivar
    'M1009-I
    cmbOferta.desactivar
    'M1009-F
    chkFMSinEspecificar.Enabled = False
    chkIPA.Enabled = False
    'M1105-I
    frmAIM.Enabled = False
    frmADS.Enabled = False
    'M1105-F
    frmENAC.Enabled = False
    
    chkFechaSolicitud.Enabled = False
    
    cmbDatos(4).Locked = True
    cmbCentroMuestra.Locked = True
    cmbReplacement.desactivar
    cmbUsuario.desactivar
    cmbDatos(5).Locked = True
    Text1(4).Locked = True
    Text1(13).Locked = True
    Text1(14).Locked = True
    opDuplicado(0).Enabled = False
    opDuplicado(1).Enabled = False
    opUrgente(0).Enabled = False
    opUrgente(1).Enabled = False
    chkOpcion(0).Enabled = False
    chkOpcion(1).Enabled = False
    chkOpcion(2).Enabled = False
    chkOpcion(3).Enabled = False
    chkOpcion(5).Enabled = False
    chkAjuste.Enabled = False
    chkConsulta.Enabled = False
    txtCONSULTA_OBSERVACIONES.Locked = True
    
    chkPRECIO_FIJADO.Enabled = False
    ' Muestreo
    cmbDatos(6).Locked = True
'    cmbbano.desactivar
    Text1(17).Locked = True
    Text1(18).Locked = True
    Text1(10).Locked = True
    ' Otros datos
    Text1(19).Locked = True
    Text1(21).Locked = True
    txtFluidoNormativa.Locked = True
    ' Botones
    cmdAnular.Enabled = True
    cmdDeter.Enabled = True
    cmdInfRegistro.Enabled = True
    cmdContra.Enabled = True
    cmdok.visible = False
    cmdVida.Enabled = True
    cmdInforme.Enabled = True
    cmdespecificas.Enabled = True
    If USUARIO.getPER_MODIFICACION = True Then
        cmdModificar.Enabled = True
    End If
    cmdListadoDeter.Enabled = True
'    Form_Load
    cmbTM.desactivar
    cmbTA.desactivar
    cmbBano.desactivar
    cmbPedidos.Locked = True
    cmdListadoDeter.Enabled = True
    cmdespecificas.Enabled = True
End Sub

Private Sub Text1_LostFocus(Index As Integer)
    Text1(Index).BackColor = vbWhite
    ' Formatear el precio
    If Index = 19 Then
        If Text1(19) <> "" Then
            If IsNumeric(Text1(19)) = True Then
                Text1(19) = Format(Text1(19), "currency")
            Else
                MsgBox "El precio debe ser numérico", vbInformation, App.Title
                Text1(19).SetFocus
            End If
        End If
    End If
    If Index = 8 Then
        Text1(8) = Replace(Text1(8).Text, """", " ")
    End If
End Sub

Private Sub cerrar_muestra()
    ' Cerrar Muestra
    Dim oMuestra As New clsMuestra
    Dim cierre As Integer
    cierre = oMuestra.comprobar_cierre(CLng(Text1(0)))
    Select Case cierre
        Case 1 ' Cerrada
           lblestado = "CERRADA"
           lblestado.BackColor = &HFF&
        Case 2 ' Pte.
           lblestado = "PTE.CIERRE"
           lblestado.BackColor = &HC0C0FF
        Case Else
           lblestado = "ABIERTA"
           lblestado.BackColor = &HC000&
    End Select
    Set oMuestra = Nothing
End Sub

Private Sub cargar_combos()
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbTM, New clsTipos_muestra, 0, frmTM_Detalle, ""
'P001-I
    llenar_combo cmbTA, New clsTipos_analisis, 0, frmTA_Detalle, ""
'P001-F
    cargar_combo cmbDatos(4), New clsformatos
    cargar_combo cmbCentroMuestra, New clsCentros
    cargar_combo cmbDatos(5), New clsEntidades_Entrega
    cargar_combo cmbDatos(6), New clsEntidades_muestreo
    llenar_combo cmbUsuario, New clsUsuarios, 0, frmUsuarios, " OR ANULADO = 1 "
    llenar_combo cmbUsuarioCierre, New clsUsuarios, 0, frmUsuarios, " OR ANULADO = 1 "
    llenar_combo cmbResponsable, New clsUsuarios, 0, frmUsuarios, " OR ANULADO = 1 "
    ' Sólo normas en estudio
    llenar_combo cmbNormas, New clsCa_normas, 0, frmCA_Normas, " AND REVISION = 1"

    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbMotivoRetraso, DECODIFICADORA.AUDITORIA_MOTIVOS_RETRASOS
    oDeco.cargar_mi_combo cmbReplacement, DECODIFICADORA.IBERIA_REPLACEMENT

End Sub

Private Sub pedidos(ID As Integer)
    Dim oPedido As New clsClientes_pedidos
    Dim anterior As Integer
    If cmbPedidos.Text <> "" Then
        anterior = cmbPedidos.BoundText
    End If
    If ID = 0 Then
        Set cmbPedidos.RowSource = oPedido.Listado_completo
    Else
        Set cmbPedidos.RowSource = oPedido.Listado_en_fecha(ID, fechaRecepcion.Value)
    End If
    cmbPedidos.ListField = "CODIGO_LARGO"
    cmbPedidos.DataField = "id_pedido"
    cmbPedidos.BoundColumn = "id_pedido"
    cmbPedidos.BoundText = anterior
End Sub
Private Sub CARGAR_COMBO_BANOS(cliente As Long, filtro As String)
    llenar_combo cmbBano, New clsBanos, cliente, frmBANO_Detalle, filtro
End Sub
Private Sub cabecera()
    With datos.ColumnHeaders
        .Add , , "Dato", 2500, lvwColumnLeft
        .Add , , "Valor", 2050, lvwColumnCenter
        .Add , , "ID", 0, lvwColumnLeft
    End With
    With listaNormas.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Nombre", listaNormas.Width * 0.99, lvwColumnLeft
'        .Add , , "Ruta", 0, lvwColumnLeft
    End With
End Sub
Private Sub cargar_determinaciones()
    frameDeterminaciones.Caption = "Determinaciones"
    deter.Width = 6045
    cmdListadoDeter.visible = True
    deter.ColumnHeaders.Clear
    With deter.ColumnHeaders
        .Add , , "Pnt", 1000, lvwColumnLeft
        .Add , , "Nombre", 3000, lvwColumnLeft
        .Add , , "Resultado", 1200, lvwColumnRight
        .Add , , "ID_TIPO_DETERMINACION", 0, lvwColumnLeft
    End With

    Dim rs As ADODB.Recordset
    deter.ListItems.Clear
    Dim oDeter As New clsDeterminaciones
    Set rs = oDeter.lista_determinaciones_muestra(CLng(Text1(0)))
    While Not rs.EOF
       With deter.ListItems.Add(, , rs(0))
          .SubItems(1) = Trim(rs(1))
          If Not IsNull(rs(2)) Then
              .SubItems(2) = rs(2)
          End If
          .SubItems(3) = rs(3)
       End With
       rs.MoveNext
    Wend
    Set oDeter = Nothing
    Set rs = Nothing
End Sub
Private Sub cargar_sellante()
    frameDeterminaciones.Caption = "Ensayos"
    deter.Width = 7350
    cmdListadoDeter.visible = False
    deter.ColumnHeaders.Clear
    deter.ListItems.Clear
    
    With deter.ColumnHeaders
        .Add , , "Ensayo", 3000, lvwColumnLeft
        .Add , , "R.Inferior", 1100, lvwColumnCenter
        .Add , , "R.Superior", 1100, lvwColumnCenter
        .Add , , "Resultado", 1000, lvwColumnRight
        .Add , , "Unidad", 1000, lvwColumnLeft
    End With

    Dim rs As ADODB.Recordset
    Dim oSe_Resultados As New clsSellantes_resultados
    Set rs = oSe_Resultados.Listado_Resultados(CLng(Text1(0)))
    If rs.RecordCount > 0 Then
        Do
            With deter.ListItems.Add(, , rs(1))
              .SubItems(1) = rs(2)
              .SubItems(2) = rs(3)
              If Trim(rs(4)) = "" Then
                  .SubItems(3) = " "
              Else
                  .SubItems(3) = Trim(rs(4))
              End If
              .SubItems(4) = rs(5)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub
Private Sub cargar_ce()
    frameDeterminaciones.Caption = "Ensayos"
    deter.Width = 7350
    cmdListadoDeter.visible = False
    deter.ColumnHeaders.Clear
    deter.ListItems.Clear
    
    With deter.ColumnHeaders
        .Add , , "Id.Canagrosa", 2000, lvwColumnLeft
        .Add , , "Id.Cliente", 2000, lvwColumnLeft
        .Add , , "Fecha", 1000, lvwColumnCenter
        .Add , , "Resultado", 1000, lvwColumnRight
        .Add , , "Conforme", 1000, lvwColumnCenter
    End With

    Dim rs As ADODB.Recordset
    Dim oCe_resultados As New clsCe_resultados
    Set rs = oCe_resultados.Listado_por_muestra(CLng(Text1(0)))
    If rs.RecordCount > 0 Then
        Do
            With deter.ListItems.Add(, , rs("identificacion_canagrosa"))
              .SubItems(1) = rs("identificacion_cliente")
              .SubItems(2) = rs("fecha")
              .SubItems(3) = rs("resultado")
              If rs("fecha") <> "" Then
                If rs("conforme") = 1 Then
                  .SubItems(4) = "CONFORME"
                Else
                  .SubItems(4) = "NO CONFORME"
                End If
            End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub
Private Sub cargar_datos_especificos()
    Dim oDatos As New clsDatos_valores
    Dim rs As ADODB.Recordset
    datos.ListItems.Clear
    Set rs = oDatos.datos_muestra_completo(CLng(Text1(0)))
    If rs.RecordCount <> 0 Then
        Do
            With datos.ListItems.Add(, , rs(0))
                 .SubItems(1) = rs(1)
                 .SubItems(2) = rs(2)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oDatos = Nothing
End Sub
