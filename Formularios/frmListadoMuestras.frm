VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmListadoMuestras 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Muestras"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   16095
   Icon            =   "frmListadoMuestras.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   10365
   ScaleWidth      =   16095
   Begin VB.Frame frmCargando 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   4680
      TabIndex        =   76
      Top             =   4815
      Visible         =   0   'False
      Width           =   6450
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Generando documento EXCEL. Por favor, espere."
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
         Height          =   330
         Index           =   1
         Left            =   540
         TabIndex        =   77
         Top             =   270
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdAIM 
      BackColor       =   &H00E0E0E0&
      Caption         =   "A.D.S."
      Height          =   870
      Left            =   12285
      Picture         =   "frmListadoMuestras.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   73
      Top             =   9450
      Width           =   1095
   End
   Begin VB.CommandButton cmdEtiquetaSoluciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Soluciones"
      Height          =   870
      Left            =   5580
      Style           =   1  'Graphical
      TabIndex        =   72
      ToolTipText     =   "Muestra el informe de registrro de la muestra"
      Top             =   9450
      Width           =   1095
   End
   Begin VB.CommandButton cmdAgrupar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Agrupar"
      Height          =   870
      Left            =   11160
      Picture         =   "frmListadoMuestras.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   65
      Top             =   9450
      Width           =   1095
   End
   Begin VB.CommandButton cmdCalibracion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Calibración"
      Height          =   870
      Left            =   10035
      Picture         =   "frmListadoMuestras.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   62
      Top             =   9450
      Width           =   1095
   End
   Begin VB.CommandButton cmdmail 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E-mail"
      Height          =   870
      Left            =   8955
      Style           =   1  'Graphical
      TabIndex        =   61
      ToolTipText     =   "Enviar informe de la ultima edición generada por E-mail"
      Top             =   9450
      Width           =   1050
   End
   Begin VB.CommandButton cmdDevolverMaterial 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Devolver Material"
      Height          =   870
      Left            =   7830
      Picture         =   "frmListadoMuestras.frx":2448
      Style           =   1  'Graphical
      TabIndex        =   44
      ToolTipText     =   "Genera una copia de la muestra seleccionada"
      Top             =   9450
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdjuntar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntar"
      Height          =   870
      Left            =   6705
      Style           =   1  'Graphical
      TabIndex        =   43
      ToolTipText     =   "Genera una copia de la muestra seleccionada"
      Top             =   9450
      Width           =   1095
   End
   Begin VB.CommandButton cmdDuplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   13725
      Style           =   1  'Graphical
      TabIndex        =   42
      ToolTipText     =   "Genera una copia de la muestra seleccionada"
      Top             =   8910
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdInformeAgrupado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe Agrupado"
      Height          =   870
      Left            =   14490
      Picture         =   "frmListadoMuestras.frx":2D12
      Style           =   1  'Graphical
      TabIndex        =   41
      ToolTipText     =   "Previsualizar informe de ensayo"
      Top             =   9225
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta"
      Height          =   870
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   39
      ToolTipText     =   "Muestra el informe de registrro de la muestra"
      Top             =   9450
      Width           =   1095
   End
   Begin VB.CommandButton cmdInfRegistro 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Doc.Registro"
      Height          =   870
      Left            =   2268
      Picture         =   "frmListadoMuestras.frx":35DC
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Muestra el informe de registrro de la muestra"
      Top             =   9450
      Width           =   1095
   End
   Begin VB.CommandButton cmdDeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registro"
      Height          =   870
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9450
      Width           =   1095
   End
   Begin VB.CommandButton cmdInforme 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe"
      Height          =   870
      Left            =   1164
      Style           =   1  'Graphical
      TabIndex        =   30
      ToolTipText     =   "Previsualizar informe de ensayo"
      Top             =   9450
      Width           =   1095
   End
   Begin VB.CommandButton cmdVida 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vida "
      Height          =   870
      Left            =   3372
      Style           =   1  'Graphical
      TabIndex        =   29
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   9450
      Width           =   1095
   End
   Begin VB.CommandButton cmdListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   870
      Left            =   13380
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   9450
      Width           =   1365
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   14760
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   9450
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2970
      Left            =   45
      TabIndex        =   0
      Top             =   330
      Width           =   16020
      Begin VB.PictureBox Picture1 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   9000
         Picture         =   "frmListadoMuestras.frx":3EA6
         ScaleHeight     =   240
         ScaleWidth      =   240
         TabIndex        =   71
         Top             =   2610
         Width           =   240
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   10125
         TabIndex        =   67
         Top             =   2610
         Width           =   2670
         Begin VB.OptionButton opUrgente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "NO"
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
            Left            =   1755
            TabIndex        =   70
            Top             =   45
            Width           =   615
         End
         Begin VB.OptionButton opUrgente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "SI"
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
            Left            =   1125
            TabIndex        =   69
            Top             =   45
            Width           =   615
         End
         Begin VB.OptionButton opUrgente 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "TODAS"
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
            Index           =   2
            Left            =   45
            TabIndex        =   68
            Top             =   45
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CheckBox chkclianulados 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar Clientes Anulados"
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
         Height          =   255
         Left            =   10755
         TabIndex        =   64
         Top             =   225
         Width           =   2595
      End
      Begin VB.CheckBox chkNoEnviadas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sólo las No Enviadas"
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
         Height          =   255
         Left            =   9765
         TabIndex        =   63
         Top             =   1305
         Width           =   2145
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   630
         TabIndex        =   50
         Top             =   270
         Visible         =   0   'False
         Width           =   2115
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todos"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   11
            Left            =   135
            TabIndex        =   58
            Top             =   225
            Value           =   -1  'True
            Width           =   1725
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Controles de Procesos"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   10
            Left            =   135
            TabIndex        =   57
            Top             =   1800
            Width           =   1950
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Controles Eficacia"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   4
            Left            =   135
            TabIndex        =   56
            Top             =   450
            Width           =   1725
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Sellantes"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   5
            Left            =   135
            TabIndex        =   55
            Top             =   675
            Width           =   1725
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Combustibles"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   6
            Left            =   135
            TabIndex        =   54
            Top             =   900
            Width           =   1725
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fluido"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   53
            Top             =   1125
            Width           =   1725
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Agroalimentario"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   8
            Left            =   135
            TabIndex        =   52
            Top             =   1350
            Width           =   1725
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ensayos Iberia"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   9
            Left            =   135
            TabIndex        =   51
            Top             =   1575
            Width           =   1725
         End
      End
      Begin VB.CheckBox chkTodosTA 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   9750
         TabIndex        =   45
         Top             =   945
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.TextBox txtreferencia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   6210
         TabIndex        =   9
         Top             =   2205
         Width           =   2235
      End
      Begin VB.CheckBox chktmanuladas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar T.Muestras Anuladas"
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
         Height          =   255
         Left            =   10755
         TabIndex        =   35
         Top             =   585
         Width           =   2910
      End
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Left            =   5715
         TabIndex        =   32
         Top             =   1575
         Width           =   705
      End
      Begin VB.TextBox txtg2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4125
         TabIndex        =   23
         Top             =   1770
         Width           =   975
      End
      Begin VB.TextBox txtg1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2835
         TabIndex        =   22
         Top             =   1770
         Width           =   975
      End
      Begin VB.TextBox txtp2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   4125
         TabIndex        =   19
         Top             =   1395
         Width           =   975
      End
      Begin VB.TextBox txtp1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2835
         TabIndex        =   18
         Top             =   1395
         Width           =   975
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
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
         Height          =   255
         Left            =   9750
         TabIndex        =   17
         Top             =   585
         Value           =   1  'Checked
         Width           =   1065
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Validez"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   13815
         TabIndex        =   2
         Top             =   135
         Width           =   2115
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Pendientes (Abiertas)"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   27
            Top             =   765
            Width           =   1860
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Muestras validas"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   240
            Value           =   -1  'True
            Width           =   1560
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Muestras anuladas"
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   495
            Width           =   1680
         End
         Begin VB.OptionButton opTipo 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Todas"
            ForeColor       =   &H80000008&
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   3
            Top             =   1035
            Width           =   1005
         End
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
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
         Height          =   255
         Left            =   9750
         TabIndex        =   1
         Top             =   225
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1890
         TabIndex        =   7
         Top             =   2160
         Width           =   1320
         _ExtentX        =   2328
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
         Format          =   52166657
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3825
         TabIndex        =   8
         Top             =   2160
         Width           =   1275
         _ExtentX        =   2249
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
         Format          =   52166657
         CurrentDate     =   38002
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   330
         Left            =   6405
         TabIndex        =   33
         Top             =   1575
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196638
         OrigLeft        =   1590
         OrigTop         =   6570
         OrigRight       =   1830
         OrigBottom      =   6975
         Max             =   2015
         Min             =   2004
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1530
         TabIndex        =   36
         Top             =   225
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbTiposMuestra 
         Height          =   330
         Left            =   1530
         TabIndex        =   37
         Top             =   585
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbTipoAnalisis 
         Height          =   330
         Left            =   1530
         TabIndex        =   46
         Top             =   945
         Width           =   7935
         _ExtentX        =   13996
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmListadoMuestras.frx":A6F8
         Height          =   315
         Left            =   6210
         TabIndex        =   48
         Top             =   2565
         Width           =   2235
         _ExtentX        =   3942
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
      Begin pryCombo.miCombo cmbReplacement 
         Height          =   345
         Left            =   1485
         TabIndex        =   59
         Top             =   2565
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbTipo 
         Height          =   345
         Left            =   10080
         TabIndex        =   74
         Top             =   2205
         Width           =   4320
         _ExtentX        =   7620
         _ExtentY        =   609
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "&Buscar"
         Default         =   -1  'True
         Height          =   1050
         Left            =   14670
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1530
         Width           =   1275
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo Ensayo"
         Height          =   195
         Index           =   8
         Left            =   9000
         TabIndex        =   75
         Top             =   2250
         Width           =   885
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "URGENTE"
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   40
         Left            =   9315
         TabIndex        =   66
         Top             =   2655
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Replacement"
         Height          =   195
         Index           =   19
         Left            =   180
         TabIndex        =   60
         Top             =   2640
         Width           =   945
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   5310
         TabIndex        =   49
         Top             =   2655
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Análisis"
         Height          =   195
         Index           =   10
         Left            =   180
         TabIndex        =   47
         Top             =   1035
         Width           =   1110
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Referencia"
         Height          =   195
         Index           =   9
         Left            =   5310
         TabIndex        =   40
         Top             =   2250
         Width           =   780
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         Height          =   315
         Index           =   0
         Left            =   5265
         TabIndex        =   34
         Top             =   1635
         Width           =   585
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
         Height          =   195
         Index           =   6
         Left            =   3885
         TabIndex        =   24
         Top             =   1830
         Width           =   135
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Por Nº de Ensayo General, desde"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   21
         Top             =   1845
         Width           =   2415
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "al"
         Height          =   195
         Index           =   7
         Left            =   3885
         TabIndex        =   20
         Top             =   1455
         Width           =   135
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Por Nº de Ensayo Particular, desde"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   16
         Top             =   1470
         Width           =   2505
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   15
         Top             =   675
         Width           =   1185
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   3330
         TabIndex        =   12
         Top             =   2205
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepcionadas desde"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   11
         Top             =   2205
         Width           =   1635
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   315
         Width           =   495
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5850
      Left            =   45
      TabIndex        =   25
      Top             =   3555
      Width           =   16035
      _ExtentX        =   28284
      _ExtentY        =   10319
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11250
      Top             =   9495
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":A73E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":B018
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":B8F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":C1CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":CAA6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":D380
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":13BE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":14079
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":1450F
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":149A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":1B208
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":21A6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":282CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":2EB2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmListadoMuestras.frx":35390
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Muestras"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   4
      Left            =   45
      TabIndex        =   14
      Top             =   0
      Width           =   16005
   End
   Begin VB.Label lblmsg 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Seleccione Criterio."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   13
      Top             =   3285
      Width           =   16035
   End
End
Attribute VB_Name = "frmListadoMuestras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const PK_ID_MUESTRA = 6
Private mvarCRITERIO_LISTADO As String
      
Private Sub chkclianulados_Click()
    cmbClientes.limpiar
    cargar_clientes
End Sub

Private Sub chkTodosTA_Click()
    If chkTodosTA.Value = Checked Then
        cmbTipoAnalisis.limpiar
        cmbTipoAnalisis.desactivar
    Else
        cmbTipoAnalisis.activar
    End If
End Sub

Private Sub cmbTiposMuestra_change()
    cargar_Ta
End Sub

Private Sub cmdAdjuntar_Click()
   On Error GoTo cmdAdjuntar_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    Dim m As String
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Selected = True Then
            m = m & lista.ListItems(i).SubItems(PK_ID_MUESTRA) & ";"
        End If
    Next
    If m = "" Then
        MsgBox "Seleccione alguna muestra.", vbExclamation, App.Title
    Else
        With frmAdjuntos
            .TOBJETO = TOBJETO.TOBJETO_MUESTRAS
            .COBJETO = 0
            .COBJETO_GRUPO_MUESTRAS = m
            .Show 1
        End With
        Set frmAdjuntos = Nothing
    End If

   On Error GoTo 0
   Exit Sub

cmdAdjuntar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntar_Click of Formulario frmListadoMuestras"
End Sub

Private Sub cmdAgrupar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
'    Dim objLitem As ListItem, objSitem As ListSubItem
'    Set objLitem = lista.ListItems(lista.selectedItem.Index)
'    Set objSitem = objLitem.ListSubItems(10)
    If esAgrupada Then
        MsgBox "La muestra ya se encuentra agrupada.", vbCritical, App.Title
    Else
        Set frmMuestras_Agrupar.listaM = lista
        frmMuestras_Agrupar.PK = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
        frmMuestras_Agrupar.Show 1
    End If
End Sub
Private Function esAgrupada() As Boolean
    If lista.ListItems.Count = 0 Then Exit Function
    Dim objLitem As ListItem, objSitem As ListSubItem
    Set objLitem = lista.ListItems(lista.selectedItem.Index)
    Set objSitem = objLitem.ListSubItems(10)
    If objSitem.ReportIcon = 11 Then
        esAgrupada = True
    Else
        esAgrupada = False
    End If
End Function
Private Sub cmdAIM_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim i As Integer
    Dim muestras As String
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Selected = True Then
            If muestras <> "" Then
                muestras = muestras & ","
            End If
            muestras = muestras & lista.ListItems(i).SubItems(PK_ID_MUESTRA)
        End If
    Next
    If muestras <> "" Then
        frmAirbus_ListadoMuestras.ID_MUESTRAS = muestras
        frmAirbus_ListadoMuestras.Show 1
    End If
End Sub

Private Sub cmdCalibracion_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim oC As New clsEquipoCalibracion
    Dim oCal As Long
    oCal = oC.CargaPorMuestraId(lista.ListItems(lista.selectedItem.Index).SubItems(6))
    If oCal = 0 Then
        MsgBox "La muestra no esta asociada a ninguna calibración.", vbCritical, App.Title
        Exit Sub
    Else
        oC.Carga oCal
        Dim oEquipo As New clsEquipos
        oEquipo.Carga oC.getEQUIPO_ID
        If oEquipo.getTIPO_EQUIPO_ID = EQ_TIPOS_EQUIPOS.TIPO_EQUIPO_TORCOMETRO Then
            MsgBox "Los TORCOMETROS deben gestionarse en GESMET.", vbExclamation, App.Title
            Exit Sub
        End If
        Dim objfrm  As New frmEquipoEdicionCalibracion
        Dim lngFila As Long, strId As String
        Dim intEstado As Integer
        strId = oC.getID_CALIBRACION
        intEstado = oC.getESTADO
        With objfrm
            Set .EQUIPO = oEquipo
            .ID = strId
            If intEstado = 0 Or intEstado = 3 Then
                .TipoEdicion = EDICION ' si no está cerrado
            Else
                .TipoEdicion = visualizar
            End If
                    
            .Show vbModal
        End With
        Unload objfrm
        Set objfrm = Nothing
    End If
End Sub

Private Sub cmdDevolverMaterial_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim cantiadad As Integer
    CANTIDAD = contarSeleccionadas
    If CANTIDAD >= 1 Then
        If MsgBox("¿Esta seguro de generar un envio de devolucion de material para las " & CANTIDAD & " muestras marcadas?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
        
        Dim texto As String
        texto = "Se devuelven a petición del cliente las piezas ensayadas correspondientes a los siguientes números de informes: "
        texto = texto & vbNewLine & vbNewLine
        Dim i As Integer
        Dim cliente As Long
        Dim oMuestra As New clsMuestra
        cliente = 0
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Selected = True Then
                If cliente = 0 Then
                    oMuestra.CargaMuestra lista.ListItems(i).SubItems(6)
                    cliente = oMuestra.getCLIENTE_ID
                End If
                texto = texto & vbNewLine & " - " & lista.ListItems(i).SubItems(5) & "/" & Year(lista.ListItems(i).SubItems(4)) & vbNewLine
            End If
        Next
        texto = texto & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine
        texto = texto & " Conforme al cliente:                            Conforme a Canagrosa: "
        texto = texto & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine & vbNewLine
        texto = texto & " Fecha. "
        
        Dim oPaquete As New clsEP_Paquetes
        Dim PAQUETE As Long
        With oPaquete
            .setASUNTO = "DEVOLUCIÓN DE MATERIAL ENSAYADO"
            .setDETALLE = texto
            .setTIPO = 0
            .setCLIENTE_ID = cliente
            .setMENSAJERIA_ID = 0
            .setFECHA_CREACION = Format(Date, "yyyy-mm-dd")
            .setHORA_CREACION = Format(Time, "hh:nn:ss")
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            PAQUETE = .Insertar
        End With
        Set oPaquete = Nothing
        
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Selected = True Then
                oMuestra.actualizarPaqueteId lista.ListItems(i).SubItems(6), PAQUETE
            End If
        Next
        Set oMuestra = Nothing
        frmEP_Paquete_Detalle.PK = PAQUETE
        frmEP_Paquete_Detalle.Show 1
    End If
End Sub

'M0926-I
Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
        If contarSeleccionadas <> 1 Then
            MsgBox "Seleccione una muestra para Duplicarla.", vbExclamation, App.Title
            Exit Sub
        End If
    
    If MsgBox("¿Esta seguro de duplicar la muestra seleccionada?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        Dim duplicamuestra As New clsMuestra_Duplicar
        idMuestra = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
        duplicamuestra.duplica_muestra (idMuestra)
        Set duplicamuestra = Nothing
        Call buscar
    End If
End Sub
'M0926-F

Private Sub cmdetiqueta_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    ReDim ETIQUETAS(contarSeleccionadas)
'    ReDim etiquetas(1)
'    etiquetas(1) = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
    Dim i As Integer
    Dim c As Integer
    c = 1
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Selected = True Then
            ETIQUETAS(c) = lista.ListItems(i).SubItems(PK_ID_MUESTRA)
            c = c + 1
        End If
    Next
    frmEtiquetas.Show 1
End Sub

Private Sub cmdEtiquetaSoluciones_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    frmSoluciones_Etiqueta.PK = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
    frmSoluciones_Etiqueta.Show 1
End Sub

Private Sub cmdInformeAgrupado_Click()
    Dim i As Integer
    Dim salida As String
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Selected = True Then
            If salida <> "" Then
                salida = salida & ","
            End If
            salida = salida & lista.ListItems(i).SubItems(PK_ID_MUESTRA)
        End If
    Next
'    MsgBox salida

    Dim objfrm As New frmReport
    Dim strCad As String
    
    With objfrm
        .iniciar
        .informe = "Informes\rptCO_Grupo"
        
        strCad = "{muestras.ID_MUESTRA} in [" & salida & "]"
        
        .criterio = strCad
        
        .imprimir = False
        .generar
        .visible = True
        
    End With
    Set objfrm = Nothing
            
End Sub

Private Sub cmdInfRegistro_Click()
    If lista.ListItems.Count = 0 Then Exit Sub
    Dim i As Integer
'    Dim muestras As String
    
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Selected = True Then
'            If muestras <> "" Then
'                muestras = muestras & ","
'            End If
'            muestras = muestras & lista.ListItems(i).SubItems(PK_ID_MUESTRA)
            Dim oMuestra As New clsMuestra
            oMuestra.Informe_Recepcion lista.ListItems(i).SubItems(PK_ID_MUESTRA), False
            Set oMuestra = Nothing
        End If
    Next
'    If muestras <> "" Then
'        Dim oMuestra As New clsMuestra
'        oMuestra.Informe_Recepcion muestras, False
'        Set oMuestra = Nothing
'    End If

'    If contarSeleccionadas <> 1 Then
'        MsgBox "Seleccione una muestra para ver el Informe de Registro.", vbExclamation, App.Title
'        Exit Sub
'    End If
'    Dim oMuestra As New clsMuestra
'    oMuestra.Informe_Recepcion lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA), False
'    Set oMuestra = Nothing
End Sub

Private Sub cmdListado_Click()
    If lista.ListItems.Count = 0 Then
        MsgBox "No existen registros para generar el listado.", vbExclamation, App.Title
        Exit Sub
    End If
    If MsgBox("¿Desea exportar a excel?", vbYesNo + vbQuestion, App.Title) = vbNo Then
        Dim objMuestras As clsMuestra
        Set objMuestras = New clsMuestra
        objMuestras.imprimir_listadomuestras mvarCRITERIO_LISTADO, Format(fdesde.Value, "dd/mm/yyyy"), Format(fhasta.Value, "dd/mm/yyyy"), chkTodos.Value = vbChecked, chkTodas.Value = vbChecked
        Set objMuestras = Nothing
    Else
        informeExcel
    End If
End Sub
Private Sub informeExcel()
   On Error GoTo informeExcel_Error
    Dim filtro As String
    filtro = filtroListado
    If filtro = "" Then Exit Sub
    ' Listado
    frmCargando.visible = True
    Me.MousePointer = 11
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "SELECT mu.id_general,concat(tm.codigo,'-',cast(mu.id_particular as char)),tm.nombre,ta.nombre, " & _
               "       cl.nombre,mu.referencia_cliente,concat(mu.fecha_recepcion,' ',mu.hora_recepcion),concat(mu.fecha_cierre,' ',mu.hora_cierre),mu.cerrada " & _
               " FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "centros as ce, " & _
                     "muestras as mu " & strSECTOR & _
               " WHERE mu.cliente_id=cl.id_cliente AND mu.tipo_muestra_id=tm.id_tipo_muestra AND mu.tipo_analisis_id=ta.id_tipo_analisis " & _
                      " and mu.centro_id = ce.id_centro " & _
                      filtro & _
                      " order by mu.id_general desc"
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLA.visible = False
        XLS.Range("1:1").HorizontalAlignment = xlCenter
        XLS.Range("1:1").VerticalAlignment = xlCenter
        'Cabecera
        XLS.Cells(1, 1) = "N.General"
        XLS.Cells(1, 2) = "Codigo"
        XLS.Cells(1, 3) = "Tipo Muestra"
        XLS.Cells(1, 4) = "Tipo Analisis"
        XLS.Cells(1, 5) = "Cliente"
        XLS.Cells(1, 6) = "Referencia Cliente"
        XLS.Cells(1, 7) = "Fecha Recepción"
        XLS.Cells(1, 8) = "Fecha Cierre"
        XLS.Cells(1, 9) = "Horas"
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 2)).ColumnWidth = 14
        XLS.Range(XLS.Cells(1, 3), XLS.Cells(1, 6)).ColumnWidth = 35
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 9)).Interior.ColorIndex = 6
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 9)).Interior.Pattern = xlSolid
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 9)).Font.ColorIndex = 3
        XLS.Range(XLS.Cells(1, 1), XLS.Cells(1, 9)).AutoFilter
        i = 2
        Do
            XLS.Cells(i, 1) = rs(0)
            XLS.Cells(i, 2) = rs(1)
            XLS.Cells(i, 3) = rs(2)
            XLS.Cells(i, 4) = rs(3)
            XLS.Cells(i, 5) = rs(4)
            XLS.Cells(i, 6) = rs(5)
            XLS.Cells(i, 7) = rs(6)
            If (rs(8) <> 0 And Trim(rs(7)) <> "0000-00-00") Then
                XLS.Cells(i, 8) = rs(7)
                XLS.Cells(i, 9) = "=(H" & i & "-G" & i & ") * 24"
            End If
            XLS.Range(XLS.Cells(i + 1, 9), XLS.Cells(i + 1, 9)).HorizontalAlignment = xlCenter
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
        XLA.visible = True
        Me.MousePointer = 0
        frmCargando.visible = False
    End If
   On Error GoTo 0
   Exit Sub

informeExcel_Error:

    frmCargando.visible = False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure informeExcel of Formulario frmListadoMuestras"
    
End Sub
Private Sub chktmanuladas_Click()
'    cmbMuestras.Text = ""
    cmbTiposMuestra.limpiar
    cargar_muestras
End Sub

Private Sub chkTodas_Click()
    If chkTodas.Value = Checked Then
        cmbTiposMuestra.limpiar
        cmbTiposMuestra.desactivar
    Else
        cmbTiposMuestra.activar
    End If
    cargar_Ta
End Sub

Private Sub chkTodos_Click()
    If chkTodos.Value = Checked Then
        cmbClientes.limpiar
        cmbClientes.desactivar
    Else
        cmbClientes.activar
    End If
End Sub


Private Function contarSeleccionadas() As Integer
    Dim i As Integer
    sel = 0
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Selected = True Then
            sel = sel + 1
        End If
    Next
    contarSeleccionadas = sel
End Function
Private Sub cmdDeter_Click()
   On Error GoTo cmdDeter_Click_Error

    If lista.ListItems.Count > 0 Then
        If contarSeleccionadas <> 1 Then
            MsgBox "Seleccione una muestra para ver el Registro.", vbExclamation, App.Title
            Exit Sub
        End If
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
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
'                ElseIf oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.DUREZA_SHORE_PIEZAS Then
'                    With frmPlasma_Dureza_Shore
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
        actualizar_lista
        gmuestra = 0
    End If

   On Error GoTo 0
   Exit Sub

cmdDeter_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdDeter_Click of Formulario frmListadoMuestras"
End Sub

Private Sub cmdInforme_Click()
    If lista.ListItems.Count > 0 Then
        If contarSeleccionadas <> 1 Then
            MsgBox "Seleccione una muestra para ver el Informe.", vbExclamation, App.Title
            Exit Sub
        End If
        If esAgrupada Then
            MsgBox "La muestra esta agrupada. Consulte la muestra Origen.", vbExclamation, App.Title
            Exit Sub
        Else
            MostrarInforme CLng(lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA))
            actualizar_lista
        End If
    End If
End Sub

Private Sub cmdmail_Click()
   On Error GoTo cmdmail_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    
    Dim i As Integer
    Dim muestras As String
    Dim cliente As Long
    Dim oMuestra As New clsMuestra
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Selected = True Then
            oMuestra.CargaMuestra CLng(lista.ListItems(i).SubItems(PK_ID_MUESTRA))
            If oMuestra.getAGRUPADA_MUESTRA_ID <> 0 Then
                MsgBox "Hay seleccionadas muestras agrupadas. Solo se deben marcar las muestras Origen.", vbExclamation, App.Title
                Exit Sub
            End If
            If oMuestra.getREVISION_USUARIO = 0 Then
                MsgBox "No se pueden enviar los informes, existen muestras sin estar REVISADAS", vbCritical, App.Title
                Exit Sub
            End If
            If oMuestra.getCLIENTE_ID <> cliente And cliente <> 0 Then
                MsgBox "No se pueden enviar los informes, están marcadas muestras de DISTINTOS CLIENTES", vbCritical, App.Title
                Exit Sub
            End If
            cliente = oMuestra.getCLIENTE_ID
            muestras = muestras & lista.ListItems(i).SubItems(PK_ID_MUESTRA) & ";"
        End If
    Next
    Me.MousePointer = 11
    enviar_informeAgrupado muestras, Me.Hwnd
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

cmdmail_Click_Error:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdmail_Click of Formulario frmListadoMuestras"

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdVida_Click()
    If lista.ListItems.Count > 0 Then
        If contarSeleccionadas <> 1 Then
            MsgBox "Seleccione una muestra para ver la Vida.", vbExclamation, App.Title
            Exit Sub
        End If
        frmVidaMuestra.PK = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
        frmVidaMuestra.Show 1
        gmuestra = 0
    End If
End Sub



Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 20
    Me.top = 20
    cabecera
    permisos
    cargar_combos
    cmbClientes.desactivar
    cmbTiposMuestra.desactivar
    cmbTipoAnalisis.desactivar
    fdesde = Date
    fhasta = Date
    txtanno = Year(Date)
    cambiar.Max = Year(Date)
    mvarCRITERIO_LISTADO = ""
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Código", 1100, lvwColumnLeft
        .Add , , "Cliente", 2500, lvwColumnLeft
        .Add , , "Tipo de Analisis/Solución", 3800, lvwColumnLeft
        .Add , , "Ref.Cliente", 4000, lvwColumnLeft
        .Add , , "Fecha", 1300, lvwColumnCenter
        .Add , , "General", 800, lvwColumnCenter
        .Add , , "ID_MUESTRA", 1, lvwColumnCenter
        .Add , , "Facturada", 1, lvwColumnCenter
        .Add , , "Centro", 1475, lvwColumnCenter
        .Add , , "", 250, lvwColumnCenter
        .Add , , "", 250, lvwColumnCenter
        .Add , , "", 250, lvwColumnCenter ' URGENTE
    End With
End Sub
Private Sub cargar_combos()
    cargar_combo cmbCentro, New clsCentros
    cargar_muestras
    cargar_clientes
    cargar_Ta
    llenar_combo cmbTipo, New clsTipos_especial, 0, Me, ""
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbReplacement, DECODIFICADORA.IBERIA_REPLACEMENT
End Sub
Private Sub cargar_clientes()
    If chkclianulados.Value = Checked Then
        llenar_combo cmbClientes, New clsCliente, 0, frmClientes, " 1 = 1 "
    Else
        llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    End If
End Sub
Private Sub cargar_muestras()
    If chktmanuladas.Value = Unchecked Then
        llenar_combo cmbTiposMuestra, New clsTipos_muestra, 0, frmTM_Detalle, "ANULADO = 0"
    Else
        llenar_combo cmbTiposMuestra, New clsTipos_muestra, 0, frmTM_Detalle, ""
    End If
End Sub
Private Sub cargar_Ta()
    cmbTipoAnalisis.limpiar
    llenar_combo cmbTipoAnalisis, New clsTipos_analisis, IIf(chkTodas.Value = Unchecked And cmbTiposMuestra.getTEXTO <> "", cmbTiposMuestra.getPK_SALIDA, 0), frmTA_Detalle, IIf(chktmanuladas.Value = Unchecked, "ANULADO = 0", "")
End Sub
Private Sub cmdBuscar_Click()
    Call buscar
End Sub
Private Function filtroListado() As String
    Dim filtro As String
   On Error GoTo filtroListado_Error

    If txtReferencia <> "" Then
        filtro = filtro & " AND mu.referencia_cliente like '%" & txtReferencia & "%'"
    End If
    ' Tipo de muestra
    If chkTodas.Value = Unchecked Then
        If cmbTiposMuestra.getTEXTO = "" Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            filtroListado = ""
            Exit Function
        End If
        filtro = filtro & " AND mu.tipo_muestra_id=" & cmbTiposMuestra.getPK_SALIDA
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.TIPO_MUESTRA_ID} = " & cmbTiposMuestra.getPK_SALIDA
    Else
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.TIPO_MUESTRA_ID} > 0"
    End If
    ' Tipo de analisis
    If chkTodosTA.Value = Unchecked Then
        If cmbTipoAnalisis.getTEXTO = "" Then
            MsgBox "Debe seleccionar un tipo de análisis.", vbExclamation, App.Title
            filtroListado = ""
            Exit Function
        End If
        filtro = filtro & " AND mu.tipo_analisis_id=" & cmbTipoAnalisis.getPK_SALIDA
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.TIPO_ANALISIS_ID} = " & cmbTipoAnalisis.getPK_SALIDA
    Else
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.TIPO_ANALISIS_ID} > 0"
    End If
    ' Clientes
    If chkTodos.Value = Unchecked Then
        If cmbClientes.getPK_SALIDA = 0 Then
            MsgBox "Debe seleccionar un cliente.", vbExclamation, App.Title
            filtroListado = ""
            Exit Function
        End If
        filtro = filtro & " AND mu.cliente_id = " & cmbClientes.getPK_SALIDA
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.CLIENTE_ID} = " & cmbClientes.getPK_SALIDA
    Else
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.CLIENTE_ID} > 0"
    End If
    ' Validez
    If opTipo(0).Value = True Then
        filtro = filtro & " AND (mu.anulada is Null or mu.anulada = 0)"
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND (ISNULL({muestras.ANULADA}) OR {muestras.ANULADA}=0) "
    ElseIf opTipo(1).Value = True Then
        filtro = filtro & " AND mu.anulada = 1"
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.ANULADA} =1"
    ElseIf opTipo(3).Value = True Then
        filtro = filtro & " AND (mu.cerrada is Null or mu.cerrada = 0) AND (mu.anulada is Null or mu.anulada = 0)"
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND (ISNULL({muestras.CERRADA}) OR {muestras.CERRADA}=0) "
    End If
    ' Tipo
    If cmbTipo.getTEXTO <> "" Then
        filtro = filtro & " AND tm.TIPO_ESPECIAL_ID = " & cmbTipo.getPK_SALIDA
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {tipos_muestra.tipo_especial_id} = " & cmbTipo.getPK_SALIDA
    End If
    ' Fechas
    Dim f_desde As String
    Dim f_hasta As String
    f_desde = Format(fdesde, "yyyy-mm-dd")
    f_hasta = Format(fhasta, "yyyy-mm-dd")
    ' Particular
    If txtp1 <> "" Or txtp2 <> "" Then
        If txtp1 = "" Or txtp2 = "" Then
            MsgBox "Debe completar los codigos de búsqueda.", vbInformation, App.Title
            filtroListado = ""
            Exit Function
        Else
            If IsNumeric(txtp1) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp1.SetFocus
                filtroListado = ""
                Exit Function
            End If
            If IsNumeric(txtp2) = False Then
                MsgBox "El codigo debe ser numérico", vbInformation, App.Title
                txtp2.SetFocus
                filtroListado = ""
                Exit Function
            End If
            f_desde = ""
            f_hasta = ""
            filtro = filtro & " AND mu.id_particular between " & CLng(txtp1) & " and " & CLng(txtp2) & " and mu.anno = " & CInt(txtanno)
        End If
    End If
    ' General
    If txtg1 <> "" Or txtg2 <> "" Then
        If IsNumeric(txtg1) = False Then
            MsgBox "El codigo debe ser numérico", vbInformation, App.Title
            txtg1.SetFocus
            Exit Function
        End If
        If IsNumeric(txtg2) = False Then
            MsgBox "El codigo debe ser numérico", vbInformation, App.Title
            txtg2.SetFocus
            Exit Function
        End If
        If txtg1 = "" Or txtg2 = "" Then
            MsgBox "Debe completar los codigos de búsqueda.", vbInformation, App.Title
            Exit Function
        Else
            f_desde = ""
            f_hasta = ""
            filtro = filtro & " AND mu.id_general between " & CLng(txtg1) & " and " & CLng(txtg2) & " and mu.anno = " & CInt(txtanno)
        End If
    End If
    ' Fecha
    If f_desde <> "" And f_hasta <> "" Then
        filtro = filtro & " AND mu.fecha_recepcion>='" & f_desde & "'"
        filtro = filtro & " AND mu.fecha_recepcion<='" & f_hasta & "'"
        mvarCRITERIO_LISTADO = "{muestras.FECHA_RECEPCION} >= Date(" & Year(f_desde) & "," & Month(f_desde) & "," & Day(f_desde) & ") AND {muestras.FECHA_RECEPCION} <= Date(" & Year(f_hasta) & "," & Month(f_hasta) & "," & Day(f_hasta) & ")"
    End If
    If cmbCentro.Text <> "" Then
        filtro = filtro & " and mu.centro_id = " & CInt(cmbCentro.BoundText)
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.centro_id} = " & CInt(cmbCentro.BoundText)
    End If
    If cmbReplacement.getTEXTO <> "" Then
        filtro = filtro & " and mu.replacement_id = " & cmbReplacement.getPK_SALIDA
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.replacement_id} = " & cmbReplacement.getPK_SALIDA
    End If
    If chkNoEnviadas.Value = Checked Then
        filtro = filtro & " and mu.enviado_correo = 0 and mu.cerrada <> 0 and mu.revision_usuario <> 0 "
        mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.enviado_correo} = 0 AND {muestras.cerrada} <> 0 AND {muestras.revision_usuario} <> 0 "
    End If
    ' Urgente
    If opUrgente(2).Value = False Then
        If opUrgente(0).Value = True Then
            filtro = filtro & " and mu.urgente = " & 0
            mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.urgente} = 0 "
        ElseIf opUrgente(1).Value = True Then
            filtro = filtro & " and mu.urgente = " & 1
            mvarCRITERIO_LISTADO = mvarCRITERIO_LISTADO & " AND {muestras.urgente} = 1 "
        End If
    End If
    filtroListado = filtro

   On Error GoTo 0
   Exit Function

filtroListado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure filtroListado of Formulario frmListadoMuestras"
End Function
Private Sub buscar()
    Dim i As Integer
    Dim consulta As String
    On Error GoTo fallo
    lista.ListItems.Clear
    Dim rs As New ADODB.Recordset
    Dim filtro As String
    filtro = filtroListado
    If filtro = "" Then
        Exit Sub
    End If
    consulta = "SELECT cl.id_cliente, concat(tm.codigo,'-',cast(mu.id_particular as char)), " & _
               "       cl.nombre,mu.tipo_analisis_id,mu.referencia_cliente,mu.fecha_recepcion,mu.id_muestra,mu.precio,ta.nombre,mu.id_general, " & _
               "       mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada,mu.revision_usuario,ce.nombre,mu.situacion,mu.agrupada, " & _
               "       mu.agrupada_muestra_id,mu.urgente,mu.consulta " & _
               " FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "centros as ce, " & _
                     "muestras as mu " & strSECTOR & _
               " WHERE mu.cliente_id=cl.id_cliente AND mu.tipo_muestra_id=tm.id_tipo_muestra AND mu.tipo_analisis_id=ta.id_tipo_analisis " & _
                      " and mu.centro_id = ce.id_centro " & _
                      filtro & _
                      " order by mu.id_general desc"
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        i = 1
        Dim objLitem As ListItem, objSI As ListSubItem
        While Not rs.EOF
            With lista.ListItems.Add(, , rs(1))
                .SubItems(1) = rs.Fields(2)
                .SubItems(2) = rs.Fields(8)
                .SubItems(3) = rs.Fields(4)
                If Not IsNull(rs.Fields(5)) Then
                .SubItems(4) = rs.Fields(5)
                End If
                If Not IsNull(rs.Fields(9)) Then
                   .SubItems(5) = Format(rs.Fields(9), "00000")
                End If
                If Not IsNull(rs.Fields(6)) Then
                    .SubItems(6) = rs.Fields(6)
                End If
                .SubItems(7) = rs(10)
                .SubItems(8) = rs(15) 'CENTRO
                    
                If rs(13) = 1 Then ' Si cerrada, bola de color
                    .ListSubItems.Add , , "", rs(16) + 7
                Else
                    .ListSubItems.Add , , "", vbNothing
                End If
                ' ICONO MUESTRA AGRUPADA
                If rs(17) = 1 Then ' AGRUPADA
                    If rs(18) = 0 Then
                        .ListSubItems.Add , , "", 10
                    Else
                        .ListSubItems.Add , , "", 11
                    End If
                Else
                    .ListSubItems.Add , , "", vbNothing
                End If
                ' ICONO URGENTE (11)
                If rs(19) = 1 Then
                    .ListSubItems.Add , , "", 12
                Else
                    .ListSubItems.Add , , "", vbNothing
                End If
           '     .SubItems(9) = rs(16) 'SITUACION
            End With
            i = lista.ListItems.Count
            lista.ListItems(i).Checked = True
            If rs.Fields(11) <> 0 Then 'ENVIADO_CORREO
                If rs(13) <> 1 Then ' Abierta
                    lista.ListItems(i).SmallIcon = 15
                    lista.ListItems(i).ToolTipText = "Determinación Pendiente"
                Else
                    lista.ListItems(i).SmallIcon = 1
                    lista.ListItems(i).ToolTipText = "Enviado Correo"
                End If
            Else
                If rs(12) <> 0 Then ' ANULADA
                    lista.ListItems(i).SmallIcon = 2
                    lista.ListItems(i).ToolTipText = "Anulada"
                Else
                    Select Case rs(13) ' Cerrada
                        Case 0 ' Abierta
                            lista.ListItems(i).SmallIcon = 5
                            lista.ListItems(i).ToolTipText = "Abierta"
                        Case 1 ' Cerrada
                            If rs(14) = 0 Then ' Revision Usuario
                                lista.ListItems(i).SmallIcon = 6
                                lista.ListItems(i).ToolTipText = "Cerrada Pendiente Revisar"
                            Else
                                lista.ListItems(i).SmallIcon = 4
                                lista.ListItems(i).ToolTipText = "Cerrada y Revisada por Usuario : " & rs(14)
                            End If
                        Case 2 ' Pdte. Cierre
                            lista.ListItems(i).SmallIcon = 3
                            lista.ListItems(i).ToolTipText = "Pdte. Cierre"
                    End Select
                End If
            End If
            ' Colorear fila si es una muestra en consulta
            If rs("CONSULTA") = 1 Then
                lista_colorear lista, i, vbRed
            End If
            rs.MoveNext
        Wend
        lblMsg.Caption = "Muestras entre el " & Format(fdesde, "dd/mm/yyyy") & " y " & Format(fhasta, "dd/mm/yyyy") & ". Total : " & rs.RecordCount
    Else
        lblMsg.Caption = "No existe ninguna muestra con esos criterios."
    End If
'    Set oAnalisis = Nothing
    Me.MousePointer = 0
    Set rs = Nothing
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al buscar las muestras.", vbCritical, Err.Description
End Sub
Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA)
        frmVerMuestra.Show 1
        actualizar_lista
    End If
End Sub

'Private Sub opTipo_Click(Index As Integer)
'    If Index = 6 Then
'        cmdInformeAgrupado.Visible = True
'        lista.MultiSelect = True
'    Else
'        cmdInformeAgrupado.Visible = False
'        lista.MultiSelect = False
'    End If
'End Sub

Private Sub txtg1_Change()
    txtp1 = ""
    txtp2 = ""
End Sub

Private Sub txtg1_GotFocus()
    txtg1.SelStart = 0
    txtg1.SelLength = Len(txtg1)
End Sub

Private Sub txtg1_LostFocus()
    txtg2 = txtg1
End Sub

Private Sub txtg2_Change()
    txtp1 = ""
    txtp2 = ""
End Sub

Private Sub txtg2_GotFocus()
    txtg2.SelStart = 0
    txtg2.SelLength = Len(txtg2)
End Sub

Private Sub txtp1_Change()
    txtg1 = ""
    txtg2 = ""
End Sub

Private Sub txtp1_GotFocus()
    txtp1.SelStart = 0
    txtp1.SelLength = Len(txtp1)
End Sub

Private Sub txtp1_LostFocus()
    txtp2 = txtp1
End Sub

Private Sub txtp2_Change()
    txtg1 = ""
    txtg2 = ""
End Sub

Public Sub actualizar_lista()
    ' Por si se ha modificado la muestra
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada,mu.revision_usuario,mu.agrupada,mu.agrupada_muestra_id,mu.urgente,mu.consulta " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.id_muestra = " & CLng(lista.ListItems(lista.selectedItem.Index).SubItems(PK_ID_MUESTRA))
    Set rs = datos_bd(consulta)
    If rs.RecordCount >= 1 Then
        Dim oMuestra As New clsMuestra
        lista.ListItems(lista.selectedItem.Index).Text = rs.Fields(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs.Fields(2)
        Dim oAnalisis As New clsTipos_analisis
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = oAnalisis.NombreAnalisis(rs.Fields(3))
        Set oAnalisis = Nothing
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs.Fields(4)
        If Not IsNull(rs.Fields(5)) Then
            lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs.Fields(5)
        End If
'        If Not IsNull(rs.Fields(7)) Then
'            lista.ListItems(lista.SelectedItem.Index).SubItems(5) = Format(rs.Fields(7), "currency")
'        End If
        lista.ListItems(lista.selectedItem.Index).SubItems(7) = rs(8)
        If rs(9) <> 0 Then ' Enviada por correo
            If rs(11) <> 1 Then ' Abierta
                lista.ListItems(lista.selectedItem.Index).SmallIcon = 15
                lista.ListItems(lista.selectedItem.Index).ToolTipText = "Determinación Pendiente"
            Else
                lista.ListItems(lista.selectedItem.Index).SmallIcon = 1
                lista.ListItems(lista.selectedItem.Index).ToolTipText = "Enviado Correo"
            End If
        Else
            If rs(10) <> 0 Then ' Anulada
                lista.ListItems(lista.selectedItem.Index).SmallIcon = 2
                lista.ListItems(lista.selectedItem.Index).ToolTipText = "Anulada"
            Else
                Select Case rs(11) ' Cerrada
                    Case 0 ' Abierta
                        lista.ListItems(lista.selectedItem.Index).SmallIcon = 5
                        lista.ListItems(lista.selectedItem.Index).ToolTipText = "Abierta"
                    Case 1 ' Cerrada
                        If rs(12) = 0 Then ' Revision Usuario
                            lista.ListItems(lista.selectedItem.Index).SmallIcon = 6
                            lista.ListItems(lista.selectedItem.Index).ToolTipText = "Cerrada Pendiente Revisar"
                        Else
                            lista.ListItems(lista.selectedItem.Index).SmallIcon = 4
                            lista.ListItems(lista.selectedItem.Index).ToolTipText = "Cerrada y Revisada por Usuario : " & rs(12)
                        End If
                    Case 2 ' Pdte. Cierre
                        lista.ListItems(lista.selectedItem.Index).SmallIcon = 3
                        lista.ListItems(lista.selectedItem.Index).ToolTipText = "Pdte. Cierre"
                End Select
            End If
        End If
        ' ICONO MUESTRA AGRUPADA
        Dim objLitem As ListItem, objSitem As ListSubItem
        Set objLitem = lista.ListItems(lista.selectedItem.Index)
        Set objSitem = objLitem.ListSubItems(10)
        If rs(13) = 1 Then ' AGRUPADA
            If rs(14) = 0 Then
                objSitem.ReportIcon = 10
            Else
                objSitem.ReportIcon = 11
            End If
        Else
            objSitem.ReportIcon = vbNothing
        End If
        ' ICONO URGENTE
        Set objSitem = objLitem.ListSubItems(11)
        If rs(15) = 1 Then
            objSitem.ReportIcon = 12
        Else
            objSitem.ReportIcon = vbNothing
        End If
        If rs("CONSULTA") = 1 Then
            lista_colorear lista, lista.selectedItem.Index, vbRed
        Else
            lista_colorear lista, lista.selectedItem.Index, vbBlack
        End If
        
    End If
    Set rs = Nothing
End Sub
Private Sub cmdListado_Click_old()
'    Dim total As Currency
'    Dim i As Integer
'    On Error GoTo fallo
'    If lista.ListItems.Count = 0 Then
'        MsgBox "No existen registros para generar el listado.", vbExclamation, App.Title
'        Exit Sub
'    End If
'    Dim rs As New ADODB.RecordSet
'    rs.Fields.Append "c1", adChar, 15, adFldUpdatable
'    rs.Fields.Append "c2", adChar, 15, adFldUpdatable
'    rs.Fields.Append "c3", adChar, 50, adFldUpdatable
'    rs.Fields.Append "c4", adChar, 50, adFldUpdatable
'    rs.Fields.Append "c5", adChar, 50, adFldUpdatable
'    rs.Open
'    total = 0
'    For i = 1 To lista.ListItems.Count
'        rs.AddNew
'        rs("c1") = Left(lista.ListItems(i).SubItems(5) & "(" & lista.ListItems(i).Text & ")", 15)
'        rs("c2") = Left(lista.ListItems(i).SubItems(4), 15)
'        rs("c3") = Left(lista.ListItems(i).SubItems(1), 50)
'        rs("c4") = Left(lista.ListItems(i).SubItems(2), 50)
'        rs("c5") = Left(lista.ListItems(i).SubItems(3), 50)
'        rs.Update
'    Next
'    ' Generar Listado
'    Dim Listado As New dataListadoMuestras
'    ' Cabecera
'    With Listado.Sections("cabecera")
'        .Controls("lbltitulo").Caption = "Listado de Muestras desde " & Format(fdesde, "dd/mm/yyyy") & " al " & Format(fhasta, "dd/mm/yyyy")
'        If chkTodos.value = Checked Then
'            .Controls("lblcliente").Caption = "Cliente : *** TODOS ***"
'        Else
'            .Controls("lblcliente").Caption = "Cliente : " & cmbClientes.getTEXTO
'        End If
'    End With
'    Set Listado.Sections("cabecera").Controls("logo").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
'    'Detalle
'    With Listado.Sections("detalle")
'        .Controls("c1").DataField = rs.Fields("c1").Name
'        .Controls("c2").DataField = rs.Fields("c2").Name
'        .Controls("c3").DataField = rs.Fields("c3").Name
'        .Controls("c4").DataField = rs.Fields("c4").Name
'        .Controls("c5").DataField = rs.Fields("c5").Name
'    End With
'    ' Pie de Pagina
''    With Listado.Sections("pie")
''        .Controls("lbltotal").Caption = Format(total, "currency")
''    End With
'    Set Listado.DataSource = rs
'    Listado.Caption = "Listado de Muestras"
'    Listado.WindowState = vbMaximized
'    Listado.Show
'    Set rs = Nothing
''    Me.Height = 7890
''    Me.Width = 12780
'    Exit Sub
'fallo:
'    MsgBox "Error al generar el listado de Analisis pendientes.", vbCritical, Err.Description
End Sub

Public Sub permisos()
'    cmdVida.Visible = False
End Sub

Private Sub txtp2_GotFocus()
    txtp2.SelStart = 0
    txtp2.SelLength = Len(txtp2)
End Sub
