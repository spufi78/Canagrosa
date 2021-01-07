VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#13.2#0"; "Codejock.CommandBars.v13.2.1.ocx"
Begin VB.Form frmMenu2 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mensajes de Usuario"
   ClientHeight    =   9795
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9480
   ControlBox      =   0   'False
   Icon            =   "frmMenu2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   9480
   Begin XtremeSuiteControls.TabControl tabControl 
      Height          =   9375
      Left            =   0
      TabIndex        =   0
      Top             =   45
      Width           =   9375
      _Version        =   851970
      _ExtentX        =   16536
      _ExtentY        =   16536
      _StockProps     =   68
      Appearance      =   6
      Color           =   16
      PaintManager.Layout=   2
      PaintManager.ShowIcons=   -1  'True
      PaintManager.LargeIcons=   -1  'True
      PaintManager.FixedTabWidth=   115
      ItemCount       =   5
      Item(0).Caption =   "Mensajes de Usuario"
      Item(0).Tooltip =   "1"
      Item(0).ImageIndex=   1
      Item(0).ControlCount=   3
      Item(0).Control(0)=   "mensajesLista"
      Item(0).Control(1)=   "Frame2"
      Item(0).Control(2)=   "lbltitulo(3)"
      Item(1).Caption =   "Equipos C/V/M"
      Item(1).ImageIndex=   2
      Item(1).ControlCount=   9
      Item(1).Control(0)=   "Frame1"
      Item(1).Control(1)=   "Label2(0)"
      Item(1).Control(2)=   "Label1"
      Item(1).Control(3)=   "lbltitulo(1)"
      Item(1).Control(4)=   "equiposchkFuera"
      Item(1).Control(5)=   "equiposchkSolo"
      Item(1).Control(6)=   "equiposLista"
      Item(1).Control(7)=   "equiposlbltotal"
      Item(1).Control(8)=   "lbltitulo(5)"
      Item(2).Caption =   "Muestras fuera de Plazo"
      Item(2).ImageIndex=   3
      Item(2).ControlCount=   6
      Item(2).Control(0)=   "lista"
      Item(2).Control(1)=   "lbltitulo(2)"
      Item(2).Control(2)=   "Label2(4)"
      Item(2).Control(3)=   "Label4"
      Item(2).Control(4)=   "cmbCentroFP"
      Item(2).Control(5)=   "lblTotalFP"
      Item(3).Caption =   "Recualificaciones"
      Item(3).ImageIndex=   4
      Item(3).ControlCount=   5
      Item(3).Control(0)=   "lbltitulo(0)"
      Item(3).Control(1)=   "recualificacionesLista"
      Item(3).Control(2)=   "recualificacionesChkTodas"
      Item(3).Control(3)=   "lblrecualificaciones(0)"
      Item(3).Control(4)=   "lblrecualificaciones(1)"
      Item(4).Caption =   "Muestras Proximas a Caducar"
      Item(4).ImageIndex=   5
      Item(4).ControlCount=   6
      Item(4).Control(0)=   "lbltitulo(4)"
      Item(4).Control(1)=   "listaProximas"
      Item(4).Control(2)=   "Label2(6)"
      Item(4).Control(3)=   "txtnum"
      Item(4).Control(4)=   "cambiar"
      Item(4).Control(5)=   "Label3"
      Begin MSDataListLib.DataCombo cmbCentroFP 
         Height          =   315
         Left            =   -69190
         TabIndex        =   43
         Top             =   1080
         Visible         =   0   'False
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   ""
      End
      Begin VB.TextBox txtnum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   -65545
         TabIndex        =   37
         Text            =   "2"
         Top             =   1035
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Timer Timer1 
         Interval        =   60000
         Left            =   6030
         Top             =   8910
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "Detalle del mensaje"
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
         Height          =   3075
         Left            =   135
         TabIndex        =   19
         Top             =   6210
         Width           =   9060
         Begin VB.TextBox txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   1365
            Index           =   0
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   1170
            Width           =   8835
         End
         Begin VB.TextBox txttexto 
            Appearance      =   0  'Flat
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
            Height          =   510
            Index           =   1
            Left            =   90
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   26
            Top             =   630
            Width           =   8835
         End
         Begin VB.TextBox txttexto 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   2
            Left            =   450
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   225
            Width           =   4290
         End
         Begin VB.TextBox txttexto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   3
            Left            =   5580
            Locked          =   -1  'True
            TabIndex        =   21
            Top             =   225
            Width           =   1365
         End
         Begin VB.TextBox txttexto 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Height          =   330
            Index           =   4
            Left            =   7695
            Locked          =   -1  'True
            TabIndex        =   20
            Top             =   225
            Width           =   1230
         End
         Begin XtremeSuiteControls.PushButton mensajesCmdDetalle 
            Height          =   435
            Left            =   6570
            TabIndex        =   28
            Top             =   2565
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Ver el detalle del Mensaje"
            BackColor       =   12632256
            Enabled         =   0   'False
            Appearance      =   5
            Picture         =   "frmMenu2.frx":6852
         End
         Begin XtremeSuiteControls.PushButton mensajesEliminar 
            Height          =   435
            Left            =   2475
            TabIndex        =   29
            Top             =   2565
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Eliminar Mensaje"
            BackColor       =   12632256
            Appearance      =   5
            Picture         =   "frmMenu2.frx":6AD3
         End
         Begin XtremeSuiteControls.PushButton mensajesCrear 
            Height          =   435
            Left            =   90
            TabIndex        =   31
            Top             =   2565
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Crear Mensaje"
            BackColor       =   12632256
            Appearance      =   5
            Picture         =   "frmMenu2.frx":D335
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "De"
            Height          =   195
            Index           =   1
            Left            =   90
            TabIndex        =   25
            Top             =   270
            Width           =   210
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Válido"
            Height          =   195
            Index           =   2
            Left            =   4905
            TabIndex        =   24
            Top             =   270
            Width           =   435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "hasta"
            Height          =   195
            Index           =   3
            Left            =   7155
            TabIndex        =   23
            Top             =   270
            Width           =   390
         End
      End
      Begin VB.CheckBox recualificacionesChkTodas 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mostrar todos las recualificaciones pendientes de todos los usuarios"
         Height          =   255
         Left            =   -69910
         TabIndex        =   17
         Top             =   1080
         Visible         =   0   'False
         Width           =   5100
      End
      Begin VB.CheckBox equiposchkSolo 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mostrar todos los equipos pendientes"
         Height          =   255
         Left            =   -69865
         TabIndex        =   9
         Top             =   1260
         Visible         =   0   'False
         Width           =   3345
      End
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H80000008&
         Height          =   1185
         Left            =   -62935
         TabIndex        =   4
         Top             =   990
         Visible         =   0   'False
         Width           =   2205
         Begin VB.OptionButton equiposOpTipo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   8
            Top             =   270
            Value           =   -1  'True
            Width           =   1725
         End
         Begin VB.OptionButton equiposOpTipo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Mantenimiento"
            Height          =   195
            Index           =   3
            Left            =   150
            TabIndex        =   7
            Top             =   930
            Width           =   1875
         End
         Begin VB.OptionButton equiposOpTipo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Calibraciones"
            Height          =   195
            Index           =   1
            Left            =   150
            TabIndex        =   6
            Top             =   510
            Width           =   1815
         End
         Begin VB.OptionButton equiposOpTipo 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Verificaciones"
            Height          =   195
            Index           =   2
            Left            =   150
            TabIndex        =   5
            Top             =   720
            Width           =   1845
         End
      End
      Begin VB.CheckBox equiposchkFuera 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Mostrar equipos F/S, B, CAU, E, I, R"
         Height          =   255
         Left            =   -69865
         TabIndex        =   3
         Top             =   1485
         Visible         =   0   'False
         Width           =   3750
      End
      Begin MSComctlLib.ListView lista 
         Height          =   7530
         Left            =   -69955
         TabIndex        =   1
         Top             =   1485
         Visible         =   0   'False
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   13282
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
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
      Begin MSComctlLib.ListView equiposLista 
         Height          =   6555
         Left            =   -69865
         TabIndex        =   10
         Top             =   2205
         Visible         =   0   'False
         Width           =   9120
         _ExtentX        =   16087
         _ExtentY        =   11562
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
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
      Begin MSComctlLib.ListView recualificacionesLista 
         Height          =   6855
         Left            =   -69910
         TabIndex        =   15
         Top             =   1395
         Visible         =   0   'False
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   12091
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
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
      Begin MSComctlLib.ListView mensajesLista 
         Height          =   5145
         Left            =   135
         TabIndex        =   18
         Top             =   1035
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   9075
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
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
         Left            =   0
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMenu2.frx":13B97
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMenu2.frx":14471
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView listaProximas 
         Height          =   6855
         Left            =   -69955
         TabIndex        =   34
         Top             =   1530
         Visible         =   0   'False
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   12091
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
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
      Begin MSComCtl2.UpDown cambiar 
         Height          =   450
         Left            =   -64854
         TabIndex        =   38
         Top             =   1035
         Visible         =   0   'False
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   794
         _Version        =   393216
         Value           =   365
         BuddyControl    =   "txtnum"
         BuddyDispid     =   196609
         OrigLeft        =   5400
         OrigTop         =   1035
         OrigRight       =   5640
         OrigBottom      =   1485
         Max             =   365
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Solo INTERNOS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Index           =   5
         Left            =   -69955
         TabIndex        =   45
         Top             =   900
         Visible         =   0   'False
         Width           =   9285
      End
      Begin VB.Label lblTotalFP 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   -64645
         TabIndex        =   44
         Top             =   1260
         Visible         =   0   'False
         Width           =   3930
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Centro"
         Height          =   240
         Left            =   -69865
         TabIndex        =   42
         Top             =   1125
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label lblrecualificaciones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "* No Caducados : 0"
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   1
         Left            =   -69910
         TabIndex        =   41
         Top             =   8505
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblrecualificaciones 
         BackColor       =   &H00FFFFFF&
         Caption         =   "* Caducados : 0"
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   -69910
         TabIndex        =   40
         Top             =   8280
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "Nº días hasta la fecha de entrega"
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
         Left            =   -68605
         TabIndex        =   39
         Top             =   1140
         Visible         =   0   'False
         Width           =   2970
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "* ROJO : Muestras fuera de plazo con IPA"
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   6
         Left            =   -69955
         TabIndex        =   36
         Top             =   8460
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   4
         Left            =   -69955
         TabIndex        =   35
         Top             =   630
         Visible         =   0   'False
         Width           =   9285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "* ROJO : Muestras fuera de plazo con IPA"
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   4
         Left            =   -69955
         TabIndex        =   33
         Top             =   9090
         Visible         =   0   'False
         Width           =   4455
      End
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mensajes de Usuario"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   3
         Left            =   45
         TabIndex        =   30
         Top             =   630
         Width           =   9285
      End
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recualificaciones Pendientes y en Formación"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   0
         Left            =   -69955
         TabIndex        =   16
         Top             =   585
         Visible         =   0   'False
         Width           =   9285
      End
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         Caption         =   "Equipos pendientes de Calibración/Verificación y Mantenimiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   1
         Left            =   -69955
         TabIndex        =   14
         Top             =   585
         Visible         =   0   'False
         Width           =   9285
      End
      Begin VB.Label equiposlbltotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   -63070
         TabIndex        =   13
         Top             =   9000
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "* Fuera de servicio o Baja"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   -69895
         TabIndex        =   12
         Top             =   8820
         Visible         =   0   'False
         Width           =   3060
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "* Fuera de fecha"
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   -69895
         TabIndex        =   11
         Top             =   9030
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lbltitulo 
         Alignment       =   2  'Center
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   2
         Left            =   -69955
         TabIndex        =   2
         Top             =   585
         Visible         =   0   'False
         Width           =   9285
      End
   End
   Begin XtremeSuiteControls.PushButton cmdMinimizar 
      Height          =   300
      Left            =   7020
      TabIndex        =   32
      Top             =   9450
      Width           =   2355
      _Version        =   851970
      _ExtentX        =   4154
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Minimizar"
      BackColor       =   12632256
      Appearance      =   5
      Picture         =   "frmMenu2.frx":14D4B
   End
   Begin XtremeCommandBars.ImageManager ImageManager1 
      Left            =   6435
      Top             =   9405
      _Version        =   851970
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmMenu2.frx":1B5AD
   End
End
Attribute VB_Name = "frmMenu2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cambiar_Change()
    muestrasProximasListado
End Sub

Private Sub cmbCentroFP_Change()
    muestrasListado
End Sub
Private Sub cmdMinimizar_Click()
    Me.WindowState = vbMinimized
End Sub

Private Sub equiposchkFuera_Click()
    equiposListado
End Sub
Private Sub equiposchkSolo_Click()
    equiposListado
End Sub
Private Sub equiposLista_DblClick()
   On Error GoTo equiposLista_DblClick_Error

    If equiposLista.ListItems.Count = 0 Then Exit Sub
'    Dim strTipo As String
'    Dim objfrm As Object
'    Dim strIdEquipo As String
'    Dim strFecha As String, strId_Evento As String
'    Dim mvarobjEquipos As New clsEquipos
'    strTipo = UCase(ClrStr(equiposLista.ListItems(equiposLista.selectedItem.Index).SubItems(6), False, True))
'    strIdEquipo = equiposLista.ListItems(equiposLista.selectedItem.Index)
'    strFecha = equiposLista.ListItems(equiposLista.selectedItem.Index).SubItems(4)
'    strId_Evento = equiposLista.ListItems(equiposLista.selectedItem.Index).SubItems(7)
'    Select Case strTipo
'        Case "2"
'            Set objfrm = New frmEquipoEdicionMtoFechasEdicion
'        Case "0"
'            Set objfrm = New frmEquipoEdicionCalibracion
'        Case "1"
'            Set objfrm = New frmEquipoEdicionVerificacion
'    End Select
'    objfrm.VieneDeCuaderno = True
'    objfrm.idEvento = CLng(strId_Evento)
'    objfrm.FechaPrevista = CDate(strFecha)
'    objfrm.idEquipo = CLng(strIdEquipo)
'    objfrm.Show vbModal
'    If objfrm.RESULTADO Then
'        equiposListado
'    End If
'    Unload objfrm
'    Set mvarobjEquipos = Nothing
'    Set objfrm = Nothing
    
    Dim objfrm As New frmEquipoEdicion
    Dim lngid As Long
    Dim objEquipo As New clsEquipos
    
    lngid = equiposLista.ListItems(equiposLista.selectedItem.Index)
    If lngid <= 0 Then Exit Sub
    
    Call objEquipo.Carga(lngid)
    
    Set objfrm.EQUIPO = objEquipo
    
    If objEquipo.getALTA_BAJA = 1 Then
        objfrm.TipoEdicion = visualizar
    Else
        objfrm.TipoEdicion = EDICION
    End If
    
    objfrm.Show vbModal
    
    Unload objfrm
    Set objfrm = Nothing

   On Error GoTo 0
   Exit Sub

equiposLista_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure equiposLista_DblClick of Formulario frmMenu2"
    
End Sub

Private Sub equiposOpTipo_Click(Index As Integer)
    equiposListado
End Sub

Private Sub Form_Load()
    log Me.Name
    Me.Left = 0
    Me.top = 0
'    Me.Left = Screen.Width - Me.Width - frmMenu.ButtonBar.Width - 80
    
    
    Set tabControl.Icons = ImageManager1.Icons
    tabControl.Item(0).Selected = True
    
    mensajesCabecera
    mensajesTimer
    cargar_combo cmbCentroFP, New clsCentros

    permisos
End Sub

Private Sub muestrasCabecera()
    lbltitulo(2) = "Muestras fuera de Plazo"
    lista.ColumnHeaders.Clear
    With lista.ColumnHeaders
        .Add , , "ID_MUESTRA", 1, lvwColumnLeft
        .Add , , "General", 900, lvwColumnCenter
        .Add , , "Particular", 900, lvwColumnCenter
        .Add , , "Cliente", 2000, lvwColumnLeft
        .Add , , "Referencia", 2000, lvwColumnLeft
        .Add , , "F.Recepcion", 1000, lvwColumnCenter
        .Add , , "F.Prevista", 1000, lvwColumnCenter
        .Add , , "Días Retraso", 900, lvwColumnCenter
    End With
End Sub

Private Sub muestrasListado()
    Dim rs As ADODB.Recordset
    Dim oMuestra As New clsMuestra
    
   On Error GoTo Carga_Error

    lista.ListItems.Clear
    Set rs = oMuestra.listadoMuestrasFueraPlazo(cmbCentroFP.BoundText)
    lblTotalFP.Caption = "Total muestras : " & rs.RecordCount
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
              .SubItems(1) = rs(1)
              .SubItems(2) = rs(2)
              .SubItems(3) = rs(3)
              .SubItems(4) = rs(4)
              .SubItems(5) = Format(rs(5), "dd-mm-yyyy")
              .SubItems(6) = Format(rs(6), "dd-mm-yyyy")
              .SubItems(7) = rs(7)
              If rs(8) = 1 Then
                colorear lista, lista.ListItems.Count, vbRed
              End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    Set oMuestra = Nothing
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

Carga_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Carga of Formulario frmCE_NoIniciados"
End Sub
Private Sub muestrasProximasCabecera()
    lbltitulo(4) = "Muestras Proximas a Caducar"

    listaProximas.ColumnHeaders.Clear
    With listaProximas.ColumnHeaders
        .Add , , "ID_MUESTRA", 1, lvwColumnLeft
        .Add , , "General", 900, lvwColumnCenter
        .Add , , "Particular", 900, lvwColumnCenter
        .Add , , "Cliente", 2000, lvwColumnLeft
        .Add , , "Referencia", 2000, lvwColumnLeft
        .Add , , "F.Recepcion", 1000, lvwColumnCenter
        .Add , , "F.Prevista", 1000, lvwColumnCenter
        .Add , , "Días Retraso", 1000, lvwColumnCenter
    End With
End Sub

Private Sub muestrasProximasListado()
    Dim rs As ADODB.Recordset
    Dim oMuestra As New clsMuestra
    
   On Error GoTo Carga_Error

    listaProximas.ListItems.Clear
    Set rs = oMuestra.listadoMuestrasProximasPlazo(txtnum)
    If rs.RecordCount > 0 Then
        Do
            With listaProximas.ListItems.Add(, , rs(0))
              .SubItems(1) = rs(1)
              .SubItems(2) = rs(2)
              .SubItems(3) = rs(3)
              .SubItems(4) = rs(4)
              .SubItems(5) = Format(rs(5), "dd-mm-yyyy")
              .SubItems(6) = Format(rs(6), "dd-mm-yyyy")
              .SubItems(7) = rs(7)
              If rs(8) = 1 Then
                colorear listaProximas, listaProximas.ListItems.Count, vbRed
              End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    Set oMuestra = Nothing
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

Carga_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Carga of Formulario muestrasProximasListado"
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

Private Sub listaProximas_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If listaProximas.ListItems.Count > 0 Then
     listaProximas.SortKey = ColumnHeader.Index - 1
     If listaProximas.SortOrder = 0 Then
        listaProximas.SortOrder = 1
     Else
        listaProximas.SortOrder = 0
     End If
     listaProximas.Sorted = True
   End If

End Sub

Private Sub listaProximas_DblClick()
    If listaProximas.ListItems.Count > 0 Then
        gmuestra = listaProximas.ListItems(listaProximas.selectedItem.Index).Text
        frmVerMuestra.Show 1
    End If

End Sub

Private Sub mensajesCmdDetalle_Click()
    If mensajesLista.ListItems.Count > 0 Then
        frmMensaje_Detalle.PK = mensajesLista.ListItems(mensajesLista.selectedItem.Index).SubItems(1)
        frmMensaje_Detalle.Show 1
    End If
End Sub

Private Sub mensajesCrear_Click()
    frmMEN_Crear.Show 1
    mensajesTimer
End Sub

Private Sub mensajesEliminar_Click()
    Dim oMensaje As New clsMensajes_usuarios
    oMensaje.Eliminar mensajesLista.ListItems(mensajesLista.selectedItem.Index).SubItems(1), USUARIO.getID_EMPLEADO
    Set oMensaje = Nothing
    mensajesTimer
End Sub

Private Sub mensajesLista_Click()
    mensajesCargarMensaje
End Sub

Private Sub mensajesLista_DblClick()
    mensajesCmdDetalle_Click
End Sub

Private Sub recualificacionesChkTodas_Click()
    recualificacionesListado
End Sub

Private Sub recualificacionesLista_DblClick()
    If recualificacionesLista.ListItems.Count > 0 Then
        frmEmpleados_Cualificaciones_Nueva.EMPLEADO_ID = recualificacionesLista.ListItems(recualificacionesLista.selectedItem.Index).SubItems(5)
        frmEmpleados_Cualificaciones_Nueva.ID_CUALIFICACION = recualificacionesLista.ListItems(recualificacionesLista.selectedItem.Index).Text
        frmEmpleados_Cualificaciones_Nueva.Show 1
    End If
End Sub

Private Sub tab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'    MsgBox Item.Index
End Sub
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).Text
        frmVerMuestra.Show 1
    End If
End Sub

Private Sub colorear(lista As ListView, fila As Integer, color As Long)
    Dim i As Integer
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
End Sub
Private Sub equiposCabecera()
    equiposLista.ColumnHeaders.Clear
    With equiposLista.ColumnHeaders
        .Add , , "Nº Equipo", 950, lvwColumnLeft
'MANTIS-824-I
        .Add , , "Estado", 800, lvwColumnCenter
'MANTIS-824-F
        .Add , , "Equipo", 2750, lvwColumnLeft
        .Add , , "Responsable", 0, lvwColumnLeft
'MANTIS-824-I
'       .Add , , "Fecha Prevista", 1500, lvwColumnLeft
'       .Add , , "Tipo", 1100, lvwColumnLeft
        .Add , , "Fecha Prevista", 1500, lvwColumnCenter
        .Add , , "Tipo", 700, lvwColumnCenter
'MANTIS-824-I
        .Add , , "idTipo", 0, lvwColumnLeft
        .Add , , "idEvento", 0, lvwColumnLeft
        .Add , , "Responsable", 2100, lvwColumnLeft
    End With
     
     Dim oParam As New clsParametros
    oParam.Carga parametros.USUARIOS_CUADERNO_AVISOS_EQUIPOS, ""
    
    If InStr(1, Replace(oParam.getVALOR, " ", ""), "," & CStr(prmIdUsuario)) > 0 Or _
       InStr(1, Replace(oParam.getVALOR, " ", ""), CStr(prmIdUsuario) & ",") > 0 Then
        equiposchkSolo.Value = Checked
    End If
    
End Sub
Private Sub equiposListado()
    
    Dim rs As ADODB.Recordset
    Dim mvarobjEquipos As New clsEquipos
   
   On Error GoTo Carga_Error
    
    Dim tipo As Integer
    If equiposOpTipo(0).Value = True Then
        tipo = 0
    ElseIf equiposOpTipo(1).Value = True Then
        tipo = 1
    ElseIf equiposOpTipo(2).Value = True Then
        tipo = 2
    Else
        tipo = 3
    End If
    Set rs = mvarobjEquipos.ListadoFueraFecha(USUARIO.getID_EMPLEADO, equiposchkSolo.Value, tipo, equiposchkFuera.Value)
    equiposLista.ListItems.Clear
    equiposlbltotal = "Total : " & rs.RecordCount
    If rs.RecordCount > 0 Then
        Do
            With equiposLista.ListItems.Add(, , Format(rs("ID_EQUIPO"), "0000"))
              .SubItems(1) = rs("ESTADO_ID")
              .SubItems(2) = rs("NOMBRE")
              .SubItems(3) = rs("ID_EMPLEADO")
              .SubItems(4) = rs("FECHA_PREVISTA")
              .SubItems(5) = rs("TIPO")
              .SubItems(6) = ""
              .SubItems(7) = ""
              .SubItems(8) = rs("RESPONSABLE_INT")
'              .SubItems(6) = rs("ID_TIPO")
'              .SubItems(7) = rs("ID_EVENTO")
'              .SubItems(8) = rs("responsable_int")
               If CInt(rs("FS")) = 1 Then
                  If rs("FECHA_PREVISTA") >= Date Then ' Fuera de servicio y en fecha
                    colorear equiposLista, equiposLista.ListItems.Count, vbBlack
                  Else
                    ' Fuera de servicio y fuera de fecha
                    colorear equiposLista, equiposLista.ListItems.Count, vbBlue
                  End If
               Else
                  If rs("FECHA_PREVISTA") >= Date Then ' Activo y en fecha
                    colorear equiposLista, equiposLista.ListItems.Count, vbBlack
                  Else
                    ' Activo y fuera de fecha
                    colorear equiposLista, equiposLista.ListItems.Count, vbRed
                  End If
               End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    Set mvarobjEquipos = Nothing
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

Carga_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Carga of Formulario frmEquipoCuadernoAvisos"
End Sub

Private Sub recualificacionesCabecera()
    recualificacionesLista.ColumnHeaders.Clear
    With recualificacionesLista.ColumnHeaders
        .Add , , "ID_CUALIFICACION", 1, lvwColumnLeft
        .Add , , "P.N.T.", 3800, lvwColumnLeft
        .Add , , "Técnico", 2000, lvwColumnLeft
        .Add , , "Formador", 2000, lvwColumnLeft
        .Add , , "F.Ult.Recu.", 1050, lvwColumnLeft
        .Add , , "ID_EMPLEADO_TECNICO", 1, lvwColumnLeft
    End With
     
'     Dim oParam As New clsParametros
'    oParam.Carga parametros.USUARIOS_CUADERNO_AVISOS_EQUIPOS, ""
'
'    If InStr(1, Replace(oParam.getVALOR, " ", ""), "," & CStr(prmIdUsuario)) > 0 Or _
'       InStr(1, Replace(oParam.getVALOR, " ", ""), CStr(prmIdUsuario) & ",") > 0 Then
'        equiposchkSolo.value = Checked
'    End If
    
End Sub
Private Sub recualificacionesListado()
    
    Dim rs As ADODB.Recordset
    Dim oEC As New clsEmpleados_cualificaciones
   
   On Error GoTo recualificacionesListado_Error

    Set rs = oEC.RecualificacionesPendientes(USUARIO.getID_EMPLEADO, recualificacionesChkTodas.Value)
    recualificacionesLista.ListItems.Clear
    Dim c1 As Integer
    Dim c2 As Integer
    c1 = 0
    c2 = 0
    lbltitulo(0).Caption = "Recualificaciones Pendientes y en Formación. Total : " & rs.RecordCount
    If rs.RecordCount > 0 Then
        Do
            With recualificacionesLista.ListItems.Add(, , Format(rs(0), "0000"))
              .SubItems(1) = rs(1)
              .SubItems(2) = rs(2)
              .SubItems(3) = rs(3)
              If Format(rs(4), "yyyy-mm-dd") <> "1900-01-01" Then
                  .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
              End If
              .SubItems(5) = rs(5)
              Dim f1 As Date
              Dim f2 As Date
              f1 = Format(rs(6), "yyyy-mm-dd")
              f2 = Format(Date - 30, "yyyy-mm-dd")
              If f1 < f2 Then
                c1 = c1 + 1
                colorear recualificacionesLista, recualificacionesLista.ListItems.Count, vbRed
              Else
                c2 = c2 + 1
              End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    lblrecualificaciones(0) = "* Caducados : " & c1
    lblrecualificaciones(1) = "* No Caducados : " & c2
    
    
    Set oEC = Nothing
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

recualificacionesListado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure recualificacionesListado of Formulario frmMenu2"

End Sub

Private Sub tabControl_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Index
    Case 0 ' Mensajes
        mensajesTimer
    Case 1 ' Equipos
        equiposCabecera
        equiposListado
    Case 2 ' Muestras no cerradas
        muestrasCabecera
        muestrasListado
    Case 3 ' Recualificaciones
        recualificacionesCabecera
        recualificacionesListado
    Case 4 ' Muestras proximas a Caducar
        muestrasProximasCabecera
        muestrasProximasListado
    End Select

End Sub
Private Sub mensajesCabecera()
    mensajesLista.ColumnHeaders.Clear
    With mensajesLista.ColumnHeaders
        .Add , , "Mensajes del usuario", 6500, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Accion", 1, lvwColumnLeft
        .Add , , "F.Desde", 1100, lvwColumnCenter
        .Add , , "F.Hasta", 1100, lvwColumnCenter
    End With
End Sub
Private Sub mensajesListado(rs As ADODB.Recordset)
    Dim oMensaje As New clsMensajes
    mensajesLista.ListItems.Clear
    Dim leido As Boolean
    leido = True
    If rs.RecordCount > 0 Then
        Do
            With mensajesLista.ListItems.Add(, , rs(1))
              .SubItems(1) = rs(0) ' ID
              .SubItems(2) = rs(3) ' ACCION
              .SubItems(3) = rs(5) ' FDESDE
              .SubItems(4) = rs(6) ' FHASTA
            End With
            If rs(2) = 0 Then
                leido = False
                mensajesLista.ListItems(mensajesLista.ListItems.Count).SmallIcon = 1
'                popupCreacion rs(0), rs(1), rs(4)
            Else
                mensajesLista.ListItems(mensajesLista.ListItems.Count).SmallIcon = 2
            End If
            rs.MoveNext
        Loop Until rs.EOF
'        If leido = False Then
'            If Not ControlPanelXP1.PanelOpen Then
'                ControlPanelXP1.PanelOpen = True
'            End If
'        End If
    End If
    Set oMensaje = Nothing
    Set rs = Nothing
End Sub
Public Sub mensajesTimer(Optional Actualizar As Boolean)
    Dim oMensaje As New clsMensajes
    Dim rs As ADODB.Recordset
    Set rs = oMensaje.Listado
    If rs.RecordCount <> mensajesLista.ListItems.Count Or Actualizar Then
        mensajesListado rs
    End If
    Set oMensaje = Nothing
    Set rs = Nothing
End Sub
Private Sub mensajesCargarMensaje()
    Dim oMensaje As New clsMensajes
    If mensajesLista.ListItems.Count = 0 Then Exit Sub
    If oMensaje.Carga(mensajesLista.ListItems(mensajesLista.selectedItem.Index).SubItems(1)) = True Then
        txttexto(1) = oMensaje.getASUNTO
        txttexto(0) = oMensaje.getTEXTO
        Dim oEmple As New clsUsuarios
        oEmple.CARGAR (oMensaje.getEMPLEADO_ID)
        txttexto(2) = oEmple.getUSUARIO
        txttexto(3) = Format(oMensaje.getFECHA_INICIO, "dd-mm-yyyy")
        txttexto(4) = Format(oMensaje.getFECHA_FIN, "dd-mm-yyyy")
        Dim omu As New clsMensajes_usuarios
        omu.Leer (mensajesLista.ListItems(mensajesLista.selectedItem.Index).SubItems(1))
        mensajesLista.ListItems(mensajesLista.selectedItem.Index).SmallIcon = 2
'        If mensajesLista.ListItems(mensajesLista.SelectedItem.Index).SubItems(2) <> "" Then
            mensajesCmdDetalle.Enabled = True
'        Else
'            mensajesCmdDetalle.Enabled = False
'        End If
    End If
    Set oMensaje = Nothing
End Sub

Private Sub Timer1_Timer()
    mensajesTimer
End Sub

Private Sub permisos()
'    If USUARIO.getPER_PLAZO_ENTREGA_LISTADO = False Then
'        tabControl.Item(2).Visible = False
'        tabControl.Item(4).Visible = False
'    End If
End Sub
