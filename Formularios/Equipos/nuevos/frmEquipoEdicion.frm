VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmEquipoEdicion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   11055
   ClientLeft      =   1905
   ClientTop       =   1485
   ClientWidth     =   13305
   ClipControls    =   0   'False
   Icon            =   "frmEquipoEdicion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11055
   ScaleWidth      =   13305
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraApertura 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Apertura/Cierre"
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   8055
      TabIndex        =   222
      Top             =   10080
      Visible         =   0   'False
      Width           =   3045
      Begin VB.OptionButton optAbrir 
         Caption         =   "ABRIR"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   224
         Top             =   240
         Width           =   1215
      End
      Begin VB.OptionButton optCerrar 
         Caption         =   "CERRAR"
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
         Height          =   405
         Left            =   1590
         Style           =   1  'Graphical
         TabIndex        =   223
         Top             =   225
         Value           =   -1  'True
         Width           =   1305
      End
   End
   Begin VB.Frame frmUnidadesLote 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Unidades de Lote Existentes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   5940
      TabIndex        =   210
      Top             =   10080
      Visible         =   0   'False
      Width           =   2850
      Begin VB.TextBox txtUnidadesLote 
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
         Height          =   285
         Left            =   1395
         MaxLength       =   100
         TabIndex        =   211
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Unidades"
         Height          =   195
         Index           =   27
         Left            =   495
         TabIndex        =   212
         Top             =   405
         Width           =   675
      End
   End
   Begin VB.TextBox txtactmantenimiento 
      Height          =   375
      Left            =   6930
      TabIndex        =   205
      Text            =   "Text1"
      Top             =   10710
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtrutaverificacion 
      Height          =   375
      Left            =   7650
      TabIndex        =   204
      Text            =   "Text1"
      Top             =   10350
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.TextBox txtactverificacion 
      Height          =   375
      Left            =   6930
      TabIndex        =   203
      Text            =   "Text1"
      Top             =   10350
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtactcalibracion 
      Height          =   375
      Left            =   6975
      TabIndex        =   202
      Text            =   "Text1"
      Top             =   10170
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtrutacalibracion 
      Height          =   375
      Left            =   7695
      TabIndex        =   201
      Text            =   "Text1"
      Top             =   10170
      Visible         =   0   'False
      Width           =   2265
   End
   Begin VB.TextBox txtIncertidumbreMax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8565
      MaxLength       =   100
      TabIndex        =   176
      Top             =   10215
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtToleranciaMax 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8520
      MaxLength       =   100
      TabIndex        =   175
      Top             =   10365
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.TextBox txtPrecision 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   8520
      MaxLength       =   100
      TabIndex        =   174
      Top             =   10695
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.CommandButton cmdEspecificaciones 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Espec. Técnicas"
      Height          =   870
      Left            =   1530
      Picture         =   "frmEquipoEdicion.frx":000C
      Style           =   1  'Graphical
      TabIndex        =   167
      Top             =   10125
      Width           =   1440
   End
   Begin VB.CommandButton cmdRecepcion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Datos Recepción"
      Height          =   870
      Left            =   45
      Picture         =   "frmEquipoEdicion.frx":08D6
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   10125
      Width           =   1440
   End
   Begin MSComctlLib.ImageList imagenes 
      Left            =   8865
      Top             =   10215
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipoEdicion.frx":11A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipoEdicion.frx":15C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEquipoEdicion.frx":1A01
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtestado 
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3645
      Locked          =   -1  'True
      TabIndex        =   147
      Text            =   "123456"
      Top             =   585
      Visible         =   0   'False
      Width           =   2970
   End
   Begin VB.Frame frmBolas 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1260
      Left            =   10875
      TabIndex        =   143
      Top             =   -135
      Width           =   2370
      Begin VB.Image imgmantenimiento 
         Height          =   480
         Left            =   -30
         Picture         =   "frmEquipoEdicion.frx":1E0E
         Top             =   855
         Width           =   480
      End
      Begin VB.Label lblmantenimiento 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "MANTENIMIENTO"
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
         Left            =   435
         TabIndex        =   146
         Top             =   975
         Width           =   1905
      End
      Begin VB.Image imgverificacion 
         Height          =   480
         Left            =   -30
         Picture         =   "frmEquipoEdicion.frx":2227
         Top             =   465
         Width           =   480
      End
      Begin VB.Label lblverificacion 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "VERIFICACION"
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
         Left            =   435
         TabIndex        =   145
         Top             =   585
         Width           =   1575
      End
      Begin VB.Image imgcalibracion 
         Height          =   480
         Left            =   -30
         Picture         =   "frmEquipoEdicion.frx":2640
         Top             =   75
         Width           =   480
      End
      Begin VB.Label lblcalibracion 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "CALIBRACION"
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
         Left            =   435
         TabIndex        =   144
         Top             =   195
         Width           =   1500
      End
   End
   Begin VB.Frame fraEstadoEquipo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Estado del Equipo"
      ForeColor       =   &H80000008&
      Height          =   795
      Left            =   6930
      TabIndex        =   140
      Top             =   10170
      Visible         =   0   'False
      Width           =   3045
      Begin VB.OptionButton optBaja 
         Caption         =   "BAJA"
         Height          =   405
         Left            =   1575
         Style           =   1  'Graphical
         TabIndex        =   142
         Top             =   225
         Width           =   1305
      End
      Begin VB.OptionButton optAlta 
         Caption         =   "ALTA"
         Height          =   375
         Left            =   330
         Style           =   1  'Graphical
         TabIndex        =   141
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdDocumentacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Documentación"
      Height          =   870
      Left            =   3015
      Picture         =   "frmEquipoEdicion.frx":2A59
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   10125
      Width           =   1440
   End
   Begin VB.CommandButton cmdImprimirFicha 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ficha Equipo"
      Height          =   855
      Left            =   10035
      Picture         =   "frmEquipoEdicion.frx":3323
      Style           =   1  'Graphical
      TabIndex        =   33
      ToolTipText     =   "Ver ficha del equipo"
      Top             =   10125
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txtIdEquipo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1305
      Locked          =   -1  'True
      MaxLength       =   11
      TabIndex        =   38
      Text            =   "123456"
      Top             =   585
      Width           =   1380
   End
   Begin VB.CommandButton cmdEtiqueta 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Etiqueta Equipo"
      Height          =   870
      Left            =   4500
      Style           =   1  'Graphical
      TabIndex        =   32
      ToolTipText     =   "Generar etiqueta"
      Top             =   10125
      Width           =   1440
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12225
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   10125
      Width           =   1050
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   9780
      Top             =   10095
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin TabDlg.SSTab tabPrincipal 
      Height          =   8805
      Left            =   45
      TabIndex        =   37
      Top             =   1245
      Width           =   13245
      _ExtentX        =   23363
      _ExtentY        =   15531
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   520
      BackColor       =   -2147483636
      ForeColor       =   -2147483641
      MouseIcon       =   "frmEquipoEdicion.frx":3BED
      TabCaption(0)   =   " Datos"
      TabPicture(0)   =   "frmEquipoEdicion.frx":3C09
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ControlPanelXP6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ControlPanelXP1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cpEstado"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cpRequiere"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "ControlPanelXP3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cpLimitaciones"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "ControlPanelXP4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "ControlPanelXP5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ControlPanelXP2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      TabCaption(1)   =   " Calibraciones "
      TabPicture(1)   =   "frmEquipoEdicion.frx":A46B
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDatos(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   " Verificaciones "
      TabPicture(2)   =   "frmEquipoEdicion.frx":10CCD
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraDatos(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   " Mantenimiento "
      TabPicture(3)   =   "frmEquipoEdicion.frx":1752F
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraDatos(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Histórico"
      TabPicture(4)   =   "frmEquipoEdicion.frx":1DD91
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraDatos(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   " Documentación/Instrucciones"
      TabPicture(5)   =   "frmEquipoEdicion.frx":245F3
      Tab(5).ControlEnabled=   0   'False
      Tab(5).ControlCount=   0
      TabCaption(6)   =   " Accesorios / Normas "
      TabPicture(6)   =   "frmEquipoEdicion.frx":2AE55
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraDatos(6)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Uso"
      TabPicture(7)   =   "frmEquipoEdicion.frx":316B7
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fraDatos(7)"
      Tab(7).ControlCount=   1
      Begin Geslab.ControlPanelXP ControlPanelXP2 
         Height          =   1875
         Left            =   6795
         TabIndex        =   118
         Top             =   6840
         Width           =   3150
         _ExtentX        =   5556
         _ExtentY        =   3307
         Caption         =   "Condiciones Ambientales"
         BackColor       =   12632256
         TextColor       =   0
         HeaderColor     =   8421504
         PanelOpen       =   0   'False
         Object.Height          =   1875
         Begin VB.CommandButton Command3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Crear Mensaje"
            Height          =   330
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   119
            Top             =   6165
            Width           =   3840
         End
         Begin VB.CheckBox chkCondicionesAmbientales 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Requiere Condiciones Ambientales"
            Height          =   240
            Left            =   180
            TabIndex        =   25
            Top             =   510
            Width           =   2820
         End
         Begin VB.Frame fraCondicionesAmbientales 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Enabled         =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1185
            Left            =   60
            TabIndex        =   120
            Top             =   510
            Visible         =   0   'False
            Width           =   3015
            Begin VB.TextBox txtHumedadMin 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1365
               MaxLength       =   100
               TabIndex        =   28
               Text            =   "0"
               Top             =   780
               Width           =   690
            End
            Begin VB.TextBox txtHumedadMax 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2175
               MaxLength       =   100
               TabIndex        =   29
               Text            =   "0"
               Top             =   780
               Width           =   690
            End
            Begin VB.TextBox txtTemperaturaMin 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1365
               MaxLength       =   100
               TabIndex        =   26
               Text            =   "0"
               Top             =   450
               Width           =   690
            End
            Begin VB.TextBox txtTemperaturaMax 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2175
               MaxLength       =   100
               TabIndex        =   27
               Text            =   "0"
               Top             =   450
               Width           =   690
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "-"
               Height          =   195
               Index           =   60
               Left            =   2085
               TabIndex        =   125
               Top             =   840
               Width           =   45
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Humedad ( % Hr)"
               Height          =   195
               Index           =   59
               Left            =   105
               TabIndex        =   124
               Top             =   825
               Width           =   1200
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "  Minimo   -   Máximo"
               Height          =   195
               Index           =   58
               Left            =   1365
               TabIndex        =   123
               Top             =   240
               Width           =   1500
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "-"
               Height          =   195
               Index           =   26
               Left            =   2085
               TabIndex        =   122
               Top             =   510
               Width           =   45
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Temperatura (º C)"
               Height          =   195
               Index           =   22
               Left            =   75
               TabIndex        =   121
               Top             =   495
               Width           =   1245
            End
         End
      End
      Begin Geslab.ControlPanelXP ControlPanelXP5 
         Height          =   3150
         Left            =   120
         TabIndex        =   126
         Top             =   405
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   5556
         Caption         =   "Datos Básicos"
         BackColor       =   12632256
         TextColor       =   0
         HeaderColor     =   8421504
         Object.Height          =   3150
         Begin VB.Frame fraDatos 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   2625
            Index           =   0
            Left            =   60
            TabIndex        =   128
            Top             =   450
            Width           =   6465
            Begin VB.TextBox txtDescripcion 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   465
               Left            =   1050
               MaxLength       =   250
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   1
               Top             =   390
               Width           =   5370
            End
            Begin VB.TextBox txtNombre 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   1050
               MaxLength       =   250
               TabIndex        =   0
               Top             =   90
               Width           =   5370
            End
            Begin VB.TextBox txtNSerie 
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
               Height          =   285
               Left            =   1050
               MaxLength       =   100
               TabIndex        =   4
               Top             =   1530
               Width           =   2145
            End
            Begin VB.TextBox txtModelo 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H00000000&
               Height          =   285
               Left            =   4290
               MaxLength       =   100
               TabIndex        =   5
               Top             =   1530
               Width           =   2115
            End
            Begin MSComCtl2.DTPicker txtFechaPuestaServicio 
               Height          =   345
               Left            =   4290
               TabIndex        =   7
               Top             =   1830
               Width           =   1395
               _ExtentX        =   2461
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
               Format          =   60293121
               CurrentDate     =   2
               MinDate         =   2
            End
            Begin MSComCtl2.DTPicker txtFechaRecepcion 
               Height          =   345
               Left            =   1050
               TabIndex        =   6
               Top             =   1830
               Width           =   1410
               _ExtentX        =   2487
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
               Format          =   60293121
               CurrentDate     =   2
               MinDate         =   2
            End
            Begin MSDataListLib.DataCombo cmbResponsable 
               Height          =   315
               Left            =   1050
               TabIndex        =   8
               Top             =   2220
               Width           =   5355
               _ExtentX        =   9446
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cmbFamilia 
               Height          =   315
               Left            =   1050
               TabIndex        =   2
               Top             =   870
               Width           =   5385
               _ExtentX        =   9499
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               Text            =   ""
            End
            Begin MSDataListLib.DataCombo cmbTipoEquipo 
               Height          =   315
               Left            =   1050
               TabIndex        =   3
               Top             =   1200
               Width           =   5385
               _ExtentX        =   9499
               _ExtentY        =   556
               _Version        =   393216
               Appearance      =   0
               Style           =   2
               Text            =   ""
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Tipo Equipo"
               Height          =   195
               Index           =   36
               Left            =   60
               TabIndex        =   137
               Top             =   1260
               Width           =   855
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nombre"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   136
               Top             =   135
               Width           =   585
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Descripción"
               Height          =   195
               Index           =   55
               Left            =   60
               TabIndex        =   135
               Top             =   495
               Width           =   840
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Area"
               Height          =   195
               Index           =   18
               Left            =   60
               TabIndex        =   134
               Top             =   915
               Width           =   330
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Nº Serie"
               Height          =   195
               Index           =   3
               Left            =   60
               TabIndex        =   133
               Top             =   1605
               Width           =   585
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "F.Recepción"
               Height          =   195
               Index           =   12
               Left            =   60
               TabIndex        =   132
               Top             =   1920
               Width           =   915
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Modelo"
               Height          =   195
               Index           =   19
               Left            =   3660
               TabIndex        =   131
               Top             =   1590
               Width           =   525
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "F. Puesta Servicio"
               Height          =   195
               Index           =   57
               Left            =   2910
               TabIndex        =   130
               Top             =   1920
               Width           =   1290
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Responsable"
               Height          =   195
               Index           =   70
               Left            =   60
               TabIndex        =   129
               Top             =   2265
               Width           =   930
            End
         End
         Begin VB.CommandButton Command4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Crear Mensaje"
            Height          =   330
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   127
            Top             =   6165
            Width           =   3840
         End
      End
      Begin Geslab.ControlPanelXP ControlPanelXP4 
         Height          =   2580
         Left            =   120
         TabIndex        =   103
         Top             =   3555
         Width           =   6630
         _ExtentX        =   11695
         _ExtentY        =   4551
         Caption         =   "Otros Datos"
         BackColor       =   12632256
         TextColor       =   0
         HeaderColor     =   8421504
         Object.Height          =   2580
         Begin VB.CheckBox chkCP 
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
            Height          =   240
            Left            =   6075
            TabIndex        =   241
            Top             =   1980
            Width           =   315
         End
         Begin VB.CheckBox chkMTL 
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
            Height          =   240
            Left            =   6075
            TabIndex        =   239
            Top             =   1755
            Width           =   315
         End
         Begin VB.CheckBox chkCritico 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "CRITICO"
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
            Height          =   285
            Left            =   90
            TabIndex        =   233
            Top             =   1845
            Width           =   1230
         End
         Begin VB.CheckBox chkEnac 
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
            Height          =   240
            Left            =   6075
            TabIndex        =   232
            Top             =   2250
            Width           =   315
         End
         Begin VB.CheckBox chkNadcap 
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
            Height          =   240
            Left            =   6075
            TabIndex        =   230
            Top             =   1530
            Width           =   315
         End
         Begin VB.CheckBox chkPrioritario 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "PRIORITARIO"
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
            Height          =   285
            Left            =   90
            TabIndex        =   206
            Top             =   2160
            Visible         =   0   'False
            Width           =   1860
         End
         Begin XtremeSuiteControls.PushButton cmdVerSituacionEnPlano 
            Height          =   435
            Left            =   135
            TabIndex        =   168
            Top             =   2115
            Visible         =   0   'False
            Width           =   2490
            _Version        =   851970
            _ExtentX        =   4392
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Ver en el Plano"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":37F19
         End
         Begin MSDataListLib.DataCombo cmbProveedor 
            Height          =   315
            Left            =   1140
            TabIndex        =   10
            Top             =   780
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin VB.TextBox txtFabricante 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   1140
            MaxLength       =   100
            TabIndex        =   9
            Top             =   480
            Width           =   5235
         End
         Begin VB.CommandButton Command2 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Crear Mensaje"
            Height          =   330
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   104
            Top             =   6165
            Width           =   3840
         End
         Begin VB.CheckBox chkMostrarEnPlano 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Mostrar en el Plano de Equipos"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   135
            TabIndex        =   14
            Top             =   1800
            Value           =   1  'Checked
            Visible         =   0   'False
            Width           =   2685
         End
         Begin VB.CheckBox chkes_accesorio 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "PUEDE SER ACCESORIO"
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
            Height          =   285
            Left            =   135
            TabIndex        =   15
            Top             =   1530
            Width           =   3255
         End
         Begin MSDataListLib.DataCombo cmbSituacion 
            Height          =   315
            Left            =   1140
            TabIndex        =   11
            Top             =   1125
            Width           =   5235
            _ExtentX        =   9234
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "- CP"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   39
            Left            =   5100
            TabIndex        =   242
            Top             =   1980
            Width           =   300
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "- MTL"
            ForeColor       =   &H00FF0000&
            Height          =   195
            Index           =   38
            Left            =   5100
            TabIndex        =   240
            Top             =   1755
            Width           =   420
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "ENAC"
            Height          =   195
            Index           =   31
            Left            =   5100
            TabIndex        =   231
            Top             =   2265
            Width           =   435
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fabricante"
            Height          =   195
            Index           =   9
            Left            =   150
            TabIndex        =   139
            Top             =   510
            Width           =   750
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   35
            Left            =   150
            TabIndex        =   138
            Top             =   825
            Width           =   735
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "NADCAP"
            Height          =   195
            Index           =   34
            Left            =   5100
            TabIndex        =   106
            Top             =   1530
            Width           =   660
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Localización"
            Height          =   195
            Index           =   8
            Left            =   150
            TabIndex        =   105
            Top             =   1215
            Width           =   885
         End
      End
      Begin VB.Frame fraDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         ForeColor       =   &H80000008&
         Height          =   8370
         Index           =   7
         Left            =   -74970
         TabIndex        =   94
         Top             =   330
         Width           =   13185
         Begin VB.TextBox txtuso 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
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
            Height          =   360
            Index           =   3
            Left            =   11160
            MaxLength       =   250
            TabIndex        =   207
            Top             =   810
            Width           =   615
         End
         Begin VB.TextBox txtuso 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Index           =   1
            Left            =   7230
            MaxLength       =   250
            TabIndex        =   101
            Top             =   120
            Visible         =   0   'False
            Width           =   1110
         End
         Begin VB.TextBox txtuso 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Index           =   2
            Left            =   11040
            Locked          =   -1  'True
            MaxLength       =   250
            TabIndex        =   99
            Top             =   120
            Width           =   1110
         End
         Begin VB.TextBox txtuso 
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
            ForeColor       =   &H00000000&
            Height          =   360
            Index           =   0
            Left            =   2070
            MaxLength       =   250
            TabIndex        =   96
            Top             =   120
            Width           =   1110
         End
         Begin MSComctlLib.ListView listaUsos 
            Height          =   7140
            Left            =   60
            TabIndex        =   95
            Top             =   1200
            Width           =   13065
            _ExtentX        =   23045
            _ExtentY        =   12594
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
         Begin XtremeSuiteControls.PushButton cmdVerificacionContador 
            Height          =   255
            Left            =   3810
            TabIndex        =   162
            Top             =   540
            Visible         =   0   'False
            Width           =   4515
            _Version        =   851970
            _ExtentX        =   7964
            _ExtentY        =   450
            _StockProps     =   79
            Caption         =   "Crear verificación para poner contador a cero"
            Appearance      =   4
            Picture         =   "frmEquipoEdicion.frx":3E77B
         End
         Begin XtremeSuiteControls.PushButton cmdModificarUsos 
            Height          =   345
            Left            =   11835
            TabIndex        =   209
            Top             =   810
            Width           =   1230
            _Version        =   851970
            _ExtentX        =   2170
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Modificar"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":44FDD
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "Usos en la Muestra"
            Height          =   195
            Index           =   25
            Left            =   9720
            TabIndex        =   208
            Top             =   900
            Width           =   1365
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Número de Usos desde la última Verificación"
            Height          =   195
            Index           =   46
            Left            =   3810
            TabIndex        =   102
            Top             =   210
            Visible         =   0   'False
            Width           =   3150
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Número Total de Usos"
            Height          =   195
            Index           =   45
            Left            =   9090
            TabIndex        =   100
            Top             =   210
            Width           =   1590
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Listado de Muestras donde se utiliza el equipo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   60
            TabIndex        =   98
            Top             =   840
            Width           =   13080
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Número Máximo de Usos"
            Height          =   195
            Index           =   44
            Left            =   120
            TabIndex        =   97
            Top             =   210
            Width           =   1770
         End
      End
      Begin VB.Frame fraDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         ForeColor       =   &H80000008&
         Height          =   8370
         Index           =   4
         Left            =   -75000
         TabIndex        =   83
         Top             =   330
         Width           =   13185
         Begin MSFlexGridLib.MSFlexGrid grdEventos 
            Height          =   7395
            Left            =   30
            TabIndex        =   84
            Top             =   480
            Width           =   13125
            _ExtentX        =   23151
            _ExtentY        =   13044
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
         End
         Begin XtremeSuiteControls.PushButton cmdAnadirEvento 
            Height          =   435
            Left            =   8370
            TabIndex        =   158
            Top             =   7920
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Añadir Evento"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":4B83F
         End
         Begin XtremeSuiteControls.PushButton cmdEliminarEvento 
            Height          =   435
            Left            =   10800
            TabIndex        =   159
            Top             =   7920
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Eliminar Evento"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":520A1
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            Caption         =   "Histórico del equipo"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   30
            TabIndex        =   160
            Top             =   120
            Width           =   13080
         End
      End
      Begin VB.Frame fraDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         ForeColor       =   &H80000008&
         Height          =   8385
         Index           =   3
         Left            =   -74955
         TabIndex        =   48
         Top             =   360
         Width           =   13140
         Begin VB.ListBox lstPlanes 
            Appearance      =   0  'Flat
            Height          =   705
            Left            =   645
            Style           =   1  'Checkbox
            TabIndex        =   87
            Top             =   405
            Width           =   5805
         End
         Begin VB.CommandButton cmdEliminarPlan 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   6165
            Picture         =   "frmEquipoEdicion.frx":58903
            Style           =   1  'Graphical
            TabIndex        =   86
            ToolTipText     =   "Eliminar accesorio"
            Top             =   45
            Width           =   315
         End
         Begin VB.CommandButton cmdAnadirPlan 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   315
            Left            =   5835
            Picture         =   "frmEquipoEdicion.frx":58A97
            Style           =   1  'Graphical
            TabIndex        =   85
            ToolTipText     =   "Añadir accesorio"
            Top             =   45
            Width           =   315
         End
         Begin VB.TextBox txtMtoFechaProxima_info 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8460
            Locked          =   -1  'True
            TabIndex        =   82
            Top             =   480
            Width           =   1395
         End
         Begin VB.ComboBox cmbMtoTipo 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmEquipoEdicion.frx":58CBC
            Left            =   10920
            List            =   "frmEquipoEdicion.frx":58CC6
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   1290
            Width           =   2235
         End
         Begin MSFlexGridLib.MSFlexGrid grdMantenimientos 
            Height          =   6165
            Left            =   30
            TabIndex        =   60
            Top             =   1695
            Width           =   13125
            _ExtentX        =   23151
            _ExtentY        =   10874
            _Version        =   393216
            FixedCols       =   0
            BackColor       =   16777215
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cmbMtoPlan 
            Height          =   315
            Left            =   645
            TabIndex        =   63
            Top             =   45
            Width           =   5160
            _ExtentX        =   9102
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker txtMtoFechaProxima 
            Height          =   315
            Left            =   9900
            TabIndex        =   78
            Top             =   480
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
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
            Format          =   60293121
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo cmbMtoResponsable 
            Height          =   315
            Left            =   8460
            TabIndex        =   79
            Top             =   120
            Width           =   4680
            _ExtentX        =   8255
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin XtremeSuiteControls.PushButton cmdAnadirMto 
            Height          =   435
            Left            =   8370
            TabIndex        =   154
            Top             =   7890
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Añadir Mantenimiento"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":58CE1
         End
         Begin XtremeSuiteControls.PushButton cmdEliminarMto 
            Height          =   435
            Left            =   10755
            TabIndex        =   155
            Top             =   7890
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Eliminar Mantenimiento"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":5F543
         End
         Begin XtremeSuiteControls.PushButton cmdGenerarFechasMto 
            Height          =   435
            Left            =   30
            TabIndex        =   157
            Top             =   7875
            Width           =   3555
            _Version        =   851970
            _ExtentX        =   6271
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Generar Fechas de Mantenimiento"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":65DA5
         End
         Begin XtremeSuiteControls.PushButton PushButton2 
            Height          =   435
            Left            =   0
            TabIndex        =   221
            Top             =   0
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Reabrir"
            Appearance      =   5
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackColor       =   &H00E0E0E0&
            Caption         =   "Listado de Mantenimientos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   2
            Left            =   60
            TabIndex        =   156
            Top             =   1290
            Width           =   13080
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha Prox."
            Height          =   195
            Index           =   24
            Left            =   7470
            TabIndex        =   81
            Top             =   555
            Width           =   855
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Resp. Intern."
            Height          =   195
            Index           =   28
            Left            =   7470
            TabIndex        =   80
            Top             =   195
            Width           =   915
         End
         Begin VB.Image imgNuevoPlanMto 
            Height          =   285
            Left            =   6600
            Picture         =   "frmEquipoEdicion.frx":6C607
            Stretch         =   -1  'True
            Top             =   480
            Width           =   285
         End
         Begin VB.Image imgBuscarPlanMto 
            Height          =   285
            Left            =   6930
            Picture         =   "frmEquipoEdicion.frx":6CED1
            Stretch         =   -1  'True
            Top             =   480
            Width           =   285
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Planes"
            Height          =   195
            Index           =   10
            Left            =   90
            TabIndex        =   61
            Top             =   90
            Width           =   480
         End
      End
      Begin VB.Frame fraDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         ForeColor       =   &H80000008&
         Height          =   8340
         Index           =   2
         Left            =   -74955
         TabIndex        =   47
         Top             =   405
         Width           =   13170
         Begin VB.ComboBox cmbTipoVerificaciones 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmEquipoEdicion.frx":6D79B
            Left            =   10590
            List            =   "frmEquipoEdicion.frx":6D7A8
            Style           =   2  'Dropdown List
            TabIndex        =   89
            Top             =   1290
            Width           =   2535
         End
         Begin pryCombo.miCombo cmbProcedimientoVer 
            Height          =   330
            Left            =   1080
            TabIndex        =   77
            Top             =   780
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   582
         End
         Begin VB.TextBox txtVerificacionFechaProxima_info 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8370
            Locked          =   -1  'True
            TabIndex        =   75
            Top             =   420
            Width           =   1395
         End
         Begin MSFlexGridLib.MSFlexGrid grdVerificaciones 
            Height          =   6060
            Left            =   30
            TabIndex        =   55
            Top             =   1740
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   10689
            _Version        =   393216
            FixedCols       =   0
            BackColor       =   16777215
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo cmbVerificacionPeriod 
            Height          =   315
            Left            =   1080
            TabIndex        =   56
            Top             =   60
            Width           =   5210
            _ExtentX        =   9181
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbVerificacionTipo 
            Height          =   315
            Left            =   1080
            TabIndex        =   57
            Top             =   420
            Width           =   5210
            _ExtentX        =   9181
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker txtVerificacionFechaProxima 
            Height          =   315
            Left            =   9810
            TabIndex        =   70
            Top             =   405
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
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
            Format          =   60293121
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo cmbVerificacionResponsable 
            Height          =   315
            Left            =   8370
            TabIndex        =   72
            Top             =   60
            Width           =   4710
            _ExtentX        =   8308
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin XtremeSuiteControls.PushButton cmdAnadirVerificacion 
            Height          =   435
            Left            =   8370
            TabIndex        =   152
            Top             =   7830
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Añadir Verificación"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":6D7C8
         End
         Begin XtremeSuiteControls.PushButton cmdEliminarVerificacion 
            Height          =   435
            Left            =   10800
            TabIndex        =   153
            Top             =   7830
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Eliminar Verificación"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":7402A
         End
         Begin XtremeSuiteControls.PushButton cmdReabrir 
            Height          =   435
            Left            =   5940
            TabIndex        =   219
            Top             =   7830
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Reabrir"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":7A88C
         End
         Begin XtremeSuiteControls.PushButton cmbduplicarVerificacion 
            Height          =   435
            Left            =   45
            TabIndex        =   243
            Top             =   7830
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Duplicar"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":810EE
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Listado de Verificaciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   1
            Left            =   60
            TabIndex        =   149
            Top             =   1290
            Width           =   13050
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Procedim."
            Height          =   195
            Index           =   40
            Left            =   120
            TabIndex        =   74
            Top             =   915
            Width           =   705
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Resp. Intern."
            Height          =   195
            Index           =   20
            Left            =   7410
            TabIndex        =   73
            Top             =   120
            Width           =   915
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha Próx."
            Height          =   195
            Index           =   15
            Left            =   7410
            TabIndex        =   71
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Periodicidad"
            Height          =   195
            Index           =   21
            Left            =   120
            TabIndex        =   59
            Top             =   150
            Width           =   870
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo"
            Height          =   195
            Index           =   23
            Left            =   120
            TabIndex        =   58
            Top             =   555
            Width           =   315
         End
      End
      Begin VB.Frame fraDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         ForeColor       =   &H80000008&
         Height          =   8340
         Index           =   1
         Left            =   -74955
         TabIndex        =   46
         Top             =   360
         Width           =   13140
         Begin XtremeSuiteControls.PushButton cmdAnadirCalibracion 
            Height          =   435
            Left            =   8355
            TabIndex        =   150
            Top             =   7845
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Añadir Calibración"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":87950
         End
         Begin VB.ComboBox cmbTipoCalibraciones 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            ItemData        =   "frmEquipoEdicion.frx":8E1B2
            Left            =   10575
            List            =   "frmEquipoEdicion.frx":8E1BF
            Style           =   2  'Dropdown List
            TabIndex        =   90
            Top             =   1290
            Width           =   2565
         End
         Begin VB.TextBox txtCalibracionFechaProxima_info 
            Appearance      =   0  'Flat
            Height          =   315
            Left            =   8400
            Locked          =   -1  'True
            TabIndex        =   68
            Top             =   420
            Width           =   1395
         End
         Begin MSDataListLib.DataCombo cmbCalibracionPeriod 
            Height          =   315
            Left            =   1080
            TabIndex        =   50
            Top             =   60
            Width           =   5200
            _ExtentX        =   9181
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cmbCalibracionTipo 
            Height          =   315
            Left            =   1080
            TabIndex        =   51
            Top             =   420
            Width           =   5200
            _ExtentX        =   9181
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin MSComCtl2.DTPicker txtCalibracionFechaProxima 
            Height          =   315
            Left            =   9840
            TabIndex        =   64
            Top             =   420
            Visible         =   0   'False
            Width           =   1560
            _ExtentX        =   2752
            _ExtentY        =   556
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
            Format          =   60293121
            CurrentDate     =   2
            MinDate         =   2
         End
         Begin MSDataListLib.DataCombo cmbCalibracionResponsable 
            Height          =   315
            Left            =   8400
            TabIndex        =   65
            Top             =   60
            Width           =   4710
            _ExtentX        =   8308
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            Style           =   2
            Text            =   ""
         End
         Begin pryCombo.miCombo cmbProcedimientoCal 
            Height          =   330
            Left            =   1080
            TabIndex        =   76
            Top             =   780
            Width           =   6255
            _ExtentX        =   11033
            _ExtentY        =   582
         End
         Begin XtremeSuiteControls.PushButton cmdEliminarCalibracion 
            Height          =   435
            Left            =   10740
            TabIndex        =   151
            Top             =   7845
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Eliminar Calibración"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":8E1DF
         End
         Begin XtremeSuiteControls.PushButton cmdReabrirCalibra 
            Height          =   435
            Left            =   5970
            TabIndex        =   220
            Top             =   7845
            Width           =   2355
            _Version        =   851970
            _ExtentX        =   4154
            _ExtentY        =   767
            _StockProps     =   79
            Caption         =   "Reabrir"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":94A41
         End
         Begin MSFlexGridLib.MSFlexGrid grdCalibraciones 
            Height          =   6060
            Left            =   30
            TabIndex        =   49
            Top             =   1740
            Width           =   13065
            _ExtentX        =   23045
            _ExtentY        =   10689
            _Version        =   393216
            FixedCols       =   0
            FocusRect       =   2
            HighLight       =   2
            FillStyle       =   1
            GridLinesFixed  =   1
            SelectionMode   =   1
            AllowUserResizing=   1
            Appearance      =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   9.75
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
            Caption         =   "Procedim."
            Height          =   195
            Index           =   37
            Left            =   120
            TabIndex        =   69
            Top             =   870
            Width           =   705
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Fecha Próx."
            Height          =   195
            Index           =   14
            Left            =   7440
            TabIndex        =   67
            Top             =   495
            Width           =   855
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Resp. Intern."
            Height          =   195
            Index           =   5
            Left            =   7440
            TabIndex        =   66
            Top             =   135
            Width           =   915
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "Listado de Calibraciones"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Index           =   0
            Left            =   60
            TabIndex        =   54
            Top             =   1290
            Width           =   13020
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Periodicidad"
            Height          =   195
            Index           =   11
            Left            =   120
            TabIndex        =   53
            Top             =   150
            Width           =   870
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Tipo"
            Height          =   195
            Index           =   6
            Left            =   120
            TabIndex        =   52
            Top             =   510
            Width           =   315
         End
      End
      Begin VB.Frame fraDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         ForeColor       =   &H80000008&
         Height          =   8370
         Index           =   6
         Left            =   -75000
         TabIndex        =   40
         Top             =   360
         Width           =   13215
         Begin VB.Frame fraAccesorios 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   8295
            Left            =   15
            TabIndex        =   43
            Top             =   0
            Width           =   6000
            Begin MSFlexGridLib.MSFlexGrid grdAccesorios 
               Height          =   7245
               Left            =   60
               TabIndex        =   44
               Top             =   450
               Width           =   5835
               _ExtentX        =   10292
               _ExtentY        =   12779
               _Version        =   393216
               FixedCols       =   0
               BackColor       =   12640511
               FocusRect       =   2
               HighLight       =   2
               FillStyle       =   1
               GridLinesFixed  =   1
               SelectionMode   =   1
               AllowUserResizing=   1
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
            End
            Begin XtremeSuiteControls.PushButton cmdAccesoriosAdd 
               Height          =   435
               Left            =   2070
               TabIndex        =   163
               Top             =   7740
               Width           =   1860
               _Version        =   851970
               _ExtentX        =   3281
               _ExtentY        =   767
               _StockProps     =   79
               Caption         =   "Añadir"
               Appearance      =   5
               Picture         =   "frmEquipoEdicion.frx":9B2A3
            End
            Begin XtremeSuiteControls.PushButton cmdAccesoriosDelete 
               Height          =   435
               Left            =   4005
               TabIndex        =   164
               Top             =   7740
               Width           =   1905
               _Version        =   851970
               _ExtentX        =   3360
               _ExtentY        =   767
               _StockProps     =   79
               Caption         =   "Eliminar"
               Appearance      =   5
               Picture         =   "frmEquipoEdicion.frx":A1B05
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "Accesorios"
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
               Left            =   60
               TabIndex        =   45
               Top             =   135
               Width           =   5850
            End
         End
         Begin VB.Frame Frame12 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            ForeColor       =   &H80000008&
            Height          =   8295
            Left            =   6030
            TabIndex        =   41
            Top             =   0
            Width           =   7125
            Begin VB.OptionButton optiponorma 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Normas"
               Height          =   195
               Index           =   1
               Left            =   1485
               TabIndex        =   92
               Top             =   7830
               Width           =   1005
            End
            Begin VB.OptionButton optiponorma 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Documentos"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   91
               Top             =   7830
               Value           =   -1  'True
               Width           =   1275
            End
            Begin pryCombo.miCombo cmbNormas 
               Height          =   330
               Left            =   90
               TabIndex        =   88
               Top             =   7335
               Visible         =   0   'False
               Width           =   7020
               _ExtentX        =   12383
               _ExtentY        =   582
            End
            Begin MSComctlLib.ListView grdNormas 
               Height          =   6765
               Left            =   90
               TabIndex        =   93
               Top             =   495
               Width           =   6960
               _ExtentX        =   12277
               _ExtentY        =   11933
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
               Left            =   90
               TabIndex        =   161
               Top             =   7335
               Width           =   6975
               _ExtentX        =   12303
               _ExtentY        =   582
            End
            Begin XtremeSuiteControls.PushButton cmdAnadirNorma 
               Height          =   435
               Left            =   3870
               TabIndex        =   165
               Top             =   7740
               Width           =   1545
               _Version        =   851970
               _ExtentX        =   2725
               _ExtentY        =   767
               _StockProps     =   79
               Caption         =   "Añadir"
               Appearance      =   5
               Picture         =   "frmEquipoEdicion.frx":A8367
            End
            Begin XtremeSuiteControls.PushButton cmdEliminarNorma 
               Height          =   435
               Left            =   5445
               TabIndex        =   166
               Top             =   7740
               Width           =   1590
               _Version        =   851970
               _ExtentX        =   2805
               _ExtentY        =   767
               _StockProps     =   79
               Caption         =   "Eliminar"
               Appearance      =   5
               Picture         =   "frmEquipoEdicion.frx":AEBC9
            End
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               Caption         =   "Conforme a las procedimientos / normas"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   60
               TabIndex        =   42
               Top             =   150
               Width           =   7005
            End
         End
      End
      Begin Geslab.ControlPanelXP cpLimitaciones 
         Height          =   2505
         Left            =   6795
         TabIndex        =   107
         Top             =   4275
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   4419
         Caption         =   "Limitaciones de Uso"
         BackColor       =   12632256
         TextColor       =   0
         HeaderColor     =   8421504
         PanelOpen       =   0   'False
         Object.Height          =   2505
         Begin VB.CommandButton cmdEliminarLimitacion 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5955
            Picture         =   "frmEquipoEdicion.frx":B542B
            Style           =   1  'Graphical
            TabIndex        =   19
            ToolTipText     =   "Eliminar accesorio"
            Top             =   2070
            Width           =   285
         End
         Begin VB.CommandButton cmdAnadirLimitacion 
            BackColor       =   &H00E0E0E0&
            Enabled         =   0   'False
            Height          =   285
            Left            =   5625
            Picture         =   "frmEquipoEdicion.frx":B55BF
            Style           =   1  'Graphical
            TabIndex        =   18
            ToolTipText     =   "Añadir accesorio"
            Top             =   2070
            Width           =   285
         End
         Begin VB.ListBox lstLimitacionesUso 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1590
            ItemData        =   "frmEquipoEdicion.frx":B57E4
            Left            =   90
            List            =   "frmEquipoEdicion.frx":B57E6
            TabIndex        =   16
            Top             =   450
            Width           =   6135
         End
         Begin VB.TextBox txtLimitacionesUso 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            ForeColor       =   &H00000000&
            Height          =   285
            Left            =   90
            MaxLength       =   100
            TabIndex        =   17
            Text            =   "(Debe guardar los datos del equipo para añadir Limitaciones de Uso)"
            Top             =   2070
            Width           =   5490
         End
         Begin VB.CommandButton cmdtodos 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Crear Mensaje"
            Height          =   330
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   108
            Top             =   6165
            Width           =   3840
         End
      End
      Begin Geslab.ControlPanelXP ControlPanelXP3 
         Height          =   1875
         Left            =   9945
         TabIndex        =   109
         Top             =   6840
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   3307
         Caption         =   "Rangos de Trabajo"
         BackColor       =   12632256
         TextColor       =   0
         HeaderColor     =   8421504
         PanelOpen       =   0   'False
         Object.Height          =   1875
         Begin VB.Frame Frame8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1425
            Left            =   30
            TabIndex        =   111
            Top             =   390
            Width           =   3165
            Begin VB.TextBox txtRangoMedidaMax 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1860
               MaxLength       =   100
               TabIndex        =   21
               Top             =   360
               Width           =   690
            End
            Begin VB.TextBox txtRangoMedidaMin 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1050
               MaxLength       =   100
               TabIndex        =   20
               Top             =   360
               Width           =   690
            End
            Begin VB.TextBox txtRangoTrabajoMax 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1860
               MaxLength       =   100
               TabIndex        =   23
               Top             =   690
               Width           =   690
            End
            Begin VB.TextBox txtRangoTrabajoMin 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1050
               MaxLength       =   100
               TabIndex        =   22
               Top             =   690
               Width           =   690
            End
            Begin pryCombo.miCombo cmbUnidades 
               Height          =   330
               Left            =   1050
               TabIndex        =   24
               Top             =   1020
               Width           =   2265
               _ExtentX        =   3995
               _ExtentY        =   582
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Unidades"
               Height          =   195
               Index           =   66
               Left            =   60
               TabIndex        =   117
               Top             =   1110
               Width           =   675
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "R. Medida"
               Height          =   195
               Index           =   65
               Left            =   60
               TabIndex        =   116
               Top             =   405
               Width           =   735
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "-"
               Height          =   195
               Index           =   64
               Left            =   1770
               TabIndex        =   115
               Top             =   420
               Width           =   45
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "  Minimo   -   Máximo"
               Height          =   195
               Index           =   63
               Left            =   1050
               TabIndex        =   114
               Top             =   150
               Width           =   1500
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "R. Trabajo"
               Height          =   195
               Index           =   62
               Left            =   60
               TabIndex        =   113
               Top             =   735
               Width           =   750
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "-"
               Height          =   195
               Index           =   61
               Left            =   1770
               TabIndex        =   112
               Top             =   750
               Width           =   45
            End
         End
         Begin VB.CommandButton Command1 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Crear Mensaje"
            Height          =   330
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   110
            Top             =   6165
            Width           =   3840
         End
      End
      Begin Geslab.ControlPanelXP cpRequiere 
         Height          =   930
         Left            =   120
         TabIndex        =   169
         Top             =   6165
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   1640
         Caption         =   "El equipo requiere"
         BackColor       =   12632256
         TextColor       =   0
         HeaderColor     =   8421504
         Object.Height          =   930
         Begin VB.CheckBox chkCon_Verificacion 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "VERIFICACIÓN"
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
            Left            =   2160
            TabIndex        =   173
            Top             =   540
            Width           =   1935
         End
         Begin VB.CheckBox chkCon_Mantenimiento 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "MANTENIMIENTO"
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
            Left            =   4185
            TabIndex        =   172
            Top             =   540
            Width           =   2235
         End
         Begin VB.CheckBox chkCon_Calibracion 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "CALIBRACIÓN"
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
            Left            =   135
            TabIndex        =   171
            Top             =   540
            Width           =   1845
         End
         Begin VB.CommandButton Command5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Crear Mensaje"
            Height          =   330
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   170
            Top             =   6165
            Width           =   3840
         End
      End
      Begin Geslab.ControlPanelXP cpEstado 
         Height          =   2235
         Left            =   6795
         TabIndex        =   180
         Top             =   1980
         Width           =   6360
         _ExtentX        =   11218
         _ExtentY        =   3942
         Caption         =   "Estado Actual del Equipo"
         BackColor       =   12632256
         TextColor       =   0
         HeaderColor     =   8421504
         Object.Height          =   2235
         Begin VB.Frame frmEstadoMantenimiento 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   135
            TabIndex        =   196
            Top             =   1530
            Width           =   6090
            Begin VB.CommandButton cmdMantenimientoMostrar 
               BackColor       =   &H00E0E0E0&
               Height          =   360
               Left            =   5550
               Picture         =   "frmEquipoEdicion.frx":B57E8
               Style           =   1  'Graphical
               TabIndex        =   197
               ToolTipText     =   "Mostrar Mantenimiento"
               Top             =   0
               Width           =   405
            End
            Begin MSComCtl2.DTPicker factmantenimiento 
               Height          =   345
               Left            =   1665
               TabIndex        =   198
               Top             =   0
               Width           =   1410
               _ExtentX        =   2487
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
               Format          =   60293121
               CurrentDate     =   2
               MinDate         =   2
            End
            Begin MSComCtl2.DTPicker fproxmantenimiento 
               Height          =   345
               Left            =   3510
               TabIndex        =   199
               Top             =   0
               Width           =   1410
               _ExtentX        =   2487
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
               Format          =   60293121
               CurrentDate     =   2
               MinDate         =   2
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Mantenimiento"
               Height          =   195
               Index           =   17
               Left            =   135
               TabIndex        =   200
               Top             =   90
               Width           =   1035
            End
         End
         Begin VB.Frame frmEstadoVerificacion 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   420
            Left            =   135
            TabIndex        =   190
            Top             =   1080
            Width           =   6090
            Begin VB.CommandButton cmdVerificacionMostrar 
               BackColor       =   &H00E0E0E0&
               Height          =   360
               Left            =   5550
               Picture         =   "frmEquipoEdicion.frx":B5A3D
               Style           =   1  'Graphical
               TabIndex        =   192
               ToolTipText     =   "Mostrar Verificación"
               Top             =   45
               Width           =   405
            End
            Begin VB.CommandButton cmdAdjuntarHojaVer 
               BackColor       =   &H00E0E0E0&
               Height          =   360
               Left            =   5130
               Picture         =   "frmEquipoEdicion.frx":B5C92
               Style           =   1  'Graphical
               TabIndex        =   191
               ToolTipText     =   "Mostrar certificado de Verificación"
               Top             =   45
               Width           =   405
            End
            Begin MSComCtl2.DTPicker factverificacion 
               Height          =   345
               Left            =   1665
               TabIndex        =   193
               Top             =   45
               Width           =   1410
               _ExtentX        =   2487
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
               Format          =   60293121
               CurrentDate     =   2
               MinDate         =   2
            End
            Begin MSComCtl2.DTPicker fproxverificacion 
               Height          =   345
               Left            =   3510
               TabIndex        =   194
               Top             =   45
               Width           =   1410
               _ExtentX        =   2487
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
               Format          =   60293121
               CurrentDate     =   2
               MinDate         =   2
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Verificación"
               Height          =   195
               Index           =   16
               Left            =   135
               TabIndex        =   195
               Top             =   135
               Width           =   825
            End
         End
         Begin VB.Frame frmEstadoCalibracion 
            BackColor       =   &H00C0C0C0&
            BorderStyle     =   0  'None
            Height          =   420
            Left            =   135
            TabIndex        =   182
            Top             =   675
            Width           =   6090
            Begin VB.CommandButton cmdAdjuntarHojaCal 
               BackColor       =   &H00E0E0E0&
               Height          =   360
               Left            =   5130
               Picture         =   "frmEquipoEdicion.frx":B5F03
               Style           =   1  'Graphical
               TabIndex        =   185
               ToolTipText     =   "Mostrar certificado de Calibración"
               Top             =   45
               Width           =   405
            End
            Begin VB.CommandButton cmdCalibracionMostrar 
               BackColor       =   &H00E0E0E0&
               Height          =   360
               Left            =   5550
               Picture         =   "frmEquipoEdicion.frx":B6174
               Style           =   1  'Graphical
               TabIndex        =   184
               ToolTipText     =   "Mostrar Calibración"
               Top             =   45
               Width           =   405
            End
            Begin MSComCtl2.DTPicker factcalibracion 
               Height          =   345
               Left            =   1665
               TabIndex        =   186
               Top             =   45
               Width           =   1410
               _ExtentX        =   2487
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
               Format          =   60293121
               CurrentDate     =   2
               MinDate         =   2
            End
            Begin MSComCtl2.DTPicker fproxcalibracion 
               Height          =   345
               Left            =   3510
               TabIndex        =   187
               Top             =   45
               Width           =   1410
               _ExtentX        =   2487
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
               Format          =   60293121
               CurrentDate     =   2
               MinDate         =   2
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Calibración"
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   183
               Top             =   135
               Width           =   780
            End
         End
         Begin VB.CommandButton Command8 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Crear Mensaje"
            Height          =   330
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   181
            Top             =   6165
            Width           =   3840
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Próxima"
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
            Left            =   3960
            TabIndex        =   189
            Top             =   450
            Width           =   675
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Actual"
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
            Left            =   2160
            TabIndex        =   188
            Top             =   450
            Width           =   555
         End
      End
      Begin Geslab.ControlPanelXP ControlPanelXP1 
         Height          =   1560
         Left            =   6795
         TabIndex        =   213
         Top             =   405
         Width           =   6345
         _ExtentX        =   11192
         _ExtentY        =   2752
         Caption         =   "Equipo de Cliente"
         BackColor       =   12632256
         TextColor       =   0
         HeaderColor     =   8421504
         Object.Height          =   1560
         Begin pryCombo.miCombo cmbTrazmet 
            Height          =   330
            Left            =   810
            TabIndex        =   234
            Top             =   1170
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   582
         End
         Begin VB.CheckBox chkInSitu 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "IN-SITU"
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
            Left            =   2655
            TabIndex        =   229
            Top             =   855
            Width           =   1080
         End
         Begin XtremeSuiteControls.PushButton cmdMetrol 
            Height          =   345
            Left            =   3915
            TabIndex        =   228
            Top             =   810
            Width           =   2310
            _Version        =   851970
            _ExtentX        =   4075
            _ExtentY        =   609
            _StockProps     =   79
            Caption         =   "Capturar Datos del Metrol"
            Appearance      =   5
            Picture         =   "frmEquipoEdicion.frx":B63C9
         End
         Begin VB.TextBox txtNumeroCliente 
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
            Height          =   330
            Left            =   810
            MaxLength       =   100
            TabIndex        =   13
            Top             =   810
            Width           =   1605
         End
         Begin pryCombo.miCombo cmbCliente 
            Height          =   330
            Left            =   810
            TabIndex        =   12
            Top             =   450
            Width           =   5445
            _ExtentX        =   9604
            _ExtentY        =   582
         End
         Begin VB.CommandButton Command6 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Crear Mensaje"
            Height          =   330
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   214
            Top             =   6165
            Width           =   3840
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Trazmet"
            Height          =   195
            Index           =   32
            Left            =   45
            TabIndex        =   235
            Top             =   1215
            Width           =   570
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "NºEquipo"
            Height          =   195
            Index           =   30
            Left            =   45
            TabIndex        =   216
            Top             =   900
            Width           =   675
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cliente"
            Height          =   195
            Index           =   29
            Left            =   45
            TabIndex        =   215
            Top             =   495
            Width           =   480
         End
      End
      Begin Geslab.ControlPanelXP ControlPanelXP6 
         Height          =   1605
         Left            =   120
         TabIndex        =   225
         Top             =   7110
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   2831
         Caption         =   "Observaciones"
         BackColor       =   12632256
         TextColor       =   0
         HeaderColor     =   8421504
         Object.Height          =   1605
         Begin VB.TextBox txtObservaciones 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H00000000&
            Height          =   1050
            Left            =   45
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   227
            Top             =   450
            Width           =   6495
         End
         Begin VB.CommandButton Command7 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Crear Mensaje"
            Height          =   330
            Left            =   45
            Style           =   1  'Graphical
            TabIndex        =   226
            Top             =   6165
            Width           =   3840
         End
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Guardar"
      Height          =   870
      Left            =   11160
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   10125
      Width           =   1050
   End
   Begin MSDataListLib.DataCombo cmbEstado 
      Height          =   360
      Left            =   3780
      TabIndex        =   217
      Top             =   585
      Width           =   3000
      _ExtentX        =   5292
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cmbCentro 
      Height          =   360
      Left            =   8010
      TabIndex        =   218
      Top             =   585
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   635
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSDataListLib.DataCombo cmbProcedencia 
      Height          =   315
      Left            =   7425
      TabIndex        =   237
      Top             =   135
      Visible         =   0   'False
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Procedencia"
      Height          =   195
      Index           =   71
      Left            =   6525
      TabIndex        =   238
      Top             =   225
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "CENTRO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   33
      Left            =   6930
      TabIndex        =   236
      Top             =   630
      Width           =   975
   End
   Begin VB.Image imgprioritario 
      Height          =   825
      Left            =   9945
      Picture         =   "frmEquipoEdicion.frx":BCC2B
      Stretch         =   -1  'True
      Top             =   180
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incertidumbre Máx. Adm."
      Height          =   195
      Index           =   67
      Left            =   6795
      TabIndex        =   179
      Top             =   10260
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tolerancia Máx."
      Height          =   195
      Index           =   68
      Left            =   6750
      TabIndex        =   178
      Top             =   10410
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Precisión"
      Height          =   195
      Index           =   69
      Left            =   6750
      TabIndex        =   177
      Top             =   10740
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ESTADO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   13
      Left            =   2790
      TabIndex        =   148
      Top             =   660
      Width           =   930
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "Nº EQUIPO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   90
      TabIndex        =   39
      Top             =   660
      Width           =   1185
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Equipo"
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
      Left            =   90
      TabIndex        =   36
      Top             =   135
      Width           =   9075
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   1170
      Left            =   0
      Top             =   0
      Width           =   13350
   End
End
Attribute VB_Name = "frmEquipoEdicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mvarlngPK As Long
Public ES_AVISO As Boolean

Private mvarenuTipoEdicion As enumTipoEdicion
Private mvarobjEquipo As clsEquipos
Private mvarblnResultado As Boolean

Private mvarblnMtoHistorico As Boolean
Private mvarobjUltCalibracion As clsEquipoCalibracion
Private mvarobjUltVerificacion As clsEquipoVerificacion
Private mvarobjUltMantenimiento As clsEquipoMantenimiento

Private Enum ColsTAB
    COL_GENERAL = 0
    COL_CALIBRACIONES = 1
    COL_VERIFICACIONES = 2
    COL_MANTENIMIENTO = 3
    COL_EVENTOS = 4
    COL_DOCUMENTACION = 5
    COL_ACCESORIOS = 6
    COLS_USO = 7
End Enum

Private Enum COLS
    COL_FECHA = 1
    COL_RESPONSABLE = 2
    COL_PROCEDIMIENTO = 3
    COL_PLAN = 4
    COL_PERIODICIDAD = 5
    COL_RESULTADOS_OBSERVACIONES = 4
    COL_ESTADO = 5
    COL_RESULTADO = 6
    col_id_estado = 8
    COL_ID_PERIODICIODAD = 9
    
    COL_EVENTO = 1
    COL_EVENTO_RAZON = 2
    COL_EVENTO_FECHAHORA = 3
    COL_EVENTO_USUARIO = 4
    COL_EVENTO_OBSERVACIONES = 5
End Enum

Private mvarlngid_plan_mantenimiento_por_defecto As Long
Private mvarblnCargando As Boolean

Private PERMITIR_CAL_VER_PREVISTAS As Boolean
Private Function comprobar_datos() As Boolean

    ' comprueba que pueda quitar la marca de accesorio del equipo
    Dim strCad As String
    Dim rs As ADODB.Recordset
    
    strCad = ""
    
'    If chkCon_Calibracion.value = vbChecked Then
'        Set rs = mvarobjEquipo.devolver_proxima_calibracion(mvarobjEquipo.getID_EQUIPO)
'        If rs.RecordCount = 0 Then
'            strCad = strCad & vbCrLf & " - Ha marcado el equipo CON CALIBRACION, pero no ha existe ningun registro previsto."
'        End If
'    End If
    
'    If chkCon_Verificacion.value = vbChecked Then
'        Set rs = mvarobjEquipo.devolver_proxima_verificacion(mvarobjEquipo.getID_EQUIPO)
'        If rs.RecordCount = 0 Then
'            strCad = strCad & vbCrLf & " - Ha marcado el equipo CON VERIFICACION, pero no ha existe ningun registro previsto."
'        End If
'    End If
        
    If Trim(txtNombre.Text) = "" Then
        strCad = strCad & vbCrLf & " - El Nombre del equipo no puede quedar vacio."
    End If
    If cmbCentro.BoundText = "" Then
        strCad = strCad & vbCrLf & " - El CENTRO del equipo no puede quedar vacio."
    End If
    If getDataComboSel(cmbFamilia) = -1 Then
        strCad = strCad & vbCrLf & " - Debe indicar la Familia del Equipo."
    End If
    
    If getDataComboSel(cmbTipoEquipo) = -1 Then
        strCad = strCad & vbCrLf & " - Debe indicar el Tipo del Equipo."
    End If
    
    If mvarenuTipoEdicion = EDICION Then
        If chkes_accesorio.Value = vbUnchecked Then
            If mvarobjEquipo.comprobar_es_accesorio_de_otros(mvarobjEquipo.getID_EQUIPO) Then
                strCad = strCad & vbCrLf & " - No puede volver el equipo a NO ACCESORIO, dado que consta como accesorio de algún otro equipo."
            End If
        End If
    End If
    'M0459-I
    If txtUnidadesLote <> "" Then
        If Not IsNumeric(txtUnidadesLote) Then
            strCad = strCad & vbCrLf & " - Las Unidades de Lote deben de ser numéricas."
        End If
    End If
    'M0459-F
    
    comprobar_datos = True
    
    If Trim(strCad) <> "" Then
        MsgBox "No se pueden guardar los cambios, por las siguientes razones: " & strCad, vbInformation, "Guardar Cambios"
        comprobar_datos = False
    End If
End Function

Private Sub desactivar_campos_info()
    cmbCalibracionPeriod.Locked = True
    cmbCalibracionResponsable.Locked = True
    cmbCalibracionTipo.Locked = True
    cmbProcedimientoCal.desactivar
    cmbVerificacionPeriod.Locked = True
    cmbVerificacionResponsable.Locked = True
    cmbVerificacionTipo.Locked = True
    cmbProcedimientoVer.desactivar
End Sub

Private Sub PresentarDatos_Accesorios()
    Dim objAcc As New clsEquipoAccesorios
    Dim rs As ADODB.Recordset
    
    Set rs = objAcc.Listado(mvarobjEquipo.getID_EQUIPO)
    
    grdAccesorios.Rows = 1
    
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        While Not rs.EOF
            grdAccesorios.Rows = grdAccesorios.Rows + 1
            grdAccesorios.TextMatrix(grdAccesorios.Rows - 1, 0) = rs("ID_ACCESORIO")
            grdAccesorios.TextMatrix(grdAccesorios.Rows - 1, 1) = rs("NOMBRE")
            If CInt(rs("EN_USO")) = 1 Then
                grdAccesorios.TextMatrix(grdAccesorios.Rows - 1, 2) = "EN USO"
            Else
                grdAccesorios.TextMatrix(grdAccesorios.Rows - 1, 2) = "FIN: " & Format(rs("FECHA_BAJA"), "DD/MM/YYYY")
            End If
            rs.MoveNext
        Wend
    End If
    
    Set objAcc = Nothing
    
End Sub

Public Property Get resultado() As Boolean

    resultado = mvarblnResultado

End Property

Public Property Let resultado(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub PresentarDatos_Generales()
    

    ' De la primera pestaña
    With mvarobjEquipo
        
        txtIdEquipo.Text = Format(.getID_EQUIPO, "00000")
        txtNombre.Text = .getNOMBRE
        'M0459-I
        txtUnidadesLote.Text = .getUNIDADES_LOTE
        'M00459-F
        lbltitulo = .getNOMBRE
        txtdescripcion.Text = .getDESCRIPCION
        cmbFamilia.BoundText = .getFAMILIA_ID
        cmbProveedor.BoundText = .getPROVEEDOR_ID
        cmbTipoEquipo.BoundText = .getTIPO_EQUIPO_ID
        txtNSerie.Text = .getSERIE
        txtModelo.Text = .getMODELO
        txtFabricante.Text = .getFABRICANTE
        If .getFECHA_RECEPCION <> "" Then _
            txtFechaRecepcion.Value = .getFECHA_RECEPCION
        If .getFECHA_SERVICIO <> "" Then _
            txtFechaPuestaServicio.Value = .getFECHA_SERVICIO
        cmbResponsable.BoundText = .getRESPONSABLE_ID
        cmbProcedencia.BoundText = .getPROCEDENCIA_ID
        chkNADCAP.Value = .getES_NADCAP
        If .getES_NADCAP = 0 Then
            lblCampos(38).visible = False
            lblCampos(39).visible = False
            chkMTL.visible = False
            chkCP.visible = False
        End If
        chkMTL.Value = .getES_MTL
        chkCP.Value = .getES_CP
        chkENAC.Value = .getES_ENAC
        chkMostrarEnPlano.Value = .getMOSTRAR_EN_PLANO
        chkPrioritario.Value = .getPRIORITARIO
        chkCritico.Value = .getCRITICO
        cmbTrazmet.MostrarElemento .getTRAZMET_ID
        chkInSitu.Value = .getINSITU
        
        If .getCONDICIONES_AMBIENTALES = 1 Then
            chkCondicionesAmbientales.Value = vbChecked
            fraCondicionesAmbientales.Enabled = True
            txtTemperaturaMin.Text = .getTEMPERATURA_MIN
            txtTemperaturaMax.Text = .getTEMPERATURA_MAX
            txtHumedadMin.Text = .getHUMEDAD_MIN
            txtHumedadMax.Text = .getHUMEDAD_MAX
        End If
        
        txtRangoMedidaMax.Text = .getRANGO_MEDIDA_MAX
        txtRangoMedidaMin.Text = .getRANGO_MEDIDA_MIN
        txtRangoTrabajoMax.Text = .getRANGO_TRABAJO_MAX
        txtRangoTrabajoMin.Text = .getRANGO_TRABAJO_MIN
        Call cmbUnidades.MostrarElemento(.getUNIDAD_ID)
        txtIncertidumbreMax.Text = .getINCERTIDUMBRE_MAXIMA
        txtToleranciaMax.Text = .getTOLERANCIA_MAXIMA
        txtPrecision.Text = .getPRECISIONN
        
        optAlta.Value = (.getALTA_BAJA <> 1)
        optBaja.Value = (.getALTA_BAJA = 1)
        
        'M1124-I
        If .getALTA_BAJA = 1 Then
            txtestado.BackColor = &HFF&
            txtestado = "BAJA"
        Else
            If .getFUERA_SERVICIO Then
                txtestado.BackColor = &HFF&
                txtestado = "FUERA DE SERVICIO"
            Else
                txtestado.BackColor = &HC000&
                txtestado = "ALTA"
            End If
        End If
        Select Case .getESTADO_ID
         Case ENUM_EQ_ESTADOS_ACTIVO
            cmbEstado.BackColor = &HC000&
         Case ENUM_EQ_ESTADOS_BAJA, ENUM_EQ_ESTADOS_FUERA_SERVICIO
            cmbEstado.BackColor = &HFF&
         Case Else
            cmbEstado.BackColor = vbWhite
        End Select
        cmbEstado.BoundText = .getESTADO_ID
        cmbCentro.BoundText = .getCENTRO_ID
        'M1124-F
        
        cmbSituacion.BoundText = .getSITUACION_ID
'J4        txtDescZonaTrabajo.Text = .getSITUACION_DESCRIPCION
        chkes_accesorio.Value = .getES_ACCESORIO
        Call PresentarDatos_LimitacionesUso
        Call PresentarDatos_Accesorios
        Call PresentarDatos_Normas
'J5        Call PresentarDatos_AdjuntosInstruccionesTecnicas
        
        txtuso(0) = .getNUMERO_USOS_MAXIMO
        txtuso(1) = .getNUMERO_USOS_CONTADOR
        'M1050-I
        cmbCliente.MostrarElemento .getCLIENTE_ID
        txtNumeroCliente = .getNUMERO_EQUIPO_CLIENTE
        If Not IsNull(.getOBSERVACIONES) Then
            txtObservaciones = .getOBSERVACIONES
        Else
            txtObservaciones = ""
        End If
        'M1050-F
    End With

End Sub

Private Sub PresentarDatos_LimitacionesUso()
    Dim objItem As clsGenericClass

    lstLimitacionesUso.Clear
    txtLimitacionesUso.Text = ""
        
    For Each objItem In mvarobjEquipo.getLIMITACIONES_USO_COL.Iterator
        If objItem.getID_AUX <> enumIdAux.ID_AUX_ELIMINADO Then
            lstLimitacionesUso.AddItem objItem.getNOMBRE
            lstLimitacionesUso.ItemData(lstLimitacionesUso.ListCount - 1) = objItem.getID
        End If
    Next objItem
    
    If lstLimitacionesUso.ListCount > 0 Then
        If cpLimitaciones.PanelOpen = False Then
            cpLimitaciones.PanelOpen = True
        End If
    End If

End Sub

Private Sub chkNADCAP_Click()
    lblCampos(38).visible = chkNADCAP.Value
    lblCampos(39).visible = chkNADCAP.Value
    chkMTL.visible = chkNADCAP.Value
    chkCP.visible = chkNADCAP.Value
    If chkNADCAP.Value = Unchecked Then
        chkMTL.Value = Unchecked
        chkCP.Value = Unchecked
    End If
End Sub

Private Sub chkPrioritario_Click()
'    If chkPrioritario.value = Checked Then
'        imgprioritario.Visible = True
'    Else
'        imgprioritario.Visible = False
'    End If
End Sub

Private Sub cmbduplicarVerificacion_Click()
    Dim lngFila As Long, strId As String
    Dim lngid_periodicidad As Long
    Dim oEV As New clsEquipoVerificacion

    lngFila = grdVerificaciones.RowSel
    If lngFila < 1 Then Exit Sub
    
    strId = grdVerificaciones.TextMatrix(lngFila, 0)
    If oEV.duplicar(CLng(strId)) = True Then
        MsgBox "La verificación se ha duplicado correctamente.", vbInformation, App.Title
        PresentarDatos_Verificacion
    Else
        MsgBox "Error al duplicar la verificación.", vbCritical, App.Title
    End If
End Sub

'J1
'Private Sub cmbCalibracionPeriod_b_Click(AREA As Integer)
'    cmbCalibracionPeriod.BoundText = cmbCalibracionPeriod_b.BoundText
'End Sub

'Private Sub cmbCalibracionPeriod_Click(AREA As Integer)
'    cmbCalibracionPeriod_b.BoundText = cmbCalibracionPeriod.BoundText
'End Sub
'Private Sub cmbCalibracionResponsable_b_Click(AREA As Integer)
'cmbCalibracionResponsable.BoundText = cmbCalibracionResponsable_b.BoundText
'End Sub

'Private Sub cmbCalibracionResponsable_Click(AREA As Integer)
'cmbCalibracionResponsable_b.BoundText = cmbCalibracionResponsable.BoundText
'End Sub

'Private Sub cmbCalibracionTipo_b_Click(AREA As Integer)
'cmbCalibracionTipo.BoundText = cmbCalibracionTipo_b.BoundText
'End Sub
'Private Sub cmbCalibracionTipo_Click(AREA As Integer)
'cmbCalibracionTipo_b.BoundText = cmbCalibracionTipo.BoundText
'End Sub
'J1
'J3
'Private Sub cmbMtoPlan_b_Click(AREA As Integer)
'    cmbMtoPlan.BoundText = cmbMtoPlan_b.BoundText
'End Sub
'Private Sub cmbMtoPlan_Click(AREA As Integer)
'    cmbMtoPlan_b.BoundText = cmbMtoPlan.BoundText
'End Sub
'Private Sub cmbMtoResponsable_b_Click(AREA As Integer)
' cmbMtoResponsable.BoundText = cmbMtoResponsable_b.BoundText
'End Sub
'Private Sub cmbMtoResponsable_Click(AREA As Integer)
' cmbMtoResponsable_b.BoundText = cmbMtoResponsable.BoundText
'End Sub
'J3

Private Sub cmbTipoCalibraciones_Click()
  
    If Not cmbTipoCalibraciones.Enabled Then Exit Sub
    
    PresentarDatos_Calibracion
End Sub

Private Sub cmbTipoVerificaciones_Click()
If Not cmbTipoVerificaciones.Enabled Then Exit Sub
PresentarDatos_Verificacion
End Sub

Private Sub cmdAccesoriosAdd_Click()
    Set frmEquipoAccesorios.EQUIPO = mvarobjEquipo
    frmEquipoAccesorios.Show vbModal
    
    
    PresentarDatos_Accesorios

End Sub

Private Sub cmdAccesoriosDelete_Click()
    Dim lngid As Long, lngFila As Long
    Dim objItem As New clsEquipoAccesorios

    With grdAccesorios
        lngFila = .RowSel
        If lngFila <= 0 Then Exit Sub
        lngid = CLng(.TextMatrix(lngFila, 0))
        objItem.Eliminar lngid, mvarobjEquipo.getID_EQUIPO
    End With

    Call PresentarDatos_Accesorios
End Sub

Private Sub cmdAdjuntarHojaCal_Click()
    Dim oD As New clsDocumentacion
    oD.CargarEquipo mvarobjEquipo.getID_EQUIPO, 0, CLng(txtactcalibracion), 2, True
    Set oD = Nothing
    
'    Dim objAI As New clsArchivoAdjunto
'    Dim destino As String, r As Double
'
'    destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\EQUIPOS\" & CStr(mvarobjEquipo.getID_EQUIPO) & "\CAL\" & txtactcalibracion & "\CERT\" & txtrutacalibracion
'
'    If destino = "" Then
'        MsgBox "No tiene certificado adjunto.", vbCritical, App.Title
'        Exit Sub
'    End If
'    If Dir(destino) = "" Then
'        MsgBox "El certificado adjunto no se encuentra en la ruta.", vbCritical, App.Title
'        Exit Sub
'    End If
'
'    ' verificar si es hoja excel
'    If UCase(Right(destino, 3)) = "XLS" Then
'        Dim XLA As excel.Application
'        Dim XLW As excel.Workbook
'        Dim XLS As excel.Worksheet
'        Set XLA = New excel.Application
'        Set XLW = XLA.Workbooks.Open(destino, , True)
'        Set XLS = XLW.Worksheets(1)
'        XLA.Visible = True
'    ElseIf Dir(destino, vbArchive) <> "" Then
'        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
'    End If

End Sub

Private Sub cmdAdjuntarHojaVer_Click()
    Dim oD As New clsDocumentacion
    oD.CargarEquipo mvarobjEquipo.getID_EQUIPO, 0, CLng(txtactverificacion), 2, True
    Set oD = Nothing
    
'    Dim destino As String, r As Double
'    destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\EQUIPOS\" & CStr(mvarobjEquipo.getID_EQUIPO) & "\VER\" & txtactverificacion & "\CERT\" & txtrutaverificacion
'
'    If destino = "" Then
'        MsgBox "No tiene certificado adjunto.", vbCritical, App.Title
'        Exit Sub
'    End If
'    If Dir(destino) = "" Then
'        MsgBox "El certificado adjunto no se encuentra en la ruta.", vbCritical, App.Title
'        Exit Sub
'    End If
'
'    ' verificar si es hoja excel
'    If UCase(Right(destino, 3)) = "XLS" Then
'        Dim XLA As excel.Application
'        Dim XLW As excel.Workbook
'        Dim XLS As excel.Worksheet
'        Set XLA = New excel.Application
'        Set XLW = XLA.Workbooks.Open(destino, , True)
'        Set XLS = XLW.Worksheets(1)
'        XLA.Visible = True
'    ElseIf Dir(destino, vbArchive) <> "" Then
'        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
'    End If
End Sub

Private Sub cmdAnadirEvento_Click()
Dim objfrm As New frmEquipoEvento
Dim lngFila As String

    Call RecogerDatos

    objfrm.PK = 0
    Set objfrm.EQUIPO = mvarobjEquipo
    
    objfrm.Show vbModal
    
    If objfrm.resultado Then
        'Set mvarobjEquipo = objfrm.Equipo
        PresentarDatos_Eventos
    End If
    
    Unload objfrm
    Set objfrm = Nothing
End Sub

Private Sub cmdAnadirLimitacion_Click()
Dim objItem As New clsGenericClass
Dim objCol As clsGenericCollection
Dim blnExiste As Boolean

    
    mvarobjEquipo.Anadir_limitacionuso_equipo txtLimitacionesUso.Text

    Call PresentarDatos_LimitacionesUso
    
End Sub

Private Sub cmdAnadirNorma_Click()
    Dim objCol As clsGenericCollection, objItem As New clsGenericClass
    Dim r As Double

    
    If optiponorma(0).Value = True Then
        If cmbDocumentos.getPK_SALIDA = 0 Then
            MsgBox "Debe seleccionar una de entre las existentes", vbOK, "Añadir Norma"
            Exit Sub
        End If
        mvarobjEquipo.Anadir_Norma_Equipo cmbDocumentos.getPK_SALIDA, 0
    Else
        If cmbNormas.getPK_SALIDA = 0 Then
            MsgBox "Debe seleccionar una de entre las existentes", vbOK, "Añadir Norma"
            Exit Sub
        End If
        mvarobjEquipo.Anadir_Norma_Equipo cmbNormas.getPK_SALIDA, 1
    End If

    Call PresentarDatos_Normas

End Sub

Private Sub cmdAnadirPlan_Click()
    Dim ID_PLAN As Long
    
    ID_PLAN = getDataComboSel(cmbMtoPlan)
    
    If ID_PLAN <= 0 Then Exit Sub
    
    
    mvarobjEquipo.anadir_plan_mantenimiento_equipo ID_PLAN
    
    
    PresentarDatos_Mantenimiento_Planes
    
End Sub

Private Sub cmdCalibracionMostrar_Click()
    Dim objfrm  As New frmEquipoEdicionCalibracion
    Dim lngFila As Long, strId As String
    
    strId = txtactcalibracion
    
    With objfrm
        Set .EQUIPO = mvarobjEquipo
        .ID = strId
        .TipoEdicion = visualizar
        .Show vbModal
    End With
    Unload objfrm
    Set objfrm = Nothing
End Sub

Private Sub cmdDocumentacion_Click()
'M0499-I
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_EQUIPO
        .COBJETO = mvarobjEquipo.getID_EQUIPO
        .Show 1
    End With
    Set frmAdjuntos = Nothing
'M0499-F
    
End Sub
Private Sub cmdEliminarCalibracion_Click()
    Dim lngFila As Long, strId As String
    Dim lng_periodicidad_id As Long
    
    lngFila = grdCalibraciones.RowSel
    If lngFila < 1 Then Exit Sub
    
    strId = grdCalibraciones.TextMatrix(lngFila, 0)
    lng_periodicidad_id = grdCalibraciones.TextMatrix(lngFila, COLS.COL_ID_PERIODICIODAD)
    
    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 Then
        MsgBox "Solo el resposable de Metrología puede eliminar la calibración.", vbExclamation, App.Title
        Exit Sub
    End If
    If MsgBox("¿Está seguro que desea eliminar la calibración seleccionada?", vbInformation + vbYesNo, "Eliminar Calibración") = vbNo Then
        Exit Sub
    End If
    mvarobjEquipo.Eliminar_Calibracion strId, lng_periodicidad_id
    Call PresentarDatos_Calibracion(True)
    PresentarDatos_Eventos
End Sub

Private Sub cmdEliminarEvento_Click()
Dim lngFila As String, lngid As Long
Dim objItem As New clsEquipoEventos

    lngFila = grdEventos.RowSel
    If lngFila < 1 Then
        Set objItem = Nothing
        Exit Sub
    End If
    ' M0496-I
    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 Then
        MsgBox "Solo el resposable de Metrología puede eliminar el evento.", vbExclamation, App.Title
        Exit Sub
    End If
    ' M0496-F
        
    lngid = CLng(grdEventos.TextMatrix(lngFila, 0))
    
    If MsgBox("Eliminar un Evento del histórico del equipo podría romper la trazabilidad del mismo. ¿Está seguro que desa continuar eliminando el Evento?", vbInformation + vbYesNo, "Eliminar Evento") = vbNo Then
        Exit Sub
    End If
       
    objItem.Eliminar lngid
    
    PresentarDatos_Eventos
    
End Sub

Private Sub cmdEliminarLimitacion_Click()
    Dim lngid As Long
    Dim objItem As clsGenericClass
    Dim objCol As clsGenericCollection

    If lstLimitacionesUso.ListIndex < 0 Then Exit Sub
    
    lngid = lstLimitacionesUso.ItemData(lstLimitacionesUso.ListIndex)
    
    mvarobjEquipo.Eliminar_LimitacionUso_equipo lngid
    
    Call PresentarDatos_LimitacionesUso

End Sub



Private Sub cmdEliminarMto_Click()
Dim lngFila As String, lng_id_plan As Long, lng_id As Long
'Dim objItem As clsEquipoMantenimiento

    lngFila = grdMantenimientos.RowSel
    
    If lngFila < 1 Then Exit Sub
    
    
    
    'Set objItem = mvarobjEquipo.Mantenimientos.Item(grdMantenimientos.TextMatrix(lngFila, 0))
    'If objItem.getESTADO > 0 Then
    '    If MsgBox("El Registro de Mantenimiento que pretende Eliminar, ya se encuentra Cerrado. ¿Está seguro que desa continuar eliminando este mantenimiento?", vbInformation + vbYesNo, "Eliminar Mantenimiento") = vbNo Then
    '        Exit Sub
    '    End If
    'End If
    
'    If CLng(grdMantenimientos.TextMatrix(lngFila, COLS.col_id_estado)) <> 0 Then
'        Call MsgBox("El Registro de Mantenimiento ya se encuentra Cerrado. No se puede eliminar.", vbInformation, "Eliminar Mantenimiento")
'        Exit Sub
'    End If
    ' M0496-I
    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 Then
        MsgBox "Solo el resposable de Metrología puede eliminar el mantenimiento.", vbExclamation, App.Title
        Exit Sub
    End If
    ' M0496-F
    
    
    lng_id = CLng(grdMantenimientos.TextMatrix(lngFila, 0))
    lng_id_plan = CLng(grdMantenimientos.TextMatrix(lngFila, COLS.COL_ID_PERIODICIODAD))
    
    'Call mvarobjEquipo.Mantenimientos.Remove(CStr(objItem.getID_MANTENIMIENTO))
    'objItem.Eliminar CLng(grdMantenimientos.TextMatrix(lngFila, 0))
    mvarobjEquipo.eliminar_mantenimiento lng_id, lng_id_plan
    'mvarobjEquipo.Carga_Mantenimiento
    
    PresentarDatos_Mantenimiento True
    PresentarDatos_Eventos
    
    Dim oMto As New clsEquipoMantenimiento
        oMto.comprobar_fin_mantenimientos_previstos mvarobjEquipo.getID_EQUIPO, lng_id_plan
    Set oMto = Nothing
    
    
End Sub

Private Sub cmdEliminarPlan_Click()
    
    Dim ID_PLAN As Long
    
    ' recoge el elemento seleccionado en la lista
    If lstPlanes.ListIndex < 0 Then Exit Sub
    
    ID_PLAN = lstPlanes.ItemData(lstPlanes.ListIndex)
    
    If ID_PLAN <= 0 Then Exit Sub
        
    mvarobjEquipo.eliminar_plan_mantenimiento_equipo ID_PLAN


    PresentarDatos_Mantenimiento_Planes

End Sub

Private Sub cmdEliminarVerificacion_Click()
    Dim lngFila As Long, strId As String
    Dim lngid_periodicidad As Long
   On Error GoTo cmdEliminarVerificacion_Click_Error
    lngFila = grdVerificaciones.RowSel
    If lngFila < 1 Then Exit Sub
    strId = grdVerificaciones.TextMatrix(lngFila, 0)
    lngid_periodicidad = grdVerificaciones.TextMatrix(lngFila, COLS.COL_ID_PERIODICIODAD)

    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 Then
        MsgBox "Solo el resposable de Metrología puede eliminar la verificación.", vbExclamation, App.Title
        Exit Sub
    End If
    If MsgBox("¿Está seguro que desea eliminar la verificación seleccionada?", vbInformation + vbYesNo, "Eliminar Verificación") = vbNo Then
        Exit Sub
    End If
    mvarobjEquipo.Eliminar_Verificacion strId, lngid_periodicidad
    
    PresentarDatos_Verificacion True
    PresentarDatos_Eventos

   On Error GoTo 0
   Exit Sub

cmdEliminarVerificacion_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminarVerificacion_Click of Formulario frmEquipoEdicion"
End Sub

Private Sub cmdEliminarVerificacion2_Click()

End Sub

Private Sub cmdEspecificaciones_Click()
    frmEquipoEspecificacionesTecnicas.PK = mvarobjEquipo.getID_EQUIPO
    frmEquipoEspecificacionesTecnicas.Show 1
End Sub

Private Sub cmdetiqueta_Click()
    If mvarobjEquipo.getID_EQUIPO = 0 Then
        If MsgBox("Los datos del equipo aun no han sido guardados. ¿Desea Guardarlos ahora?", vbInformation + vbYesNo, "Imprimir Ficha de Equipo") = vbYes Then
            mvarobjEquipo.Insertar
        Else
            Exit Sub
        End If
    End If
        
    Call mvarobjEquipo.ImprimirEtiqueta(mvarobjEquipo.getID_EQUIPO)
End Sub

Private Sub cmdImprimirFicha_Click()
    If mvarobjEquipo.getID_EQUIPO = 0 Then
        If MsgBox("Los datos del equipo aun no han sido guardados. ¿Desea Guardarlos ahora?", vbInformation + vbYesNo, "Imprimir Ficha de Equipo") = vbYes Then
            mvarobjEquipo.Insertar
        Else
            Exit Sub
        End If
    End If
        
    Call mvarobjEquipo.ImprimirFichaEquipo(mvarobjEquipo.getID_EQUIPO)
End Sub

'Private Sub cmdInsertaNorma_Click()
'    gID = 0
'    If optiponorma(0).value = True Then
'        frmCA_Listado_Documentos.VINCULAR = True
'        frmCA_Listado_Documentos.Show 1
'    Else
'        frmCA_Listado_Normas.VINCULAR = True
'        frmCA_Listado_Normas.Show 1
'    End If
'    If gID <> 0 Then
'        If optiponorma(0).value = True Then
'            cmbDocumentos.MostrarElemento gID
'        Else
'            cmbNormas.MostrarElemento gID
'        End If
'    End If
'End Sub

Private Sub cmdMantenimientoMostrar_Click()
    If txtactmantenimiento <> "" Then
        If IsNumeric(txtactmantenimiento) Then
    Dim objfrm  As New frmEquipoEdicionMtoFechasEdicion
    Dim lngFila As String

    objfrm.TipoEdicion = visualizar
    Set objfrm.EQUIPO = mvarobjEquipo
    objfrm.id_mantenimiento = CLng(txtactmantenimiento)
    objfrm.MostrarCierre = True
    
    objfrm.Show vbModal
    
    Unload objfrm
    Set objfrm = Nothing
    End If
    End If
End Sub

Private Sub cmdMetrol_Click()
    If txtNumeroCliente = "" Then
        MsgBox "Introduzca el código del equipo.", vbExclamation, App.Title
        Exit Sub
    End If
    ' Validar que el equipo no exista
    Dim rs As ADODB.Recordset
    Dim c As String
    c = "select id_equipo from equipos where numero_equipo_cliente = '" & Trim(txtNumeroCliente) & "' limit 1"
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        MsgBox "El equipo ya existe con número : " & rs(0), vbExclamation, App.Title
        Exit Sub
    End If
    ' Captura de datos
'    If CrearConexionGlobal_metrol() = True Then
        c = "select * from equipos where codigo = '" & Trim(txtNumeroCliente) & "' limit 1"
        Set rs = datos_bd_metrol(c)
        If rs.RecordCount > 0 Then
            ' Familia
            Dim rsFamilia As ADODB.Recordset
            c = "select * from familias where id_familia = " & rs("familia_id")
            Set rsFamilia = datos_bd_metrol(c)
            If rsFamilia.RecordCount > 0 Then
                txtNombre = rsFamilia("NOMBRE")
            End If
            Set rsFamilia = Nothing
            ' Fabricante
            Dim rsFab As ADODB.Recordset
            c = "select * from fabricantes where id_fabricante = " & rs("fabricante_id")
            Set rsFab = datos_bd_metrol(c)
            If rsFab.RecordCount > 0 Then
                txtFabricante = rsFab("NOMBRE")
            End If
            Set rsFab = Nothing
            txtObservaciones = ""
            txtObservaciones = txtObservaciones & "CAPACIDAD  : " & rs("CAPACIDAD") & vbNewLine
            ' Intervalo
            Dim rsIntervalo As ADODB.Recordset
            c = "select * from intervalos where id_intervalo = " & rs("INTERVALO_ID")
            Set rsIntervalo = datos_bd_metrol(c)
            If rsIntervalo.RecordCount > 0 Then
                txtObservaciones = txtObservaciones & "INTERVALO  : " & rsIntervalo("CODIGO") & vbNewLine
            End If
            Set rsIntervalo = Nothing
            ' Datos tecnicos
            Dim rsDT As ADODB.Recordset
            c = "select * from equipos_dt where equipo_id = " & rs("ID_EQUIPO") & " limit 1"
            Set rsDT = datos_bd_metrol(c)
            If rsDT.RecordCount > 0 Then
            ' Unidad
                Dim rsUnidad As ADODB.Recordset
                c = "select * from unidades where id_unidad = " & rsDT("UNIDAD_ID")
                Set rsUnidad = datos_bd_metrol(c)
                If rsUnidad.RecordCount > 0 Then
                    txtObservaciones = txtObservaciones & "UNIDAD  : " & rsUnidad("NOMBRE") & vbNewLine
                End If
                Set rsUnidad = Nothing
                txtObservaciones = txtObservaciones & "TOLERANCIA  : " & rsDT("TOLERANCIA") & vbNewLine
                txtObservaciones = txtObservaciones & "RESOLUCION  : " & rsDT("RESOLUCION") & vbNewLine
                'INSITU
                chkInSitu.Value = rsDT("INSITU")
            End If
            Set rsDT = Nothing
            txtNSerie = rs("SERIE")
            txtModelo = rs("MODELO")
            ' Cliente
            Dim opar As New clsParametros
            opar.Carga 270, ""
            cmbCliente.MostrarElemento opar.getVALOR
            ' Responsable
            opar.Carga 271, ""
            cmbResponsable.BoundText = opar.getVALOR
'JGM            chkCon_Calibracion.value = Checked
            ' Tipo de Equipo
            Select Case rs("AREA_ID")
            Case "D"
                cmbFamilia.BoundText = 6
            Case "E"
                cmbFamilia.BoundText = 5
            Case "M"
                cmbFamilia.BoundText = 16
            Case "Q"
                cmbFamilia.BoundText = 17
            Case "T"
                cmbFamilia.BoundText = 3
            Case "V"
                cmbFamilia.BoundText = 1
            End Select
            ' Area
            Select Case rs("PLANTA_ID")
            Case "P"
                cmbSituacion.BoundText = 37
            Case "C"
                cmbSituacion.BoundText = 35
            Case "K"
                cmbSituacion.BoundText = 40
            Case "S"
                cmbSituacion.BoundText = 36
            End Select
            ' Localizacion (Planta)
            opar.Carga 272, ""
            cmbTipoEquipo.BoundText = opar.getVALOR
'        End If
'        conn_metrol.Close
    End If
End Sub

Private Sub cmdModificarUsos_Click()
    If txtuso(3) <> "" Then
        If CInt(txtuso(3)) Then
            Dim oEq_usos As New clsEq_usos
            oEq_usos.Modificar txtIdEquipo, listaUsos.ListItems(listaUsos.selectedItem.Index).SubItems(6), txtuso(3)
            listaUsos.ListItems(listaUsos.selectedItem.Index).SubItems(8) = txtuso(3)
        End If
    End If
        
End Sub

Private Sub cmdok_Click()

On Error GoTo cmdok_Click_Error

    If Not comprobar_datos Then Exit Sub
    
    Call GuardarEquipo
        
    cmdVerSituacionEnPlano.Enabled = True
    mvarblnResultado = True
    
On Error GoTo 0
    Exit Sub
cmdok_Click_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmEquipoEdicion"
End Sub
Private Sub chkCondicionesAmbientales_Click()
    fraCondicionesAmbientales.visible = (chkCondicionesAmbientales.Value = vbChecked)
    fraCondicionesAmbientales.Enabled = (chkCondicionesAmbientales.Value = vbChecked)
End Sub

Private Sub cmdReabrir_Click()
    Dim lngFila As Long, strId As String
    Dim lngid_periodicidad As Long
    Dim verificacion As New clsEquipoVerificacion

    lngFila = grdVerificaciones.RowSel
    If lngFila < 1 Then Exit Sub

    strId = grdVerificaciones.TextMatrix(lngFila, 0)
    lngid_periodicidad = grdVerificaciones.TextMatrix(lngFila, COLS.COL_ID_PERIODICIODAD)
 
    verificacion.Reabrir strId
    MsgBox "La verificación ha sido cambiada a 'Prevista'", vbInformation, App.Title
    PresentarDatos_Verificacion
End Sub

Private Sub cmdReabrirCalibra_Click()
'M1146
    Dim lngFila As Long, strId As String
    Dim lng_periodicidad_id As Long
    Dim calibracion As New clsEquipoCalibracion
    
    lngFila = grdCalibraciones.RowSel
    If lngFila < 1 Then Exit Sub
    
    strId = grdCalibraciones.TextMatrix(lngFila, 0)
    lng_periodicidad_id = grdCalibraciones.TextMatrix(lngFila, COLS.COL_ID_PERIODICIODAD)
 
    calibracion.Reabrir strId
    MsgBox "La calibración ha sido cambiada a 'Prevista'", vbInformation, App.Title
    Call PresentarDatos_Calibracion(True)
    PresentarDatos_Eventos
End Sub

Private Sub cmdRecepcion_Click()
    frmEquipoRecepcion.PK = mvarobjEquipo.getID_EQUIPO
    frmEquipoRecepcion.Show 1
End Sub

Private Sub cmdVerificacionContador_Click()
    Dim objfrm  As New frmEquipoEdicionVerificacion
    Dim strId As String, fila As Long
    
    strId = "0"
    Set objfrm.EQUIPO = mvarobjEquipo
    objfrm.ID = strId
    objfrm.TipoEdicion = Alta
    objfrm.idVerificadorInternoInicial = USUARIO.getID_EMPLEADO
    objfrm.idProcedmientoInicial = cmbProcedimientoVer.getPK_SALIDA
    If txtVerificacionFechaProxima.Value = CDate("1900-01-01") Then
        objfrm.FechaProximaInicial = Now
    Else
        objfrm.FechaProximaInicial = txtVerificacionFechaProxima.Value
    End If
    objfrm.IdPeriodoInicial = getDataComboSel(cmbVerificacionPeriod)
    objfrm.IdTipoVerificacionIncial = getDataComboSel(cmbVerificacionTipo)
    
    'MANTIS-810-I (MUCHO OJO CON ESTA INSTRUCCION)
    'objfrm.copiarUltimaVerificacionPeriodo = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_ENSAYO
    objfrm.copiarUltimaVerificacionPeriodo = ENUM_EQUIPOS_PERIODICIDAD.ENUM_EQUIPOS_PERIODICIDAD_ANTES_CADA_ENSAYO
    'MANTIS-810-F
    objfrm.Show vbModal
    
    If Not objfrm.resultado Then
        Unload objfrm
        Set objfrm = Nothing
        Exit Sub
    End If
    
    Set mvarobjEquipo = objfrm.EQUIPO
    
    PresentarDatos_Usos
End Sub

Private Sub cmdVerificacionMostrar_Click()
    Dim objfrm  As New frmEquipoEdicionVerificacion
    Dim lngFila As Long, strId As String
    Dim intEstado As Integer
    strId = txtactverificacion
    With objfrm
        Set .EQUIPO = mvarobjEquipo
        .ID = strId
        .TipoEdicion = visualizar
        .Show vbModal
    End With
    Unload objfrm
    Set objfrm = Nothing
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'M0496-I
    Select Case KeyCode
        Case 27
            cmdcancel_Click
        Case 116
            frmUnidadesLote.visible = Not frmUnidadesLote.visible
            lblCampos(46).visible = Not lblCampos(46).visible
            txtuso(1).visible = Not txtuso(1).visible
            cmdVerificacionContador.visible = Not cmdVerificacionContador.visible
    End Select
    'M0496-F
End Sub

Private Sub grdAccesorios_DblClick()

Dim objfrm As frmEquipoAccesorios
Dim lngFila As Long

    lngFila = grdAccesorios.RowSel

    If lngFila <= 0 Then
        Set objfrm = Nothing
        Exit Sub
    End If

    Set objfrm = New frmEquipoAccesorios
    Set objfrm.EQUIPO = mvarobjEquipo
    objfrm.PK = CInt(grdAccesorios.TextMatrix(lngFila, 0))
    
    objfrm.Show vbModal

    PresentarDatos_Accesorios

End Sub


Private Sub grdCalibraciones_Click()
    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 Then
       cmdReabrirCalibra.Enabled = False
    Else
       cmdReabrirCalibra.Enabled = True
    End If
End Sub

Private Sub grdCalibraciones_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If grdCalibraciones.MouseRow <> 0 Then Exit Sub
    OrdernarColumnaMsGrid grdCalibraciones, IIf(grdCalibraciones.MouseCol = 3, grdCalibraciones.COLS - 1, grdCalibraciones.MouseCol)
End Sub


Private Sub grdEventos_DblClick()
    Dim objfrm  As New frmEquipoEvento
    Dim lngFila As String

    lngFila = grdEventos.RowSel
    If lngFila < 1 Then Exit Sub
    
    Call RecogerDatos

    Set objfrm.EQUIPO = mvarobjEquipo
    objfrm.PK = CLng(grdEventos.TextMatrix(lngFila, 0))
    
    objfrm.Show vbModal
    
    
    Unload objfrm
    Set objfrm = Nothing
End Sub

Private Sub grdEventos_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Si no es fila 0, nada
    If grdEventos.MouseRow <> 0 Then Exit Sub

    ' Ordena por la columna donde se hace Click.
    OrdernarColumnaMsGrid grdEventos, IIf(grdEventos.MouseCol = COLS.COL_EVENTO_FECHAHORA, grdEventos.COLS - 1, grdEventos.MouseCol)

End Sub


Private Sub grdMantenimientos_DblClick()

    Dim objfrm  As New frmEquipoEdicionMtoFechasEdicion
    Dim lngFila As String

    lngFila = grdMantenimientos.RowSel
    If lngFila < 1 Then Exit Sub
    
    Call RecogerDatos

    objfrm.TipoEdicion = IIf(mvarblnMtoHistorico, enumTipoEdicion.visualizar, enumTipoEdicion.EDICION)
    Set objfrm.EQUIPO = mvarobjEquipo
    objfrm.id_mantenimiento = CLng(grdMantenimientos.TextMatrix(lngFila, 0))
    objfrm.MostrarCierre = True
    
    objfrm.Show vbModal
    
    If objfrm.resultado Then
'        Set mvarlngidEquipo = objfrm.EQUIPO
        mvarobjEquipo.Carga mvarobjEquipo.getID_EQUIPO
        Call PresentarDatos_Mantenimiento
        PresentarDatos_Eventos
    End If
    
    Unload objfrm
    Set objfrm = Nothing
End Sub

Private Sub grdNormas_DblClick()
    Dim strDest As String
    Dim r As Double

    If grdNormas.ListItems.Count > 0 Then
        strDest = Replace(grdNormas.ListItems(grdNormas.selectedItem.Index).SubItems(2), "/", "\")
        If Trim(strDest) = "" Then
            MsgBox "El documento no tiene nada vinculado.", vbExclamation, App.Title
        Else
            If Dir(strDest, vbArchive) <> "" Then
                r = Shell("rundll32.exe url.dll,FileProtocolHandler " & strDest, vbNormalFocus)
            End If
        End If
    End If
End Sub

Private Sub grdVerificaciones_Click()
    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 Then
       cmdReabrir.Enabled = False
    Else
       cmdReabrir.Enabled = True
    End If
End Sub

Private Sub grdVerificaciones_DblClick()
    If cmbTipoEquipo.BoundText = EQ_TIPOS_EQUIPOS.TIPO_EQUIPO_TORCOMETRO Then
        MsgBox "Los TORCOMETROS deben gestionarse en GESMET.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim objfrm  As New frmEquipoEdicionVerificacion
    Dim lngFila As Long, strId As String
    Dim intEstado As Integer
    
    lngFila = grdVerificaciones.RowSel
    If lngFila < 1 Then Exit Sub
    
    strId = grdVerificaciones.TextMatrix(lngFila, 0)
    intEstado = CLng(grdVerificaciones.TextMatrix(lngFila, COLS.col_id_estado))
    With objfrm
        Set .EQUIPO = mvarobjEquipo
        .ID = strId
        
        If intEstado = 0 Then
            .TipoEdicion = EDICION ' si no está cerrado
        Else
            .TipoEdicion = visualizar
        End If
        
        
        .Show vbModal
        
        If .resultado Then
            Set mvarobjEquipo = .EQUIPO
            Call PresentarDatos_Verificacion(True)
            PresentarDatos_Eventos
            ' Por si han cambiado las limitaciones de uso
            Call PresentarDatos_LimitacionesUso
    
        End If
    End With
    Unload objfrm
    Set objfrm = Nothing
End Sub
Private Sub grdVerificaciones_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' Si no es fila 0, nada
    If grdVerificaciones.MouseRow <> 0 Then Exit Sub
    ' Ordena por la columna donde se hace Click.
    OrdernarColumnaMsGrid grdVerificaciones, IIf(grdVerificaciones.MouseCol = 3, grdVerificaciones.COLS - 1, grdVerificaciones.MouseCol)
End Sub

Private Sub imgBuscarPlanMto_Click()
    Dim objfrm As New frmEquipoPlanMtoEdicion
    Dim ID As Long
    ID = getDataComboSel(cmbMtoPlan)
    
    If ID = -1 Then
        If MsgBox("No ha seleccionado ningún plan. ¿Desea crear uno nuevo?", vbInformation + vbYesNo, "Crear Plan de Mantenimiento de Equipos") = vbYes Then
            objfrm.TipoEdicion = Alta
        Else
            Set objfrm = Nothing
            Exit Sub
        End If
    Else
        objfrm.TipoEdicion = EDICION
        
    End If
    objfrm.ID_PLAN_MTO = ID
    objfrm.Show vbModal
    If objfrm.resultado = True Then
        
        If objfrm.TipoEdicion = Alta Then
            ID = objfrm.ID_PLAN_MTO
            cargar_combo cmbMtoPlan, New clsPlanMantenimiento
            cmbMtoPlan.BoundText = CStr(ID)
        End If
    End If
    Unload objfrm
    Set objfrm = Nothing
End Sub

Private Sub imgNuevoPlanMto_Click()
    Dim objfrm As New frmEquipoPlanMtoEdicion
    Dim ID As Long


    objfrm.TipoEdicion = Alta
        
    objfrm.Show vbModal

    If objfrm.resultado = True Then
        ID = objfrm.ID_PLAN_MTO
        cargar_combo cmbMtoPlan, New clsPlanMantenimiento
'J3        cargar_combo cmbMtoPlan_b, New clsPlanMantenimiento
        cmbMtoPlan.BoundText = CStr(ID)
'J3        cmbMtoPlan_b.BoundText = CStr(ID)
    End If

    Unload objfrm
    Set objfrm = Nothing

End Sub

Private Sub listaUsos_Click()
    txtuso(3) = ""
    If listaUsos.ListItems.Count > 0 Then
        txtuso(3) = listaUsos.ListItems(listaUsos.selectedItem.Index).SubItems(8)
    End If
End Sub

Private Sub listaUsos_DblClick()
    If listaUsos.ListItems.Count > 0 Then
        gmuestra = listaUsos.ListItems(listaUsos.selectedItem.Index).SubItems(6)
        frmVerMuestra.Show 1
    End If
End Sub

Private Sub lstLimitacionesUso_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then cmdEliminarLimitacion_Click
    End Sub

Private Sub lstPlanes_ItemCheck(Item As Integer)

    If mvarblnCargando Then Exit Sub
    If lstPlanes.ListCount <= 0 Then Exit Sub
    If Item < 0 Then Exit Sub
    
    If lstPlanes.Selected(Item) Then
        mvarobjEquipo.actualizar_plan_mantenimiento_por_defecto lstPlanes.ItemData(Item)
    Else
        If lstPlanes.ItemData(Item) = mvarlngid_plan_mantenimiento_por_defecto Then
            lstPlanes.Selected(Item) = True
        End If
    End If
    
    Call PresentarDatos_Mantenimiento_Planes

End Sub

Private Sub optAlta_Click()
    If mvarobjEquipo.getALTA_BAJA = 1 Then
        If optAlta.Value Then
            If MsgBox("¿Desea volver a poner de ALTA este equipo?", vbInformation + vbYesNo) = vbYes Then
                If mvarobjEquipo.volver_dar_alta_equipo Then
                    mvarobjEquipo.Carga mvarobjEquipo.getID_EQUIPO
                    des_activar_campos True
                    Form_Load
                End If
            Else
                optBaja.Value = True
            End If
        End If
    End If
End Sub
'M1326-I
Private Sub optAbrir_Click()
    If Not UltimoMovimientoApertura Then
        If optAbrir.Value Then
            If MsgBox("Se ha solicitado la apertura del equipo. ¿Desea continuar?", vbInformation + vbYesNo) = vbYes Then
                If mvarobjEquipo.abrir_cerrar_equipo(1) Then
                    mvarobjEquipo.Carga mvarobjEquipo.getID_EQUIPO
                    des_activar_campos True
                    Form_Load
                End If
            Else
                optCerrar.Value = True
            End If
        End If
    End If
End Sub
Private Sub optCerrar_Click()
    If UltimoMovimientoApertura Then
        If optCerrar.Value Then
            If MsgBox("Se ha solicitado el cierre del equipo. ¿Desea continuar?", vbInformation + vbYesNo) = vbYes Then
                If mvarobjEquipo.abrir_cerrar_equipo(0) Then
                    mvarobjEquipo.Carga mvarobjEquipo.getID_EQUIPO
                    des_activar_campos True
                    Form_Load
                End If
            Else
                optAbrir.Value = True
            End If
        End If
    End If
End Sub

Private Function UltimoMovimientoApertura() As Boolean
    Dim rs As ADODB.Recordset
    UltimoMovimientoApertura = False
    Set rs = mvarobjEquipo.ultimo_movimiento_apertura()
    If rs.RecordCount > 0 Then
       If rs("RAZON_ID") = EQUIPOS_EVENTOS_MOTIVOS.EVTR_APERTURA Then
          UltimoMovimientoApertura = True
       End If
    End If
End Function
'M1326-F


Private Sub optiponorma_Click(Index As Integer)
'    cmbNormas.Limpiar
    If Index = 0 Then
        cmbDocumentos.visible = True
        cmbNormas.visible = False
'        llenar_combo cmbNormas, New clsCa_documentos, 0, frmCA_Documento, ""
    Else
        cmbDocumentos.visible = False
        cmbNormas.visible = True
'        llenar_combo cmbNormas, New clsCa_normas, 0, frmCA_Normas, ""
    End If
End Sub


Private Sub tabPrincipal_Click(PreviousTab As Integer)
    If tabPrincipal.Tab = 0 Then Exit Sub
    
'    If mvarenuTipoEdicion = ALTA Then
'        If MsgBox("Antes de Continuar debe guardar los datos principales del equipo. ¿Desea guardar ahora?", vbInformation + vbYesNo, "Guardar Datos de Equipo") = vbNo Then
'            tabPrincipal.Tab = 0
'        Else
'            cmdok_Click
'        End If
'    End If
    
    Select Case tabPrincipal.Tab
        Case ColsTAB.COL_CALIBRACIONES  ' Calibraciones
            PresentarDatos_Calibracion
        Case ColsTAB.COL_VERIFICACIONES  ' Verificaciones
            PresentarDatos_Verificacion
        Case ColsTAB.COL_MANTENIMIENTO  ' Mantenimientos
            PresentarDatos_Mantenimiento
        Case ColsTAB.COL_EVENTOS  ' Eventos
            PresentarDatos_Eventos
        Case ColsTAB.COLS_USO  ' Usos
            PresentarDatos_Usos
    End Select
End Sub

'J1
'Private Sub txtCalibracionFechaProxima_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
'txtCalibracionFechaProxima_b.Text = Format(txtCalibracionFechaProxima.value, "dd/mm/yyyy")
'End Sub

'Private Sub txtConformeNormas_execQuery(rst As ADODB.RecordSet, ByVal Text As String)
'Dim strConsulta As String
'
'strConsulta = "SELECT * FROM ca_normas WHERE concat(lower(nombre),lower(codigo)) like '%" & LCase(Text) & "%'"
'Set rst = datos_bd(strConsulta)
'
'End Sub

Private Sub cmbMtoTipo_Click()
    Call cmbMtoTipo_Change
End Sub

Private Sub cmbMtoTipo_Change()

    mvarblnMtoHistorico = (cmbMtoTipo.ItemData(cmbMtoTipo.ListIndex) = 0)
    ' Solo se habilita el botón cuando no es histórico
    cmdAnadirMto.Enabled = Not mvarblnMtoHistorico
    Call PresentarDatos_Mantenimiento
    'JGM PresentarDatos_Eventos
End Sub

Private Sub cmdAnadirCalibracion_Click()
Dim objfrm  As New frmEquipoEdicionCalibracion
Dim strId As String, fila As Long

strId = "0"

Set objfrm.EQUIPO = mvarobjEquipo
objfrm.ID = strId
objfrm.TipoEdicion = Alta
objfrm.idCalibradorInternoInicial = getDataComboSel(cmbCalibracionResponsable)
objfrm.idProcedmientoInicial = cmbProcedimientoCal.getPK_SALIDA

If txtCalibracionFechaProxima.Value = CDate("1900-01-01") Then
    objfrm.FechaProximaInicial = Now
Else
    objfrm.FechaProximaInicial = txtCalibracionFechaProxima.Value
End If
objfrm.IdPeriodoInicial = getDataComboSel(cmbCalibracionPeriod)
objfrm.IdTipoCalibracionIncial = getDataComboSel(cmbCalibracionTipo)

    objfrm.Show vbModal

If Not objfrm.resultado Then
    Unload objfrm
    Set objfrm = Nothing
    Exit Sub
End If

Set mvarobjEquipo = objfrm.EQUIPO

Call PresentarDatos_Calibracion(True)
PresentarDatos_Eventos
' Por si han cambiado las limitaciones de uso
Call PresentarDatos_LimitacionesUso
 



End Sub

Private Sub cmdAnadirMto_Click()
Dim objfrm  As New frmEquipoEdicionMtoFechasEdicion
Dim lngFila As String

    If lstPlanes.ListCount <= 0 Then
        ' Debe Señalar el plan
        MsgBox "Para añadir un mantenimiento debe seleccionar previamente algún Plan de Mantenimiento", vbInformation, "Añadir Plan de Mantenimiento"
        Exit Sub
    End If
        
    
    Call RecogerDatos
    

    objfrm.TipoEdicion = Alta
    Set objfrm.EQUIPO = mvarobjEquipo
    objfrm.id_mantenimiento = 0
    
    objfrm.MostrarCierre = True
    
    objfrm.Show vbModal
    
    If objfrm.resultado Then
        Set mvarobjEquipo = objfrm.EQUIPO
        PresentarDatos_Mantenimiento True
        PresentarDatos_Eventos
    End If
    
    Unload objfrm
    Set objfrm = Nothing

End Sub

Private Sub cmdAnadirVerificacion_Click()
Dim objfrm  As New frmEquipoEdicionVerificacion
Dim strId As String, fila As Long

strId = "0"


Set objfrm.EQUIPO = mvarobjEquipo
objfrm.ID = strId
objfrm.TipoEdicion = Alta
objfrm.idVerificadorInternoInicial = getDataComboSel(cmbVerificacionResponsable)
objfrm.idProcedmientoInicial = cmbProcedimientoVer.getPK_SALIDA
If txtVerificacionFechaProxima.Value = CDate("1900-01-01") Then
    objfrm.FechaProximaInicial = Now
Else
    objfrm.FechaProximaInicial = txtVerificacionFechaProxima.Value
End If
objfrm.IdPeriodoInicial = getDataComboSel(cmbVerificacionPeriod)
objfrm.IdTipoVerificacionIncial = getDataComboSel(cmbVerificacionTipo)

    objfrm.Show vbModal

If Not objfrm.resultado Then
    Unload objfrm
    Set objfrm = Nothing
    Exit Sub
End If

Set mvarobjEquipo = objfrm.EQUIPO

PresentarDatos_Verificacion True
PresentarDatos_Eventos
' Por si han cambiado las limitaciones de uso
Call PresentarDatos_LimitacionesUso

End Sub


Private Sub cmdGenerarFechasMto_Click()
Dim objfrm As New frmEquipoEdicionMtoFechas
Dim objPlan As clsPlanMantenimiento
Dim lngIdPlan As Long
Dim objItem As New clsEquipoMantenimiento, objCol As New clsGenericCollection

    
    If lstPlanes.ListCount <= 0 Then
        MsgBox "Debe añadir al menos un Plan de Mantenimiento para generar las fechas de Mantenimiento.", vbInformation, "Generar Fechas Plan Mantenimiento"
        Set objfrm = Nothing
        Exit Sub
    End If
    
    Call RecogerDatos
       
    Set objfrm.EQUIPO = mvarobjEquipo
    
    objfrm.idResponsable = getDataComboSel(cmbMtoResponsable)
    
    objfrm.Show vbModal
    
    If objfrm.resultado Then
    
        Set mvarobjEquipo = objfrm.EQUIPO
        
        Call PresentarDatos_Mantenimiento
        PresentarDatos_Eventos
    End If

If objfrm.resultado Then
End If

Unload objfrm
Set objfrm = Nothing

End Sub


Private Sub cmdVerSituacionEnPlano_Click()
    Dim objfrm As New frmEquipoPlanoLocalizacion
    
    If chkMostrarEnPlano.Value = vbUnchecked Then
        MsgBox "Este equipo no tiene ubicación en el Plano de Situación", vbInformation, "Ver Equipo en Plano de Situación"
        Exit Sub
    End If
    
    objfrm.PK = txtIdEquipo.Text
    
    Set objfrm.EQUIPO = mvarobjEquipo
    objfrm.VerTodo = False
    objfrm.Show vbModal
    
'    If objfrm.RESULTADO Then
'        Set mvarobjEquipo = objfrm.EQUIPO
'    End If
    
    Unload objfrm
    Set objfrm = Nothing

End Sub
Private Sub chkCon_Calibracion_Click()
    Call estado_frame_calibracion
End Sub

Private Sub chkCon_Verificacion_Click()
    Call estado_frame_verificacion
End Sub

Private Sub chkCon_Mantenimiento_Click()
    Call estado_frame_mantenimiento
End Sub

Private Sub estado_frame_calibracion()
    lblcalibracion.visible = (chkCon_Calibracion.Value = vbChecked)
    imgcalibracion.visible = (chkCon_Calibracion.Value = vbChecked)
           
    cmdAnadirCalibracion.Enabled = (chkCon_Calibracion.Value = vbChecked)
    cmdEliminarCalibracion.Enabled = (chkCon_Calibracion.Value = vbChecked)
    
    cmbCalibracionPeriod.Enabled = (chkCon_Calibracion.Value = vbChecked)
    cmbCalibracionTipo.Enabled = (chkCon_Calibracion.Value = vbChecked)
    cmbCalibracionResponsable.Enabled = (chkCon_Calibracion.Value = vbChecked)
    txtCalibracionFechaProxima_info.Enabled = (chkCon_Calibracion.Value = vbChecked)
    
    If chkCon_Calibracion.Value = vbChecked Then
        cmbProcedimientoCal.activar
    Else
        cmbProcedimientoCal.desactivar
    End If
    
    tabPrincipal.TabVisible(ColsTAB.COL_CALIBRACIONES) = (chkCon_Calibracion.Value = vbChecked)
    frmEstadoCalibracion.visible = (chkCon_Calibracion.Value = vbChecked)
End Sub

Private Sub estado_frame_verificacion()
    lblverificacion.visible = (chkCon_Verificacion.Value = vbChecked)
    imgverificacion.visible = (chkCon_Verificacion.Value = vbChecked)

    cmdAnadirVerificacion.Enabled = (chkCon_Verificacion.Value = vbChecked)
    cmdEliminarVerificacion.Enabled = (chkCon_Verificacion.Value = vbChecked)
    
    cmbVerificacionPeriod.Enabled = (chkCon_Verificacion.Value = vbChecked)
    cmbVerificacionTipo.Enabled = (chkCon_Verificacion.Value = vbChecked)
    cmbVerificacionResponsable.Enabled = (chkCon_Verificacion.Value = vbChecked)
    txtVerificacionFechaProxima_info.Enabled = (chkCon_Verificacion.Value = vbChecked)
    
    If chkCon_Verificacion.Value = vbChecked Then
        cmbProcedimientoVer.activar
    Else
        cmbProcedimientoVer.desactivar
    End If
    
    tabPrincipal.TabVisible(ColsTAB.COL_VERIFICACIONES) = (chkCon_Verificacion.Value = vbChecked)
    frmEstadoVerificacion.visible = (chkCon_Verificacion.Value = vbChecked)
    
End Sub

Private Function estado_frame_mantenimiento()
    lblmantenimiento.visible = (chkCon_Mantenimiento.Value = vbChecked)
    imgmantenimiento.visible = (chkCon_Mantenimiento.Value = vbChecked)
    
    cmdAnadirMto.Enabled = (chkCon_Mantenimiento.Value = Checked)
    cmdEliminarMto.Enabled = (chkCon_Mantenimiento.Value = Checked)
    cmdGenerarFechasMto.Enabled = (chkCon_Mantenimiento.Value = Checked)
    cmbMtoPlan.Enabled = (chkCon_Mantenimiento.Value = Checked)
    cmdAnadirPlan.Enabled = (chkCon_Mantenimiento.Value = Checked)
    cmdEliminarPlan.Enabled = (chkCon_Mantenimiento.Value = Checked)
    
    imgBuscarPlanMto.Enabled = (chkCon_Mantenimiento.Value = Checked)
    imgNuevoPlanMto.Enabled = (chkCon_Mantenimiento.Value = Checked)
    
    tabPrincipal.TabVisible(ColsTAB.COL_MANTENIMIENTO) = (chkCon_Mantenimiento.Value = vbChecked)
    frmEstadoMantenimiento.visible = (chkCon_Mantenimiento.Value = vbChecked)
    
End Function

Private Sub cmbProveedor_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    
    configuraGrid
    
    tabPrincipal.TabVisible(ColsTAB.COL_DOCUMENTACION) = False

    Call cargar_combos
    
    Call PresentarDatos
        
    Call desactivar_campos_info
    'M1146-I
    cmdReabrir.Enabled = False
    cmdReabrirCalibra.Enabled = False
    'M1146-F
    'M1326-I (FORMATO DE PREGUNTA PROVISIONAL. 1213 Cámara de niebla ASCOTT)
    habilitar_apertura
    'M1326-F
    
    tabPrincipal.Tab = ColsTAB.COL_GENERAL
End Sub
'M1326-I
Private Sub habilitar_apertura()
    Dim opar As New clsParametros
    Dim Equipos() As String
    Dim EQUIPO As Long
    Dim i As Integer
    Dim equipoEnLista As Boolean
    EQUIPO = mvarobjEquipo.getID_EQUIPO
    'JGM-I
'    If (opar.Carga(300, "")) Then
    If (opar.Carga(PARAM_EQUIPOS_APERTURA_CIERRE, "")) Then
        Equipos = Split(opar.getVALOR, ";")
'    End If
    'JGM-F
        equipoEnLista = False
        For i = LBound(Equipos) To UBound(Equipos) - 1
            If EQUIPO = CLng(Equipos(i)) Then
                equipoEnLista = True
            End If
        Next i
        
        If equipoEnLista Then
            fraApertura.visible = True
        
            If Not UltimoMovimientoApertura Then
               optCerrar.Value = True
            Else
               optAbrir.Value = True
            End If
        End If
    End If
    Set opar = Nothing
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set mvarobjEquipo = Nothing
End Sub
Private Sub grdCalibraciones_DblClick()
    If cmbTipoEquipo.BoundText = EQ_TIPOS_EQUIPOS.TIPO_EQUIPO_TORCOMETRO Then
        MsgBox "Los TORCOMETROS deben gestionarse en GESMET.", vbExclamation, App.Title
        Exit Sub
    End If
    
    Dim objfrm  As New frmEquipoEdicionCalibracion
    Dim lngFila As Long, strId As String
    Dim intEstado As Integer
    
    lngFila = grdCalibraciones.RowSel
    If lngFila < 1 Then Exit Sub
    
    strId = grdCalibraciones.TextMatrix(lngFila, 0)
    intEstado = CLng(grdCalibraciones.TextMatrix(lngFila, COLS.col_id_estado))
    
    With objfrm
        Set .EQUIPO = mvarobjEquipo
        .ID = strId
        If intEstado = 0 Or intEstado = 3 Then
            .TipoEdicion = EDICION ' si no está cerrado
        Else
            .TipoEdicion = visualizar
        End If
                
        .Show vbModal
        
        If .resultado Then
            Set mvarobjEquipo = .EQUIPO
            'mvarobjEquipo.Carga_Calibraciones
            Call PresentarDatos_Calibracion(True)
            PresentarDatos_Eventos
            ' Por si han cambiado las limitaciones de uso
            Call PresentarDatos_LimitacionesUso
        End If
    End With
    Unload objfrm
    Set objfrm = Nothing
End Sub

Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora

    'M1124-I
    cargar_combo cmbEstado, New clsEq_Estados
    cargar_combo cmbCentro, New clsCentros
    'M1124-F
    oDeco.cargar_combo cmbFamilia, DECODIFICADORA.EQ_FAMILIAS
    oDeco.cargar_combo cmbSituacion, DECODIFICADORA.EQ_SITUACIONES
    oDeco.cargar_combo cmbProcedencia, DECODIFICADORA.EQ_PROCEDENCIA_EQUIPOS
    oDeco.cargar_combo cmbCalibracionPeriod, DECODIFICADORA.EQ_periodicidad
    oDeco.cargar_combo cmbVerificacionPeriod, DECODIFICADORA.EQ_periodicidad
    oDeco.cargar_combo cmbTipoEquipo, DECODIFICADORA.EQ_TIPOS_EQUIPO
    oDeco.cargar_combo cmbCalibracionTipo, DECODIFICADORA.EQ_TIPO_CALIBRACION
    oDeco.cargar_combo cmbVerificacionTipo, DECODIFICADORA.EQ_TIPO_CALIBRACION
    cargar_combo cmbMtoPlan, New clsPlanMantenimiento
    llenar_combo cmbUnidades, New clsUnidades, 0, Me, ""
    cargar_combo cmbProveedor, New clsProveedor
'M1050-I Cambiar las combos de usuario y decodificadoras para mejorar rendimiento
'    cargar_combo cmbResponsable, New clsUsuarios
'    cargar_combo cmbCalibracionResponsable, New clsUsuarios
'    cargar_combo cmbVerificacionResponsable, New clsUsuarios
'    cargar_combo cmbMtoResponsable, New clsUsuarios
    Dim rs As ADODB.Recordset
    Dim oUsuario As New clsUsuarios
    Set rs = oUsuario.Listado_Combo
    
    Set cmbResponsable.RowSource = rs
    cmbResponsable.ListField = rs(1).Name
    cmbResponsable.BoundColumn = rs(0).Name
    Set cmbCalibracionResponsable.RowSource = rs
    cmbCalibracionResponsable.ListField = rs(1).Name
    cmbCalibracionResponsable.BoundColumn = rs(0).Name
    Set cmbVerificacionResponsable.RowSource = rs
    cmbVerificacionResponsable.ListField = rs(1).Name
    cmbVerificacionResponsable.BoundColumn = rs(0).Name
    Set cmbMtoResponsable.RowSource = rs
    cmbMtoResponsable.ListField = rs(1).Name
    cmbMtoResponsable.BoundColumn = rs(0).Name
    Set rs = Nothing
    
    llenar_combo cmbCliente, New clsCliente, 0, frmClientes, ""
    llenar_combo cmbTrazmet, New clsEq_trazmet, 0, Me, ""
'M1050-F
    llenar_combo cmbProcedimientoCal, New clsCa_documentos, 0, frmCA_Documento, " codigo like '%PNT C%'"
    llenar_combo cmbProcedimientoVer, New clsCa_documentos, 0, frmCA_Documento, ""
    
    optiponorma(0).Value = True
    llenar_combo cmbNormas, New clsCa_normas, 0, frmCA_Normas, ""
    llenar_combo cmbDocumentos, New clsCa_documentos, 0, frmCA_Documento, ""
    
    cmbTipoCalibraciones.Enabled = False
    cmbTipoVerificaciones.Enabled = False
    
    cmbTipoCalibraciones.ItemData(0) = 0 ' Historico
    cmbTipoCalibraciones.ItemData(1) = 1 ' Previsto
    cmbTipoCalibraciones.ItemData(2) = 2 ' Todos
    
    cmbTipoVerificaciones.ItemData(0) = 0
    cmbTipoVerificaciones.ItemData(1) = 1
    cmbTipoVerificaciones.ItemData(2) = 2
    cmbMtoTipo.ItemData(0) = 0
    cmbMtoTipo.ItemData(1) = 1

    ' JGM: Ver Todos
    cmbTipoCalibraciones.ListIndex = 2
    cmbTipoVerificaciones.ListIndex = 2
    cmbMtoTipo.ListIndex = 1
    
    cmbTipoCalibraciones.Enabled = True
    cmbTipoVerificaciones.Enabled = True
    cmbMtoTipo.Enabled = True
    
    Set mvarobjUltCalibracion = New clsEquipoCalibracion
    Set mvarobjUltMantenimiento = New clsEquipoMantenimiento
    Set mvarobjUltVerificacion = New clsEquipoVerificacion
    
End Sub

Private Sub des_activar_campos(ByVal prmEstado As Boolean)

    fraEstadoEquipo.Enabled = prmEstado
'    fraDatos(0).Enabled = prmEstado
'    fraDatos(3).Enabled = prmEstado
'    fraDatos(4).Enabled = prmEstado

    'Accesorios
    cmdAccesoriosAdd.Enabled = prmEstado
    cmdAccesoriosDelete.Enabled = prmEstado
    ' Normas
    If prmEstado Then
        cmbNormas.activar
    Else
        cmbNormas.desactivar
    End If
'    cmdInsertaNorma.Enabled = prmEstado
    cmdAnadirNorma.Enabled = prmEstado
    cmdEliminarNorma.Enabled = prmEstado
    ' PARTE CALIBRACIONES
    cmbCalibracionPeriod.Enabled = prmEstado
    cmbCalibracionResponsable.Enabled = prmEstado
    cmbCalibracionResponsable.Enabled = prmEstado
    If prmEstado Then
        cmbProcedimientoCal.activar
    Else
        cmbProcedimientoCal.desactivar
    End If
    cmdAnadirCalibracion.Enabled = prmEstado
    cmdEliminarCalibracion.Enabled = prmEstado
 '   cmbTipoCalibraciones.Enabled = prmEstado
    
    ' PARTE verificaciones
    cmbVerificacionPeriod.Enabled = prmEstado
    cmbVerificacionResponsable.Enabled = prmEstado
    cmbVerificacionResponsable.Enabled = prmEstado
    If prmEstado Then
        cmbProcedimientoVer.activar
    Else
        cmbProcedimientoVer.desactivar
    End If
    cmdAnadirVerificacion.Enabled = prmEstado
    cmdEliminarVerificacion.Enabled = prmEstado
'    cmbTipoVerificaciones.Enabled = prmEstado
    
    chkCon_Calibracion.Enabled = prmEstado
    chkCon_Mantenimiento.Enabled = prmEstado
    chkCon_Verificacion.Enabled = prmEstado
    
'    cmdetiqueta.Enabled = prmEstado
'    cmdImprimirFicha.Enabled = prmEstado
'     cmdDocumentacion.Enabled = prmEstado
'    cmdParametrosResultadosCalib.Enabled = prmEstado
'    cmdParametrosResultadosVerif.Enabled = prmEstado
    cmdok.Enabled = prmEstado
    'M1146-I
'JGM    If Not USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 0 And prmEstado = True Then
        cmdReabrir.Enabled = prmEstado
        cmdReabrirCalibra.Enabled = prmEstado
'JGM    Else
'JGM        cmdReabrir.Enabled = prmEstado
'JGM        cmdReabrirCalibra.Enabled = prmEstado
'JGM    End If
    'M1146-F
End Sub

Private Sub cmdEliminarNorma_Click()

    If grdNormas.ListItems.Count = 0 Then Exit Sub
    
    If MsgBox("¿Esta seguro de eliminar el documento?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        mvarobjEquipo.Eliminar_Norma_equipo grdNormas.ListItems(grdNormas.selectedItem.Index).Text, grdNormas.ListItems(grdNormas.selectedItem.Index).SubItems(3)
        Call PresentarDatos_Normas
    End If

End Sub

Public Property Get PK() As Long

    PK = mvarlngPK
    
    ' Por defecto, se edita
    mvarenuTipoEdicion = EDICION
    
    Set mvarobjEquipo = New clsEquipos
    
    mvarobjEquipo.Carga mvarlngPK
    
End Property

Public Property Let PK(ByVal lngPK As Long)

    mvarlngPK = lngPK

    Set mvarobjEquipo = New clsEquipos
    Call mvarobjEquipo.Carga(lngPK)
    
    mvarenuTipoEdicion = EDICION
    Call Form_Load

End Property

Private Sub configuraGrid()

    With grdAccesorios
        .COLS = 3
        .ColWidth(0) = .Width * 0.15
        .ColWidth(1) = .Width * 0.55
        .ColWidth(2) = .Width * 0.25
        .TextMatrix(0, 0) = "Nº Eq."
        .TextMatrix(0, 1) = "Accesorio"
        .TextMatrix(0, 2) = "Estado"
        .Rows = 1
    End With

    grdNormas.ColumnHeaders.Clear
    With grdNormas.ColumnHeaders
        .Add , , "ID", 0, lvwColumnLeft
        .Add , , "Nombre", grdNormas.Width * 0.99, lvwColumnLeft
        .Add , , "Ruta", 0, lvwColumnLeft
        .Add , , "Tipo", 0, lvwColumnLeft
    End With

    With grdCalibraciones
        .COLS = 10
        .ColWidth(0) = 0
        
        .ColWidth(1) = .Width * 0.1  ' Fecha
        .ColWidth(2) = .Width * 0.21 ' Responsable
        .ColWidth(3) = .Width * 0.3 ' Procedimiento
        .ColWidth(4) = .Width * 0.12 ' Periodicidad
        .ColWidth(5) = .Width * 0.12 ' Estado
        .ColWidth(6) = .Width * 0.12 ' Resultado
        .ColWidth(7) = 0 ' Fecha Para ordenar
        .ColWidth(8) = 0 ' id_estado
        .ColWidth(9) = 0 ' PERIODICIDAD_ID
        
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Responsable"
        .TextMatrix(0, 3) = "Procedimiento"
        .TextMatrix(0, 4) = "Periodicidad"
        .TextMatrix(0, 5) = "Estado"
        .TextMatrix(0, 6) = "Resultado"
        
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .Rows = 1
    End With
    
    With grdVerificaciones
        .COLS = 10
        .ColWidth(0) = 0
        .ColWidth(1) = .Width * 0.1 ' Fecha
        .ColWidth(2) = .Width * 0.2 ' Responsable
        .ColWidth(3) = .Width * 0.3 ' Procedimiento
        .ColWidth(4) = .Width * 0.12 ' pERIODICIDAD
        .ColWidth(5) = .Width * 0.12 ' Estado
        .ColWidth(6) = .Width * 0.12 ' Resultado
        .ColWidth(7) = 0 ' Fecha Para ordenar
        .ColWidth(8) = 0 ' id_estado
        .ColWidth(9) = 0 ' PERIODICIDAD_ID
        
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Responsable"
        .TextMatrix(0, 3) = "Procedimiento"
        .TextMatrix(0, 4) = "Periodicidad"
        .TextMatrix(0, 5) = "Estado"
        .TextMatrix(0, 6) = "Resultado"
        
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignCenterCenter
        
        .Rows = 1
        
    End With
    
    With grdMantenimientos
        .COLS = 10
        .ColWidth(0) = 0
        .ColWidth(1) = .Width * 0.1  ' Fecha
        .ColWidth(2) = .Width * 0.2 ' Responsable
        .ColWidth(3) = .Width * 0.3  ' Procedimiento
        .ColWidth(4) = .Width * 0.2 ' Plan
        .ColWidth(5) = .Width * 0.16 ' Periodicidad
        .ColWidth(6) = 0 ' Fecha Para ordenar
        .ColWidth(7) = 0 ' id_estado
        .ColWidth(8) = 0 ' ID_PLAN_MTO
        .ColWidth(9) = 0 ' ID_PLAN_MTO
        
        .TextMatrix(0, 1) = "Fecha"
        .TextMatrix(0, 2) = "Responsable"
        .TextMatrix(0, 3) = "Procedimiento"
        .TextMatrix(0, 4) = "Plan"
        .TextMatrix(0, 5) = "Periodicidad"
        .Rows = 1
    End With
    
    With grdEventos
        .COLS = 6
        .ColWidth(0) = 0
        .ColWidth(1) = .Width * 0.2 ' Evento
        .ColWidth(2) = .Width * 0.2 ' Motivo
        .ColWidth(3) = .Width * 0.15 ' Fecha/hora
        .ColWidth(4) = .Width * 0.15 ' Responsable
        .ColWidth(5) = .Width * 0.2 ' Observaciones
        
        .TextMatrix(0, 1) = "Evento"
        .TextMatrix(0, 2) = "Motivo"
        .TextMatrix(0, 3) = "Fecha/Hora"
        .TextMatrix(0, 4) = "Resp. Evento"
        .TextMatrix(0, 5) = "Observaciones"
        .Rows = 1
        
    End With
    
End Sub


Private Sub PresentarDatos()
    mvarblnMtoHistorico = True
    
    If mvarenuTipoEdicion = Alta Then
        mvarobjEquipo.CrearID
        txtIdEquipo.Text = Format(mvarobjEquipo.getID_EQUIPO, "00000")
        txtFechaRecepcion.Value = Date
        txtestado.BackColor = &H80C0FF
        txtestado.Text = "PDTE.ALTA"
        cpRequiere.PanelOpen = True
        cpEstado.PanelOpen = False
        tabPrincipal.TabVisible(ColsTAB.COL_CALIBRACIONES) = False
        tabPrincipal.TabVisible(ColsTAB.COL_VERIFICACIONES) = False
        tabPrincipal.TabVisible(ColsTAB.COL_EVENTOS) = False
        tabPrincipal.TabVisible(ColsTAB.COLS_USO) = False
        frmBolas.visible = False
        
        txtFechaPuestaServicio.visible = False
        lblCampos(57).visible = False
        
        cmdRecepcion.visible = False
        cmdEspecificaciones.visible = False
        cmdEtiqueta.visible = False
'        cmdImprimirFicha.Visible = False
        cmdDocumentacion.visible = False
        
        fraCondicionesAmbientales.visible = False
        
        Set mvarobjUltCalibracion = New clsEquipoCalibracion
        Set mvarobjUltVerificacion = New clsEquipoVerificacion
        Set mvarobjUltMantenimiento = New clsEquipoMantenimiento
        Exit Sub
    End If
    
    frmBolas.visible = True
        
    cmdRecepcion.visible = True
    cmdEspecificaciones.visible = True
    cmdEtiqueta.visible = True
'    cmdImprimirFicha.Visible = True
    cmdDocumentacion.visible = True
    
    ' Habilita las limitaciones de uso
    txtLimitacionesUso.Text = ""
    txtLimitacionesUso.Enabled = True
    cmdAnadirLimitacion.Enabled = True
    cmdEliminarLimitacion.Enabled = True
    
    ' cuando es edicion, se puede ver en plano.
    cmdVerSituacionEnPlano.Enabled = True
    
    PresentarDatos_Generales
    PresentarDatos_Calibracion_Siguiente
    PresentarDatos_Verificacion_Siguiente
    PresentarDatos_Mantenimiento_Siguiente
    
    PERMITIR_CAL_VER_PREVISTAS = True

    If mvarobjEquipo.getFUERA_SERVICIO = 1 Or mvarobjEquipo.getALTA_BAJA = 1 Then
        PERMITIR_CAL_VER_PREVISTAS = False
        des_activar_campos False
        ' a la responsable de metrología le permite poner de alta el equipo de nuevo
        If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 1 Then
            fraEstadoEquipo.Enabled = True
        End If
    End If
    
        
    If USUARIO.getRESPONSABLE_DEPARTAMENTOS(enumDPTO.METROLOGIA) = 1 Then
        cpRequiere.CanExpand = True
    Else
        cpRequiere.CanExpand = False
    End If
    
    
    If mvarobjEquipo.getFUERA_SERVICIO = 1 Then
        fraDatos(4).Enabled = True
        ' deja hacer verificaciones y modificaciones previstas
        PERMITIR_CAL_VER_PREVISTAS = True
        cmbTipoCalibraciones.Enabled = True
        cmdAnadirCalibracion.Enabled = True
        cmbTipoVerificaciones.Enabled = True
        cmdAnadirVerificacion.Enabled = True
    End If

    If chkCon_Calibracion.Value = Unchecked And _
       chkCon_Verificacion.Value = Unchecked And _
       chkCon_Mantenimiento.Value = Unchecked Then
        If cpEstado.PanelOpen = True Then
            cpEstado.PanelOpen = False
        End If
    End If
End Sub

Public Property Get TipoEdicion() As enumTipoEdicion

    TipoEdicion = mvarenuTipoEdicion

End Property

Public Property Let TipoEdicion(ByVal enuTipoEdicion As enumTipoEdicion)

    mvarenuTipoEdicion = enuTipoEdicion

End Property

'Private Sub PresentarDatos_ResumenCalVerMto()
'    With mvarobjEquipo
'        chkCon_Calibracion.Value = .getCON_CALIBRACION
'        chkCon_Verificacion.Value = .getCON_VERIFICACION
'        chkCon_Mantenimiento.Value = .getCON_MANTENIMIENTO
'        cmbCalibracionPeriod.BoundText = .getPERIODICIDAD_CALIBRACION_ID
'        cmbCalibracionTipo.BoundText = .getTIPO_CALIBRACION_ID
'        If .getFECHA_PROX_CALIBRACION <> "" Then txtCalibracionFechaProxima.Value = .getFECHA_PROX_CALIBRACION
'        cmbCalibracionResponsable.BoundText = .getCALIBRADOR_INTERNO_ID
'        cmbProcedimientoCal.MostrarElemento .getPROCEDIMIENTO_CALIBRACION_ID
'
'        cmbVerificacionTipo.BoundText = .getTIPO_VERIFICACION_ID
'        cmbVerificacionPeriod.BoundText = .getPERIODICIDAD_VERIFICACION_ID
'        If .getFECHA_PROX_VERIFICACION <> "" Then txtVerificacionFechaProxima.Value = .getFECHA_PROX_VERIFICACION
'        cmbVerificacionResponsable.BoundText = .getVERIFICADOR_INTERNO_ID
'        cmbProcedimientoVer.MostrarElemento .getPROCEDIMIENTO_VERIFICACION_ID
'
'        cmbMtoPlan.BoundText = .getPLAN_MANTENIMIENTO_ID
'        If .getFECHA_PROX_MANTENIMIENTO <> "" Then txtMtoFechaProxima.Value = .getFECHA_PROX_MANTENIMIENTO
'        txtMtoFechaProxima_info.Text = Format(.getFECHA_PROX_MANTENIMIENTO, "dd/mm/yyyy")
'        cmbMtoResponsable.BoundText = .getMANTENEDOR_ID
'
'        chkCon_Calibracion_Click
'        chkCon_Verificacion_Click
'        chkCon_Mantenimiento_Click
'    End With
'End Sub

Public Property Get EQUIPO() As clsEquipos

    Set EQUIPO = mvarobjEquipo

End Property

Public Property Set EQUIPO(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

Private Sub PresentarDatos_Calibracion(Optional ByVal Actualizar_Fecha_Prox_bd As Boolean = False)
        
    Dim fecha_prox As String, fecha_ult As String, responsable_ult As String, responsable_ult_id As String
    Dim rs As ADODB.Recordset, rs_prev As ADODB.Recordset
        
   On Error GoTo PresentarDatos_Calibracion_Error

    grdCalibraciones.Rows = 1
    
    Set rs = mvarobjEquipo.devolver_lista_calibraciones(mvarobjEquipo.getID_EQUIPO, fecha_ult, responsable_ult, responsable_ult_id, cmbTipoCalibraciones.ItemData(cmbTipoCalibraciones.ListIndex))
        
    If rs.RecordCount = 0 Then Exit Sub
        
    rs.MoveFirst
    With grdCalibraciones
        .Redraw = False
        While Not rs.EOF
            
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rs!ID_CALIBRACION
            .TextMatrix(.Rows - 1, COLS.COL_FECHA) = Format(rs!fecha_actual, "dd/mm/yyyy")
            .TextMatrix(.Rows - 1, COLS.COL_RESPONSABLE) = rs!responsable
            .TextMatrix(.Rows - 1, COLS.COL_PROCEDIMIENTO) = rs!PROCEDIMIENTO
            .TextMatrix(.Rows - 1, COLS.COL_RESULTADOS_OBSERVACIONES) = rs!PERIODICIDAD
            .TextMatrix(.Rows - 1, COLS.COL_ID_PERIODICIODAD) = rs!PERIODICIDAD_ID
            .TextMatrix(.Rows - 1, COLS.COL_ESTADO) = rs!ESTADO_TEXTO
            If CInt(rs!ESTADO) <> CAL_ESTADOS.CAL_ESTADO_PREVISTA Then
                .TextMatrix(.Rows - 1, COLS.COL_RESULTADO) = rs!RESULTADO_TEXTO
            Else
                .TextMatrix(.Rows - 1, COLS.COL_RESULTADO) = ""
            End If
            
            Select Case CInt(rs!ESTADO)
            Case CAL_ESTADOS.CAL_ESTADO_PREVISTA
                colorear_fila grdCalibraciones, &H80FF80
            Case CAL_ESTADOS.CAL_ESTADO_REALIZADA
                If CInt(rs!resultado) = CAL_RESULTADOS.CAL_RESULTADO_NO_CONFORME Then
                    colorear_fila grdCalibraciones, vbRed
                ElseIf CInt(rs!resultado) = CAL_RESULTADOS.CAL_RESULTADO_REQ_AJUSTE Then
                    colorear_fila grdCalibraciones, vbBlue
                End If
            End Select
            .TextMatrix(.Rows - 1, .COLS - 2) = Format(rs!fecha_actual, "yyyy/mm/dd")
            .TextMatrix(.Rows - 1, COLS.col_id_estado) = CStr(CInt(rs!ESTADO))
            rs.MoveNext
        Wend
        .Redraw = True
    End With
    chkCon_Calibracion_Click
    PresentarDatos_Calibracion_Siguiente

   On Error GoTo 0
   Exit Sub

PresentarDatos_Calibracion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_Calibracion of Formulario frmEquipoEdicion"
End Sub
Private Sub PresentarDatos_Calibracion_Siguiente()
        
'    Dim fecha_prox As String, fecha_ult As String, responsable_ult As String, responsable_ult_id As String
'    Dim rs_prev As ADODB.Recordset
        
    If mvarobjEquipo.getCON_CALIBRACION = 1 Or chkCon_Calibracion.Value = Checked Then
        chkCon_Calibracion.Value = vbChecked
        ' Datos generales de la calibracion
        cmbCalibracionPeriod.BoundText = mvarobjEquipo.getPERIODICIDAD_CALIBRACION_ID
        cmbCalibracionResponsable.BoundText = mvarobjEquipo.getCALIBRADOR_INTERNO_ID
        cmbCalibracionTipo.BoundText = mvarobjEquipo.getTIPO_CALIBRACION_ID
        cmbProcedimientoCal.MostrarElemento mvarobjEquipo.getPROCEDIMIENTO_CALIBRACION_ID
        ' Proxima calibracion
        If Not IsNull(mvarobjEquipo.getFECHA_PROX_CALIBRACION) And IsDate(mvarobjEquipo.getFECHA_PROX_CALIBRACION) Then
            txtCalibracionFechaProxima.Value = mvarobjEquipo.getFECHA_PROX_CALIBRACION
            txtCalibracionFechaProxima_info.Text = mvarobjEquipo.getFECHA_PROX_CALIBRACION
            fproxcalibracion = mvarobjEquipo.getFECHA_PROX_CALIBRACION
            If CDate(mvarobjEquipo.getFECHA_PROX_CALIBRACION) > Date Then
                Set imgcalibracion.Picture = imagenes.ListImages(1).Picture
            Else
                Set imgcalibracion.Picture = imagenes.ListImages(2).Picture
            End If
        Else
            Set imgcalibracion.Picture = imagenes.ListImages(3).Picture
        End If
        ' Ultima Calibracion
        Dim idc As Long
        idc = mvarobjEquipo.devolver_actual_calibracion(mvarobjEquipo.getID_EQUIPO)
        If idc <> 0 Then
            Dim oEC As New clsEquipoCalibracion
            oEC.Carga idc
            txtactcalibracion = idc
            txtrutacalibracion = oEC.getRUTA_CERTIFICADO
            If IsDate(oEC.getFECHA_ACTUAL) Then
                factcalibracion = oEC.getFECHA_ACTUAL
            End If
        End If
        
'        Set rs_prev = mvarobjEquipo.devolver_proxima_calibracion(mvarobjEquipo.getID_EQUIPO)
'        If rs_prev.RecordCount <> 0 Then
'            With rs_prev
'                ' fechas
'                txtCalibracionFechaProxima.Value = !fecha_prevista
'                If IsDate(!fecha_prevista) And !fecha_prevista <> "" Then
'                    If CDate(!fecha_prevista) > Date Then
'                        Set imgcalibracion.Picture = imagenes.ListImages(1).Picture
'                    Else
'                        Set imgcalibracion.Picture = imagenes.ListImages(2).Picture
'                    End If
'                Else
'                    Set imgcalibracion.Picture = imagenes.ListImages(3).Picture
'                End If
'                txtCalibracionFechaProxima_info.Text = !fecha_prevista
'                cmbCalibracionPeriod.BoundText = !PERIODICIDAD_ID
'                cmbCalibracionResponsable.BoundText = !CALIBRADOR_INTERNO_ID
'                cmbCalibracionTipo.BoundText = !TIPO_ID
'                cmbProcedimientoCal.MostrarElemento !PROCEDIMIENTO_ID
'                Dim idc As Long
'                idc = mvarobjEquipo.devolver_actual_calibracion(mvarobjEquipo.getID_EQUIPO)
'                If idc <> 0 Then
'                    Dim oEC As New clsEquipoCalibracion
'                    oEC.Carga idc
'                    txtactcalibracion = idc
'                    txtrutacalibracion = oEC.getRUTA_CERTIFICADO
'                    If IsDate(oEC.getFECHA_ACTUAL) Then
'                        factcalibracion = oEC.getFECHA_ACTUAL
'                    End If
'                    If IsDate(oEC.getFECHA_PROXIMA) Then
'                        fproxcalibracion = oEC.getFECHA_PROXIMA
'                    End If
'                End If
'            End With
'        Else
'            txtCalibracionFechaProxima.Value = "1900-01-01"
'            txtCalibracionFechaProxima_info.Text = ""
'            cmbCalibracionPeriod.BoundText = ""
'            cmbCalibracionResponsable.BoundText = ""
'            cmbCalibracionTipo.BoundText = ""
'            cmbProcedimientoCal.MostrarElemento 0
'            frmEstadoCalibracion.visible = False
'        End If
    End If
    Call estado_frame_calibracion
End Sub

Private Sub PresentarDatos_Verificacion(Optional ByVal Actualizar_Fecha_Prox_bd As Boolean = False)
        
    Dim fecha_prox As String, fecha_ult As String, responsable_ult As String, responsable_ult_id As String
    Dim rs As ADODB.Recordset, rs_prev As ADODB.Recordset
        
    Set rs = mvarobjEquipo.devolver_lista_verificaciones(mvarobjEquipo.getID_EQUIPO, fecha_ult, responsable_ult, responsable_ult_id, cmbTipoVerificaciones.ItemData(cmbTipoVerificaciones.ListIndex))
    grdVerificaciones.Rows = 1
    
    If rs.RecordCount = 0 Then Exit Sub
        
    rs.MoveFirst
    fecha_ult = DateSerial(1900, 1, 1)
    responsable_ult = ""
    
    With grdVerificaciones
        .Redraw = False
        While Not rs.EOF
            
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, 0) = rs!ID_VERIFICACION
            .TextMatrix(.Rows - 1, COLS.COL_FECHA) = Format(rs!fecha_actual, "dd/mm/yyyy")
            .TextMatrix(.Rows - 1, COLS.COL_RESPONSABLE) = rs!responsable
            .TextMatrix(.Rows - 1, COLS.COL_PROCEDIMIENTO) = rs!PROCEDIMIENTO
            .TextMatrix(.Rows - 1, COLS.COL_RESULTADOS_OBSERVACIONES) = rs!PERIODICIDAD
            .TextMatrix(.Rows - 1, COLS.COL_ID_PERIODICIODAD) = rs!PERIODICIDAD_ID
            .TextMatrix(.Rows - 1, COLS.COL_ESTADO) = IIf(Not IsNull(rs!ESTADO_TEXTO), rs!ESTADO_TEXTO, "")
            
            If CInt(rs!ESTADO) <> VER_ESTADOS.VER_ESTADO_PREVISTA Then
                .TextMatrix(.Rows - 1, COLS.COL_RESULTADO) = IIf(Not IsNull(rs!RESULTADO_TEXTO), rs!RESULTADO_TEXTO, "")
            Else
                .TextMatrix(.Rows - 1, COLS.COL_RESULTADO) = ""
            End If
            Select Case CInt(rs!ESTADO)
            Case VER_ESTADOS.VER_ESTADO_PREVISTA
                colorear_fila grdVerificaciones, &H80FF80
            Case VER_ESTADOS.VER_ESTADO_REALIZADA
                If CInt(rs!resultado) = VER_RESULTADOS.VER_RESULTADO_NO_CONFORME Then
                    colorear_fila grdVerificaciones, vbRed
                ElseIf CInt(rs!resultado) = VER_RESULTADOS.VER_RESULTADO_REQ_AJUSTE Then
                    colorear_fila grdVerificaciones, vbBlue
                End If
            End Select
            .TextMatrix(.Rows - 1, .COLS - 2) = Format(rs!fecha_actual, "yyyy/mm/dd")
            .TextMatrix(.Rows - 1, COLS.col_id_estado) = CStr(CInt(rs!ESTADO))
            rs.MoveNext
        Wend
        .Redraw = True
    End With
    chkCon_Verificacion_Click
    PresentarDatos_Verificacion_Siguiente
End Sub
Private Sub PresentarDatos_Verificacion_Siguiente()
        
'    Dim fecha_prox As String, fecha_ult As String, responsable_ult As String, responsable_ult_id As String
'    Dim rs_prev As ADODB.Recordset
        
    If mvarobjEquipo.getCON_VERIFICACION = 1 Or chkCon_Verificacion.Value = Checked Then
        chkCon_Verificacion.Value = vbChecked
        ' Datos generales de la calibracion
        cmbVerificacionPeriod.BoundText = mvarobjEquipo.getPERIODICIDAD_VERIFICACION_ID
        cmbVerificacionResponsable.BoundText = mvarobjEquipo.getVERIFICADOR_INTERNO_ID
        cmbVerificacionTipo.BoundText = mvarobjEquipo.getTIPO_VERIFICACION_ID
        cmbProcedimientoVer.MostrarElemento mvarobjEquipo.getPROCEDIMIENTO_VERIFICACION_ID
        ' Proxima verificacion
        If Not IsNull(mvarobjEquipo.getFECHA_PROX_VERIFICACION) And IsDate(mvarobjEquipo.getFECHA_PROX_VERIFICACION) Then
            txtVerificacionFechaProxima.Value = mvarobjEquipo.getFECHA_PROX_VERIFICACION
            txtVerificacionFechaProxima_info.Text = mvarobjEquipo.getFECHA_PROX_VERIFICACION
            fproxverificacion = mvarobjEquipo.getFECHA_PROX_VERIFICACION
            If CDate(mvarobjEquipo.getFECHA_PROX_VERIFICACION) > Date Then
                Set imgverificacion.Picture = imagenes.ListImages(1).Picture
            Else
                Set imgverificacion.Picture = imagenes.ListImages(2).Picture
            End If
        Else
            Set imgverificacion.Picture = imagenes.ListImages(3).Picture
        End If
        ' Ultima Calibracion
        Dim idc As Long
        idc = mvarobjEquipo.devolver_actual_verificacion(mvarobjEquipo.getID_EQUIPO)
        If idc <> 0 Then
            Dim oEC As New clsEquipoVerificacion
            oEC.Carga idc
            txtactverificacion = idc
            txtrutaverificacion = oEC.getRUTA_CERTIFICADO
            If IsDate(oEC.getFECHA_ACTUAL) Then
                factverificacion = oEC.getFECHA_ACTUAL
            End If
        End If
'        Set rs_prev = mvarobjEquipo.devolver_proxima_verificacion(mvarobjEquipo.getID_EQUIPO)
'
'        If rs_prev.RecordCount <> 0 Then
'            With rs_prev
'                ' fechas
'                txtVerificacionFechaProxima.Value = !fecha_prevista
'
'                If IsDate(!fecha_prevista) And !fecha_prevista <> "" Then
'                    If CDate(!fecha_prevista) > Date Then
'                        Set imgverificacion.Picture = imagenes.ListImages(1).Picture
'                    Else
'                        Set imgverificacion.Picture = imagenes.ListImages(2).Picture
'                    End If
'                Else
'                    Set imgverificacion.Picture = imagenes.ListImages(3).Picture
'                End If
'                txtVerificacionFechaProxima_info.Text = !fecha_prevista
'                cmbVerificacionPeriod.BoundText = !PERIODICIDAD_ID
'                cmbVerificacionResponsable.BoundText = !VERIFICADOR_INTERNO_ID
'                cmbVerificacionTipo.BoundText = !TIPO_ID
'                cmbProcedimientoVer.MostrarElemento !PROCEDIMIENTO_ID
'                Dim idc As Long
'                idc = mvarobjEquipo.devolver_actual_verificacion(mvarobjEquipo.getID_EQUIPO)
'                If idc <> 0 Then
'                    Dim oEC As New clsEquipoVerificacion
'                    oEC.Carga idc
'                    txtactverificacion = idc
'                    txtrutaverificacion = oEC.getRUTA_CERTIFICADO
'                    factverificacion = oEC.getFECHA_ACTUAL
'                    fproxverificacion = oEC.getFECHA_PROXIMA
'                End If
'
'            End With
'
'        Else
'                txtVerificacionFechaProxima.Value = "1900-01-01"
'                txtVerificacionFechaProxima_info.Text = ""
'                cmbVerificacionPeriod.BoundText = ""
'                cmbVerificacionResponsable.BoundText = ""
'                cmbVerificacionTipo.BoundText = ""
'                cmbProcedimientoVer.MostrarElemento 0
'        End If
    End If
    
    estado_frame_verificacion
End Sub






Private Sub PresentarDatos_Mantenimiento(Optional ByVal Actualizar_Fecha_Prox_bd As Boolean = False)
    Dim objCol As clsGenericCollection, objItem As clsEquipoMantenimiento
    Dim dtmFechaProxima As Date, strPlan_Manteniemiento As String, strRESPONSABLE As String, strProcedimiento As String
    Dim rs As ADODB.Recordset

On Error GoTo PresentarDatos_Mantenimiento_Error

'    If mvarobjEquipo.getCON_MANTENIMIENTO = 1 Or chkCon_Mantenimiento.value = Checked Then
'        chkCon_Mantenimiento.value = vbChecked
'    End If
'    chkCon_Mantenimiento_Click
    'Set objCol = mvarobjEquipo.Mantenimientos
    
    mvarblnMtoHistorico = (cmbMtoTipo.ItemData(cmbMtoTipo.ListIndex) = 0)
    If mvarblnMtoHistorico Then
        Set rs = mvarobjEquipo.devolver_lista_mantenimientos_historicos
    Else
        Set rs = mvarobjEquipo.devolver_lista_mantenimientos_pendientes
    End If
   
    PresentarDatos_Mantenimiento_Planes
    
    grdMantenimientos.Rows = 1
    
    dtmFechaProxima = "31/12/2050"
    strPlan_Manteniemiento = ""
    strRESPONSABLE = ""
    strProcedimiento = ""
    
    If rs.RecordCount <> 0 Then
    grdMantenimientos.Redraw = False
    rs.MoveFirst
        While Not rs.EOF
            grdMantenimientos.Rows = grdMantenimientos.Rows + 1
            grdMantenimientos.TextMatrix(grdMantenimientos.Rows - 1, 0) = rs("ID_MANTENIMIENTO")
            grdMantenimientos.TextMatrix(grdMantenimientos.Rows - 1, COLS.COL_FECHA) = rs("FECHA_ACTUAL")
            grdMantenimientos.TextMatrix(grdMantenimientos.Rows - 1, grdMantenimientos.COLS - 1) = Format(CDate(rs("FECHA_ACTUAL")), "yyyy/mm/dd")
            grdMantenimientos.TextMatrix(grdMantenimientos.Rows - 1, COLS.COL_RESPONSABLE) = rs("RESPONSABLE")
            grdMantenimientos.TextMatrix(grdMantenimientos.Rows - 1, COLS.COL_PROCEDIMIENTO) = rs("PROTOCOLO")
            grdMantenimientos.TextMatrix(grdMantenimientos.Rows - 1, COLS.COL_PLAN) = rs("PLAN_MANTENIMIENTO")
            grdMantenimientos.TextMatrix(grdMantenimientos.Rows - 1, COLS.COL_PERIODICIDAD) = rs("PERIODICIDAD")
            grdMantenimientos.TextMatrix(grdMantenimientos.Rows - 1, COLS.col_id_estado) = rs("estado")
            grdMantenimientos.TextMatrix(grdMantenimientos.Rows - 1, COLS.COL_ID_PERIODICIODAD) = rs("PLANMTO_ID")
            
            Select Case CInt(rs!ESTADO)
            Case 0
                If Format(rs("FECHA_ACTUAL"), "YYYY-MM-DD") < Format(Date, "YYYY-MM-DD") Then
                    colorear_fila grdMantenimientos, &H8080FF
                End If
                colorear_fila grdCalibraciones, &H80FF80
            Case 1
                colorear_fila grdMantenimientos, &H80FF80
            End Select
            rs.MoveNext
        Wend
    End If
    
    PresentarDatos_Mantenimiento_Siguiente
    grdMantenimientos.Redraw = True

On Error GoTo 0
    Exit Sub
PresentarDatos_Mantenimiento_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_Mantenimiento of Formulario frmEquipoEdicion"
    Resume Next
End Sub
Private Sub PresentarDatos_Mantenimiento_Siguiente()
'    Dim objCol As clsGenericCollection, objItem As clsEquipoMantenimiento
'    Dim dtmFechaProxima As Date, strPlan_Manteniemiento As String, strRESPONSABLE As String, strProcedimiento As String
    Dim rs As ADODB.Recordset

    If mvarobjEquipo.getCON_MANTENIMIENTO = 1 Or chkCon_Mantenimiento.Value = Checked Then
        chkCon_Mantenimiento.Value = vbChecked
    End If
    chkCon_Mantenimiento_Click
    
    If IsNull(mvarobjEquipo.getFECHA_ULT_MANTENIMIENTO) Or mvarobjEquipo.getFECHA_ULT_MANTENIMIENTO = "" Then
        factmantenimiento.visible = False
    Else
        factmantenimiento.visible = True
        factmantenimiento.Value = Replace(mvarobjEquipo.getFECHA_ULT_MANTENIMIENTO, "'", "")
        Dim idc As Long

        idc = mvarobjEquipo.devolver_actual_mantenimiento(mvarobjEquipo.getID_EQUIPO)
        If idc <> 0 Then
            txtactmantenimiento = idc
        End If
    End If
    
    If IsNull(mvarobjEquipo.getFECHA_PROX_MANTENIMIENTO) Or mvarobjEquipo.getFECHA_PROX_MANTENIMIENTO = "" Then
        txtMtoFechaProxima.Value = Now
        txtMtoFechaProxima_info.Text = ""
        fproxmantenimiento.visible = False
        Set imgmantenimiento.Picture = imagenes.ListImages(1).Picture
    Else
        txtMtoFechaProxima.Value = Replace(mvarobjEquipo.getFECHA_PROX_MANTENIMIENTO, "'", "")
        txtMtoFechaProxima_info.Text = Replace(mvarobjEquipo.getFECHA_PROX_MANTENIMIENTO, "'", "")
        fproxmantenimiento.visible = True
        fproxmantenimiento = Replace(mvarobjEquipo.getFECHA_PROX_MANTENIMIENTO, "'", "")
        If mvarobjEquipo.getFECHA_PROX_MANTENIMIENTO > Date Then
            Set imgmantenimiento.Picture = imagenes.ListImages(1).Picture
        Else
           Set imgmantenimiento.Picture = imagenes.ListImages(2).Picture
        End If
        
    End If
    If Not IsNull(mvarobjEquipo.getRESPONSABLE_ID) And mvarobjEquipo.getRESPONSABLE_ID > 0 Then
        cmbMtoResponsable.BoundText = mvarobjEquipo.getRESPONSABLE_ID
    End If
    
    
'    Set rs = mvarobjEquipo.devolver_fecha_prox_mantenimiento()
'    Dim idc As Long
'    On Error Resume Next ' Puede que devuelva solamente una fila con todos los campos a nulos
'    If rs("FECHA_ACTUAL") = "" Then
'        txtMtoFechaProxima.Value = Now
'        txtMtoFechaProxima_info.Text = ""
'
'        Set imgmantenimiento.Picture = imagenes.ListImages(1).Picture
'        factmantenimiento = mvarobjEquipo.getFECHA_ULT_MANTENIMIENTO
'        If Not IsNull(mvarobjEquipo.getFECHA_PROX_MANTENIMIENTO) And mvarobjEquipo.getFECHA_PROX_MANTENIMIENTO <> "" Then
'            fproxmantenimiento = mvarobjEquipo.getFECHA_PROX_MANTENIMIENTO
'        Else
'            fproxmantenimiento.visible = False
'        End If
'            idc = mvarobjEquipo.devolver_actual_mantenimiento(mvarobjEquipo.getID_EQUIPO)
'            If idc <> 0 Then
'                txtactmantenimiento = idc
'            End If
'    Else
'        txtMtoFechaProxima.Value = rs("FECHA_ACTUAL")
'        txtMtoFechaProxima_info.Text = rs("FECHA_ACTUAL")
'
'        If IsDate(rs("FECHA_ACTUAL")) Then
'            If CDate(rs("FECHA_ACTUAL")) > Date Then
'                Set imgmantenimiento.Picture = imagenes.ListImages(1).Picture
'            Else
'                Set imgmantenimiento.Picture = imagenes.ListImages(2).Picture
'            End If
'        End If
'                idc = mvarobjEquipo.devolver_actual_mantenimiento(mvarobjEquipo.getID_EQUIPO)
'                If idc <> 0 Then
'                    Dim oEC As New clsEquipoMantenimiento
'                    oEC.carga idc
'                    txtactmantenimiento = idc
'                    factmantenimiento = oEC.getFECHA_ACTUAL
'                    fproxmantenimiento = rs("FECHA_ACTUAL")
'                End If
'
'    End If
'    cmbMtoResponsable.BoundText = rs("MANTENEDOR_ID")
    estado_frame_mantenimiento
End Sub


Private Sub PresentarDatos_Eventos()

On Error GoTo PresentarDatos_Eventos_Error
    
    Dim objEventos As New clsEquipoEventos, rs As ADODB.Recordset

    Set rs = objEventos.Listado(mvarobjEquipo.getID_EQUIPO)
    Set objEventos = Nothing

    grdEventos.Rows = 1
        
    If rs Is Nothing Then Exit Sub
    If rs.RecordCount = 0 Then Exit Sub
    
    With grdEventos
        .Redraw = False
        rs.MoveFirst
        
        While Not rs.EOF
            ' Presenta el histórico, los que sean distintos de 0
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = rs("ID_EVENTOEQUIPO")
            .TextMatrix(.Rows - 1, COLS.COL_EVENTO) = rs("EVENTO")
            .TextMatrix(.Rows - 1, COLS.COL_EVENTO_RAZON) = rs("RAZON")
            .TextMatrix(.Rows - 1, COLS.COL_EVENTO_FECHAHORA) = Format(rs("TS"), "dd/mm/yyyy Hh:Nn")
            .TextMatrix(.Rows - 1, COLS.COL_EVENTO_USUARIO) = rs("USUARIO")
            If Len(rs("observaciones")) > 100 Then
                .TextMatrix(.Rows - 1, COLS.COL_EVENTO_OBSERVACIONES) = Left(rs("OBSERVACIONES"), 100) & "..."
            Else
                .TextMatrix(.Rows - 1, COLS.COL_EVENTO_OBSERVACIONES) = rs("OBSERVACIONES")
            End If
            rs.MoveNext
        Wend
        .Redraw = True
    End With

On Error GoTo 0
    Exit Sub
PresentarDatos_Eventos_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PresentarDatos_Eventos of Formulario frmEquipoEdicion"
End Sub

Private Sub PresentarDatos_Usos()
    If listaUsos.ColumnHeaders.Count = 0 Then
        With listaUsos.ColumnHeaders
            .Add , , "Código", 900, lvwColumnLeft
            .Add , , "Cliente", 2200, lvwColumnLeft
            .Add , , "Tipo de Analisis/Solución", 3400, lvwColumnLeft
            .Add , , "Ref.Cliente", 3400, lvwColumnLeft
            .Add , , "Fecha", 1100, lvwColumnCenter
            .Add , , "General", 800, lvwColumnCenter
            .Add , , "ID_MUESTRA", 1, lvwColumnCenter
            .Add , , "Facturada", 1, lvwColumnCenter
            .Add , , "Nº Usos", 800, lvwColumnCenter
        End With
    End If
        
        Dim eq As New clsEq_usos
        txtuso(2) = eq.Numero_Total_Usos(mvarobjEquipo.getID_EQUIPO)
        
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT cl.id_cliente, " & _
                   "concat(tm.codigo,'-',cast(mu.id_particular as char)), " & _
                   "cl.nombre, " & _
                   "mu.tipo_analisis_id, " & _
                   "mu.referencia_cliente, " & _
                   "mu.fecha_recepcion, " & _
                   "mu.id_muestra, " & _
                   "mu.precio, " & _
                   "ta.nombre, " & _
                   "mu.id_general, " & _
                   "mu.documento_pago,mu.enviado_correo,mu.anulada,mu.cerrada,equ.usos " & _
                   "FROM clientes as cl, " & _
                         "tipos_muestra as tm, " & _
                         "tipos_analisis as ta, " & _
                         "muestras as mu,eq_usos as equ " & _
                   "WHERE mu.cliente_id=cl.id_cliente AND " & _
                          " mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                          " mu.tipo_analisis_id=ta.id_tipo_analisis AND " & _
                          " mu.id_muestra = equ.muestra_id AND " & _
                          " equ.equipo_id = " & mvarobjEquipo.getID_EQUIPO & _
                          " order by mu.id_muestra desc"
        Me.MousePointer = 11
        Set rs = datos_bd(consulta)
        If rs.RecordCount >= 1 Then
            listaUsos.ListItems.Clear
            While Not rs.EOF
                With listaUsos.ListItems.Add(, , rs(1))
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
                .SubItems(8) = rs(14)
                End With
                rs.MoveNext
            Wend
        End If
        Me.MousePointer = 0
        
End Sub
Private Sub PresentarDatos_Normas()
   
    Dim rs As ADODB.Recordset
    Set rs = mvarobjEquipo.listado_normas_equipos()
    grdNormas.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With grdNormas.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = rs(3)
            End With
            If rs(4) = 1 Then
                colorear grdNormas, grdNormas.ListItems.Count, vbRed
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub

Private Sub colorear(grid As ListView, fila As Integer, color As Long)
    Dim i As Integer
    grid.ListItems(fila).ForeColor = color
    For i = 1 To grid.ColumnHeaders.Count - 1
        grid.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
End Sub
Private Sub RecogerDatos()
    Dim objFamilia As New clsFamiliasEquipos
    On Error GoTo RecogerDatos_Error
    With mvarobjEquipo
        .setNOMBRE = txtNombre.Text
        'M1124-I
        .setESTADO_ID = cmbEstado.BoundText
        .setCENTRO_ID = cmbCentro.BoundText
        'M1124-F
        'M0459-I
        If Trim(txtUnidadesLote.Text) = "" Then
            .setUNIDADES_LOTE = 0
        Else
            .setUNIDADES_LOTE = txtUnidadesLote.Text
        End If
        'M0459-F
        .setDESCRIPCION = txtdescripcion.Text
        .setFAMILIA_ID = getDataComboSel(cmbFamilia)
        .setPROVEEDOR_ID = getDataComboSel(cmbProveedor)
        Set objFamilia = New clsFamiliasEquipos
        Call objFamilia.Carga(getDataComboSel(cmbFamilia))
        Set .setFAMILIA = objFamilia
        .setSERIE = txtNSerie.Text
        .setMODELO = txtModelo.Text
        .setFABRICANTE = txtFabricante.Text
        .setFECHA_RECEPCION = Format(txtFechaRecepcion.Value, "dd/mm/yyyy")
        .setFECHA_SERVICIO = Format(txtFechaPuestaServicio.Value, "dd/mm/yyyy")
        .setRESPONSABLE_ID = getDataComboSel(cmbResponsable)
        .setPROCEDENCIA_ID = getDataComboSel(cmbProcedencia)
        .setES_NADCAP = chkNADCAP.Value
        .setES_MTL = chkMTL.Value
        .setES_CP = chkCP.Value
        .setES_ENAC = chkENAC.Value
        .setMOSTRAR_EN_PLANO = chkMostrarEnPlano.Value
        .setPRIORITARIO = chkPrioritario.Value
        .setCRITICO = chkCritico.Value
        If cmbTrazmet.getTEXTO = "" Then
            .setTRAZMET_ID = 0
        Else
            .setTRAZMET_ID = cmbTrazmet.getPK_SALIDA
        End If
        .setINSITU = chkInSitu.Value
        .setES_ACCESORIO = chkes_accesorio.Value
        
        If IsNumeric(cmbTipoEquipo.BoundText) Then
            .setTIPO_EQUIPO_ID = CInt(cmbTipoEquipo.BoundText)
        Else
            .setTIPO_EQUIPO_ID = 0
        End If
        
        .setCONDICIONES_AMBIENTALES = chkCondicionesAmbientales.Value
            If Not IsNumeric(txtTemperaturaMax.Text) Then txtTemperaturaMax.Text = "0"
            If Not IsNumeric(txtTemperaturaMin.Text) Then txtTemperaturaMin.Text = "0"
            .setTEMPERATURA_MAX = CCur(txtTemperaturaMax.Text)
            .setTEMPERATURA_MIN = CCur(txtTemperaturaMin.Text)
            
            If Not IsNumeric(txtHumedadMax.Text) Then txtHumedadMax.Text = "0"
            If Not IsNumeric(txtHumedadMin.Text) Then txtHumedadMin.Text = "0"
            .setHUMEDAD_MAX = CCur(txtHumedadMax.Text)
            .setHUMEDAD_MIN = CCur(txtHumedadMin.Text)
        
        .setALTA_BAJA = IIf(optAlta.Value, 0, 1)
        .setSITUACION_ID = getDataComboSel(cmbSituacion)
'J4        .setSITUACION_DESCRIPCION = txtDescZonaTrabajo.Text
        
        ' Rangos
        
        .setRANGO_MEDIDA_MAX = IIf(IsNumeric(txtRangoMedidaMax.Text), txtRangoMedidaMax.Text, 0)
        .setRANGO_MEDIDA_MIN = IIf(IsNumeric(txtRangoMedidaMin.Text), txtRangoMedidaMin.Text, 0)
        .setRANGO_TRABAJO_MAX = IIf(IsNumeric(txtRangoTrabajoMax.Text), txtRangoTrabajoMax.Text, 0)
        .setRANGO_TRABAJO_MIN = IIf(IsNumeric(txtRangoTrabajoMin.Text), txtRangoTrabajoMin.Text, 0)
        .setUNIDAD_ID = cmbUnidades.getPK_SALIDA
        .setINCERTIDUMBRE_MAXIMA = IIf(IsNumeric(txtIncertidumbreMax.Text), CStr(txtIncertidumbreMax.Text), 0)
        .setTOLERANCIA_MAXIMA = IIf(IsNumeric(txtToleranciaMax.Text), CStr(txtToleranciaMax.Text), 0)
        .setPRECISIONN = IIf(IsNumeric(txtPrecision.Text), CStr(txtPrecision.Text), 0)
        
        .setCON_CALIBRACION = chkCon_Calibracion.Value
        .setCON_VERIFICACION = chkCon_Verificacion.Value
        .setCON_MANTENIMIENTO = chkCon_Mantenimiento.Value
        
        'Datos Calibracion
        If .getCON_CALIBRACION Then
'            .setFECHA_ULT_CALIBRACION = mvarobjUltCalibracion.getFECHA_PROXIMA
            .setFECHA_ULT_CALIBRACION = "'" & Format(factcalibracion, "yyyy-mm-dd") & "'"
            .setFECHA_PROX_CALIBRACION = "'" & Format(fproxcalibracion, "yyyy-mm-dd") & "'"
            .setCALIBRADOR_INTERNO_ID = getDataComboSel(cmbCalibracionResponsable)
            .setTIPO_CALIBRACION_ID = getDataComboSel(cmbCalibracionTipo)
            .setPERIODICIDAD_CALIBRACION_ID = getDataComboSel(cmbCalibracionPeriod)
            .setPROCEDIMIENTO_CALIBRACION_ID = cmbProcedimientoCal.getPK_SALIDA
        Else
            .setFECHA_ULT_CALIBRACION = "NULL"
            .setFECHA_PROX_CALIBRACION = "NULL"
            .setCALIBRADOR_INTERNO_ID = -1
            .setTIPO_CALIBRACION_ID = -1
            .setPERIODICIDAD_CALIBRACION_ID = -1
            .setPROCEDIMIENTO_CALIBRACION_ID = 0
        End If
        
        If .getCON_VERIFICACION Then
            .setFECHA_ULT_VERIFICACION = "'" & Format(factverificacion, "yyyy-mm-dd") & "'"
            .setFECHA_PROX_VERIFICACION = "'" & Format(fproxverificacion, "yyyy-mm-dd") & "'"
            .setVERIFICADOR_INTERNO_ID = getDataComboSel(cmbVerificacionResponsable)
            .setTIPO_VERIFICACION_ID = getDataComboSel(cmbVerificacionTipo)
            .setPERIODICIDAD_VERIFICACION_ID = getDataComboSel(cmbVerificacionPeriod)
            .setPROCEDIMIENTO_VERIFICACION_ID = cmbProcedimientoVer.getPK_SALIDA
        Else
            .setFECHA_ULT_VERIFICACION = "null"
            .setFECHA_PROX_VERIFICACION = "null"
            .setVERIFICADOR_INTERNO_ID = -1
            .setTIPO_VERIFICACION_ID = -1
            .setPERIODICIDAD_VERIFICACION_ID = -1
            .setPROCEDIMIENTO_VERIFICACION_ID = 0
        End If
        
        If .getCON_MANTENIMIENTO Then
            .setFECHA_ULT_MANTENIMIENTO = "'" & Format(factmantenimiento, "yyyy-mm-dd") & "'"
            .setFECHA_PROX_MANTENIMIENTO = "'" & Format(fproxmantenimiento, "yyyy-mm-dd") & "'"
            .setMANTENEDOR_ID = getDataComboSel(cmbMtoResponsable)
        Else
            .setFECHA_ULT_MANTENIMIENTO = "null"
            .setFECHA_PROX_MANTENIMIENTO = "null"
            .setMANTENEDOR_ID = -1
            .setPLAN_MANTENIMIENTO_ID = -1
        End If
        If txtuso(0) <> "" Then
            .setNUMERO_USOS_MAXIMO = txtuso(0)
        Else
            .setNUMERO_USOS_MAXIMO = 0
        End If
        .setNUMERO_USOS_CONTADOR = txtuso(1)
        'M1050-I
        If cmbCliente.getTEXTO = "" Then
            .setCLIENTE_ID = 0
        Else
            .setCLIENTE_ID = cmbCliente.getPK_SALIDA
        End If
        .setNUMERO_EQUIPO_CLIENTE = txtNumeroCliente
        .setOBSERVACIONES = txtObservaciones
        'M1050-F
    End With



On Error GoTo 0
    Exit Sub
RecogerDatos_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure RecogerDatos of Formulario frmEquipoEdicion"
    ' Para depuracion
    
End Sub

Private Sub txtHumedadMax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtTemperaturaMin, KeyAscii, True)

End Sub


Private Sub txtHumedadMin_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtTemperaturaMin, KeyAscii, True)

End Sub


Private Sub txtIncertidumbreMax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtIncertidumbreMax, KeyAscii, True)

End Sub


Private Sub txtLimitacionesUso_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then Call cmdAnadirLimitacion_Click
End Sub

Private Sub txtMtoFechaProxima_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
'J3    txtMtoFechaProxima_b.Text = Format(txtMtoFechaProxima.value, "dd/mm/yyyy")
    txtMtoFechaProxima_info.Text = Format(txtMtoFechaProxima.Value, "dd/mm/yyyy")
End Sub


Private Sub txtMtoFechaProxima_Change()
'J3    txtMtoFechaProxima_b.Text = Format(txtMtoFechaProxima.value, "dd/mm/yyyy")
    txtMtoFechaProxima_info.Text = Format(txtMtoFechaProxima.Value, "dd/mm/yyyy")
    
End Sub


Private Sub txtnombre_Change()
    lbltitulo = txtNombre
End Sub

Private Sub txtPrecision_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtPrecision, KeyAscii, True)
End Sub


Private Sub txtRangoMedidaMax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtRangoMedidaMax, KeyAscii, True)
End Sub


Private Sub txtRangoMedidaMin_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtRangoMedidaMin, KeyAscii, True)

End Sub


Private Sub txtRangoTrabajoMax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtRangoTrabajoMax, KeyAscii, True)
End Sub


Private Sub txtRangoTrabajoMin_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtRangoTrabajoMin, KeyAscii, True)

End Sub


Private Sub txtTemperaturaMax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtTemperaturaMin, KeyAscii, True)

End Sub


Private Sub txtTemperaturaMin_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtTemperaturaMin, KeyAscii, True)

End Sub


Private Sub txtToleranciaMax_KeyPress(KeyAscii As Integer)
KeyAscii = KeyAscii_SoloDecimal(txtToleranciaMax, KeyAscii, True)

End Sub


Private Sub txtuso_LostFocus(Index As Integer)
    If Index = 0 Then
        If Not IsNumeric(txtuso(Index)) Then
            MsgBox "El numero de usos debe ser numérico.", vbCritical, App.Title
            txtuso(Index) = ""
            txtuso(Index).SetFocus
        End If
    End If
End Sub
'J2
'Private Sub txtVerificacionFechaProxima_CallbackKeyDown(ByVal KeyCode As Integer, ByVal Shift As Integer, ByVal CallbackField As String, CallbackDate As Date)
'    txtVerificacionFechaProxima_b.Text = Format(txtVerificacionFechaProxima.value, "dd/mm/yyyy")
'End Sub
'J2


Private Sub GuardarEquipo()
    Dim objEquipo As New clsEquipos

    Dim cambioEstado As Boolean
    cambioEstado = False
    If mvarobjEquipo.getESTADO_ID <> cmbEstado.BoundText Then
        cambioEstado = True
    End If

    Call RecogerDatos
    Dim PROCESO_ALTA As Boolean
    PROCESO_ALTA = False
    If mvarenuTipoEdicion = Alta Then
        PROCESO_ALTA = True
        mvarobjEquipo.Insertar
    Else
        Call mvarobjEquipo.Modificar(mvarobjEquipo.getID_EQUIPO)
    End If
    
    mvarenuTipoEdicion = EDICION
    
    If cambioEstado Then
        Dim oEV As New clsEquipoEventos
        With oEV
            .setEQUIPO_ID = mvarobjEquipo.getID_EQUIPO
            .setEVENTO_ID = EQUIPOS_EVENTOS.EVT_CAMBIO_ESTADO_EQUIPO
            .setRAZON_ID = EQUIPOS_EVENTOS_MOTIVOS.EVTRZ_OTROS
            .setOBSERVACIONES = "CAMBIO A ESTADO : " & cmbEstado.Text
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .Insertar
        End With
        Set oEV = Nothing
    End If
    ' Recarga el Equipo
    objEquipo.Carga mvarobjEquipo.getID_EQUIPO
    Set mvarobjEquipo = objEquipo
    
    Form_Load
    
    MsgBox "El equipo se ha guardado Correctamente.", vbInformation, "Guardar Datos Equipo"
    
    If PROCESO_ALTA Then
        cmdRecepcion_Click
        cmdEspecificaciones_Click
        cmdDocumentacion_Click
        cmdetiqueta_Click
        
        MsgBox "Proceso de Alta completado.", vbInformation, "Guardar Datos Equipo"
    End If
End Sub

Private Sub PresentarDatos_Mantenimiento_Planes()

' Muestra la lista de planes de mantenimiento que tiene un equipo
' de ellos, señala el preferido

    
    mvarlngid_plan_mantenimiento_por_defecto = 0
    
    Dim rs As ADODB.Recordset
    
    Set rs = mvarobjEquipo.devolver_lista_planes_mantenimiento()

    lstPlanes.Clear

    If rs.RecordCount = 0 Then Exit Sub
    
    rs.MoveFirst
    mvarblnCargando = True
    While Not rs.EOF
        lstPlanes.AddItem rs("NOMBRE")
        lstPlanes.ItemData(lstPlanes.ListCount - 1) = rs("plan_mantenimiento_id")
        If CInt(rs("por_defecto")) = 1 Then
            lstPlanes.Selected(lstPlanes.ListCount - 1) = True
            mvarlngid_plan_mantenimiento_por_defecto = rs("plan_mantenimiento_id")
        End If
        rs.MoveNext
    Wend
    mvarblnCargando = False

End Sub

Private Sub colorear_fila(grid As MSFlexGrid, color As Long)
    Dim icol As Integer
    With grid
        .Redraw = False
        .Row = .Rows - 1
        For icol = 0 To .COLS - 1
            .Col = icol
            .CellBackColor = color
        Next
        .Redraw = True
    End With
End Sub
