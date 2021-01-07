VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmVerMuestra 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Otros"
   ClientHeight    =   8040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13290
   Icon            =   "frmVerMuestra.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8040
   ScaleWidth      =   13290
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame7 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
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
      Height          =   6765
      Left            =   11790
      TabIndex        =   66
      Top             =   0
      Width           =   1410
      Begin VB.CommandButton cmdRecarga 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Recarga"
         Height          =   870
         Left            =   90
         Picture         =   "frmVerMuestra.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   72
         ToolTipText     =   "Enviar informe por E-mail"
         Top             =   3060
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdAdjuntos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntos"
         Height          =   915
         Left            =   90
         Picture         =   "frmVerMuestra.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   71
         ToolTipText     =   "Enviar informe por E-mail"
         Top             =   2115
         Width           =   1230
      End
      Begin VB.CommandButton cmdEADS 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Esp. EADS"
         Height          =   915
         Left            =   90
         Picture         =   "frmVerMuestra.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   70
         ToolTipText     =   "Informe especial sin determinaciones de EADS"
         Top             =   3960
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.CommandButton cmdVida 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Vida "
         Height          =   915
         Left            =   90
         Picture         =   "frmVerMuestra.frx":17A8
         Style           =   1  'Graphical
         TabIndex        =   69
         ToolTipText     =   "Enviar informe por E-mail"
         Top             =   1170
         Width           =   1230
      End
      Begin VB.CommandButton cmdContra 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Contradictorio"
         Height          =   915
         Left            =   90
         Picture         =   "frmVerMuestra.frx":2072
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Enviar informe por E-mail"
         Top             =   225
         Width           =   1230
      End
      Begin VB.CommandButton cmdCert 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Certificado"
         Height          =   870
         Left            =   90
         Picture         =   "frmVerMuestra.frx":293C
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Certificado de pinturas"
         Top             =   4905
         Visible         =   0   'False
         Width           =   1230
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2595
      Left            =   5910
      TabIndex        =   52
      Top             =   4170
      Width           =   5835
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00004080&
         Height          =   315
         Index           =   17
         Left            =   1665
         Locked          =   -1  'True
         TabIndex        =   54
         Top             =   855
         Width           =   4065
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1290
         Index           =   18
         Left            =   1665
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   53
         Top             =   1200
         Width           =   4065
      End
      Begin MSComCtl2.DTPicker fechaMuestreo 
         Height          =   330
         Left            =   1650
         TabIndex        =   55
         Top             =   150
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   6
         Left            =   1665
         TabIndex        =   56
         Top             =   510
         Width           =   4050
         _ExtentX        =   7144
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
         TabIndex        =   60
         Top             =   540
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   16
         Left            =   90
         TabIndex        =   59
         Top             =   195
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   17
         Left            =   90
         TabIndex        =   58
         Top             =   1680
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalles"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   18
         Left            =   90
         TabIndex        =   57
         Top             =   885
         Width           =   1995
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3105
      Left            =   5910
      TabIndex        =   40
      Top             =   645
      Width           =   5805
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   13
         Left            =   1650
         Locked          =   -1  'True
         TabIndex        =   42
         Top             =   1545
         Width           =   4065
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1155
         Index           =   14
         Left            =   1650
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   1890
         Width           =   4065
      End
      Begin MSComCtl2.DTPicker fechaRecepcion 
         Height          =   330
         Left            =   1635
         TabIndex        =   43
         Top             =   180
         Width           =   1605
         _ExtentX        =   2831
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Bindings        =   "frmVerMuestra.frx":2C46
         Height          =   315
         Index           =   4
         Left            =   1650
         TabIndex        =   44
         Top             =   870
         Width           =   4065
         _ExtentX        =   7170
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
         Left            =   1650
         TabIndex        =   45
         Top             =   1215
         Width           =   4065
         _ExtentX        =   7170
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
      Begin MSDataListLib.DataCombo cmbDatos 
         Bindings        =   "frmVerMuestra.frx":2C8C
         Height          =   315
         Index           =   3
         Left            =   1650
         TabIndex        =   78
         Top             =   540
         Width           =   4065
         _ExtentX        =   7170
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
         Caption         =   "Recepcionada por"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   9
         Left            =   90
         TabIndex        =   51
         Top             =   540
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   10
         Left            =   90
         TabIndex        =   50
         Top             =   180
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Entregada por"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   11
         Left            =   90
         TabIndex        =   49
         Top             =   1215
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Envase"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   12
         Left            =   90
         TabIndex        =   48
         Top             =   855
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   13
         Left            =   90
         TabIndex        =   47
         Top             =   2295
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Detalles"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   14
         Left            =   90
         TabIndex        =   46
         Top             =   1560
         Width           =   1995
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2190
      Left            =   30
      TabIndex        =   32
      Top             =   4575
      Width           =   5805
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
         Left            =   135
         Locked          =   -1  'True
         TabIndex        =   82
         Top             =   1800
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.CommandButton cmdfactura 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ver factura"
         Height          =   285
         Left            =   4185
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   1755
         Width           =   1275
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
         Left            =   2655
         Locked          =   -1  'True
         TabIndex        =   80
         Top             =   1755
         Width           =   1410
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
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   65
         Top             =   180
         Width           =   1410
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
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   34
         Top             =   180
         Width           =   1410
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   750
         Index           =   21
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   33
         Top             =   915
         Width           =   4290
      End
      Begin MSComCtl2.DTPicker FechaEntrega 
         Height          =   330
         Left            =   1440
         TabIndex        =   35
         Top             =   510
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
         Format          =   16449537
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Se encuentra en la factura número"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   6
         Left            =   45
         TabIndex        =   79
         Top             =   1800
         Width           =   2760
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio por Deter."
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   23
         Left            =   3060
         TabIndex        =   64
         Top             =   225
         Width           =   1200
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Precio"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   19
         Left            =   90
         TabIndex        =   38
         Top             =   210
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Incidencias"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   20
         Left            =   45
         TabIndex        =   37
         Top             =   1215
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha entrega"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   21
         Left            =   90
         TabIndex        =   36
         Top             =   525
         Width           =   1995
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2610
      Left            =   30
      TabIndex        =   19
      Top             =   1620
      Width           =   5835
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   315
         Index           =   8
         Left            =   1395
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   1530
         Width           =   4380
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
         TabIndex        =   21
         Top             =   1890
         Value           =   -1  'True
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
         TabIndex        =   20
         Top             =   1890
         Width           =   615
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   2
         Left            =   1395
         TabIndex        =   23
         Top             =   840
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo cmbDatos 
         Bindings        =   "frmVerMuestra.frx":2CD2
         Height          =   315
         Index           =   0
         Left            =   1395
         TabIndex        =   24
         Top             =   180
         Width           =   4365
         _ExtentX        =   7699
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
         Index           =   1
         Left            =   1395
         TabIndex        =   25
         Top             =   510
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin MSDataListLib.DataCombo cmbPedidos 
         Bindings        =   "frmVerMuestra.frx":2D02
         Height          =   315
         Left            =   1395
         TabIndex        =   76
         Top             =   2160
         Width           =   4365
         _ExtentX        =   7699
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
         Index           =   7
         Left            =   1395
         TabIndex        =   83
         Top             =   1170
         Width           =   4380
         _ExtentX        =   7726
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
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
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pedido"
         Height          =   195
         Index           =   4
         Left            =   75
         TabIndex        =   77
         Top             =   2205
         Width           =   525
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   2
         Left            =   90
         TabIndex        =   31
         Top             =   210
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de muestra"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   3
         Left            =   90
         TabIndex        =   30
         Top             =   540
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre Baño"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   4
         Left            =   90
         TabIndex        =   29
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de analisis"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   5
         Left            =   90
         TabIndex        =   28
         Top             =   870
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Analisis duplicado"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   7
         Left            =   90
         TabIndex        =   27
         Top             =   1845
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ref. muestra"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   8
         Left            =   90
         TabIndex        =   26
         Top             =   1560
         Width           =   1995
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   30
      TabIndex        =   13
      Top             =   645
      Width           =   5835
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   5
         Left            =   900
         Locked          =   -1  'True
         TabIndex        =   62
         Top             =   180
         Width           =   1230
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   0
         Left            =   4860
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   180
         Width           =   870
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   2925
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   180
         Width           =   780
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
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   3
         Left            =   3690
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   180
         Width           =   795
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ID"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   22
         Left            =   4545
         TabIndex        =   61
         Top             =   225
         Width           =   285
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Numero"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   0
         Left            =   105
         TabIndex        =   18
         Top             =   225
         Width           =   1995
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Codigo"
         ForeColor       =   &H00000000&
         Height          =   285
         Index           =   1
         Left            =   2250
         TabIndex        =   17
         Top             =   225
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   960
      Left            =   10575
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6930
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   960
      Left            =   11880
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   6930
      Width           =   1230
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1245
      Left            =   30
      TabIndex        =   5
      Top             =   6750
      Width           =   10425
      Begin VB.CommandButton cmdEtiqueta 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Etiqueta"
         Height          =   960
         Left            =   9135
         Picture         =   "frmVerMuestra.frx":2D48
         Style           =   1  'Graphical
         TabIndex        =   75
         ToolTipText     =   "Muestra el informe de registrro de la muestra"
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdespecificas 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Dat. Especificos"
         Height          =   960
         Left            =   5265
         Picture         =   "frmVerMuestra.frx":3052
         Style           =   1  'Graphical
         TabIndex        =   74
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton cmdListadoDeter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Determinaciones"
         Height          =   960
         Left            =   3915
         Picture         =   "frmVerMuestra.frx":31E7
         Style           =   1  'Graphical
         TabIndex        =   73
         Top             =   180
         Width           =   1320
      End
      Begin VB.CommandButton cmdInfRegistro 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Doc. Registro"
         Height          =   960
         Left            =   7875
         Picture         =   "frmVerMuestra.frx":3A35
         Style           =   1  'Graphical
         TabIndex        =   63
         ToolTipText     =   "Muestra el informe de registrro de la muestra"
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdInforme 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informe"
         Height          =   960
         Left            =   6615
         Picture         =   "frmVerMuestra.frx":42FF
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Previsualizar informe de ensayo"
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdDeter 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Registro"
         Height          =   960
         Left            =   2655
         Picture         =   "frmVerMuestra.frx":4609
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdAnular 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Anular"
         Height          =   960
         Left            =   1395
         Picture         =   "frmVerMuestra.frx":4ED3
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   180
         Width           =   1230
      End
      Begin VB.CommandButton cmdModificar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Modificar"
         Height          =   960
         Left            =   120
         Picture         =   "frmVerMuestra.frx":51DD
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   1230
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Códigos de Registro"
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
      Height          =   300
      Index           =   0
      Left            =   30
      TabIndex        =   4
      Top             =   345
      Width           =   5835
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Recepción"
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
      Height          =   300
      Index           =   8
      Left            =   5910
      TabIndex        =   2
      Top             =   345
      Width           =   5835
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Otros datos"
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
      Height          =   315
      Index           =   7
      Left            =   30
      TabIndex        =   39
      Top             =   4275
      Width           =   5805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Consulta de Muestras"
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
      Index           =   5
      Left            =   30
      TabIndex        =   0
      Top             =   0
      Width           =   11715
   End
   Begin VB.Label lblestado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   8940
      TabIndex        =   10
      Top             =   0
      Width           =   2805
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Muestreo"
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
      Height          =   315
      Index           =   9
      Left            =   5910
      TabIndex        =   3
      Top             =   3840
      Width           =   5820
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Datos del Registro"
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
      Height          =   300
      Index           =   6
      Left            =   30
      TabIndex        =   1
      Top             =   1335
      Width           =   5865
   End
End
Attribute VB_Name = "frmVerMuestra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim frecepcion As Date
Dim fmuestreo As Date
Dim fentrega As Date
Private Sub cmbDatos_Click(Index As Integer, Area As Integer)
  On Error GoTo fallo:
  Dim oBANO As New clsBanos
  Select Case Index
  Case 0:
       ' Pedidos
       cmbPedidos.Text = ""
       If cmbDatos(0).Text <> "" Then
            pedidos (cmbDatos(0).BoundText)
       End If
  Case 1: 'muestra los analisis para el tipo de muestra seleccionado
       Dim tipo As New clsMuestra
       If cmbDatos(1).Text <> "" Then
        If Not tipo.esBano(cmbDatos(1).BoundText) Then  'es un baño Id_Espacial
        Dim oAnalisis As New clsTipos_analisis
         ' Es una determinacion de muestra
         cmbDatos(7).Text = ""
         cmbDatos(7).Enabled = False
         Label2(5).Caption = "Tipo de análisis"

         If oAnalisis.AnalisisAsociadosMuestra(cmbDatos(1).BoundText).RecordCount <> 0 Then 'existe un registro al menos
            Set cmbDatos(2).RowSource = oAnalisis.AnalisisAsociadosMuestra(cmbDatos(1).BoundText) 'comboanalisis
            cmbDatos(2).ListField = "nombre" 'lo que enseña
            cmbDatos(2).DataField = "id_tipo_analisis" 'campo asociado
            cmbDatos(2).BoundColumn = "id_tipo_analisis" 'lo que realmente
            cmbDatos(2).Text = oAnalisis.AnalisisAsociadosMuestra(cmbDatos(1).BoundText).Fields("nombre").value
            cmbDatos(2).Enabled = True
         Else 'si no recupero ningun registro
            cmbDatos(2).Text = ""
            cmbDatos(2).Enabled = False
         End If
        Else
         ' Es un baño especial
            Label2(5).Caption = "Solución"
            If cmbDatos(0).Text = "" Then
                MsgBox "Seleccione primero un cliente.", vbInformation, App.Title
                cmbDatos(1).Text = ""
                cmbDatos(0).SetFocus
                Exit Sub
            End If
            cmbDatos(2).Text = ""
            cmbDatos(2).Enabled = False
            Dim rsbano As New ADODB.RecordSet
            Set rsbano = oBANO.banos_cliente(cmbDatos(0).BoundText, cmbDatos(1).BoundText)
            If rsbano.RecordCount = 0 Then
                MsgBox "No hay baños para el cliente y tipo de muestra seleccionado.", vbInformation, App.Title
                cmbDatos(1).Text = ""
                Exit Sub
            Else
                cargar_combo_banos cmbDatos(0).BoundText, cmbDatos(1).BoundText
                cmbDatos(7).Enabled = True
            End If
            Set oBANO = Nothing
        End If 'fin del esBanno
        Set tipo = Nothing
        Set oAnalisis = Nothing
       End If
    Case 7:
        If cmbDatos(7).Text <> "" Then
         If IsNumeric(cmbDatos(7).BoundText) Then
'            Dim oBANO As New clsBanos
            oBANO.cargar_bano (cmbDatos(7).BoundText)
            cmbDatos(2).BoundText = oBANO.getID_SOLUCION
         End If
        End If
        Text1(8) = cmbDatos(7).Text
    End Select
  cmbDatos(2).ToolTipText = cmbDatos(2).Text
  Exit Sub
fallo:
    MsgBox "Error al decodificar los campos", vbCritical, Err.Description

End Sub

Private Sub cmdAdjuntos_Click()
    frmMuestras_Adjuntos.Show 1
    imprimir_recepcion
    consulta_muestra
End Sub

Private Sub cmdAnular_Click()
    If MsgBox("Va a anular la muestra. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        frmMotivo.Show 1
        If Trim(motivo) = "" Then
            MsgBox "Para anular la muestra es necesario introducir un motivo.", vbInformation, App.Title
            Exit Sub
        End If
        Dim omuestra As New clsMuestra
        If omuestra.Anular(CLng(Text1(0)), Trim(motivo)) Then
            consulta_muestra
        End If
        Set omuestra = Nothing
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdCert_Click()
    On Error GoTo fallo
    Dim doc As String
    doc = NOMBRE_DOCUMENTO(Text1(0), False) & ".doc"
    If Dir(doc) = "" Then
        Dim appword As Word.Application
        Dim docword As Word.Document
        ' Crear copia para su uso
        Set appword = CreateObject("word.application")
        Set docword = appword.Documents.Open(copiar_plantilla("certificacion", Text1(0), 1))
        appword.Visible = False
        appword.WindowState = wdWindowStateMinimize
        ' Cabecera
        Dim omuestra As New clsMuestra
        Dim rs As ADODB.RecordSet
        Set rs = omuestra.datos_cabecera_documento(Text1(0))
        With docword.Sections(1).Headers(1).Range.Tables(1)
            .Rows(2).Cells(1).Range.Text = rs(0)
            .Rows(3).Cells(1).Range.Text = rs(1)
            If rs(3) = "" Then
                .Rows(4).Cells(1).Range.Text = rs(2) & " " & rs(3)
            Else
                .Rows(4).Cells(1).Range.Text = rs(2) & " " & rs(3) & " (" & Trim(rs(4)) & ")"
            End If
            .Rows(5).Cells(1).Range.Text = rs(5)
            If Trim(rs(6)) <> "" Then
                .Rows(6).Cells(1).Range.Text = "A/A de " & rs(6)
            End If
            .Rows(2).Cells(3).Range.Text = rs(7)
            .Rows(3).Cells(3).Range.Text = rs(8)
            If Not IsNull(rs(9)) Then
            .Rows(4).Cells(3).Range.Text = rs(9)
            End If
        End With
        ' Ensayo y mensaje de edición
        With docword.Sections(1).Headers(1).Range.Tables(2)
            .Rows(1).Cells(1).Range.InsertAfter (rs(17) & "/" & Format(rs(7), "yyyy") & "/Edición " & rs(10) + 1)
            Dim mensaje_edicion As String
            If rs(10) = 0 Then
                mensaje_edicion = ""
            Else
                mensaje_edicion = ReadINI(App.Path + "\config.ini", "edicion", "mensaje") & "-" & ReadINI(App.Path + "\config.ini", "edicion", "ingles")
            End If
            .Rows(2).Cells(1).Range.Text = mensaje_edicion
        End With
        With docword.Sections(1).Headers(1).Range.Tables(3)
            .Rows(1).Cells(2).Range.InsertAfter cmbDatos(1).Text ' Tipo Muestra (Proceso)
            .Rows(3).Cells(2).Range.InsertAfter Text1(8).Text  ' Ref (Pintor)
        End With
        docword.Save
        appword.Visible = True
        Set docword = Nothing
        Set appword = Nothing
    Else
        ver_documento_word (doc)
    End If
    Me.MousePointer = 0
    Exit Sub
fallo:
    On Error Resume Next
    appword.Documents.Close (0)
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
    Me.MousePointer = 0
    MsgBox "Error al generar el documento: " & Err.Description, vbCritical, App.Title
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
        .setEMPLEADO_ID = usuario.getID_EMPLEADO
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
        .setFECHA_COMIENZO = 0
        .setFECHA_CIERRE = 0
        .setCERRADA = 0
        .setDOCUMENTO_PAGO = 0
        .setULT_EDICION_IMP = 0
        .guardarMuestra
     End With
     ' Insertar determinaciones
     Dim rs_deter As ADODB.RecordSet
     Dim oDeter As New clsDeterminaciones
     Dim oDatosDet As New clsDatos_determinaciones
     Dim ocampos As New clsFormulas_campos
     Dim rscampos As ADODB.RecordSet
     Dim determinacion As Long
     Set rs_deter = oDeter.lista_contradictorio(CLng(Text1(0)))
     If rs_deter.RecordCount <> 0 Then
        Do
            oDeter.setMUESTRA_ID = omuestra_destino.getID_MUESTRA
            oDeter.setTIPO_DETERMINACION_ID = rs_deter("tipo_determinacion_id")
            oDeter.setORDEN = rs_deter("orden")
            oDeter.setFORMULA_ID = rs_deter("formula_id")
            oDeter.setES_DUPLICADO = rs_deter("es_duplicado")
            oDeter.setSITUACION = rs_deter("situacion")
            determinacion = oDeter.InsertarDeterminacion
            ' Recuperar formulas_camposs (CAMPO_ID)
            Set rscampos = ocampos.ListaFormulas(rs_deter("formula_id"))
            ' Insertar Datos_Determinaciones
            If rscampos.RecordCount <> 0 Then
              Do
               oDatosDet.setDETERMINACION_ID = determinacion
               oDatosDet.setCAMPO_ID = rscampos("id_campo")
               oDatosDet.setVALOR_1 = "I-1"
               oDatosDet.setVALOR_2 = "I-2"
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
     Dim rs_oval As ADODB.RecordSet
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
     imprimir omuestra_destino.getID_MUESTRA, 10, False
     MsgBox "El contradictorio ha sido registrado con el Nº: " & omuestra_destino.getID_GENERAL & " y código: " & omuestra_destino.CodigoParticular(omuestra_destino.getID_MUESTRA), vbInformation, App.Title
    End If
    Exit Sub
fallo:
    MsgBox "Error al generar el contradictorio. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdDeter_Click()
        gmuestra = CLng(Text1(0))
        Dim omuestra As New clsMuestra
        omuestra.CargaMuestra (gmuestra)
        Select Case omuestra.getANALISIS_MODIFICADO
            Case 2 ' Control de eficacia
                frmCE_Resultados.Show 1
            Case 3 ' Sellante
                frmSE_Resultados.Show 1
            Case Else
                frmDeterminaciones.Show 1
        End Select
        consulta_muestra
End Sub

Private Sub cmdEADS_Click()
    On Error GoTo fallo
    Dim doc As String
    doc = NOMBRE_DOCUMENTO(Text1(0), False) & ".doc"
    If Dir(doc) = "" Then
        Dim appword As Word.Application
        Dim docword As Word.Document
        ' Crear copia para su uso
        Dim oTipo_documento_analisis As New clsTipos_Documentos_Analisis
        If oTipo_documento_analisis.CARGAR(cmbDatos(2).BoundText) = True Then
            Set appword = CreateObject("word.application")
            Set docword = appword.Documents.Open(copiar_plantilla(oTipo_documento_analisis.getPLANTILLA, Text1(0), 1))
            docword.Save
            appword.Visible = True
            Set docword = Nothing
            Set appword = Nothing
        End If
    Else
        ver_documento_word (doc)
    End If
    Me.MousePointer = 0
    Exit Sub
fallo:
    On Error Resume Next
    appword.Documents.Close (0)
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
    Me.MousePointer = 0
    MsgBox "Error al generar el documento: " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdespecificas_Click()
    Dim omuestra As New clsMuestra
    omuestra.CargaMuestra (CLng(Text1(0)))
    frmDatosEspecificos.PK_MUESTRA = CLng(Text1(0))
    frmDatosEspecificos.PK_BANO = omuestra.getBANO_ID
    frmDatosEspecificos.Show 1
End Sub

Private Sub cmdEtiqueta_Click()
    ReDim etiquetas(1)
    etiquetas(1) = gmuestra
    frmEtiquetas.Show 1
End Sub

Private Sub cmdFactura_Click()
   On Error GoTo cmdfactura_Click_Error
    Dim oDoc_pago As New clsDocs_pago
    oDoc_pago.generar_factura CLng(Text1(7)), False, False, ""
   On Error GoTo 0
   Exit Sub

cmdfactura_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdfactura_Click of Formulario frmVerMuestra"
End Sub

Private Sub cmdInforme_Click()
    Me.MousePointer = 11
    gmuestra = CLng(Text1(0))
    frmPrevisualizar.Show 1
    Me.MousePointer = 0
End Sub

Private Sub cmdInfRegistro_Click()
    On Error GoTo fallo
    Dim doc As String
    If UCase(usuario.getNOMBRE) = "PRUEBA" Then
        doc = ReadINI(App.Path + "\config.ini", "documentos", "recepcion") & "\prueba\" & Year(fechaRecepcion) & "\" & Text1(0) & ".doc"
    Else
        doc = ReadINI(App.Path + "\config.ini", "documentos", "recepcion") & "\" & Year(fechaRecepcion) & "\" & Text1(0) & ".doc"
    End If
    If Dir(doc) <> "" Then
        FileCopy doc, App.Path & "\recepcion.doc"
        ver_documento_word (App.Path & "\recepcion.doc")
    Else
        Me.MousePointer = 11
        If imprimir(Text1(0), 10, True) = True Then
            FileCopy doc, App.Path & "\recepcion.doc"
            ver_documento_word (App.Path & "\recepcion.doc")
        Else
            Me.MousePointer = 0
            MsgBox "Se ha producido un error al generar el documento.", vbCritical, Err.Description
        End If
    End If
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al generar el documento.", vbCritical, App.Title
End Sub

Private Sub cmdListadoDeter_Click()
    gmuestra = CLng(Text1(0))
    frmVerDeterminaciones.Show 1
    consulta_muestra
    gmuestra = 0
End Sub

Private Sub cmdModificar_Click()
    Dim color As Single
    color = &H80C0FF
    'Titulo
'    Label1(5).BackColor = color
    Label1(5).Caption = "MODIFICACION de la Muestra " & Text1(5) & " (" & Text1(1) & "-" & Text1(3) & ")"
    'Registro
    Text1(8).Locked = False
    'Recepcion
    cmbDatos(0).Locked = False
    cmbDatos(1).Enabled = True
    cmbDatos(2).Enabled = True
    cmbDatos(3).Locked = False
    cmbDatos(4).Locked = False
    cmbDatos(5).Locked = False
    Text1(13).Locked = False
    Text1(14).Locked = False
    opDuplicado(0).Enabled = True
    opDuplicado(1).Enabled = True
    ' Muestreo
    cmbDatos(6).Locked = False
    Text1(17).Locked = False
    Text1(18).Locked = False
    ' Otros datos
    If usuario.getPER_FACTURACION = True Then
        Text1(19).Locked = False
    End If
    Text1(21).Locked = False
    ' Botones
    cmdAnular.Enabled = False
    cmdDeter.Enabled = False
    cmdInfRegistro.Enabled = False
    cmdContra.Enabled = False
    cmdok.Visible = True
    cmdModificar.Enabled = False
    cmdVida.Enabled = False
    cmdListadoDeter.Enabled = False
    cmdInforme.Enabled = False
    cmdespecificas.Enabled = False
    Dim s As String
    s = cmbDatos(2).Text
'    Call cmbdatos_Click(1, 1)
    cmbDatos(2).Text = s
    cmbPedidos.Locked = False
    ' No modificar ref si cerrada
    Dim omuestra As New clsMuestra
    omuestra.CargaMuestra (CLng(Text1(0)))
    If omuestra.getCERRADA = 1 Then
        MsgBox "INFORMACION : Con la muestra cerrada, no se puede modificar la referencia.", vbInformation, App.Title
        Text1(8).Locked = True
    End If
    Text1(8).SetFocus
End Sub
Public Function validar() As Boolean
    validar = True
    If IsNumeric(Text1(19)) = False Then
        MsgBox "El precio debe ser numérico.", vbCritical, "Error"
        Text1(19).SetFocus
        validar = False
        Exit Function
    End If
    If cmbDatos(0).BoundText = "" Then
        MsgBox "El cliente debe estar informado.", vbInformation, App.Title
        validar = False
        cmbDatos(0).SetFocus
        Exit Function
    End If
    If cmbDatos(1).BoundText = "" Then
        MsgBox "El tipo de muestra debe estar informado.", vbInformation, App.Title
        validar = False
        cmbDatos(1).SetFocus
        Exit Function
    End If
    If cmbDatos(2).BoundText = "" Then
        MsgBox "El tipo de análisis o el baño deben estar informados.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If cmbDatos(3).BoundText = "" Then
        MsgBox "El usuario que recepciona debe estar informado.", vbInformation, App.Title
        validar = False
        cmbDatos(3).SetFocus
        Exit Function
    End If
End Function
Private Sub cmdok_Click()
    On Error GoTo fallo
    Dim s As String
    Dim consulta As String
    If validar = False Then
        Exit Sub
    End If
    ' Modificamos la muestra
    Me.MousePointer = 11
    Dim omuestra As New clsMuestra
    With omuestra
        .setREFERENCIA_CLIENTE = Trim(Text1(8))
        .setFECHA_RECEPCION = Format(fechaRecepcion.value, "yyyy-mm-dd")
        If cmbDatos(4).BoundText <> "" Then
            .setFORMATO_ID = cmbDatos(4).BoundText
        Else
            .setFORMATO_ID = 0
        End If
        .setENTIDAD_ENTREGA_ID = cmbDatos(5).BoundText
        .setDETALLE_ENTREGA = Trim(Text1(13))
        .setOBSERVACIONES_ENTREGA = Trim(Text1(14))
        .setFECHA_MUESTREO = Format(fechaMuestreo.value, "yyyy-mm-dd")
        .setENTIDAD_MUESTREO_ID = cmbDatos(6).BoundText
        .setCLIENTE_ID = cmbDatos(0).BoundText
        .setDETALLE_MUESTREO = Trim(Text1(17))
        .setOBSERVACIONES_MUESTREO = Trim(Text1(18))
        .setFECHA_PREV_FIN = Format(FechaEntrega.value, "yyyy-mm-dd")
        .setPRECIO = Replace(Format(Text1(19), "####0.00"), ",", ".")
        .setOBSERVACIONES = Trim(Text1(21))
        If opDuplicado(0).value = True Then
           .setANALISIS_DUPLICADO = 1
        Else
           .setANALISIS_DUPLICADO = 0
        End If
        If cmbPedidos.Text = "" Then
            .setPEDIDO_ID = 0
        Else
            .setPEDIDO_ID = cmbPedidos.BoundText
        End If
        .setEMPLEADO_ID = cmbDatos(3).BoundText
        .Modificar CLng(Text1(0))
    End With
    omuestra.CargaMuestra (CLng(Text1(0)))
    ' Verificamos si se modifica el tipo de la muestra
    Dim TIPO_MUESTRA_ID As Long
    Dim ID_PARTICULAR As Long
    Dim TIPO_ANALISIS_ID As Integer
    Dim BANO_ID As Integer
    ' Tipo analisis
    If cmbDatos(2).Text <> "" Then
        TIPO_ANALISIS_ID = cmbDatos(2).BoundText
    Else
        TIPO_ANALISIS_ID = 0
    End If
    ' Baño
    If cmbDatos(7).Text = "" Then
        BANO_ID = 0
    Else
        BANO_ID = cmbDatos(7).BoundText
        Dim oBANO As New clsBanos
        oBANO.cargar_bano (BANO_ID)
        TIPO_ANALISIS_ID = oBANO.getID_SOLUCION
    End If
    If omuestra.getTIPO_MUESTRA_ID <> cmbDatos(1).BoundText Or _
      (omuestra.getTIPO_ANALISIS_ID <> TIPO_ANALISIS_ID And omuestra.getBANO_ID = 0) Or _
      (omuestra.getBANO_ID <> BANO_ID) Then
        TIPO_MUESTRA_ID = cmbDatos(1).BoundText
'        Dim omuestra_auxiliar As New clsMuestra
        If omuestra.getTIPO_MUESTRA_ID <> cmbDatos(1).BoundText Then
            omuestra.CrearIdCodigoParticular (cmbDatos(1).BoundText)
            ID_PARTICULAR = omuestra.getID_PARTICULAR
        Else
            ID_PARTICULAR = omuestra.getID_PARTICULAR
        End If
        ' Modificamos la muestra
        consulta = "update muestras set " & _
                    " tipo_muestra_id = " & TIPO_MUESTRA_ID & "," & _
                    " id_particular = " & ID_PARTICULAR & "," & _
                    " tipo_analisis_id = " & TIPO_ANALISIS_ID & "," & _
                    " bano_id = " & BANO_ID & _
                    " where id_muestra = " & CLng(Text1(0))
        execute_bd consulta
        ' Borramos las determinaciones
        Dim oDeter As New clsDeterminaciones
        oDeter.Eliminar_Por_Muestra (CLng(Text1(0)))
        ' Insertamos las determinaciones por defecto
'        If TIPO_ANALISIS_ID <> 0 Then
        oDeter.Insertar_determinaciones_por_defecto (CLng(Text1(0)))
'        End If
        ' Datos especificos
        Dim ode As New clsDatos_valores
        ode.Eliminar_datos_especificos_vacios CLng(Text1(0))
        ode.Insertar_datos_especificos_por_defecto CLng(Text1(0))
    End If
    imprimir_recepcion
    Set omuestra = Nothing
    consulta_muestra
    proteger_campos
    Me.MousePointer = 0
    MsgBox "Datos modificados correctamente.", vbInformation, App.Title
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox Err.Description, vbCritical, Err.Number
End Sub

Private Sub cmdRecarga_Click()
    gmuestra = CLng(Text1(0))
    frmEads_Recarga.Show 1
    gmuestra = 0
End Sub

Private Sub cmdVida_Click()
    gmuestra = CLng(Text1(0))
    frmVidaMuestra.Show 1
    gmuestra = 0
End Sub

Private Sub FechaEntrega_Change()
    If cmdok.Visible = False Then
        FechaEntrega = fentrega
    End If
End Sub
Private Sub fechaMuestreo_Change()
    If cmdok.Visible = False Then
        fechaMuestreo = fmuestreo
    End If
End Sub
Private Sub fechaRecepcion_Change()
    If cmdok.Visible = False Then
        fechaRecepcion = frecepcion
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Me.MousePointer = 0
    Select Case KeyCode
        Case 27
            cmdcancel_Click
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    permisos
    consulta_muestra
End Sub
Public Sub consulta_muestra()
    Dim omuestra As New clsMuestra
    Dim codigo As String
    Dim Pos As Integer
    With omuestra
    If .CargaMuestra(gmuestra) Then
        Text1(0) = .getID_MUESTRA
        Text1(5) = .getID_GENERAL
        codigo = .CodigoParticular(.getID_MUESTRA)
        Pos = InStr(1, codigo, "-", vbTextCompare)
        Text1(1) = Mid(.CodigoParticular(.getID_MUESTRA), 1, Pos - 1)
        Text1(3) = .getID_PARTICULAR
        ' Cliente
        cmbDatos(0).BoundText = .getCLIENTE_ID
        ' Pedidos
        pedidos (.getCLIENTE_ID)
        cmbPedidos.BoundText = .getPEDIDO_ID
        cmbDatos(1).BoundText = .getTIPO_MUESTRA_ID
        ' Certificados especiales Canagrosa
        ' Certif. de pintura
        If .getTIPO_MUESTRA_ID = 71 Or .getTIPO_MUESTRA_ID = 72 Then
            cmdCert.Visible = True
            If .getCERRADA <> 1 Then
                cmdCert.Enabled = False
            End If
        End If
        ' Cert. especiales de banos
        Dim oTipo_documento_analisis As New clsTipos_Documentos_Analisis
        If oTipo_documento_analisis.CARGAR(.getTIPO_ANALISIS_ID) = True Then
            cmdEADS.Visible = True
        End If
        cmbDatos(2).BoundText = .getTIPO_ANALISIS_ID
        ' Baño
        If .getBANO_ID = 0 Then
'            Text1(4) = ""
            Label2(5).Caption = "Tipo de análisis"
            cmbDatos(7).Text = ""
        Else
            Label2(5).Caption = "Solución"
            cargar_combo_banos .getCLIENTE_ID, .getTIPO_MUESTRA_ID
            cmbDatos(7).BoundText = .getBANO_ID
            
'            Dim oBANO As New clsBanos
'            oBANO.cargar_bano (.getBANO_ID)
'            Text1(4) = oBANO.getNOMBRE
'            Set oBANO = Nothing
            cmdRecarga.Visible = True
        End If
        If .getANALISIS_DUPLICADO = 0 Then
            opDuplicado(1).value = True
        Else
            opDuplicado(0).value = True
        End If
        Text1(8) = .getREFERENCIA_CLIENTE
        fechaRecepcion = Format(.getFECHA_RECEPCION, "dd/mm/yyyy")
        frecepcion = fechaRecepcion.value
        ' Empleado
        cmbDatos(3).BoundText = .getEMPLEADO_ID
'        Dim oempleado As New clsUsuarios
'        oempleado.cargar (.getEMPLEADO_ID)
'        Text1(2) = oempleado.getNOMBRE & " " & oempleado.getAPELLIDOS
'        Set oempleado = Nothing
        cmbDatos(4).BoundText = .getFORMATO_ID
        cmbDatos(5).BoundText = .getENTIDAD_ENTREGA_ID
        Text1(13) = .getDETALLE_ENTREGA
        Text1(14) = .getOBSERVACIONES_ENTREGA
        If IsNull(.getFECHA_MUESTREO) = False And Trim(.getFECHA_MUESTREO) <> "" Then
            fechaMuestreo = Format(.getFECHA_MUESTREO, "dd/mm/yyyy")
            fmuestreo = fechaMuestreo.value
        Else
            fechaMuestreo = Format(Date, "dd/mm/yyyy")
            fmuestreo = fechaMuestreo.value
        End If
        cmbDatos(6).BoundText = .getENTIDAD_MUESTREO_ID
        Text1(17) = .getDETALLE_MUESTREO
        Text1(18) = .getOBSERVACIONES_MUESTREO
        Text1(19) = Format(.getPRECIO, "currency")
        Text1(6) = Format(omuestra.ImporteMuestraPorDeterminaciones(gmuestra, .getCLIENTE_ID), "currency")
        If IsNull(.getFECHA_PREV_FIN) = False And .getFECHA_PREV_FIN <> "" Then
            FechaEntrega = Format(.getFECHA_PREV_FIN, "dd/mm/yyyy")
        Else
            FechaEntrega = Format(Date, "dd/mm/yyyy")
        End If
        fentrega = FechaEntrega.value
        Text1(21) = .getOBSERVACIONES
        Select Case .getCERRADA
        Case 1
            lblestado = "CERRADA"
'            Label1(5).Width = 8895
        Case 2
            lblestado = "PTE.CIERRE"
            Label1(5).Width = 8895
        Case Else
'            If .getANULADA = 0 Then
                cerrar_muestra
'            End If
        End Select
        If .getANULADA <> 0 Then
            lblestado = "ANULADA"
            Label1(5).Width = 8895
            cmdAnular.Enabled = False
            cmdModificar.Enabled = False
            cmdDeter.Enabled = False
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
    End If
    End With
    Label1(5).BackColor = &HC0FFFF
    Label1(5).Caption = "Consulta de Muestra " & Text1(5) & " (" & Text1(1) & "-" & Text1(3) & ")"
    Set omuestra = Nothing
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

Public Sub permisos()
    If usuario.getPER_FACTURACION = False Then
        Text1(19).Locked = True
    End If
    If usuario.getPER_MODIFICACION = False Then
        cmdModificar.Enabled = False
    End If
'    cmdVida.Visible = False
End Sub
Public Sub proteger_campos()
    Dim color As Single
    color = &HC0FFFF
    'Titulo
    Label1(5).BackColor = color
    Label1(5).Caption = "Consulta de Muestra " & Text1(5) & " (" & Text1(1) & "-" & Text1(3) & ")"
    'Registro
    Text1(8).Locked = True
    'Recepcion
    cmbDatos(0).Locked = True
    cmbDatos(4).Locked = True
    cmbDatos(3).Locked = True
    cmbDatos(5).Locked = True
    Text1(13).Locked = True
    Text1(14).Locked = True
    opDuplicado(0).Enabled = False
    opDuplicado(1).Enabled = False
    ' Muestreo
    cmbDatos(6).Locked = True
    cmbDatos(7).Enabled = False
'    cmbDatos(7).Locked = True
    Text1(17).Locked = True
    Text1(18).Locked = True
    ' Otros datos
    Text1(19).Locked = True
    Text1(21).Locked = True
    ' Botones
    cmdAnular.Enabled = True
    cmdDeter.Enabled = True
    cmdInfRegistro.Enabled = True
    cmdContra.Enabled = True
    cmdok.Visible = False
    cmdVida.Enabled = True
    cmdInforme.Enabled = True
    cmdespecificas.Enabled = True
    If usuario.getPER_MODIFICACION = True Then
        cmdModificar.Enabled = True
    End If
    cmdListadoDeter.Enabled = True
'    Form_Load
    cmbDatos(1).Enabled = False
    cmbDatos(2).Enabled = False
    cmbPedidos.Locked = True
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

Public Sub cerrar_muestra()
    ' Cerrar Muestra
    Dim omuestra As New clsMuestra
    Dim cierre As Integer
    cierre = omuestra.comprobar_cierre(CLng(Text1(0)))
    Select Case cierre
        Case 1 ' Cerrada
           lblestado = "CERRADA"
           Label1(5).Width = 8895
        Case 2 ' Pte.
           lblestado = "PTE.CIERRE"
           Label1(5).Width = 8895
    End Select
    Set omuestra = Nothing
End Sub

Public Sub cargar_combos()
    Cargar_Combo cmbDatos(0), New clsCliente
    Cargar_Combo cmbDatos(1), New clsTipos_muestra
    Cargar_Combo cmbDatos(2), New clsTipos_analisis
    Cargar_Combo cmbDatos(4), New clsformatos
    Cargar_Combo cmbDatos(5), New clsEntidades_Entrega
    Cargar_Combo cmbDatos(6), New clsEntidades_muestreo
    Cargar_Combo cmbDatos(3), New clsUsuarios
    pedidos (0)
End Sub

Public Sub imprimir_recepcion()
    imprimir CLng(Text1(0)), 10, False
End Sub
Public Sub pedidos(ID As Integer)
    Dim oPedido As New clsClientes_pedidos
    Dim anterior As Integer
    If cmbPedidos.Text <> "" Then
        anterior = cmbPedidos.BoundText
    End If
    If ID = 0 Then
        Set cmbPedidos.RowSource = oPedido.Listado_completo
    Else
'        Set cmbPedidos.RowSource = oPedido.Listado_por_Cliente(ID)
        Set cmbPedidos.RowSource = oPedido.Listado_en_fecha(ID, Date)
    End If
    cmbPedidos.ListField = "CODIGO_LARGO"
    cmbPedidos.DataField = "id_pedido"
    cmbPedidos.BoundColumn = "id_pedido"
    cmbPedidos.BoundText = anterior
End Sub
Private Sub cargar_combo_banos(cliente As Long, tipo_muestra As Long)
    ' Cargamos la combo de baños del cliente
    Dim obanos As New clsBanos
    Set cmbDatos(7).RowSource = obanos.banos_cliente(cliente, tipo_muestra)
    cmbDatos(7).ListField = "nombre"
    cmbDatos(7).BoundColumn = "id_bano"
    cmbDatos(7).DataField = "id_bano" 'campo asociado
    Set obanos = Nothing
End Sub
