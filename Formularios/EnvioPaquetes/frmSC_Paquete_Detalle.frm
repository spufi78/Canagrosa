VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmSC_Paquete_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Subcontratación de ensayos - Detalle"
   ClientHeight    =   8775
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12585
   Icon            =   "frmSC_Paquete_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8775
   ScaleWidth      =   12585
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdFacturacion 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Facturas"
      Height          =   915
      Left            =   7380
      Style           =   1  'Graphical
      TabIndex        =   47
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   7830
      Width           =   1275
   End
   Begin VB.Frame frmHistoria 
      Height          =   4110
      Left            =   3330
      TabIndex        =   26
      Top             =   1755
      Width           =   6315
      Begin VB.Frame Frame3 
         Caption         =   "Petición"
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
         Height          =   1005
         Left            =   45
         TabIndex        =   38
         Top             =   450
         Width           =   6225
         Begin VB.TextBox txtUsuarioPeticion 
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
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   40
            Top             =   225
            Width           =   4920
         End
         Begin VB.TextBox txtFechaPeticion 
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
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   39
            TabStop         =   0   'False
            Top             =   585
            Width           =   1380
         End
         Begin VB.Label Label3 
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   405
            TabIndex        =   42
            Top             =   270
            Width           =   645
         End
         Begin VB.Label Label6 
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   495
            TabIndex        =   41
            Top             =   630
            Width           =   510
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Trámite"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   1005
         Left            =   45
         TabIndex        =   33
         Top             =   1440
         Width           =   6225
         Begin VB.TextBox txtUsuarioTramite 
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
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   35
            Top             =   180
            Width           =   4920
         End
         Begin VB.TextBox txtFechaTramite 
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
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   540
            Width           =   1380
         End
         Begin VB.Label Label4 
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   405
            TabIndex        =   37
            Top             =   270
            Width           =   645
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   495
            TabIndex        =   36
            Top             =   630
            Width           =   510
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Recepción"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   45
         TabIndex        =   28
         Top             =   2430
         Width           =   6225
         Begin VB.TextBox txtUsuarioRecepcion 
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
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            TabIndex        =   30
            Top             =   225
            Width           =   4920
         End
         Begin VB.TextBox txtFechaRecepcion 
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
            Height          =   285
            Left            =   1125
            Locked          =   -1  'True
            MaxLength       =   100
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   585
            Width           =   1380
         End
         Begin VB.Label Label5 
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   405
            TabIndex        =   32
            Top             =   270
            Width           =   645
         End
         Begin VB.Label Label8 
            Caption         =   "Fecha:"
            Height          =   195
            Left            =   495
            TabIndex        =   31
            Top             =   630
            Width           =   510
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Cerrar"
         Height          =   465
         Left            =   2475
         TabIndex        =   27
         Top             =   3510
         Width           =   1455
      End
      Begin VB.Label lblSubtitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         Caption         =   "Historia"
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
         Index           =   0
         Left            =   45
         TabIndex        =   43
         Top             =   135
         Width           =   6225
      End
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   0
      MaxLength       =   100
      TabIndex        =   25
      TabStop         =   0   'False
      Text            =   "Importe Total Presupuestado:"
      Top             =   5940
      Width           =   11310
   End
   Begin VB.TextBox txtSuma 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   11295
      MaxLength       =   100
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   5940
      Width           =   1275
   End
   Begin VB.CommandButton CMDACEPTAR 
      Height          =   330
      Left            =   10935
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6300
      Width           =   375
   End
   Begin VB.TextBox txtNuevoPrecio 
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
      Height          =   330
      Left            =   11295
      MaxLength       =   100
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6300
      Width           =   1275
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   1095
      Index           =   3
      Left            =   45
      MaxLength       =   255
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   22
      Top             =   6660
      Width           =   12495
   End
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   915
      Left            =   8685
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   7830
      Width           =   1275
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar Determinación"
      Height          =   915
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7830
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Subcontratación"
      Height          =   1365
      Left            =   0
      TabIndex        =   14
      Top             =   315
      Width           =   12570
      Begin VB.TextBox txtEdicion 
         Height          =   285
         Left            =   5895
         TabIndex        =   48
         Text            =   "1"
         Top             =   225
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox chkFactura 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Factura"
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   945
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.TextBox txtFactura 
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
         Height          =   285
         Left            =   1035
         MaxLength       =   100
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   945
         Visible         =   0   'False
         Width           =   1380
      End
      Begin VB.TextBox txtDatos 
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
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   2
         Left            =   7740
         MaxLength       =   100
         TabIndex        =   6
         Top             =   630
         Width           =   4350
      End
      Begin VB.TextBox txtDatos 
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
         Height          =   285
         Index           =   1
         Left            =   1035
         MaxLength       =   100
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   270
         Width           =   1650
      End
      Begin MSComCtl2.DTPicker datFecha 
         Height          =   315
         Left            =   3825
         TabIndex        =   1
         TabStop         =   0   'False
         Top             =   225
         Width           =   1590
         _ExtentX        =   2805
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
         Format          =   52494337
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbUsuario 
         Height          =   330
         Left            =   7740
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker datFechaFactura 
         Height          =   315
         Left            =   3825
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   945
         Visible         =   0   'False
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
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
         Format          =   52494337
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbSubcontratas 
         Height          =   330
         Left            =   1035
         TabIndex        =   45
         Top             =   585
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   582
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmSC_Paquete_Detalle.frx":08CA
         Height          =   315
         Left            =   7740
         TabIndex        =   50
         Top             =   990
         Width           =   1875
         _ExtentX        =   3307
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
         Caption         =   "Centro"
         Height          =   195
         Index           =   22
         Left            =   6795
         TabIndex        =   51
         Top             =   1035
         Width           =   465
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha F."
         Height          =   240
         Left            =   3060
         TabIndex        =   21
         Top             =   990
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   2
         Left            =   6795
         TabIndex        =   19
         Top             =   270
         Width           =   540
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   240
         Left            =   3195
         TabIndex        =   18
         Top             =   270
         Width           =   510
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Presupuesto"
         Height          =   195
         Index           =   6
         Left            =   6795
         TabIndex        =   17
         Top             =   630
         Width           =   885
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcontrata"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   16
         Top             =   630
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código SC"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   15
         Top             =   315
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   915
      Left            =   11295
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Salir"
      Top             =   7830
      Width           =   1230
   End
   Begin MSComctlLib.ListView lstMuestras 
      Height          =   4260
      Left            =   0
      TabIndex        =   11
      Top             =   1710
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   7514
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
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
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   915
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Modificar paquete"
      Top             =   7830
      Width           =   1275
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Doble Click para ver la Muestra"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   3780
      TabIndex        =   49
      Top             =   6255
      Width           =   4785
   End
   Begin VB.Label lblTramite 
      BackColor       =   &H80000009&
      Caption         =   "Necesita permisos como tramitador si desea modificarlo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   240
      Left            =   2430
      TabIndex        =   46
      Top             =   7875
      Visible         =   0   'False
      Width           =   4785
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "F5 - HISTORIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   240
      Left            =   10980
      TabIndex        =   44
      Top             =   0
      Width           =   1530
   End
   Begin VB.Label lblObservaciones 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones:"
      Height          =   285
      Left            =   90
      TabIndex        =   20
      Top             =   6345
      Width           =   1140
   End
   Begin VB.Label lblSubtitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle Subcontratación"
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
      Index           =   2
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   12570
   End
End
Attribute VB_Name = "frmSC_Paquete_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long
'M1274-i
Public EDICION As Long
'M1274-F

Public Sub cabecera()
    With lstMuestras.ColumnHeaders
        .Add , , "Código", 1000, lvwColumnLeft              ' Muestra
        .Add , , "Ref. Cliente", 3300, lvwColumnLeft
        .Add , , "Determinacion", 3000, lvwColumnLeft       ' Determinación
        .Add , , "Valor ref.", 2000, lvwColumnLeft
        .Add , , "Normativa aplic.", 2000, lvwColumnLeft
        .Add , , "ID_DETERMINACION", 1, lvwColumnLeft       ' ID_DETERMINACION
        .Add , , "PRECIO", 1000, lvwColumnCenter        'PRECIO PRESUPUESTADO
        .Add , , "ID_MUESTRA", 0, lvwColumnCenter        'MUESTRA
    End With
End Sub

Private Sub chkFactura_Click()
    If chkFactura.value = 0 Then
        datFechaFactura.Enabled = False
'M1163-I
        txtFactura.Enabled = False
'M1163-F
    Else
        datFechaFactura.Enabled = True
'M1163-I
        txtFactura.Enabled = True
'M1163-F
    End If
End Sub

Private Sub cmdAceptar_Click()
     
    If lstMuestras.ListItems.Count = 0 Or lstMuestras.selectedItem.Selected = False Then
        Exit Sub
    End If
    If Not IsNumeric(Trim(txtNuevoPrecio.Text)) Then
       txtNuevoPrecio.Text = "0"
    End If
    lstMuestras.ListItems(lstMuestras.selectedItem.Index).SubItems(6) = moneda(Trim(txtNuevoPrecio.Text))
       
    'DEVOLVEMOS LA APARIENCIA ORIGINAL A LA CELDA DE IMPORTE
    'RECALCULAMOS LAS SUMAS
    
    txtNuevoPrecio.BackColor = &HFFFFFF
    txtNuevoPrecio.ForeColor = vbBlue
    txtNuevoPrecio.Text = ""
    txtSuma = SumarImportes
    txtDatos(2) = txtSuma
    txtDatos(2).Enabled = False
End Sub

Private Sub cmdAdjuntos_Click()
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_PAQUETE_SUBCONTRATA
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
End Sub

Private Sub cmdEliminar_Click()
    If lstMuestras.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("¿Desea eliminar la determinación?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim opd As New clsSC_Paquetes_Detalle
        opd.Eliminar_Determinacion PK, EDICION, lstMuestras.ListItems(lstMuestras.selectedItem.Index).SubItems(5)
        lstMuestras.ListItems.Remove lstMuestras.selectedItem.Index
    End If
End Sub


Private Sub Form_Load()
    log (Me.Name)
    Me.top = 1700
    Me.Left = 300
    cargar_botones Me
    Call cabecera
    Call cargar_combo_subcontratas
    llenar_combo cmbUsuario, New clsUsuarios, 0, frmUsuarios, ""
    cargar_combo cmbCentro, New clsCentros
    cmbUsuario.desactivar
'M0959-I
    datFechaFactura.value = Date
    txtDatos(2).Enabled = False
'M0959-F
    frmHistoria.Visible = False
    
    If PK <> 0 Then
        CARGAR
    End If
'MXXXX-I
'M0959-I
'    If txtFactura.Text = "" Then
'        datFechaFactura.Enabled = False
'    Else
'        datFechaFactura.Enabled = True
'    End If
'M0959-F
    If chkFactura.value = 0 Then
        datFechaFactura.Enabled = False
'M1163-I
        txtFactura.Enabled = False
'M1163-F
    Else
        datFechaFactura.Enabled = True
'M1163-I
        txtFactura.Enabled = True
'M1163-F
    End If
'MXXXX-F
'M1171-I
    If Not USUARIO.getPER_TRAMITACION_CONTRATA Then
        cmdok.Enabled = False
        cmdEliminar.Enabled = False
        cmdAceptar.Enabled = False
        txtNuevoPrecio.Locked = True
        lblTramite.Visible = True
    End If
'M1171-F
    If Not USUARIO.getPER_TESORERIA_FP Then
        cmdFacturacion.Visible = False
    End If
        
End Sub

' botones
'M1271-I
Private Sub cmdFacturacion_Click()
        Dim oPaquete As New clsSC_Paquetes
        oPaquete.Carga PK, EDICION
        frmProveedores_Facturas.TOBJETO = TOBJETO.TOBJETO_SC_DETERMINACIONES
        frmProveedores_Facturas.COBJETO = PK
        frmProveedores_Facturas.PK = oPaquete.getSUBCONTRATA_ID
        frmProveedores_Facturas.Show 1
        Set oPaquete = Nothing
End Sub
'M1271-F

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If datos_correctos Then
        Dim oSC_Paquete As New clsSC_Paquetes
        Dim lngPaquete As Long
        
        With oSC_Paquete
              .setCENTRO_ID = cmbCentro.BoundText
              .setPRESUPUESTO = moneda_bd(txtDatos(2))
              .setOBSERVACIONES = txtDatos(3)
              .setSUBCONTRATA_ID = cmbSubcontratas.getPK_SALIDA
              .setEDICION = EDICION
'M0959-I
              If txtFactura.Text <> "" Then
                 .setFACTURA_RECIBIDA = 1
                 .setFFACTURA = Format(datFechaFactura.value, "yyyy-mm-dd")
                 .setNFACTURA = Trim(txtFactura)
              End If
'M0959-F
        End With

        If MsgBox("Va a modificar el paquete. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'M0957-I
            Dim opd As New clsSC_Paquetes_Detalle
            Dim i As Integer
'M1171-I
            Dim oTipoContratas As New clsTipos_determinacion_contratas
            Dim oDeterminacion As New clsDeterminaciones
'M1171-F
            For i = 1 To lstMuestras.ListItems.Count
                opd.Carga PK, EDICION, lstMuestras.ListItems(i).SubItems(5)
                opd.setPRECIO = moneda_bd(Trim(lstMuestras.ListItems(i).SubItems(6)))
                opd.setEDICION = EDICION
                opd.Modificar
'M1171-I
                'ACTUALIZAMOS EL PRECIO DE LA DETERMINACIÓN
                oDeterminacion.CargarDeterminacion lstMuestras.ListItems(i).SubItems(5)
                oTipoContratas.CargaContrataDeter cmbSubcontratas.getPK_SALIDA, oDeterminacion.getTIPO_DETERMINACION_ID
                oTipoContratas.setPRECIO = moneda_bd(lstMuestras.ListItems(i).SubItems(6))
                oTipoContratas.Modificar
'M1171-F
            Next i
'M0957-F
            oSC_Paquete.Modificar PK, EDICION
            lngPaquete = PK
            Set oTipoContratas = Nothing
            Set oDeterminacion = Nothing
        Else
            Exit Sub
        End If
        MsgBox "El paquete se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
        Unload Me
        
      End If
'      frmSC_Ensayos_subcontratan_listado.cargar_lista

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdOk_Click of Formulario frmSC_Paquete_Detalle"
End Sub

Private Sub cmdcancel_Click()
    PK = 0
    Unload Me
End Sub
' --------------------------

Private Sub lstMuestras_Click()

    If lstMuestras.ListItems.Count > 0 Then
        txtNuevoPrecio.BackColor = &H80FFFF
        txtNuevoPrecio.ForeColor = vbBlue
        txtNuevoPrecio.Text = moneda(lstMuestras.ListItems(lstMuestras.selectedItem.Index).SubItems(6))
    End If
    
End Sub

Private Sub lstMuestras_DblClick()
   On Error GoTo lstMuestras_DblClick_Error

    If lstMuestras.ListItems.Count > 0 Then
        If IsLoadForm("frmVerMuestra") Then
            MsgBox "La ventana de detalle de muestra ya esta abierta. No esta permitido abrirla dos veces.", vbCritical, App.Title
        Else
            gmuestra = lstMuestras.ListItems(lstMuestras.selectedItem.Index).SubItems(7)
            frmVerMuestra.Show 1
        End If
    End If

   On Error GoTo 0
   Exit Sub

lstMuestras_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lstMuestras_DblClick of Formulario frmSC_Paquete_Detalle"
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    Select Case Index
        Case 2, 3:
            
        Case Else
            KeyAscii = 0
    End Select
End Sub

' funciones auxiliares del formulario


Public Sub CARGAR()
    Dim oSC_Paquete As New clsSC_Paquetes
    Dim rs As ADODB.Recordset
    Dim usu As New clsUsuarios
    Dim lngTotalDeterminacionesPaquete As Long
    
    Me.MousePointer = vbHourglass
'M1274-I
'    If oSC_Paquete.Carga(PK) = True Then
     If oSC_Paquete.Carga(PK, EDICION) = True Then
'M1274-F
        With oSC_Paquete
            txtDatos(1) = .getCODIGO_SC
            cmbCentro.BoundText = .getCENTRO_ID
            lblSubtitulo(2) = "Detalle del paquete: " & .getCODIGO_SC
            txtDatos(2) = moneda(.getPRESUPUESTO)
            txtDatos(3) = .getOBSERVACIONES
'M1274-I
            txtEdicion = EDICION 'CE y DETER solo contempla modificaciones sin cambio de versión
'M1274-F
            cmbSubcontratas.MostrarElemento .getSUBCONTRATA_ID
            cmbUsuario.MostrarElemento .getUSUARIO_ID
'M0959-I
            txtFactura = .getNFACTURA
            If .getFFACTURA <> "" Then
                datFechaFactura = .getFFACTURA
            End If
'M0959-F
            If IsDate(.getFECHA_CREACION) Then
                datFecha = .getFECHA_CREACION
                txtFechaPeticion.Text = Format(.getFECHA_CREACION, "yyyy-mm-dd")
                If usu.CARGAR(.getUSUARIO_ID) Then
                    txtUsuarioPeticion.Text = usu.getNOMBRE & " " & usu.getAPELLIDOS
                Else
                    txtUsuarioPeticion.Text = "N/A"
                End If
            Else
                txtFechaPeticion.Text = " -- "
                txtUsuarioPeticion.Text = "#Error en la fecha de recepción#"
            End If
            If IsDate(.getFECHA_APROBACION) Then
                txtFechaTramite.Text = Format(.getFECHA_APROBACION, "yyyy-mm-dd")
                If usu.CARGAR(.getAPROBADOR_ID) Then
                    txtUsuarioTramite.Text = usu.getNOMBRE & " " & usu.getAPELLIDOS
                Else
                    txtUsuarioTramite.Text = "N/A"
                End If
            Else
                txtFechaTramite.Text = "N/A"
                txtUsuarioTramite.Text = "N/A"
            End If
            If IsDate(.getFECHA_RECEPCION) Then
                txtFechaRecepcion.Text = Format(.getFECHA_RECEPCION, "yyyy-mm-dd")
                If usu.CARGAR(.getRECEPTOR_ID) Then
                    txtUsuarioRecepcion.Text = usu.getNOMBRE & " " & usu.getAPELLIDOS
                Else
                    txtUsuarioRecepcion.Text = "N/A"
                End If
            Else
                txtFechaRecepcion.Text = "N/A"
                txtUsuarioRecepcion.Text = "N/A"
            End If
            
            lstMuestras.ListItems.Clear
'M1274-I
'            Set RS = oSC_Paquete.Listado_muestras_determinaciones(.getID_PAQUETE)
            Set rs = oSC_Paquete.Listado_muestras_determinaciones(.getID_PAQUETE, EDICION)
            lngTotalDeterminacionesPaquete = rs.RecordCount
            If rs.RecordCount <> 0 Then
                Do
                    With lstMuestras.ListItems.Add(, , rs(0)) ' Código
                        .SubItems(1) = rs(1)                  ' ref Cliente
                        .SubItems(2) = rs(2)                  ' Determinación
                        .SubItems(3) = rs(3)                  ' Valor de referencia
                        .SubItems(4) = rs(4)                  ' Normativa aplicable
                        .SubItems(5) = rs(5)                  ' ID_DETERMINACION
                        .SubItems(6) = moneda(Trim(rs(6)))    ' PRECIO
                        .ListSubItems(6).ForeColor = vbBlue
                        .ListSubItems(6).bold = True
                        .SubItems(7) = rs(7)                  ' ID_MUESTRA
                        
                    End With
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            
        End With
    End If
    SC_bloqueaModificacion (oSC_Paquete.getESTADO)
    Me.MousePointer = vbNormal
    Set oSC_Paquete = Nothing
    txtSuma = SumarImportes
     
    lblSubtitulo(2) = lblSubtitulo(2) & " - Nº determinaciones: " & lngTotalDeterminacionesPaquete

End Sub
Public Sub SC_bloqueaModificacion(estadoPaquete As Long)
    If Not USUARIO.getPER_TRAMITACION_CONTRATA Then
        Select Case estadoPaquete
        Case SC_ESTADO_RECIBIDO
            cmdEliminar.Enabled = False
            cmdok.Enabled = False
            cmdAceptar.Enabled = False
            txtNuevoPrecio.Enabled = False
        Case SC_ESTADO_HISTORICO
            cmdEliminar.Enabled = False
            cmdok.Enabled = False
            cmdAceptar.Enabled = False
            txtNuevoPrecio.Enabled = False
        End Select
    End If
End Sub
Public Function datos_correctos() As Boolean
    datos_correctos = True
    
    If Trim(txtDatos(2)) = "" Then ' presupuesto
        If MsgBox("No ha indicado ningún presupuesto. ¿Modificar el paquete sin presupuesto?", vbYesNo + vbInformation, App.Title) = vbNo Then
        datos_correctos = False
        txtDatos(2).SetFocus
        Exit Function
        End If
    End If
    If Trim(txtDatos(3)) = "" Then ' observaciones
        If MsgBox("No ha indicado ningúna observación. ¿Modificar el paquete sin observaciones?", vbYesNo + vbInformation, App.Title) = vbNo Then
        datos_correctos = False
        txtDatos(3).SetFocus
        Exit Function
        End If
    End If
    If cmbCentro.BoundText = "" Then
        MsgBox "No ha indicado el centro.", vbExclamation, App.Title
        datos_correctos = False
        cmbCentro.SetFocus
        Exit Function
    End If
End Function

Private Sub cargar_combo_subcontratas()
    llenar_combo cmbSubcontratas, New clsProveedor, 0, frmProveedores_Detalle, " ES_SUBCONTRATA = 1 "
End Sub

Private Function SumarImportes() As String
    If lstMuestras.ListItems.Count = 0 Then
        SumarImportes = 0
        Exit Function
    End If
    Dim indice As Integer
    Dim Suma As Double
    
    Suma = 0
    For indice = 1 To lstMuestras.ListItems.Count
        If lstMuestras.ListItems(indice).SubItems(6) <> "" Then
            Suma = Suma + CDbl(lstMuestras.ListItems(indice).SubItems(6))
        End If
    Next indice
    SumarImportes = moneda(CStr(Suma))
End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    On Error GoTo Form_KeyUp_Error

    Me.MousePointer = 0
    Select Case KeyCode
        Case 116 ' F5 Datos especiales
            If frmHistoria.Visible = False Then
                Frame3.ForeColor = SC_COLOR_PENDIENTE
                Frame4.ForeColor = SC_COLOR_TRAMITADO
                Frame5.ForeColor = SC_COLOR_RECIBIDO
                frmHistoria.Visible = True
            Else
                frmHistoria.Visible = False
            End If
    End Select

   On Error GoTo 0
   Exit Sub

Form_KeyUp_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_KeyUp of Formulario frmEmpleados_Matriz"

End Sub

Private Sub Command1_Click()
    frmHistoria.Visible = False
End Sub

