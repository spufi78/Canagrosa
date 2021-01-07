VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmREX_Bote_Recepcion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recepción de Bote de Reactivo"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7980
   Icon            =   "frmREX_Bote_Recepcion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   7980
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNumeroBotes 
      Alignment       =   2  'Center
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
      Height          =   375
      Left            =   1710
      TabIndex        =   46
      Top             =   7650
      Width           =   1005
   End
   Begin VB.TextBox txtIdAnt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   43
      Top             =   8955
      Visible         =   0   'False
      Width           =   2325
   End
   Begin VB.TextBox txtRutaAnt 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Top             =   9000
      Visible         =   0   'False
      Width           =   3810
   End
   Begin VB.TextBox txtproveedor 
      Height          =   330
      Left            =   6840
      TabIndex        =   41
      Text            =   "0"
      Top             =   810
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.CheckBox chkCert 
      Caption         =   "Check1"
      Height          =   195
      Left            =   5895
      TabIndex        =   38
      Top             =   8235
      Visible         =   0   'False
      Width           =   1590
   End
   Begin VB.TextBox datos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DataSource      =   "Adodc1"
      Height          =   330
      IMEMode         =   3  'DISABLE
      Index           =   1
      Left            =   1665
      TabIndex        =   37
      Top             =   8325
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Selección del Certificado Externo"
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
      Height          =   1230
      Left            =   45
      TabIndex        =   35
      Top             =   6390
      Width           =   7875
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar"
         Height          =   825
         Left            =   6930
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   270
         Width           =   810
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         DataSource      =   "Adodc1"
         Height          =   825
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   540
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   270
         Width           =   4620
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   825
         Index           =   0
         Left            =   5220
         Picture         =   "frmREX_Bote_Recepcion.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   270
         Width           =   810
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escaner"
         Height          =   825
         Left            =   6075
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   270
         Width           =   825
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ruta"
         BeginProperty Font 
            Name            =   "Verdana"
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
         TabIndex        =   36
         Top             =   585
         Width           =   405
      End
   End
   Begin VB.CheckBox chkEtiqueta 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Imprimir etiquetas para los reactivos"
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
      Height          =   375
      Left            =   45
      TabIndex        =   19
      Top             =   8010
      Value           =   1  'Checked
      Width           =   6675
   End
   Begin VB.CheckBox chktodos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Recepcionar todos los botes con los mismos datos"
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
      Height          =   375
      Left            =   45
      TabIndex        =   18
      Top             =   7650
      Width           =   6675
   End
   Begin VB.TextBox txtboteanterior 
      Height          =   285
      Left            =   4140
      TabIndex        =   34
      Top             =   8730
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.TextBox txttotal 
      Height          =   285
      Left            =   2025
      TabIndex        =   33
      Top             =   8640
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8460
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   6900
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   8460
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos comúnes de recepción"
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
      Height          =   5655
      Left            =   45
      TabIndex        =   22
      Top             =   720
      Width           =   7875
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   4950
         TabIndex        =   47
         Top             =   2385
         Width           =   1950
         Begin VB.OptionButton opHenkel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "No"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   0
            Left            =   1080
            TabIndex        =   49
            Top             =   270
            Value           =   -1  'True
            Width           =   795
         End
         Begin VB.OptionButton opHenkel 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Si"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Index           =   1
            Left            =   315
            TabIndex        =   48
            Top             =   270
            Width           =   555
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "HENKEL"
            Height          =   195
            Index           =   10
            Left            =   585
            TabIndex        =   50
            Top             =   0
            Width           =   645
         End
      End
      Begin VB.CheckBox chkNoCaduca 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "No caduca"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3150
         TabIndex        =   7
         Top             =   2205
         Width           =   1275
      End
      Begin VB.TextBox txtdias 
         Height          =   375
         Left            =   5805
         TabIndex        =   39
         Text            =   "0"
         Top             =   2160
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   315
         Index           =   4
         Left            =   1485
         TabIndex        =   4
         Top             =   1395
         Width           =   6240
      End
      Begin VB.CheckBox chkAbrir 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Abrir bote"
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3150
         TabIndex        =   9
         Top             =   2610
         Width           =   1275
      End
      Begin VB.OptionButton opAceptado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   2190
         TabIndex        =   13
         Top             =   5220
         Width           =   795
      End
      Begin VB.OptionButton opAceptado 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Si"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   1500
         TabIndex        =   12
         Top             =   5220
         Value           =   -1  'True
         Width           =   555
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1455
         Index           =   3
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   3735
         Width           =   7575
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Index           =   2
         Left            =   1485
         TabIndex        =   5
         Top             =   1755
         Width           =   6240
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   0
         Left            =   1485
         TabIndex        =   0
         Top             =   315
         Width           =   1425
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   1
         Left            =   4905
         TabIndex        =   1
         Top             =   315
         Width           =   2805
      End
      Begin MSComCtl2.DTPicker fcaducidad 
         Height          =   390
         Left            =   1470
         TabIndex        =   6
         Top             =   2130
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   52101121
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fapertura 
         Height          =   390
         Left            =   1470
         TabIndex        =   8
         Top             =   2580
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   52101121
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   1485
         TabIndex        =   3
         Top             =   1035
         Width           =   6240
         _ExtentX        =   11007
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
      Begin MSComCtl2.DTPicker frecepcion 
         Height          =   390
         Left            =   1470
         TabIndex        =   10
         Top             =   3015
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   52101121
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Bindings        =   "frmREX_Bote_Recepcion.frx":0BD4
         Height          =   315
         Left            =   1485
         TabIndex        =   2
         Top             =   675
         Width           =   6240
         _ExtentX        =   11007
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
         Left            =   180
         TabIndex        =   45
         Top             =   720
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepción"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   44
         Top             =   3060
         Width           =   780
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   9
         Left            =   180
         TabIndex        =   31
         Top             =   1080
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   30
         Top             =   1440
         Width           =   840
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aceptado"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   29
         Top             =   5310
         Width           =   690
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   5
         Left            =   240
         TabIndex        =   28
         Top             =   3510
         Width           =   1065
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Apertura"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   27
         Top             =   2640
         Width           =   600
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Caducidad"
         Height          =   195
         Index           =   13
         Left            =   180
         TabIndex        =   26
         Top             =   2235
         Width           =   765
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Lote"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   25
         Top             =   1815
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F. Pedido"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   24
         Top             =   390
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   3
         Left            =   4275
         TabIndex        =   23
         Top             =   345
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   180
      Top             =   8730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recepción de Bote de Reactivo"
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
      Height          =   195
      Left            =   90
      TabIndex        =   40
      Top             =   360
      Width           =   7125
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Recepción de Bote de Reactivo"
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
      Height          =   300
      Index           =   2
      Left            =   90
      TabIndex        =   32
      Top             =   45
      Width           =   5805
      WordWrap        =   -1  'True
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   8055
   End
End
Attribute VB_Name = "frmREX_Bote_Recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nbote As Integer
Dim escaner As Boolean

'Private Sub bttDateReset_Click()
'    frecepcion.value = Date
'End Sub

'Private Sub bttTimeReset_Click()
'    TimePicker.Day = Date
'    TimePicker.value = Time
'End Sub

Private Sub chkNoCaduca_Click()
    If chkNoCaduca.Value = Checked Then
        fcaducidad.Enabled = False
    Else
        fcaducidad.Enabled = True
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdEscaner_Click()
    escaner = True
    frmEscaner.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = ""
        nombreNuevo = InputBox("Introduzca nombre para el archivo Escaneado, sin ninguna extensión (SOLO LETRAS Y NUMEROS).", "Escaneando Archivo", , Screen.Width / 3, (Screen.Height / 3))
        nombreNuevo = Eliminar_Caracteres_Archivo(nombreNuevo)
        If Trim(nombreNuevo) <> "" Then
            datos(0).Text = documento_escaner
            datos(1).Text = nombreNuevo & ".pdf"
'            cmdAdjuntar_Click
        End If
    End If
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar = True Then
        Dim BOTE As Long
        Dim oPB As New clsPedidos_bote_ex
        Dim obt As New clsTipos_bote_ex
        Dim frecepcion As String
        'M1340-I
        Dim obe As New clsBotes_ex
        Dim strBotes As String
        'M1340-F
        oPB.cargar_con_bote gpedido, gTipo_Bote
        obt.cargar (oPB.getTIPO_BOTE_EX_ID)
        
        If chktodos.Value = Unchecked Then
            cmdok.Enabled = False
            'M1340-I
            log "REX-BOTE-RECEPCION (1)"
            insertar_bote nbote, oPB.getTIPO_BOTE_EX_ID, True
            log "REX-BOTE-RECEPCION (2)"
            'M1340-F
            oPB.recibir_bote gpedido, gTipo_Bote
            log "REX-BOTE-RECEPCION (3)"
   
'            If opb.getCANTIDAD > opb.getCANTIDAD_RECIBIDA Then
            If oPB.getCANTIDAD * obt.getCANTIDAD_UNIDAD_PEDIDO > oPB.getCANTIDAD_RECIBIDA Then
                log "REX-BOTE-RECEPCION (4). opb.getCANTIDAD : " & oPB.getCANTIDAD
                log "REX-BOTE-RECEPCION (4). opb.getCANTIDAD_RECIBIDA : " & oPB.getCANTIDAD_RECIBIDA
                If oPB.getCANTIDAD * obt.getCANTIDAD_UNIDAD_PEDIDO - (oPB.getCANTIDAD_RECIBIDA + 1) > 0 Then
                    MsgBox "Bote insertado correctamente. Quedan " & oPB.getCANTIDAD * obt.getCANTIDAD_UNIDAD_PEDIDO - (oPB.getCANTIDAD_RECIBIDA + 1), vbOKOnly + vbInformation, App.Title
                    Label1(2).Caption = "Recepción de Bote de Reactivo nº " & nbote & " de " & txttotal
                    txtDatos(2).SetFocus
                    log "REX-BOTE-RECEPCION (5)"
                End If
            End If
        Else
            If txtNumeroBotes = "" Then
                MsgBox "Indique el número de botes a recepcionar con los mismo datos.", vbExclamation, App.Title
                txtNumeroBotes.SetFocus
                Exit Sub
            End If
            If Not IsNumeric(txtNumeroBotes) Then
                MsgBox "El número de botes a recepcionar no es correcto.", vbExclamation, App.Title
                txtNumeroBotes.SetFocus
                Exit Sub
            End If
            If CInt(txtNumeroBotes) > CInt(txttotal) Then
                MsgBox "El número de botes a recepcionar no puede ser mayor que lo pendiente.", vbExclamation, App.Title
                txtNumeroBotes.SetFocus
                Exit Sub
            End If
            log "REX-BOTE-RECEPCION (6)"
            Dim CANTIDAD As Integer
            cmdok.Enabled = False
            For i = 1 To CInt(txtNumeroBotes)
                Label1(2).Caption = "Recepción de Bote de Reactivo nº " & nbote & " de " & txttotal
                BOTE = insertar_bote(nbote, oPB.getTIPO_BOTE_EX_ID, False)
                If chkEtiqueta.Value = Checked Then
                    If i = 1 Then
                        strBotes = CStr(BOTE)
                    Else
                        strBotes = strBotes & "," & CStr(BOTE)
                    End If
                End If
'20170425-I
                CANTIDAD = CANTIDAD + 1
                If obt.getCANTIDAD_UNIDAD_PEDIDO = CANTIDAD Then
                    CANTIDAD = 0
'20170425-F
                    oPB.recibir_bote gpedido, gTipo_Bote
'20170425-I
                End If
'20170425-F
                DoEvents
            Next
            'M1340-I
            'Llamada a generación de etiquetas con IN
            If chkEtiqueta.Value = Checked And strBotes <> "" Then
                log "REX-BOTE-RECEPCION (8)"
                obe.imprimir_etiqueta strBotes 'Listado de reactivos
                DoEvents
                log "REX-BOTE-RECEPCION (9)"
            End If
            'M1340-F
            log "REX-BOTE-RECEPCION (7). strBotes : " & strBotes
        End If
        log "REX-BOTE-RECEPCION (10)"
        oPB.cargar_con_bote gpedido, gTipo_Bote
        log "REX-BOTE-RECEPCION (11)"
        If oPB.getCANTIDAD = oPB.getCANTIDAD_RECIBIDA Then
'        If opb.getCANTIDAD * obt.getCANTIDAD_UNIDAD_PEDIDO = opb.getCANTIDAD_RECIBIDA Then
            log "REX-BOTE-RECEPCION (12)"
            oPB.Recibir gpedido, gTipo_Bote, True
            log "REX-BOTE-RECEPCION (13)"
            MsgBox "Registro de Botes de Reactivos completado.", vbOKOnly + vbInformation, App.Title
        End If
        cmdok.Enabled = True
        log "REX-BOTE-RECEPCION (14)"
        If txtproveedor <> "" Then
            If IsNumeric(txtproveedor) Then
                log "REX-BOTE-RECEPCION (15)"
                frmProveedores_Evaluacion.PK = txtproveedor
                frmProveedores_Evaluacion.L_DESCRIPCION = "PEDIDO " & txtDatos(4)
                frmProveedores_Evaluacion.L_FECHA = frecepcion
                frmProveedores_Evaluacion.Show 1
            End If
        End If
        log "REX-BOTE-RECEPCION (16)"
        Unload Me
        
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmREX_Bote_Recepcion"
End Sub
Private Sub datos_GotFocus(Index As Integer)
    datos(Index).BackColor = &HC0FFFF
End Sub

Private Sub datos_LostFocus(Index As Integer)
    datos(Index).BackColor = vbWhite
End Sub

Private Sub Form_Activate()
'    log "Activate. nbote : " & nbote
'    If Not escaner Then
'        If nbote = 1 Then
'            cargar_pedido
'        End If
'    End If
'    escaner = False
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    escaner = False
    nbote = 1
'    Cargar_Combo cmbtipo, New clsTipos_m_referencia
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, DECODIFICADORA.REX_TIPOS
    cargar_combo cmbCentro, New clsCentros
    fcaducidad.Value = Date
    frecepcion.Value = Date
    fapertura.Value = Date
    txtUsuario = USUARIO.getUSUARIO
    
    cargar_pedido

End Sub
Private Sub cargar_pedido()
    Dim oPE As New clsPedidos_bote_ex
    Dim obt As New clsTipos_bote_ex
    Dim ore As New clsTipos_reactivo_ex
    oPE.cargar_con_bote gpedido, gTipo_Bote
    obt.cargar (oPE.getTIPO_BOTE_EX_ID)
    ' Cantidad a recibir sera la cantidad pedida * unidades de cada paquete
'    txttotal = ope.getCANTIDAD - (ope.getCANTIDAD_RECIBIDA)
    If obt.getCANTIDAD_UNIDAD_PEDIDO = 1 Then
        lblsubtitulo = "Se han pedido " & oPE.getCANTIDAD & " Unidad/es."
    Else
        lblsubtitulo = "Se han pedido " & oPE.getCANTIDAD & " Paquetes. Cada paquete contiene " & obt.getCANTIDAD_UNIDAD_PEDIDO & " Unidades"
    End If
    txttotal = (oPE.getCANTIDAD * obt.getCANTIDAD_UNIDAD_PEDIDO) - (oPE.getCANTIDAD_RECIBIDA)
    txtNumeroBotes = txttotal
    Label1(2).Caption = "Recepción de Bote de Reactivo nº " & nbote & " de " & txttotal
    txtDatos(0) = Format(oPE.getFECHA, "dd/mm/yyyy")
    txtDatos(1) = obt.getCODIGO
    cmbTipo.BoundText = obt.getTIPO_M_REFERENCIA_ID
    cmbCentro.BoundText = oPE.getCENTRO_ID
    ' Si es un material certificado, incluir la ruta del certificado
    Select Case obt.getTIPO_M_REFERENCIA_ID
        Case 2, 3, 6
            If USUARIO.getPER_PEDIDOS_REACTIVOS = False Then
                MsgBox "Usted no tiene permiso para recepcionar reactivos certificados.", vbExclamation, App.Title
                Unload Me
                Exit Sub
            End If
            MsgBox "La recepción pertenece a botes de reactivos certificados. Añada la ruta del pdf con la certificación.", vbInformation, App.Title
            chkCert.Value = Checked
'            datos(0).SetFocus
'        Case Else
'            datos(0).Enabled = False
'            cmdEXplorar(0).Enabled = False
'            cmdMostrar.Enabled = False
    End Select
    
    txtproveedor.Text = oPE.getPROVEEDOR_ID
    
    ore.cargar (obt.getTIPO_REACTIVO_EX_ID)
    txtDatos(4) = ore.getNOMBRE
    If obt.getTIPO_CADUCIDAD_ID <> 0 Then
        Dim oTC As New clsTipos_caducidad
        oTC.cargar obt.getTIPO_CADUCIDAD_ID
        txtdias = oTC.getDIAS
        fcaducidad = frecepcion + CInt(txtdias)
    End If
End Sub
Private Sub frecepcion_Change()
    fcaducidad = frecepcion + CInt(txtdias)
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(2)) = "" Then
        MsgBox "Debe introducir el lote.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If chkNoCaduca.Value = Unchecked Then
        If fcaducidad <= frecepcion Then
            MsgBox "La fecha de caducidad debe ser mayor que la de recepción.", vbInformation, App.Title
            validar = False
            Exit Function
        End If
    End If
    ' Si es un material certificado, validar que se ha insertado el certificado
'    If datos(0).Enabled = True Then
    If chkCert.Value = Checked Then
        If Trim(datos(0)) = "" Then
            MsgBox "Introduzca el certificado externo del reactivo.", vbInformation, App.Title
            validar = False
            Exit Function
        Else
            If Dir(datos(0)) = "" Then
                MsgBox "El certificado externo del reactivo no existe.", vbInformation, App.Title
                validar = False
                Exit Function
            End If
        End If
    End If
    'M1166-I
    Dim oTIPO As New clsTipos_bote_ex
    Dim oPB As New clsPedidos_bote_ex
    oPB.cargar_con_bote gpedido, gTipo_Bote
    oTIPO.cargar oPB.getTIPO_BOTE_EX_ID
    If oTIPO.getRESPONSABLE_ID = 0 And (oTIPO.getTIPO_M_REFERENCIA_ID = 2 Or oTIPO.getTIPO_M_REFERENCIA_ID = 3 Or oTIPO.getTIPO_M_REFERENCIA_ID = 6 Or oTIPO.getTIPO_M_REFERENCIA_ID = 7) Then
        MsgBox "Debe definir primero un responsable para este tipo de reactivo (M.R ó M.R.C)", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    Set oTIPO = Nothing
    Set oPB = Nothing
    'M1166-F
    If cmbCentro.Text = "" Then
        MsgBox "Debe indicar el Centro de destino del reactivo.", vbCritical, App.Title
        validar = False
        Exit Function
    End If
End Function

'M1340: Añadimos el control para que genere o no la etiqueta desde aqui
Private Function insertar_bote(numero_bote As Integer, TIPO_BOTE As Long, ETIQUETA As Boolean) As Long
    Dim obe As New clsBotes_ex
    Dim oEvaluacion As clsRex_botes_certificados
    log "insertar_bote-I"
    Dim BOTE As Long
    BOTE = 0
    With obe
       .setTIPO_BOTE_EX_ID = TIPO_BOTE
       .setLOTE = txtDatos(2)
       .setFECHA_PEDIDO = Format(txtDatos(0), "yyyy-mm-dd")
       If chkNoCaduca.Value = Unchecked Then
           .setFECHA_CADUCIDAD = Format(fcaducidad.Value, "yyyy-mm-dd")
           .setNO_CADUCA = 0
       Else
           .setFECHA_CADUCIDAD = "0000-00-00"
           .setNO_CADUCA = 1
       End If
       If opHenkel(1).Value = True Then
        .setHENKEL = 1
       Else
        .setHENKEL = 0
       End If
       .setFECHA_RECEPCION = Format(frecepcion.Value, "yyyy-mm-dd")
       If chkAbrir.Value = Checked Then
            .setFECHA_APERTURA = Format(fapertura.Value, "yyyy-mm-dd")
            .setABIERTO = 1
       Else
            .setFECHA_APERTURA = "0000-00-00"
            .setABIERTO = 0
       End If
       .setOBSERVACIONES = txtDatos(3)
       If opAceptado(0).Value = True Then
        .setANULADO = 0
       Else
        .setANULADO = 1
       End If
'       .setCERTIFICADO_EXTERNO = datos(0)
       .setCERTIFICADO_EXTERNO = ""
       'M1076-I
       .setPEDIDO_BOTE_EX_ID = gpedido
       'M1076-F
       'M1166-I
       'JGM-I
       Dim oTIPO As New clsTipos_bote_ex
       oTIPO.cargar TIPO_BOTE
       .setFECHA_CERTIFICACION = "0000-00-00"
       'JGM-F
        .setUSUARIO_CERTIFICADOR = oTIPO.getRESPONSABLE_ID
       'M1166-F
       .setCENTRO_ID = cmbCentro.BoundText
       BOTE = .Insertar
       If BOTE > 0 Then
          ' Certificado externo
          If datos(0) <> "" Then
              log "frmREX_Evaluacion_Parametros-adjuntar-I"
              adjuntar CLng(BOTE)
              log "frmREX_Evaluacion_Parametros-adjuntar-F"
          End If
          ' Si es un material certificado, validar que se ha insertado el certificado
          log "frmREX_Evaluacion_Parametros, numero_bote : " & numero_bote
          If numero_bote = 1 Then
'             If datos(0).Enabled = True Then
             If chkCert.Value = Checked Then
             'M1166-I
             '    frmREX_evaluacion.BOTE_EX_ID = BOTE
             '    frmREX_evaluacion.consulta = False
             '    frmREX_evaluacion.Show 1
                  log "frmREX_Evaluacion_Parametros-I"
                  frmREX_Evaluacion_Parametros.BOTE_EX_ID = BOTE
                  frmREX_Evaluacion_Parametros.Show 1
             'M1166-F
                  txtboteanterior = BOTE
                  log "frmREX_Evaluacion_Parametros-F"
             End If
          Else ' Si no es el primero y es certificado, copio sus datos
'             If datos(0).Enabled = True Then
             If chkCert.Value = Checked Then
                log "oEvaluacion-I"
                Set oEvaluacion = New clsRex_botes_certificados
                oEvaluacion.Carga CLng(txtboteanterior)
                oEvaluacion.setBOTE_EX_ID = BOTE
                oEvaluacion.setC03_INVENTARIO = BOTE
                oEvaluacion.Insertar
                Set oEvaluacion = Nothing
                log "oEvaluacion-F"
             End If
          End If
         nbote = nbote + 1
         'M1340-I
         'If chkEtiqueta.value = Checked Then
         If chkEtiqueta.Value = Checked And ETIQUETA = True Then
         'M1340-F
             obe.imprimir_etiqueta CStr(BOTE)
         End If
       End If
    End With
    log "insertar_bote-F"
    Set obe = Nothing
    insertar_bote = BOTE
End Function

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
'    cd.InitDir = "c:\"
    cd.ShowOpen
    If cd.FileName <> "" Then
'        datos(4).Text = cd.FileTitle 'cd.FileName  '
        datos(0).Text = cd.FileName
        datos(1).Text = cd.FileTitle
    End If
End Sub

Private Sub cmdMostrar_Click()
    On Error GoTo fallo
    If datos(0) <> "" Then
        If Dir(datos(0)) <> "" Then
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & datos(0), vbMaximizedFocus)
        End If
    End If
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title

End Sub

Private Sub adjuntar(BOTE As Long)
   On Error GoTo adjuntar_Error

    On Error Resume Next
    
    'M0601-I
'    Dim RUTA As String
'    RUTA = ReadINI(App.Path + "\config.ini", "Documentos", "rex_certificados")
'    MkDir RUTA & "\" & CStr(BOTE)
'   On Error GoTo adjuntar_Error
'    FileCopy datos(0), RUTA & "\" & CStr(BOTE) & "\" & datos(1)
'    Dim oBote As New clsBotes_ex
'    oBote.setCERTIFICADO_EXTERNO = RUTA & "\" & CStr(BOTE) & "\" & datos(1)
'    oBote.InformarRutaCertificado BOTE
'    Set oBote = Nothing
'    datos(0) = RUTA & "\" & CStr(BOTE) & "\" & datos(1)
    ' insertar adjunto
    Dim oAdjunto As New clsAdjuntos
    Dim adjunto As Long
    If datos(0) = txtRutaAnt Then
        adjunto = txtIdAnt
    Else
        adjunto = 0
    End If
    With oAdjunto
        .setTIPO = TOBJETO.TOBJETO_REX_CERTIFICADOS
        .setCODIGO = BOTE
        .setCODIGO_DECODIFICADORA = 0
        .setTIPO_DOCUMENTO_ID = 9
        .setOBSERVACIONES = ""
        .setFICHERO_NOMBRE = datos(1)
        .setFICHERO_RUTA = datos(0)
        adjunto = .Insertar(adjunto, False)
    End With
    If adjunto > 0 Then
        ' Actualizar el codigo del adjunto
        Dim oBote As New clsBotes_ex
        oBote.setCERTIFICADO_EXTERNO = adjunto & " - " & datos(1)
        oBote.InformarRutaCertificado BOTE
        Set oBote = Nothing
        ' Almacenar la ruta anterior
        txtRutaAnt = datos(0)
        txtIdAnt = adjunto
    End If
    Set oAdjunto = Nothing
    'M0601-F

   On Error GoTo 0
   Exit Sub

adjuntar_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure adjuntar of Formulario frmREX_Bote_Recepcion"
End Sub

