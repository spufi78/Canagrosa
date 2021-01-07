VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmProyectos_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7440
   ClientLeft      =   15
   ClientTop       =   -15
   ClientWidth     =   10185
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmProyectos_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   10185
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Asignación de usuarios y dedicación"
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
      Left            =   90
      TabIndex        =   22
      Top             =   3915
      Width           =   10005
      Begin VB.CommandButton cmdDesadignar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Quitar del Proyecto"
         Height          =   780
         Left            =   2070
         Picture         =   "frmProyectos_Detalle.frx":000C
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   1260
         Width           =   1500
      End
      Begin VB.CommandButton cmdasignar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Asingar al Proyecto"
         Height          =   780
         Left            =   270
         Picture         =   "frmProyectos_Detalle.frx":08D6
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   1260
         Width           =   1500
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2280
         Left            =   4050
         TabIndex        =   23
         Top             =   180
         Width           =   5880
         _ExtentX        =   10372
         _ExtentY        =   4022
         View            =   3
         LabelEdit       =   1
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
      Begin MSDataListLib.DataCombo cmbUsuAsig 
         Bindings        =   "frmProyectos_Detalle.frx":11A0
         Height          =   315
         Left            =   90
         TabIndex        =   25
         Top             =   675
         Width           =   3855
         _ExtentX        =   6800
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
         Caption         =   "Usuario"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   24
         Top             =   405
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9045
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6525
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7965
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6525
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del proyecto"
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
      Height          =   3300
      Left            =   45
      TabIndex        =   7
      Top             =   540
      Width           =   10080
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   5
         Left            =   5355
         MaxLength       =   255
         TabIndex        =   20
         Top             =   2475
         Width           =   1545
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   4
         Left            =   1650
         MaxLength       =   255
         TabIndex        =   18
         Top             =   2475
         Width           =   1545
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         Enabled         =   0   'False
         Height          =   330
         Index           =   3
         Left            =   5355
         MaxLength       =   255
         TabIndex        =   15
         Top             =   2115
         Width           =   1545
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   1650
         MaxLength       =   255
         TabIndex        =   2
         Top             =   2115
         Width           =   1545
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1005
         Index           =   1
         Left            =   1650
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   8265
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
         ForeColor       =   &H000000FF&
         Height          =   330
         Index           =   0
         Left            =   1650
         TabIndex        =   0
         Top             =   270
         Width           =   8265
      End
      Begin MSComCtl2.DTPicker f1 
         Height          =   345
         Left            =   1650
         TabIndex        =   3
         Top             =   2835
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   48234497
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker f2 
         Height          =   345
         Left            =   5355
         TabIndex        =   4
         Top             =   2835
         Width           =   1575
         _ExtentX        =   2778
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
         Format          =   48234497
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSDataListLib.DataCombo cmbUsuario 
         Bindings        =   "frmProyectos_Detalle.frx":11E6
         Height          =   315
         Left            =   1650
         TabIndex        =   17
         Top             =   1710
         Width           =   8265
         _ExtentX        =   14579
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
         Caption         =   "Inporte Real"
         Height          =   195
         Index           =   6
         Left            =   3645
         TabIndex        =   21
         Top             =   2520
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Importe Planificado"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   19
         Top             =   2520
         Width           =   1350
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Horas Dedicadas"
         Height          =   195
         Index           =   2
         Left            =   3645
         TabIndex        =   16
         Top             =   2160
         Width           =   1230
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1755
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de Finalización"
         Height          =   195
         Index           =   0
         Left            =   3645
         TabIndex        =   13
         Top             =   2925
         Width           =   1545
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de Inicio"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   12
         Top             =   2925
         Width           =   1095
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Horas Planificadas"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   11
         Top             =   2160
         Width           =   1320
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Comentarios"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   990
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   315
         Width           =   840
      End
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9585
      Picture         =   "frmProyectos_Detalle.frx":122C
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0E0FF&
      Caption         =   "Detalle de Proyecto"
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
      TabIndex        =   10
      Top             =   120
      Width           =   2085
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00C0E0FF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   10305
   End
End
Attribute VB_Name = "frmProyectos_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Private Sub chkop1_Click()
'    If chkop1.value = Checked Then
'        f1.Enabled = True
'    Else
'        f1.Enabled = False
'    End If
End Sub

Private Sub chkop2_Click()
'    If chkop2.value = Checked Then
'        f2.Enabled = True
'    Else
'        f2.Enabled = False
'    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
      Dim oEC As New clsEquipos_calibracion
      With oEC
        .Eliminar PK
        .setEQUIPO_ID = PK
        .setMODALIDAD = txtDatos(0)
        .setPROCEDIMIENTO = txtDatos(1)
        .setRESPONSABLE = txtDatos(2)
        .setPERIODO = txtDatos(3)
'        If chkop1.value = Checked Then
'            .setFECHA_CALIBRACION = Format(f1, "yyyy-mm-dd")
'        Else
'            .setFECHA_CALIBRACION = FNULA
'        End If
'        If chkop2.value = Checked Then
'            .setFECHA_SIGUIENTE_CALIBRACION = Format(f2, "yyyy-mm-dd")
'        Else
'            .setFECHA_SIGUIENTE_CALIBRACION = FNULA
'        End If
        .Insertar
      End With
      MsgBox "La calibración del equipo se ha actualizado correctamente.", vbOKOnly + vbInformation, App.Title
      Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Dim titulo As String
    If PK <> 0 Then
        cargar
    End If
End Sub
Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub cargar()
    Dim oEquipo As New clsEquipos
    If oEquipo.Carga(PK) = True Then
        lbltitulo = "Calibración del Equipo : " & oEquipo.getNOMBRE
        Me.Caption = lbltitulo
        Dim oEC As New clsEquipos_calibracion
        If oEC.Carga(PK) Then
            With oEC
                txtDatos(0) = .getMODALIDAD
                txtDatos(1) = .getPROCEDIMIENTO
                txtDatos(2) = .getRESPONSABLE
                txtDatos(3) = .getPERIODO
 '               If .getFECHA_CALIBRACION = FNULA Or Not IsDate(.getFECHA_CALIBRACION) Then
 '                   chkop1.value = Unchecked
 '                   f1.Enabled = False
 '               Else
 '                   chkop1.value = Checked
 '                   f1 = Format(.getFECHA_CALIBRACION, "dd-mm-yyyy")
 '               End If
 '               If .getFECHA_SIGUIENTE_CALIBRACION = FNULA Or Not IsDate(.getFECHA_SIGUIENTE_CALIBRACION) Then
 '                   chkop2.value = Unchecked
 '                   f2.Enabled = False
 '               Else
 '                   chkop2.value = Checked
 '                   f2 = Format(.getFECHA_SIGUIENTE_CALIBRACION, "dd-mm-yyyy")
 '               End If
            End With
        End If
    End If
    Set oEquipo = Nothing
End Sub
