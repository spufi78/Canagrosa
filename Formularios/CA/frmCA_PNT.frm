VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmCA_PNT 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de nuevo PNT"
   ClientHeight    =   10785
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12255
   Icon            =   "frmCA_PNT.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10785
   ScaleWidth      =   12255
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkesPNT 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Es PNT"
      Height          =   195
      Left            =   9990
      TabIndex        =   42
      Top             =   225
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdrevisado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "2. Revisado"
      Height          =   825
      Left            =   5220
      Picture         =   "frmCA_PNT.frx":2AFA
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   9855
      Width           =   1140
   End
   Begin VB.CommandButton cmdterminado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "1. Creado"
      Height          =   825
      Left            =   2880
      Picture         =   "frmCA_PNT.frx":33C4
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   9855
      Width           =   1140
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos de la Modificación (Nuevas ediciones)"
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
      Height          =   2310
      Left            =   45
      TabIndex        =   34
      Top             =   3015
      Width           =   12165
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
         Left            =   1035
         MaxLength       =   255
         TabIndex        =   12
         Top             =   1530
         Width           =   1290
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   870
         Index           =   3
         Left            =   3645
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1350
         Width           =   7245
      End
      Begin pryCombo.miCombo cmbusuario 
         Height          =   330
         Index           =   4
         Left            =   3645
         TabIndex        =   7
         Top             =   270
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbusuario 
         Height          =   330
         Index           =   6
         Left            =   3645
         TabIndex        =   11
         Top             =   990
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Index           =   5
         Left            =   1035
         TabIndex        =   8
         Top             =   630
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbusuario 
         Height          =   330
         Index           =   5
         Left            =   3645
         TabIndex        =   9
         Top             =   630
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Index           =   4
         Left            =   1035
         TabIndex        =   6
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Index           =   6
         Left            =   1035
         TabIndex        =   10
         Top             =   990
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aprobación"
         Height          =   195
         Index           =   9
         Left            =   135
         TabIndex        =   47
         Top             =   1035
         Width           =   810
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Creación"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   46
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Revisado por"
         Height          =   195
         Index           =   5
         Left            =   2565
         TabIndex        =   45
         Top             =   675
         Width           =   945
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edición"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   39
         Top             =   1575
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Revisión"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   38
         Top             =   675
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modificación"
         Height          =   195
         Index           =   4
         Left            =   2565
         TabIndex        =   37
         Top             =   1575
         Width           =   900
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realizado por"
         Height          =   195
         Index           =   16
         Left            =   2565
         TabIndex        =   36
         Top             =   315
         Width           =   975
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aprobado por"
         Height          =   195
         Index           =   14
         Left            =   2565
         TabIndex        =   35
         Top             =   1035
         Width           =   960
      End
   End
   Begin VB.CommandButton cmdModificado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver/Modificar"
      Height          =   825
      Left            =   90
      Picture         =   "frmCA_PNT.frx":3C8E
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   9855
      Width           =   1140
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fechas y usuarios de creación (Edición 1)"
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
      Height          =   1410
      Left            =   45
      TabIndex        =   24
      Top             =   1575
      Width           =   12165
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Index           =   3
         Left            =   1035
         TabIndex        =   4
         Top             =   990
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Index           =   2
         Left            =   1035
         TabIndex        =   2
         Top             =   630
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Index           =   1
         Left            =   1035
         TabIndex        =   0
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
         Format          =   16515073
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin pryCombo.miCombo cmbusuario 
         Height          =   330
         Index           =   1
         Left            =   3645
         TabIndex        =   1
         Top             =   270
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbusuario 
         Height          =   330
         Index           =   2
         Left            =   3645
         TabIndex        =   3
         Top             =   630
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   582
      End
      Begin pryCombo.miCombo cmbusuario 
         Height          =   330
         Index           =   3
         Left            =   3645
         TabIndex        =   5
         Top             =   990
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   582
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aprobación"
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   30
         Top             =   1035
         Width           =   810
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Aprobado por"
         Height          =   195
         Index           =   12
         Left            =   2565
         TabIndex        =   29
         Top             =   1035
         Width           =   960
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Revisión"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   28
         Top             =   675
         Width           =   615
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Revisado por"
         Height          =   195
         Index           =   10
         Left            =   2565
         TabIndex        =   27
         Top             =   675
         Width           =   945
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Creación"
         Height          =   195
         Index           =   7
         Left            =   90
         TabIndex        =   26
         Top             =   315
         Width           =   630
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Realizado por"
         Height          =   195
         Index           =   0
         Left            =   2565
         TabIndex        =   25
         Top             =   315
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdgenera 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar Edición"
      Height          =   825
      Left            =   9360
      Picture         =   "frmCA_PNT.frx":4558
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   9855
      Width           =   1410
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   10800
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   9855
      Width           =   1410
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
      Height          =   960
      Index           =   1
      Left            =   45
      TabIndex        =   21
      Top             =   585
      Width           =   12150
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H00FF0000&
         Height          =   285
         Index           =   1
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   20
         Top             =   270
         Width           =   10875
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Height          =   285
         Index           =   2
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   585
         Width           =   10875
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   6
         Left            =   90
         TabIndex        =   23
         Top             =   315
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   22
         Top             =   630
         Width           =   840
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4095
      Left            =   45
      TabIndex        =   14
      Top             =   5670
      Width           =   12165
      _ExtentX        =   21458
      _ExtentY        =   7223
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
   Begin XtremeSuiteControls.PushButton cmdEliminarEdicion 
      Height          =   300
      Left            =   10620
      TabIndex        =   43
      Top             =   5355
      Width           =   1590
      _Version        =   851970
      _ExtentX        =   2805
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Eliminar Edición"
      Appearance      =   5
      Picture         =   "frmCA_PNT.frx":4E22
   End
   Begin XtremeSuiteControls.PushButton cmdModificarEdicion 
      Height          =   300
      Left            =   8820
      TabIndex        =   44
      Top             =   5355
      Width           =   1770
      _Version        =   851970
      _ExtentX        =   3122
      _ExtentY        =   529
      _StockProps     =   79
      Caption         =   "Modificar Edición"
      Appearance      =   5
      Picture         =   "frmCA_PNT.frx":B684
   End
   Begin VB.CommandButton cmdAprobado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "3. Aprobado"
      Height          =   825
      Left            =   7425
      Picture         =   "frmCA_PNT.frx":11EE6
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   9855
      Width           =   1140
   End
   Begin VB.Shape Shape1 
      Height          =   915
      Left            =   2835
      Top             =   9810
      Width           =   5775
   End
   Begin VB.Image img2 
      Height          =   465
      Left            =   6660
      Picture         =   "frmCA_PNT.frx":127B0
      Stretch         =   -1  'True
      Top             =   9990
      Width           =   465
   End
   Begin VB.Image img1 
      Height          =   495
      Left            =   4410
      Picture         =   "frmCA_PNT.frx":1307A
      Stretch         =   -1  'True
      Top             =   9990
      Width           =   495
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Rellene todos los campos para la creación de un nuevo PNT"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   33
      Top             =   285
      Width           =   4320
   End
   Begin VB.Image imagen 
      Height          =   435
      Left            =   11700
      Picture         =   "frmCA_PNT.frx":13944
      Stretch         =   -1  'True
      Top             =   45
      Width           =   435
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
      TabIndex        =   32
      Top             =   30
      Width           =   2760
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Modificaciones"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
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
      TabIndex        =   31
      Top             =   5355
      Width           =   11820
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   555
      Left            =   0
      Top             =   0
      Width           =   12350
   End
End
Attribute VB_Name = "frmCA_PNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Public traza As String
Private Enum COLS
    C_DOC_ID = 0
    C_EDICION = 1
    C_fecha = 2
    C_MODIFICACION = 3
    C_RESPONSABLE = 4
    C_F_REVISION = 5
    C_REVISOR = 6
    C_F_APROBACION = 7
    C_APROBADOR = 8
    C_ID_RESPONSABLE = 9
    C_ID_REVISOR = 10
    C_ID_APROBADOR = 11
End Enum

Private Sub cmbUsuario_Change(Index As Integer)

    Dim oPNT As New clsCa_pnt
    Dim EDICION As Integer
    If oPNT.Carga_Ultima_edicion(PK) Then
        EDICION = oPNT.getEDICION
    Else
        EDICION = 1
    End If
    If Index = 1 And EDICION = 1 Then
        cmbusuario(4).MostrarElemento cmbusuario(1).getPK_SALIDA
    End If
    If Index = 2 And EDICION = 1 Then
        cmbusuario(5).MostrarElemento cmbusuario(2).getPK_SALIDA
    End If
    If Index = 3 And EDICION = 1 Then
        cmbusuario(6).MostrarElemento cmbusuario(3).getPK_SALIDA
    End If
End Sub

Private Sub cmdAprobado_Click()
   On Error GoTo cmdAprobado_Click_Error
    'M2643-I
    ' Requisito NADCAP
    Dim oCA_NADCAP As New clsCa_documentos
    oCA_NADCAP.carga PK
    If oCA_NADCAP.getNADCAP = 1 Or oCA_NADCAP.getMTL = 1 Then
        If MsgBox("¿Ha verificado que se cumplen todos los requisitos NADCAP/MTL?", vbQuestion + vbYesNo, App.Title) = vbNo Then
            Exit Sub
        End If
    End If
    Set oCA_NADCAP = Nothing
    'M2643-F
    'M2841-I
    Dim strFecha As String
    strFecha = InputBox("Fecha de Aprobación del documento (dd/mm/yyyy) : ", App.Title, Date)
    If Trim(strFecha) = "" Then
        MsgBox "Es necesario indicar la fecha de Aprobación", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    If Not IsDate(strFecha) Then
        MsgBox "La fecha de Aprobación no tiene formato correcto", vbCritical + vbOKOnly, App.Title
        Exit Sub
    End If
    'M2841-F
    Dim oPNT As New clsCa_pnt
    Dim oCA As New clsCa_documentos
    oPNT.Carga_Ultima_edicion PK
    oCA.Modificar_Estado PK, C_CA_DOCUMENTOS_ESTADOS.CA_VIGOR
    ' Al aprobar, informamos la fecha de ultima modificación a la fecha del
    ' sistema, igual que la de aprobación, para que coincida la fecha de
    ' arriba con la de abajo.
    'JGM oPNT.Informar_Modificacion PK, oPNT.getEDICION, Format(Date, "yyyy-mm-dd")
    ' Informamos la fecha de aprobación con la fecha del sistema
    ' si es primera edición
    'M2841-I
'    oPNT.Informar_Aprobacion PK, oPNT.getEDICION, Date
    oPNT.Informar_Aprobacion PK, oPNT.getEDICION, CDate(strFecha)
    'M2841-F
    ' Ponemos en vigor el documento con la fecha de aprobación
    oCA.setEDICION = oPNT.getEDICION
    'M2841-I
'    oCA.setFECHA = Format(Date, "yyyy-mm-dd")
    oCA.setFECHA = Format(strFecha, "yyyy-mm-dd")
    'M2841-F
    oCA.Modificar_Edicion PK
    ' Informamos la ruta del fichero pdf o de trabajo para excel
    Dim documento As String
    Dim RUTA_TRABAJO As String
    Dim RUTA_PDF As String
    Dim EXTENSION As String
    Dim oDeco As New clsDecodificadora
    RUTA_TRABAJO = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Documentos\Trabajo\"
    RUTA_PDF = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Documentos\PDF\"
    ' Cargamos la descripción de la familia
    oCA.carga PK
    oDeco.Carga_valor DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS, oCA.getFAMILIA_ID
    ' Creamos la carpeta de la familia por si no existe
    RUTA_TRABAJO = RUTA_TRABAJO & oDeco.getDESCRIPCION
    RUTA_PDF = RUTA_PDF & oDeco.getDESCRIPCION
    ' Cargamos el tipo de plantilla
    oDeco.Carga_valor DECODIFICADORA.CALIDAD_PLANTILLAS_DOCUMENTOS, oCA.getPLANTILLA_ID
    Dim s() As String
    s = Split(oDeco.getPARAMETROS, ".")
    EXTENSION = "." & s(1)
'    EXTENSION = Right(oDeco.getPARAMETROS, 4)
    ' Nombre del documento
    If UCase(EXTENSION) = ".XLS" Or UCase(EXTENSION) = ".XLSX" Then
        documento = Replace(Eliminar_Caracteres_Archivo(Trim(oCA.getCODIGO)), ".", " ") & EXTENSION
        oCA.Informar_ruta PK, RUTA_TRABAJO & "\" & documento
    Else
        documento = Replace(Eliminar_Caracteres_Archivo(Trim(oCA.getCODIGO)), ".", " ") & ".pdf"
        oCA.Informar_ruta PK, RUTA_PDF & "\" & documento
    End If
    ' Generamos el word del pnt
    Dim oVida As New clsCa_documentos_vida
    With oVida
        .setIDENTIFICADOR = PK
        .setMOTIVO = "Generación de Nueva edición : " & oPNT.getEDICION
        .setTIPO_ID = CALIDAD_VIDA_TIPOS.CALIDAD_VIDA_TIPOS_DOCUMENTO
        .setSUBTIPO_ID = CALIDAD_VIDA_SUBTIPOS.CALIDAD_VIDA_SUBTIPOS_NUEVA_EDICION
        .setUSUARIO = USUARIO.getID_EMPLEADO
        .Insertar
    End With
    ' Enviar correo de distribución al aprobar una edición
    Dim destinatario As String
    Dim mensaje As String
    Dim ASUNTO As String
    Dim oParametro As New clsParametros
    oParametro.carga parametros.ENVIO_CORREO_PNT, ""
    If oParametro.getVALOR = 1 Then
        oParametro.carga parametros.CORREO_DISTRIBUCION_CALIDAD, ""
        destinatario = oParametro.getVALOR
        If destinatario <> "" Then
            ASUNTO = "Nueva edición de documento de calidad, código : " & txtDatos(1)
            mensaje = "Se ha aprobado una edición del siguiente documento de calidad: " & vbNewLine & vbNewLine
            mensaje = mensaje & " Generación de Nueva edición : " & oPNT.getEDICION
            mensaje = mensaje & " Código : " & txtDatos(1) & vbNewLine
            mensaje = mensaje & " Descripción : " & txtDatos(2) & vbNewLine
            mensaje = mensaje & " Aprobado por : " & USUARIO.getUSUARIO & vbNewLine
            
            Dim rs As ADODB.Recordset
            Dim c As String
            c = "SELECT A.ID_EQUIPO, A.NOMBRE, A.SERIE,A.MODELO " & _
                " FROM EQUIPOS A, EQ_NORMAS_EQUIPOS B " & _
                " WHERE A.ID_EQUIPO = B.EQUIPO_ID " & _
                "   AND B.DOCUMENTO_ID = " & PK & _
                "   AND TIPO = 0 " & _
                " ORDER BY A.NOMBRE "
            Set rs = datos_bd(c)
            If rs.RecordCount > 0 Then
                mensaje = mensaje & vbNewLine
                mensaje = mensaje & vbNewLine & "        ** LISTA DE EQUIPOS A REVISAR AL TENER LA NORMA ASIGNADA **"
                mensaje = mensaje & vbNewLine & "---------------------------------------------------------------------------------"
                mensaje = mensaje & vbNewLine & "NºEQUIPO                NOMBRE                       SERIE           MODELO "
                mensaje = mensaje & vbNewLine & "---------------------------------------------------------------------------------"
                Do
                    mensaje = mensaje & vbNewLine
                    mensaje = mensaje & Format(rs(0), "!" & String(8, "@")) & " "
                    mensaje = mensaje & Format(Left(rs(1), 40), "!" & String(40, "@")) & " "
                    mensaje = mensaje & Format(Left(rs(2), 15), "!" & String(15, "@")) & " "
                    mensaje = mensaje & Format(Left(rs(3), 15), "!" & String(15, "@")) & " "
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            
            mensaje = mensaje & vbNewLine
            'M1384-I: documentos asociados/vinculados
            'mensaje = mensaje & vbNewLine
            Dim oCP As New clsCa_documentos_PNT
            Set rs = oCP.Listado(PK)
            If rs.RecordCount > 0 Then
                mensaje = mensaje & vbNewLine
                mensaje = mensaje & vbNewLine & "---------------------------------------------------------------------------------"
                mensaje = mensaje & vbNewLine & "           POR FAVOR, REVISE LA SIGUIENTE LISTA DE DOCUMENTOS ASOCIADOS        **"
                mensaje = mensaje & vbNewLine & "---------------------------------------------------------------------------------"
'JGM                mensaje = mensaje & vbNewLine & "COD_DOCUMENTO                DESCRIPCION                                   "
                mensaje = mensaje & vbNewLine & "                                 DESCRIPCION                                   "
                mensaje = mensaje & vbNewLine & "---------------------------------------------------------------------------------"
                Do
                    mensaje = mensaje & vbNewLine
'JGM                    mensaje = mensaje & Format(Left(rs(0), 18), "!" & String(8, "@")) & " "
                    mensaje = mensaje & Format(Left(rs(1), 80), "!" & String(80, "@")) & " "
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            Set oCP = Nothing
            mensaje = mensaje & vbNewLine
            mensaje = mensaje & vbNewLine
            'M1384-F
            mensaje = mensaje & " Mensaje enviado automáticamente desde Geslab, el " & Format(Date, "dd-mm-yyyy") & " a las " & Format(Time, "hh:mm:ss")
            ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
        End If
    End If
    Set oParametro = Nothing
    crear_pnt PK, True
    MsgBox "Se ha aprobado el documento correctamente.", vbInformation, App.Title
    Unload Me

   On Error GoTo 0
   Exit Sub

cmdAprobado_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAprobado_Click of Formulario frmCA_PNT"
End Sub

Private Sub cmdEliminarEdicion_Click()
    If lista.ListItems.Count > 0 Then
      If MsgBox("Va a ELIMINAR la edición seleccionada. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
      End If
        Dim oCA_PNT As New clsCa_pnt
        oCA_PNT.Eliminar lista.ListItems(lista.selectedItem.Index), lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_EDICION)
        Set oCA_PNT = Nothing
        lista.ListItems.Remove lista.selectedItem.Index
    End If
End Sub

Private Sub cmdgenera_Click()
   On Error GoTo cmdgenera_Click_Error

    escribe_traza "Pulsado Genera edición."
    If validar = True Then
      If MsgBox("Va a generar la edición " & txtDatos(0) & " del documento. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
      End If
      Dim oPNT As New clsCa_pnt
      With oPNT
        .setDOCUMENTO_ID = PK
        .setEDICION = txtDatos(0)
        .setCODIGO = txtDatos(1)
        .setDESCRIPCION = txtDatos(2)
        .setMODIFICACION = txtDatos(3)
        .setFECHA_CREACION = Format(fecha(1), "yyyy-mm-dd")
        .setFECHA_REVISION = Format(fecha(2), "yyyy-mm-dd")
        .setFECHA_APROBACION = Format(fecha(3), "yyyy-mm-dd")
        .setUSUARIO_CREACION = cmbusuario(1).getPK_SALIDA
        .setUSUARIO_REVISION = cmbusuario(2).getPK_SALIDA
        .setUSUARIO_APROBACION = cmbusuario(3).getPK_SALIDA
        
        .setMODIFICACION_FECHA = Format(fecha(4), "yyyy-mm-dd")
        .setMODIFICACION_FECHA_REVISION = Format(fecha(5), "yyyy-mm-dd")
        .setMODIFICACION_FECHA_APROBACION = Format(fecha(6), "yyyy-mm-dd")
        .setMODIFICACION_USUARIO_ = cmbusuario(4).getPK_SALIDA
        .setMODIFICACION_USUARIO_REVISION = cmbusuario(5).getPK_SALIDA
        .setMODIFICACION_USUARIO_APROBACION = cmbusuario(6).getPK_SALIDA
        escribe_traza "ID : " & PK
        escribe_traza "CODIGO : " & txtDatos(1)
        escribe_traza "EDICION : " & txtDatos(0)
        Dim oCA As New clsCa_documentos
        Dim oPNT_AUX As New clsCa_pnt
        If oPNT_AUX.carga(PK, txtDatos(0)) Then ' Existe la edición, la modificamos
            escribe_traza "Existe la edición"
            If .Modificar(PK, txtDatos(0)) = True Then
              oCA.Modificar_Estado PK, C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_MODIFICACION
            End If
        Else
            escribe_traza "Insertar edición"
            If .Insertar > 0 Then
              If oCA.carga(PK) Then
                If CInt(oCA.getEDICION) <= CInt(txtDatos(0)) Then
                    oCA.setEDICION = CInt(txtDatos(0))
                    oCA.setFECHA = Format(fecha(4), "dd-mm-yyyy")
                    oCA.Modificar_Estado PK, C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_MODIFICACION
                End If
              End If
            End If
        End If
        escribe_traza "Edición generada correctamente"
        MsgBox "Datos almacenados correctamente.", vbInformation, App.Title
        crear_pnt PK, False
      End With
      Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdgenera_Click_Error:

    error_grave_jgm "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdgenera_Click of Formulario frmCA_PNT" & vbNewLine & traza

End Sub
Private Sub cmdModificado_Click()
    Dim documento As String
   On Error GoTo cmdModificado_Click_Error
   escribe_traza "Pulsado botón Ver/Modificar"
    documento = calidad_ruta_documento_trabajo(PK)
    escribe_traza "Documento : " & documento
    
'1018-I (Numero de anotaciones)
    Dim ocada As New clsCa_documentos_anotaciones
    Dim cont As Integer
    cont = ocada.anotaciones(PK)
    If cont <> 0 Then
        MsgBox "Existen anotaciones del documento. Recuerde introducir los cambios.", vbExclamation, App.Title
    End If
    Set ocada = Nothing
'1018-F

    If Dir(documento) <> "" Then
        If UCase(Right(documento, 3)) <> "DOC" Then
            escribe_traza "Ver documento No Word : " & documento
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & documento, vbMaximizedFocus)
            escribe_traza "Terminado Ver documento No Word : " & documento
        Else
            escribe_traza "Ver documento Word : " & documento
            ver_documento_word documento
            escribe_traza "Terminado Ver documento Word : " & documento
        End If
    Else
        MsgBox "No localizo el documento con ese código.", vbExclamation, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cmdModificado_Click_Error:

    error_grave_jgm "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificado_Click of Formulario frmCA_PNT" & vbNewLine & traza

End Sub

Private Sub cmdModificarEdicion_Click()
   On Error GoTo cmdModificarEdicion_Click_Error

    If lista.ListItems.Count > 0 Then
      If MsgBox("Va a MODIFICAR la edición seleccionada. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbNo Then
        Exit Sub
      End If
        escribe_traza "Pulsado MODIFICAR edición."
        If validar = True Then
          Dim oPNT As New clsCa_pnt
          With oPNT
            .setDOCUMENTO_ID = PK
            .setEDICION = txtDatos(0)
            .setCODIGO = txtDatos(1)
            .setDESCRIPCION = txtDatos(2)
            .setMODIFICACION = txtDatos(3)
            .setFECHA_CREACION = Format(fecha(1), "yyyy-mm-dd")
            .setFECHA_REVISION = Format(fecha(2), "yyyy-mm-dd")
            .setFECHA_APROBACION = Format(fecha(3), "yyyy-mm-dd")
            .setUSUARIO_CREACION = cmbusuario(1).getPK_SALIDA
            .setUSUARIO_REVISION = cmbusuario(2).getPK_SALIDA
            .setUSUARIO_APROBACION = cmbusuario(3).getPK_SALIDA
            
            .setMODIFICACION_FECHA = Format(fecha(4), "yyyy-mm-dd")
            .setMODIFICACION_FECHA_REVISION = Format(fecha(5), "yyyy-mm-dd")
            .setMODIFICACION_FECHA_APROBACION = Format(fecha(6), "yyyy-mm-dd")
            .setMODIFICACION_USUARIO_ = cmbusuario(4).getPK_SALIDA
            .setMODIFICACION_USUARIO_REVISION = cmbusuario(5).getPK_SALIDA
            .setMODIFICACION_USUARIO_APROBACION = cmbusuario(6).getPK_SALIDA
            escribe_traza "ID : " & PK
            escribe_traza "CODIGO : " & txtDatos(1)
            escribe_traza "EDICION : " & txtDatos(0)
                
            .Modificar PK, txtDatos(0)
            escribe_traza "Edición modificada correctamente"
            cargarLista PK
            MsgBox "Datos almacenados correctamente.", vbInformation, App.Title
          End With
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdModificarEdicion_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdModificarEdicion_Click of Formulario frmCA_PNT"
End Sub

Private Sub cmdrevisado_Click()
    If MsgBox("Va a dar por Revisado el PNT. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        'M2643-I
        ' Requisito NADCAP
        Dim oCA_NADCAP As New clsCa_documentos
        oCA_NADCAP.carga PK
        If oCA_NADCAP.getNADCAP = 1 Or oCA_NADCAP.getMTL = 1 Then
            If MsgBox("¿Ha verificado que se cumplen todos los requisitos NADCAP/MTL?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
        End If
        Set oCA_NADCAP = Nothing
        'M2643-F
    
    
        Dim oPNT As New clsCa_pnt
        oPNT.Carga_Ultima_edicion PK
'        oPNT.Informar_Revision PK, Date
        oPNT.Informar_Revision PK, oPNT.getEDICION, Date
        ' Enviar correo al revisor
        Dim destinatario As String
        Dim ASUNTO As String
        Dim mensaje As String
        Dim oUsuario As New clsUsuarios
'JGM        oUsuario.cargar cmbusuario(3).getPK_SALIDA
        If oPNT.getEDICION = 1 Then
          oUsuario.CARGAR oPNT.getUSUARIO_APROBACION
        Else
          oUsuario.CARGAR oPNT.getMODIFICACION_USUARIO_APROBACION
        End If
        destinatario = oUsuario.getEMAIL
        ASUNTO = "Aprobación de " & txtDatos(1)
        mensaje = "Tiene usted pendiente la Aprobación de : " & vbNewLine & vbNewLine
        mensaje = mensaje & " Código : " & txtDatos(1) & vbNewLine
        mensaje = mensaje & " Descripción : " & txtDatos(2) & vbNewLine
        mensaje = mensaje & " Creado por : " & cmbusuario(1).getTEXTO & vbNewLine
        mensaje = mensaje & " Revisado por : " & cmbusuario(2).getTEXTO & vbNewLine
        mensaje = mensaje & " Mensaje enviado automáticamente desde Geslab, el " & Format(Date, "dd-mm-yyyy") & " a las " & Format(Time, "hh:mm:ss")
        If destinatario <> "" Then
            Dim oParametro As New clsParametros
            oParametro.carga parametros.ENVIO_CORREO_PNT, ""
            If oParametro.getVALOR = 1 Then
                ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
            End If
            Set oParametro = Nothing
        End If
        ' Mensaje en geslab
'JGM        enviar_mensaje ASUNTO, mensaje, cmbusuario(3).getPK_SALIDA
        If oPNT.getEDICION = 1 Then
            enviar_mensaje ASUNTO, mensaje, oPNT.getUSUARIO_APROBACION
        Else
            enviar_mensaje ASUNTO, mensaje, oPNT.getMODIFICACION_USUARIO_APROBACION
        End If
        ' Marcar el PNT como PDTE. DE APROBACION
        Dim oCA As New clsCa_documentos
        oCA.Modificar_Estado PK, C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_APROBACION
        Unload Me
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdterminado_Click()
    If MsgBox("Va a dar por terminadas las modificaciones del PNT. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        'M2643-I
        ' Si tiene documentos vinculados, indicar mensaje de revisión
        Dim oCPNT As New clsCa_documentos_PNT
        Dim rsPNT As ADODB.Recordset
        Set rsPNT = oCPNT.Listado(PK)
        If rsPNT.RecordCount > 0 Then
            MsgBox "Existen documentos vinculados. Recuerde que debe revisarlos.", vbExclamation, App.Title
        End If
        Set rsPNT = Nothing
        Set oCPNT = Nothing
        ' Requisito NADCAP
        Dim oCA_NADCAP As New clsCa_documentos
        oCA_NADCAP.carga PK
        If oCA_NADCAP.getNADCAP = 1 Or oCA_NADCAP.getMTL = 1 Then
            If MsgBox("¿Ha verificado que se cumplen todos los requisitos NADCAP/MTL?", vbQuestion + vbYesNo, App.Title) = vbNo Then
                Exit Sub
            End If
        End If
        Set oCA_NADCAP = Nothing
        'M2643-F
    
     ' Si es primera versión, revisor y aprobador
        ' Enviar correo al revisor
        Dim oPNT As New clsCa_pnt
        oPNT.Carga_Ultima_edicion PK
        oPNT.Informar_Modificacion PK, oPNT.getEDICION, Date

        Dim destinatario As String
        Dim ASUNTO As String
        Dim mensaje As String
        Dim oUsuario As New clsUsuarios
        If oPNT.getEDICION = 1 Then
          oUsuario.CARGAR oPNT.getUSUARIO_REVISION
        Else
          oUsuario.CARGAR oPNT.getMODIFICACION_USUARIO_REVISION
        End If
        Dim tipo As String
'JGM        If oPNT.getEDICION = 1 And chkesPNT.Value = Checked Then
        If chkesPNT.Value = Checked Then
            tipo = "Revisión"
        Else
            tipo = "Aprobación"
        End If
        destinatario = oUsuario.getEMAIL
        ASUNTO = tipo & " de " & txtDatos(1)
        mensaje = "Tiene usted pendiente la " & tipo & " de : " & vbNewLine & vbNewLine
        mensaje = mensaje & " Código : " & txtDatos(1) & vbNewLine
        mensaje = mensaje & " Descripción : " & txtDatos(2) & vbNewLine
        If oPNT.getEDICION = 1 Then
            mensaje = mensaje & " Creado por : " & cmbusuario(1).getTEXTO & vbNewLine
        Else
            mensaje = mensaje & " Creado por : " & cmbusuario(4).getTEXTO & vbNewLine
        End If
        mensaje = mensaje & " Mensaje enviado automáticamente desde Geslab, el " & Format(Date, "dd-mm-yyyy") & " a las " & Format(Time, "hh:mm:ss")
        If destinatario <> "" Then
            Dim oParametro As New clsParametros
            oParametro.carga parametros.ENVIO_CORREO_PNT, ""
            If oParametro.getVALOR = 1 Then
                ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
            End If
            Set oParametro = Nothing
        End If
        ' Mensaje en geslab
'        If oPNT.getEDICION = 1 And chkesPNT.Value = Checked Then
        If chkesPNT.Value = Checked Then
            enviar_mensaje ASUNTO, mensaje, oPNT.getUSUARIO_REVISION
        Else
            enviar_mensaje ASUNTO, mensaje, oPNT.getMODIFICACION_USUARIO_REVISION
        End If
        ' Marcar el PNT como PDTE. REVISION
        Dim oCA As New clsCa_documentos
'JGM        If oPNT.getEDICION = 1 And chkesPNT.Value = Checked Then
        If chkesPNT.Value = Checked Then
            oCA.Modificar_Estado PK, C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_REVISION
        Else
            oCA.Modificar_Estado PK, C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_APROBACION
        End If
        ' Cerrar el formulario
        Unload Me
    End If
End Sub

Private Sub fecha_Change(Index As Integer)
    If Index = 1 Then
        fecha(4).Value = fecha(1).Value
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    traza = ""
    cargar_botones Me
'    txtDatos(0) = "1"
'    fecha(0) = Date
'    cargar_combos
    cabecera
    cargar_datos
    activar_botones
    'JGM-I
    If USUARIO.getPER_ADMIN_PNT = True Then
        cmdModificarEdicion.visible = True
        cmdEliminarEdicion.visible = True
        
    Else
        cmdModificarEdicion.visible = False
        cmdEliminarEdicion.visible = False
    End If
    If chkesPNT.Value = Unchecked Then
        fecha(5).visible = False
        cmbusuario(5).visible = False
        lblCampos(3).visible = False
        lblCampos(5).visible = False
    End If
    'JGM-F
End Sub

Private Sub img1_Click()
    If MsgBox("Va a dar enviar a modificar nuevamente el PNT. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
     ' Si es primera versión, revisor y aprobador
        ' Enviar correo al revisor
        Dim oPNT As New clsCa_pnt
        oPNT.Carga_Ultima_edicion PK
        Dim destinatario As String
        Dim ASUNTO As String
        Dim mensaje As String
        Dim oUsuario As New clsUsuarios
        If oPNT.getEDICION = 1 Then
          oUsuario.CARGAR oPNT.getUSUARIO_CREACION
        Else
'JGM          oUsuario.cargar oPNT.getMODIFICACION_USUARIO_REVISION
          oUsuario.CARGAR oPNT.getMODIFICACION_USUARIO_
        End If
        destinatario = oUsuario.getEMAIL
        ASUNTO = "Corrección de " & txtDatos(1)
        mensaje = "Tiene usted pendiente la corrección de : " & vbNewLine & vbNewLine
        mensaje = mensaje & " Código : " & txtDatos(1) & vbNewLine
        mensaje = mensaje & " Descripción : " & txtDatos(2) & vbNewLine
        mensaje = mensaje & " Mensaje enviado automáticamente desde Geslab, el " & Format(Date, "dd-mm-yyyy") & " a las " & Format(Time, "hh:mm:ss")
        If destinatario <> "" Then
            Dim oParametro As New clsParametros
            oParametro.carga parametros.ENVIO_CORREO_PNT, ""
            If oParametro.getVALOR = 1 Then
                ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
            End If
        End If
        ' Mensaje en geslab
        enviar_mensaje ASUNTO, mensaje, oUsuario.getID_EMPLEADO
        ' Marcar el PNT como PDTE. REVISION
        Dim oCA As New clsCa_documentos
        oCA.Modificar_Estado PK, C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_MODIFICACION
        ' Cerrar el formulario
        Unload Me
    End If

End Sub

Private Sub img2_Click()
    If MsgBox("Va a enviar a modificar nuevamente el PNT. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
     ' Si es primera versión, revisor y aprobador
        ' Enviar correo al revisor
        Dim oPNT As New clsCa_pnt
        oPNT.Carga_Ultima_edicion PK
        Dim destinatario As String
        Dim ASUNTO As String
        Dim mensaje As String
        Dim oUsuario As New clsUsuarios
        Dim tipo As String
        If oPNT.getEDICION = 1 Then
            If chkesPNT.Value = Checked Then
              oUsuario.CARGAR oPNT.getUSUARIO_REVISION
              tipo = "Revisión tras rechazo"
            Else
              oUsuario.CARGAR oPNT.getUSUARIO_CREACION
              tipo = "Corrección"
'jgm              oUsuario.cargar oPNT.getMODIFICACION_USUARIO_REVISION
            End If
        Else
            If chkesPNT.Value = Checked Then
                oUsuario.CARGAR oPNT.getMODIFICACION_USUARIO_REVISION
                tipo = "Revisión"
            Else
                oUsuario.CARGAR oPNT.getMODIFICACION_USUARIO_
                tipo = "Corrección"
            End If
        End If
'        If oPNT.getEDICION = 1 Then
'            If chkesPNT.Value = Checked Then
'                tipo = "Revisión"
'            Else
'                tipo = "Corrección"
'            End If
'        Else
'            tipo = "Corrección"
'        End If
        destinatario = oUsuario.getEMAIL
        ASUNTO = tipo & " de " & txtDatos(1)
        mensaje = "Tiene usted pendiente la " & tipo & " de : " & vbNewLine & vbNewLine
        mensaje = mensaje & " Código : " & txtDatos(1) & vbNewLine
        mensaje = mensaje & " Descripción : " & txtDatos(2) & vbNewLine
        mensaje = mensaje & " Mensaje enviado automáticamente desde Geslab, el " & Format(Date, "dd-mm-yyyy") & " a las " & Format(Time, "hh:mm:ss")
        If destinatario <> "" Then
            Dim oParametro As New clsParametros
            oParametro.carga parametros.ENVIO_CORREO_PNT, ""
            If oParametro.getVALOR = 1 Then
                ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
            End If
        End If
        ' Mensaje en geslab
        enviar_mensaje ASUNTO, mensaje, oUsuario.getID_EMPLEADO
        ' Marcar el PNT como PDTE. REVISION
        Dim oCA As New clsCa_documentos
'        If oPNT.getEDICION = 1 Then
            If chkesPNT.Value = Checked Then
                oCA.Modificar_Estado PK, C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_REVISION
            Else
                oCA.Modificar_Estado PK, C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_MODIFICACION
            End If
'        Else
'            If chkesPNT.Value = Checked Then
'                oCA.Modificar_Estado PK, C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_REVISION
'            Else
'                oCA.Modificar_Estado PK, C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_MODIFICACION
'            End If
'        End If
        ' Cerrar el formulario
        Unload Me
    End If

End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        ' Sólo el superusuario puede generar ediciones anteriores
'        If usuario.getPER_ADMIN_PNT = True Or lista.selectedItem.Index = 1 Then
        If USUARIO.getPER_ADMIN_PNT = True Then
            
            txtDatos(0) = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_EDICION)
            txtDatos(3) = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_MODIFICACION)
            
            If lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_fecha) <> "" Then
                fecha(4) = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_fecha)
            End If
            If lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_F_REVISION) <> "" Then
                fecha(5) = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_F_REVISION)
            Else
                fecha(5) = "14/01/2004"
            End If
            If lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_F_APROBACION) <> "" Then
                fecha(6) = lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_F_APROBACION)
            Else
                fecha(6) = "14/01/2004"
            End If
            If lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_ID_RESPONSABLE) <> "" Then
                cmbusuario(4).MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_ID_RESPONSABLE)
            Else
                cmbusuario(4).limpiar
            End If
            If lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_ID_REVISOR) <> "" Then
                cmbusuario(5).MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_ID_REVISOR)
            Else
                cmbusuario(5).limpiar
            End If
            If lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_ID_APROBADOR) <> "" Then
                cmbusuario(6).MostrarElemento lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_ID_APROBADOR)
            Else
                cmbusuario(6).limpiar
            End If
        End If
    End If
End Sub
Private Sub lista_DblClick()
   On Error GoTo lista_DblClick_Error

    If lista.ListItems.Count > 0 Then
        Dim oCA_Documento As New clsCa_documentos
        oCA_Documento.mostrarEdicion lista.ListItems(lista.selectedItem.Index).Text, CInt(lista.ListItems(lista.selectedItem.Index).SubItems(COLS.C_EDICION)), True
        Set oCA_Documento = Nothing
    End If

   On Error GoTo 0
   Exit Sub

lista_DblClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lista_DblClick of Formulario frmCA_PNT"
End Sub


Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub

Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 10 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(3)) = "" Then
        MsgBox "Debe indicar los datos de modificación/creación.", vbExclamation, App.Title
        txtDatos(3).SetFocus
        validar = False
        Exit Function
    End If
    ' Si es primera versión, validar los datos de creación, sino los de modificación
    If CInt(txtDatos(0)) = 1 And chkesPNT.Value = Checked Then
        If cmbusuario(1).getTEXTO = "" Then
            MsgBox "Debe seleccionar un usuario de creación.", vbExclamation, App.Title
            cmbusuario(1).SetFocus
            validar = False
            Exit Function
        End If
        If cmbusuario(2).getTEXTO = "" Then
            MsgBox "Debe seleccionar un usuario de revisión.", vbExclamation, App.Title
            cmbusuario(2).SetFocus
            validar = False
            Exit Function
        End If
        If cmbusuario(3).getTEXTO = "" Then
            MsgBox "Debe seleccionar un usuario de aprobación.", vbExclamation, App.Title
            cmbusuario(3).SetFocus
            validar = False
            Exit Function
        End If
    Else
        If cmbusuario(4).getTEXTO = "" Then
            MsgBox "Debe seleccionar un usuario de modificación.", vbExclamation, App.Title
            cmbusuario(4).SetFocus
            validar = False
            Exit Function
        End If
        If cmbusuario(5).visible = True Then
            If cmbusuario(5).getTEXTO = "" Then
                MsgBox "Debe seleccionar un usuario de revisión.", vbExclamation, App.Title
                cmbusuario(5).SetFocus
                validar = False
                Exit Function
            End If
        End If
        If cmbusuario(6).getTEXTO = "" Then
            MsgBox "Debe seleccionar un usuario de aprobación.", vbExclamation, App.Title
            cmbusuario(6).SetFocus
            validar = False
            Exit Function
        End If
    End If
    ' Si no es primera edición, validar que el documento existe
    If CInt(txtDatos(0)) > 1 Then
        If calidad_ruta_documento_trabajo(PK) = "" Then
            MsgBox "El documento de trabajo no existe. ¿Ha creado la primera edición?", vbExclamation, App.Title
            validar = False
            Exit Function
        End If
    End If
End Function

Private Sub cargar_combos(filtro As String)
    llenar_combo cmbusuario(1), New clsUsuarios, 0, Me, filtro
    llenar_combo cmbusuario(2), New clsUsuarios, 0, Me, filtro
    llenar_combo cmbusuario(3), New clsUsuarios, 0, Me, filtro
    llenar_combo cmbusuario(4), New clsUsuarios, 0, Me, filtro
    llenar_combo cmbusuario(5), New clsUsuarios, 0, Me, filtro
    llenar_combo cmbusuario(6), New clsUsuarios, 0, Me, filtro
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "DOCUMENTO_ID", 1, lvwColumnLeft
        .Add , , "Edición", 750, lvwColumnCenter
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Modificación", 4000, lvwColumnLeft
        .Add , , "Responsable", 1300, lvwColumnCenter
        .Add , , "F.Revisión", 1050, lvwColumnCenter
        .Add , , "Revisor", 1300, lvwColumnCenter
        .Add , , "F.Aprobación", 1050, lvwColumnCenter
        .Add , , "Aprobador", 1300, lvwColumnCenter
        .Add , , "ID_RESPONSABLE", 1, lvwColumnCenter
        .Add , , "ID_REVISOR", 1, lvwColumnCenter
        .Add , , "ID_APROBADOR", 1, lvwColumnCenter
    End With
End Sub
Private Sub crear_pnt(ID As Long, pdf As Boolean)
    On Error GoTo fallo
    ' Crear copia para su uso
    Dim oPNT_Modificaciones As New clsCa_pnt
    Dim oPNT As New clsCa_documentos
    Dim oDeco As New clsDecodificadora
    Dim EDICION As Integer
    escribe_traza "---------------"
    escribe_traza "Creación de PNT"
    escribe_traza "---------------"
    If oPNT_Modificaciones.Carga_Ultima_edicion(ID) Then
        EDICION = oPNT_Modificaciones.getEDICION
    End If
    If oPNT.carga(ID) Then
        ' Informamos las rutas del documento
        Dim documento As String
        Dim COPIA As String
        Dim RUTA_TRABAJO As String
        Dim RUTA_VERSIONES As String
        Dim EXTENSION As String
        RUTA_TRABAJO = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Documentos\Trabajo\"
        RUTA_VERSIONES = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Documentos\Trabajo\Versiones\"
        ' Cargamos la descripción de la familia
        oDeco.Carga_valor DECODIFICADORA.CA_DOCUMENTOS_FAMILIAS, oPNT.getFAMILIA_ID
        ' Creamos la carpeta de la familia por si no existe
        RUTA_TRABAJO = RUTA_TRABAJO & oDeco.getDESCRIPCION
        RUTA_VERSIONES = RUTA_VERSIONES & oDeco.getDESCRIPCION
        escribe_traza "RUTA_TRABAJO : " & RUTA_TRABAJO
        escribe_traza "RUTA_VERSIONES : " & RUTA_VERSIONES
        ' Cargamos el tipo de plantilla
        oDeco.Carga_valor DECODIFICADORA.CALIDAD_PLANTILLAS_DOCUMENTOS, oPNT.getPLANTILLA_ID
        Dim s() As String
        s = Split(oDeco.getPARAMETROS, ".")
        EXTENSION = "." & s(1)
'        EXTENSION = Right(oDeco.getPARAMETROS, 4)
        
        escribe_traza "EXTENSION : " & EXTENSION
        ' Nombre del documento y su copia (version)
        documento = Replace(Eliminar_Caracteres_Archivo(Trim(oPNT.getCODIGO)), ".", " ") & EXTENSION
        COPIA = Replace(Eliminar_Caracteres_Archivo(Trim(oPNT.getCODIGO)), ".", " ") & " Ed." & EDICION - 1 & EXTENSION
        escribe_traza "DOCUMENTO : " & documento
        escribe_traza "COPIA : " & COPIA
        On Error Resume Next
        MkDir RUTA_TRABAJO
        MkDir RUTA_VERSIONES
        On Error GoTo fallo
        ' Validar que existe el documento
        ' Verificamos si se esta regenerando la edición, en tal caso,
        ' copiamos el documento en lugar de la plantilla
        If EDICION = 1 Then
            If Dir(RUTA_TRABAJO & "\" & documento) = "" Then
                ' Copiamos la plantilla correspondiente
                escribe_traza "EDICION 1.NO EXISTE"
                Dim PLANTILLA As String
                PLANTILLA = ReadINI(App.Path + "\config.ini", "Documentos", "Plantillas") & "\" & Replace(oDeco.getPARAMETROS, "/", "\")
                FileCopy PLANTILLA, RUTA_TRABAJO & "\" & documento
                escribe_traza "EDICION 1.COPIA CORRECTA"
            Else
                escribe_traza "EDICION 1. EXISTE"
            End If
        Else
            ' Creamos la copia del documento
            If Dir(RUTA_TRABAJO & "\" & documento) <> "" Then
                FileCopy RUTA_TRABAJO & "\" & documento, RUTA_VERSIONES & "\" & COPIA
                escribe_traza "EDICION DISTINTA 1.COPIA CORRECTA"
            Else
                MsgBox "Ojo, no existe el documento en su ruta : " & RUTA_TRABAJO & "\" & documento
                escribe_traza "Ojo, no existe el documento en su ruta : " & RUTA_TRABAJO & "\" & documento
            End If
        End If
        ' Si es un documento word, continuamos el proceso, sino abrimos el focumento
        If UCase(Right(documento, 3)) <> "DOC" And UCase(Right(documento, 4)) <> "DOCX" Then
            If Dir(RUTA_TRABAJO & "\" & documento) <> "" Then
                r = Shell("rundll32.exe url.dll,FileProtocolHandler " & RUTA_TRABAJO & "\" & documento, vbMaximizedFocus)
            End If
            escribe_traza "VISUALIZAR DOCUMENTO DISTINTO A WORD"
            Exit Sub
        End If
        ' Abrimos el word
        If Dir(RUTA_TRABAJO & "\" & documento) <> "" Then
            Dim appword As Word.Application
            Dim docword As Word.Document
            Set appword = CreateObject("word.application")
            Set docword = appword.Documents.Open(RUTA_TRABAJO & "\" & documento)
            escribe_traza "ABRE WORD"
            If pdf = False Then
                appword.visible = True
            End If
            If chkesPNT.Value = Unchecked Then
                docword.Save
                docword.Close
                appword.Quit
                escribe_traza "NO ES PNT, SE MANDA A IMPRIMIR"
                imprimir ID, 40, False
                escribe_traza "NO ES PNT, SE MANDA A IMPRIMIR TERMINA"
                Exit Sub
            End If
        End If
    End If
    Dim oUsuario As New clsUsuarios
    If Dir(RUTA_TRABAJO & "\" & documento) <> "" Then
     If oPNT_Modificaciones.carga(ID, 1) Then
        escribe_traza "COMIENZO EDICION PLANTILLA PNT"
        ' Código y Título
        With docword.Tables(2)
            .Rows(1).Cells(1).Range.Text = oPNT.getCODIGO
            .Rows(2).Cells(1).Range.Text = oPNT.getNOMBRE
        End With
        escribe_traza "CODIGO Y TITULO"
        ' Pie del PNT
        With docword.Sections(1).Footers(1).Range.Tables(1)
            .Rows(1).Cells(2).Range.Text = oPNT.getNOMBRE
        End With
        escribe_traza "PIE"
        ' Cabecera
        With docword.Sections(1).Headers(1).Range.Tables(1)
            .Rows(2).Cells(3).Range.Text = Format(oPNT.getFECHA, "dd/mm/yyyy")
            .Rows(3).Cells(3).Range.Text = oPNT.getEDICION
        End With
        escribe_traza "CABECERA"
        If EDICION = 1 Then
           ' Datos de elaboración
           escribe_traza "COMIENZO EDICION 1"
           With docword.Tables(1)
               ' Fechas
               .Rows(3).Cells(1).Range.Text = "Fecha : " & Format(oPNT_Modificaciones.getFECHA_CREACION, "dd/mm/yyyy")
               .Rows(3).Cells(2).Range.Text = "Fecha : " & Format(oPNT_Modificaciones.getFECHA_REVISION, "dd/mm/yyyy")
               .Rows(3).Cells(3).Range.Text = "Fecha : " & Format(oPNT_Modificaciones.getFECHA_APROBACION, "dd/mm/yyyy")
               escribe_traza "FECHAS"
               ' Usuarios
               oUsuario.CARGAR oPNT_Modificaciones.getUSUARIO_CREACION
               .Rows(5).Cells(1).Range.Text = oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
               'Inc-740 Solo se incluye la firma si pdf=true (viene de aprobación)
               If oUsuario.getFIRMA <> "" And pdf Then
                   If Dir(oUsuario.getFIRMA) <> "" Then
'                       .Rows(4).Cells(1).Range.Cells.Delete
                       .Rows(4).Cells(1).Range.Delete
                       .Rows(4).Cells(1).Range.InlineShapes.AddPicture oUsuario.getFIRMA
                   End If
               End If
               escribe_traza "FIRMA 1"
               oUsuario.CARGAR oPNT_Modificaciones.getUSUARIO_REVISION
               .Rows(5).Cells(2).Range.Text = oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
               If oUsuario.getFIRMA <> "" And pdf Then
                   If Dir(oUsuario.getFIRMA) <> "" Then
                       .Rows(4).Cells(2).Range.Delete
                       .Rows(4).Cells(2).Range.InlineShapes.AddPicture oUsuario.getFIRMA
                   End If
               End If
               escribe_traza "FIRMA 2"
               oUsuario.CARGAR oPNT_Modificaciones.getUSUARIO_APROBACION
               .Rows(5).Cells(3).Range.Text = oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
               If oUsuario.getFIRMA <> "" And pdf Then
                   If Dir(oUsuario.getFIRMA) <> "" Then
                       .Rows(4).Cells(3).Range.Delete
                       .Rows(4).Cells(3).Range.InlineShapes.AddPicture oUsuario.getFIRMA
                   End If
               End If
               escribe_traza "FIRMA 3"
           End With
        End If
        ' Modificaciones
        Dim i As Integer
        escribe_traza "COMIENZO TABLA MODIFICACIONES"
        While docword.Tables(3).Rows.Count <= EDICION
            docword.Tables(3).Rows.Add
        Wend
        Dim fila As Integer
        fila = docword.Tables(3).Rows.Count
        If oPNT_Modificaciones.Carga_Ultima_edicion(ID) Then
            With docword.Tables(3)
                .Rows(fila).Cells(1).Range.Text = oPNT_Modificaciones.getEDICION
' Igualar la fecha de la edición en la aprobación
                If oPNT.getESTADO_ID = C_CA_DOCUMENTOS_ESTADOS.CA_VIGOR Then
                    .Rows(fila).Cells(2).Range.Text = Format(oPNT.getFECHA, "dd/mm/yyyy")
                Else
                    .Rows(fila).Cells(2).Range.Text = oPNT_Modificaciones.getMODIFICACION_FECHA
                End If

                .Rows(fila).Cells(3).Range.Text = oPNT_Modificaciones.getMODIFICACION
                oUsuario.CARGAR oPNT_Modificaciones.getMODIFICACION_USUARIO_
                'Inc-740 Solo se incluye la firma si pdf=true (viene de aprobación)
                If oUsuario.getFIRMA <> "" And pdf Then
                    If Dir(oUsuario.getFIRMA) <> "" Then
                        .Rows(fila).Cells(4).Range.Delete
                        .Rows(fila).Cells(4).Range.InlineShapes.AddPicture oUsuario.getFIRMA
                       'Inc-739 Se incluye el nombre debajo de la firma en la lista de modificaciones del PNT
                        .Rows(fila).Cells(4).Range.InsertAfter Chr(13) & oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
                    End If
       'Inc-739 Se incluye el nombre debajo de la firma en la lista de modificaciones del PNT
                Else
                    .Rows(fila).Cells(4).Range.Delete
                    .Rows(fila).Cells(4).Range.InsertAfter Chr(13) & oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
                End If
'                oUsuario.cargar oPNT_Modificaciones.getMODIFICACION_USUARIO_REVISION
                oUsuario.CARGAR oPNT_Modificaciones.getMODIFICACION_USUARIO_APROBACION
                If oUsuario.getFIRMA <> "" And pdf Then
                    If Dir(oUsuario.getFIRMA) <> "" Then
                        .Rows(fila).Cells(5).Range.Delete
                        .Rows(fila).Cells(5).Range.InlineShapes.AddPicture oUsuario.getFIRMA
                        'Inc-739 Se incluye el nombre debajo de la firma en la lista de modificaciones del PNT
                        .Rows(fila).Cells(5).Range.InsertAfter Chr(13) & oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
                    End If
       'Inc-739 Se incluye el nombre debajo de la firma en la lista de modificaciones del PNT
                Else
                
                    .Rows(fila).Cells(5).Range.Delete
                    .Rows(fila).Cells(5).Range.InsertAfter Chr(13) & oUsuario.getNOMBRE & " " & oUsuario.getAPELLIDOS
                    
                End If
            End With
        End If
        escribe_traza "FIN TABLA MODIFICACIONES"
     End If
     escribe_traza "GRABANDO WORD"
     docword.Save
     escribe_traza "FIN GRABANDO WORD"
     If pdf Then
        docword.Close
        ' Generamos el pdf
        escribe_traza "IMPRESION PDF"
        imprimir ID, 40, False
        escribe_traza "FIN IMPRESION PDF"
     End If
    End If
    escribe_traza "----------------------------"
    escribe_traza "FINALIZACION DE CREACION PNT"
    escribe_traza "----------------------------"
    If Dir(RUTA_TRABAJO & "\" & documento) <> "" Then
        Set docword = Nothing
        Set appword = Nothing
    End If
    Exit Sub
fallo:
    error_grave_jgm "Error al generar el PNT : " & Err.Description & vbNewLine & traza
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
End Sub
Private Sub cargarLista(ID As Long)
    Dim oPNT As New clsCa_pnt
    lista.ListItems.Clear
    Dim rs As ADODB.Recordset
    Set rs = oPNT.Listado(PK)
        If rs.RecordCount <> 0 Then
            Do
                With lista.ListItems.Add(, , PK)
                 .SubItems(1) = rs(0) ' EDICION
                 .SubItems(2) = Format(rs(1), "dd/mm/yyyy") ' FECHA
                 .SubItems(3) = rs(2) ' MODIFICACION
                 .SubItems(4) = rs(3) ' RESPONSABLE
                 If chkesPNT.Value = Checked Then
                    If Not IsNull(rs(4)) Then ' F.REVISION
                       If rs(4) <> "14/01/2004" Then
                            .SubItems(5) = Format(rs(4), "dd/mm/yyyy")
                       End If
                    End If
                 Else
                    .SubItems(6) = ""
                    .SubItems(7) = ""
                 End If
                 .SubItems(6) = rs(5) ' REVISOR
                 If Not IsNull(rs(6)) Then ' F.APROBACION
                    If rs(6) <> "14/01/2004" Then
                        .SubItems(7) = Format(rs(6), "dd/mm/yyyy")
                    End If
                 End If
                 .SubItems(8) = rs(7) ' APROBADOR
                 .SubItems(9) = rs(8) ' ID_RESPONSABLE
                 .SubItems(10) = rs(9) ' ID_REVISOR
                 .SubItems(11) = rs(10) ' ID_APROBADOR
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
End Sub
Private Sub cargar_datos()
    Dim oCA_Documento As New clsCa_documentos
    If oCA_Documento.esPNT(PK) Then
        chkesPNT.Value = Checked
        Frame2.visible = True
    Else
        chkesPNT.Value = Unchecked
        Frame2.visible = False
        lista.ColumnHeaders(COLS.C_F_REVISION).Width = 1
        lista.ColumnHeaders(COLS.C_ID_REVISOR).Width = 1
    End If
    Set oCA_Documento = Nothing
    Dim oPNT As New clsCa_pnt
    With oPNT
     cargar_combos ""
     If .carga(PK, 1) Then
'       cargar_combos " OR ANULADO = 1"
        fecha(1).Value = .getFECHA_CREACION
        fecha(2).Value = .getFECHA_REVISION
        fecha(3).Value = .getFECHA_APROBACION
        cmbusuario(1).MostrarElemento .getUSUARIO_CREACION
        cmbusuario(2).MostrarElemento .getUSUARIO_REVISION
        cmbusuario(3).MostrarElemento .getUSUARIO_APROBACION
        ' Cargar lista
        cargarLista PK
     Else
'        cargar_combos ""
        cmbusuario(1).MostrarElemento USUARIO.getID_EMPLEADO
        fecha(1) = Date
        fecha(2) = Date
        fecha(3) = Date
     End If
    End With
    fecha(4) = Date
    cmbusuario(4).MostrarElemento USUARIO.getID_EMPLEADO
    cmbusuario(5).limpiar
    cmbusuario(6).limpiar
End Sub

Private Sub activar_botones()
    ' Evaluar Estado del documento y activar botones
'    cmdModificado.Enabled = False
   On Error GoTo activar_botones_Error

    cmdterminado.Enabled = False
    cmdrevisado.Enabled = False
    cmdAprobado.Enabled = False
    cmdgenera.Enabled = False
    cmdModificado.Enabled = False
    img1.visible = False
    img2.visible = False
    Dim oCA As New clsCa_documentos
    Dim oPNT As New clsCa_pnt
    With oCA
        If .carga(PK) Then
                oPNT.Carga_Ultima_edicion PK
                Select Case .getESTADO_ID
                Case C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_CREACION
                    cmdgenera.Enabled = True
                Case C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_MODIFICACION
'Inc-741 Solo se activa el botón si el usuario asignado es el que está logado.

                    If oPNT.getMODIFICACION_USUARIO_ = USUARIO.getID_EMPLEADO Or _
                           USUARIO.getPER_ADMIN_PNT = True Then
                           
                            cmdgenera.Enabled = True
                            cmdterminado.Enabled = True
                            cmdModificado.Enabled = True
                            
                    End If
                Case C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_REVISION
                    If CInt(oPNT.getEDICION) <= 1 Then
                        If oPNT.getUSUARIO_REVISION = USUARIO.getID_EMPLEADO Or _
                           USUARIO.getPER_ADMIN_PNT = True Then
                            cmdrevisado.Enabled = True
                            cmdModificado.Enabled = True
                            img1.visible = True
                        End If
                    Else
                        If oPNT.getMODIFICACION_USUARIO_REVISION = USUARIO.getID_EMPLEADO Or _
                           USUARIO.getPER_ADMIN_PNT = True Then
                            cmdrevisado.Enabled = True
                            cmdModificado.Enabled = True
                            img1.visible = True
                        End If
                    
                    End If
                Case C_CA_DOCUMENTOS_ESTADOS.CA_PDTE_APROBACION
                    If CInt(oPNT.getEDICION) <= 1 Then
                        If chkesPNT.Value = Checked Then
                            If oPNT.getUSUARIO_APROBACION = USUARIO.getID_EMPLEADO Or _
                               USUARIO.getPER_ADMIN_PNT = True Then
                                cmdAprobado.Enabled = True
                                cmdModificado.Enabled = True
                                img2.visible = True
                            End If
                        Else
                            If oPNT.getMODIFICACION_USUARIO_REVISION = USUARIO.getID_EMPLEADO Or _
                               USUARIO.getPER_ADMIN_PNT = True Then
                                cmdAprobado.Enabled = True
                                cmdModificado.Enabled = True
                                img2.visible = True
                            End If
                        End If
                    Else
'                        If oPNT.getMODIFICACION_USUARIO_REVISION = USUARIO.getID_EMPLEADO
                        If oPNT.getMODIFICACION_USUARIO_APROBACION = USUARIO.getID_EMPLEADO Or _
                           USUARIO.getPER_ADMIN_PNT = True Then
                            cmdAprobado.Enabled = True
                            cmdModificado.Enabled = True
                            img2.visible = True
                        End If
                    End If
                Case C_CA_DOCUMENTOS_ESTADOS.CA_VIGOR
                    cmdgenera.Enabled = True
                End Select
                ' Si no es primera edicion y no es superusuario, protegemos los usuarios
'                If oPNT.getEDICION > 1 And usuario.getPER_ADMIN_PNT = False Then
                If (oCA.getEDICION > 1 Or (oCA.getEDICION = 1 And oCA.getESTADO_ID = C_CA_DOCUMENTOS_ESTADOS.CA_VIGOR)) And USUARIO.getPER_ADMIN_PNT = False Then
                    Frame2.Enabled = False
                End If
            End If
    End With
    If USUARIO.getPER_ADMIN_PNT Then
        txtDatos(0).Locked = False
    Else
        txtDatos(0).Locked = True
    End If
   On Error GoTo 0
   Exit Sub

activar_botones_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure activar_botones of Formulario frmCA_PNT"
End Sub
Private Sub enviar_mensaje(ASUNTO As String, texto As String, destinatario As Integer)
    ' Enviar aviso
    Dim oMensaje As New clsMensajes
    With oMensaje
        .setASUNTO = ASUNTO
        .setTEXTO = texto
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setFECHA_INICIO = Format(Date, "yyyy-mm-dd")
        .setFECHA_FIN = Format(Date + 15, "yyyy-mm-dd")
        .setACCION = "frmCA_Documento;" & CStr(PK)
        
        .setHORA_INICIO = Format(Time, "hh:mm:ss")
        .setHORA_FIN = Format(Time, "hh:mm:ss")
        .setDURACION = 0
        .setCATEGORIA = MENSAJES_CATEGORIAS.MENSAJES_CATEGORIAS_CALIDAD
        
        mens = .Insertar
        If mens > 0 Then
            Dim omu As New clsMensajes_usuarios
            omu.setEMPLEADO_ID = destinatario
            omu.setMENSAJE_ID = mens
            omu.Insertar
        End If
        frmCalendario.cargar_eventos
    End With
End Sub
Private Sub escribe_traza(texto As String)
    traza = traza & texto & vbNewLine
End Sub

