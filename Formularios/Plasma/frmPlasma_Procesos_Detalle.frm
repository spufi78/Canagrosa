VERSION 5.00
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmPlasma_Procesos_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Detalle de Ficha de Plasma"
   ClientHeight    =   10365
   ClientLeft      =   45
   ClientTop       =   300
   ClientWidth     =   11070
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmPlasma_Procesos_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   11070
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmDureza 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "ENSAYO DE DUREZA"
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
      Height          =   2175
      Left            =   45
      TabIndex        =   56
      Top             =   2205
      Width           =   10950
      Begin VB.TextBox txtDurezaEsp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   330
         Left            =   1170
         TabIndex        =   6
         Top             =   1620
         Width           =   8520
      End
      Begin VB.TextBox txtDurezaReq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         Height          =   330
         Left            =   1170
         Locked          =   -1  'True
         MaxLength       =   255
         TabIndex        =   5
         Top             =   1170
         Width           =   8520
      End
      Begin pryCombo.miCombo cmbFichaDureza 
         Height          =   375
         Left            =   1170
         TabIndex        =   3
         Top             =   270
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbDureza 
         Height          =   375
         Left            =   1170
         TabIndex        =   4
         Top             =   720
         Width           =   9600
         _ExtentX        =   16933
         _ExtentY        =   661
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Specification"
         Height          =   195
         Index           =   19
         Left            =   135
         TabIndex        =   63
         Top             =   1695
         Width           =   915
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Requirement"
         Height          =   195
         Index           =   17
         Left            =   135
         TabIndex        =   59
         Top             =   1215
         Width           =   900
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo"
         Height          =   195
         Index           =   16
         Left            =   135
         TabIndex        =   58
         Top             =   765
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ficha"
         Height          =   195
         Index           =   15
         Left            =   135
         TabIndex        =   57
         Top             =   315
         Width           =   390
      End
   End
   Begin VB.Frame frmPlasma 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   7260
      Left            =   0
      TabIndex        =   39
      Top             =   2160
      Width           =   11085
      Begin VB.Frame frmBond 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Bond Coat"
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
         Height          =   3525
         Left            =   45
         TabIndex        =   48
         Top             =   45
         Width           =   10950
         Begin VB.TextBox txtBondSpecification 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1125
            TabIndex        =   8
            Top             =   675
            Width           =   8625
         End
         Begin VB.Frame Frame2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ensayos"
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
            Height          =   2355
            Left            =   135
            TabIndex        =   49
            Top             =   1035
            Width           =   10260
            Begin VB.CheckBox chkMicroestructura 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Check1"
               Height          =   240
               Index           =   0
               Left            =   180
               TabIndex        =   9
               Top             =   315
               Width           =   240
            End
            Begin VB.CheckBox chkTraccion 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Check1"
               Height          =   240
               Index           =   0
               Left            =   180
               TabIndex        =   11
               Top             =   675
               Width           =   240
            End
            Begin VB.CheckBox chkMacroDureza 
               BackColor       =   &H00C0C0C0&
               Caption         =   "chkMacroDureza"
               Height          =   240
               Index           =   0
               Left            =   180
               TabIndex        =   13
               Top             =   1080
               Width           =   240
            End
            Begin VB.CheckBox chkMicroDureza 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Check1"
               Height          =   240
               Index           =   0
               Left            =   180
               TabIndex        =   15
               Top             =   1485
               Width           =   240
            End
            Begin VB.CheckBox chkEspesor 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Check1"
               Height          =   240
               Index           =   0
               Left            =   180
               TabIndex        =   17
               Top             =   1890
               Width           =   240
            End
            Begin pryCombo.miCombo cmbMicroestructura 
               Height          =   375
               Index           =   0
               Left            =   1665
               TabIndex        =   10
               Top             =   270
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   661
            End
            Begin pryCombo.miCombo cmbTraccion 
               Height          =   375
               Index           =   0
               Left            =   1665
               TabIndex        =   12
               Top             =   675
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   661
            End
            Begin pryCombo.miCombo cmbMacro 
               Height          =   375
               Index           =   0
               Left            =   1665
               TabIndex        =   14
               Top             =   1080
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   661
            End
            Begin pryCombo.miCombo cmbmicro 
               Height          =   375
               Index           =   0
               Left            =   1665
               TabIndex        =   16
               Top             =   1485
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   661
            End
            Begin pryCombo.miCombo cmbEspesor 
               Height          =   375
               Index           =   0
               Left            =   1665
               TabIndex        =   18
               Top             =   1890
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   661
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Micro Dureza"
               Height          =   195
               Index           =   6
               Left            =   495
               TabIndex        =   54
               Top             =   1530
               Width           =   945
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Macro Dureza"
               Height          =   195
               Index           =   4
               Left            =   495
               TabIndex        =   53
               Top             =   1125
               Width           =   1005
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Tracción"
               Height          =   195
               Index           =   1
               Left            =   495
               TabIndex        =   52
               Top             =   720
               Width           =   630
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Microestructura"
               Height          =   195
               Index           =   0
               Left            =   495
               TabIndex        =   51
               Top             =   315
               Width           =   1095
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Espesor"
               Height          =   195
               Index           =   12
               Left            =   495
               TabIndex        =   50
               Top             =   1935
               Width           =   570
            End
         End
         Begin pryCombo.miCombo cmbBond 
            Height          =   375
            Left            =   1125
            TabIndex        =   7
            Top             =   315
            Width           =   9690
            _ExtentX        =   17092
            _ExtentY        =   661
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Specification"
            Height          =   195
            Index           =   14
            Left            =   135
            TabIndex        =   61
            Top             =   750
            Width           =   915
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ficha"
            Height          =   195
            Index           =   3
            Left            =   135
            TabIndex        =   55
            Top             =   360
            Width           =   390
         End
      End
      Begin VB.Frame frmTop 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Top Coat"
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
         Height          =   3525
         Left            =   45
         TabIndex        =   40
         Top             =   3645
         Width           =   10950
         Begin VB.TextBox txtTopSpecification 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   330
            Left            =   1125
            TabIndex        =   20
            Top             =   675
            Width           =   8625
         End
         Begin VB.Frame Frame5 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ensayos"
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
            Height          =   2355
            Left            =   135
            TabIndex        =   41
            Top             =   1035
            Width           =   10260
            Begin VB.CheckBox chkMicroestructura 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Check1"
               Height          =   240
               Index           =   1
               Left            =   180
               TabIndex        =   21
               Top             =   315
               Width           =   240
            End
            Begin VB.CheckBox chkTraccion 
               BackColor       =   &H00C0C0C0&
               Caption         =   "chkTraccion"
               Height          =   240
               Index           =   1
               Left            =   180
               TabIndex        =   23
               Top             =   675
               Width           =   240
            End
            Begin VB.CheckBox chkMacroDureza 
               BackColor       =   &H00C0C0C0&
               Caption         =   "chkMacroDureza"
               Height          =   240
               Index           =   1
               Left            =   180
               TabIndex        =   25
               Top             =   1080
               Width           =   240
            End
            Begin VB.CheckBox chkMicroDureza 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Check1"
               Height          =   240
               Index           =   1
               Left            =   180
               TabIndex        =   27
               Top             =   1485
               Width           =   240
            End
            Begin VB.CheckBox chkEspesor 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Check1"
               Height          =   240
               Index           =   1
               Left            =   180
               TabIndex        =   29
               Top             =   1890
               Width           =   240
            End
            Begin pryCombo.miCombo cmbMicroestructura 
               Height          =   375
               Index           =   1
               Left            =   1665
               TabIndex        =   22
               Top             =   270
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   661
            End
            Begin pryCombo.miCombo cmbTraccion 
               Height          =   375
               Index           =   1
               Left            =   1665
               TabIndex        =   24
               Top             =   675
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   661
            End
            Begin pryCombo.miCombo cmbMacro 
               Height          =   375
               Index           =   1
               Left            =   1665
               TabIndex        =   26
               Top             =   1080
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   661
            End
            Begin pryCombo.miCombo cmbmicro 
               Height          =   375
               Index           =   1
               Left            =   1665
               TabIndex        =   28
               Top             =   1485
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   661
            End
            Begin pryCombo.miCombo cmbEspesor 
               Height          =   375
               Index           =   1
               Left            =   1665
               TabIndex        =   30
               Top             =   1890
               Width           =   8160
               _ExtentX        =   14393
               _ExtentY        =   661
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Microestructura"
               Height          =   195
               Index           =   5
               Left            =   495
               TabIndex        =   46
               Top             =   315
               Width           =   1095
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Tracción"
               Height          =   195
               Index           =   7
               Left            =   495
               TabIndex        =   45
               Top             =   720
               Width           =   630
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Macro Dureza"
               Height          =   195
               Index           =   8
               Left            =   495
               TabIndex        =   44
               Top             =   1125
               Width           =   1005
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Micro Dureza"
               Height          =   195
               Index           =   9
               Left            =   495
               TabIndex        =   43
               Top             =   1530
               Width           =   945
            End
            Begin VB.Label lblCampos 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Espesor"
               Height          =   195
               Index           =   13
               Left            =   495
               TabIndex        =   42
               Top             =   1935
               Width           =   570
            End
         End
         Begin pryCombo.miCombo cmbTop 
            Height          =   375
            Left            =   1125
            TabIndex        =   19
            Top             =   315
            Width           =   9690
            _ExtentX        =   17092
            _ExtentY        =   661
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Specification"
            Height          =   195
            Index           =   18
            Left            =   135
            TabIndex        =   62
            Top             =   750
            Width           =   915
         End
         Begin VB.Label lblCampos 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0C0C0&
            Caption         =   "Ficha"
            Height          =   195
            Index           =   10
            Left            =   135
            TabIndex        =   47
            Top             =   360
            Width           =   390
         End
      End
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   7515
      Style           =   1  'Graphical
      TabIndex        =   31
      Top             =   9450
      Width           =   1365
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9990
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   9450
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   8910
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   9450
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle"
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
      Height          =   1455
      Left            =   45
      TabIndex        =   34
      Top             =   675
      Width           =   10935
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
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
         Index           =   0
         Left            =   945
         TabIndex        =   0
         Top             =   225
         Width           =   8760
      End
      Begin pryCombo.miCombo cmbFabricante 
         Height          =   375
         Left            =   945
         TabIndex        =   1
         Top             =   630
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   661
      End
      Begin pryCombo.miCombo cmbTest 
         Height          =   345
         Left            =   945
         TabIndex        =   2
         Top             =   1035
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   609
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Test"
         Height          =   195
         Index           =   20
         Left            =   90
         TabIndex        =   60
         Top             =   1065
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fabricante"
         Height          =   195
         Index           =   11
         Left            =   90
         TabIndex        =   38
         Top             =   675
         Width           =   750
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proceso"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   35
         Top             =   300
         Width           =   585
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Ficha de Plasma"
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
      TabIndex        =   37
      Top             =   45
      Width           =   2895
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de Ficha de Plasma"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   36
      Top             =   330
      Width           =   1935
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   11100
   End
End
Attribute VB_Name = "frmPlasma_Procesos_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Private Sub chkEspesor_Click(Index As Integer)
    If chkEspesor(Index).Value = Checked Then
        cmbEspesor(Index).activar
    Else
        cmbEspesor(Index).desactivar
    End If
End Sub

Private Sub chkMacroDureza_Click(Index As Integer)
    If chkMacroDureza(Index).Value = Checked Then
        cmbMacro(Index).activar
    Else
        cmbMacro(Index).desactivar
    End If
End Sub
Private Sub chkMicroDureza_Click(Index As Integer)
    If chkMicroDureza(Index).Value = Checked Then
        cmbmicro(Index).activar
    Else
        cmbmicro(Index).desactivar
    End If
End Sub

Private Sub chkMicroestructura_Click(Index As Integer)
    If chkMicroestructura(Index).Value = Checked Then
        cmbMicroestructura(Index).activar
    Else
        cmbMicroestructura(Index).desactivar
    End If
End Sub

Private Sub chkTraccion_Click(Index As Integer)
    If chkTraccion(Index).Value = Checked Then
        cmbTraccion(Index).activar
    Else
        cmbTraccion(Index).desactivar
    End If
End Sub

Private Sub cmbBond_change()
    If cmbBond.getTEXTO <> "" Then
        Dim oPF As New clsPlasma_ficha
        With oPF
            .Carga cmbBond.getPK_SALIDA
            cmbMicroestructura(0).MostrarElemento .getMICROESTRUCTURA
            cmbTraccion(0).MostrarElemento .getTRACCION
            cmbMacro(0).MostrarElemento .getMACRO_DUREZA
            cmbmicro(0).MostrarElemento .getMICRO_DUREZA
            ' MICROESTRUCTURA
            If .getMICROESTRUCTURA = 0 Then
                chkMicroestructura(0).Enabled = False
                chkMicroestructura(0).Value = Unchecked
                cmbMicroestructura(0).desactivar
            Else
                chkMicroestructura(0).Enabled = True
                chkMicroestructura(0).Value = Checked
                cmbMicroestructura(0).activar
            End If
            ' TRACCION
            If .getTRACCION = 0 Then
                chkTraccion(0).Enabled = False
                chkTraccion(0).Value = Unchecked
                cmbTraccion(0).desactivar
            Else
                chkTraccion(0).Enabled = True
                chkTraccion(0).Value = Checked
                cmbTraccion(0).activar
            End If
            ' MACRO_DUREZA
            If .getMACRO_DUREZA = 0 Then
                chkMacroDureza(0).Enabled = False
                chkMacroDureza(0).Value = Unchecked
                cmbMacro(0).desactivar
            Else
                chkMacroDureza(0).Enabled = True
                chkMacroDureza(0).Value = Checked
                cmbMacro(0).activar
            End If
            ' MICRO_DUREZA
            If .getMICRO_DUREZA = 0 Then
                chkMicroDureza(0).Enabled = False
                chkMicroDureza(0).Value = Unchecked
                cmbmicro(0).desactivar
            Else
                chkMicroDureza(0).Enabled = True
                chkMicroDureza(0).Value = Checked
                cmbmicro(0).activar
            End If
            ' ESPESOR
            If .getESPESOR = 0 Then
                chkEspesor(0).Enabled = False
                chkEspesor(0).Value = Unchecked
                cmbEspesor(0).desactivar
            Else
                chkEspesor(0).Enabled = True
                chkEspesor(0).Value = Checked
                cmbEspesor(0).activar
            End If
        End With
        Set oPF = Nothing
    Else
        cmbMicroestructura(0).limpiar
        cmbTraccion(0).limpiar
        cmbMacro(0).limpiar
        cmbmicro(0).limpiar
        cmbEspesor(0).limpiar
    End If
End Sub

Private Sub cmbFabricante_change()
    cmbBond.limpiar
    cmbTop.limpiar
    Dim i As Integer
    For i = 0 To 1
        cmbMicroestructura(i).limpiar
        cmbTraccion(i).limpiar
        cmbMacro(i).limpiar
        cmbmicro(i).limpiar
        chkMicroestructura(i).Value = Unchecked
        chkTraccion(i).Value = Unchecked
        chkMacroDureza(i).Value = Unchecked
        chkMicroDureza(i).Value = Unchecked
        chkEspesor(i).Value = Unchecked
        cmbMicroestructura(i).desactivar
        cmbTraccion(i).desactivar
        cmbMacro(i).desactivar
        cmbmicro(i).desactivar
        cmbEspesor(i).desactivar
    Next
    If cmbFabricante.getTEXTO = "" Then
        frmTop.Enabled = False
        frmBond.Enabled = False
    Else
        frmTop.Enabled = True
        frmBond.Enabled = True
        llenar_combo cmbBond, New clsPlasma_ficha, 0, frmPlasma_Ficha_Detalle, " FABRICANTE_ID = " & cmbFabricante.getPK_SALIDA
        llenar_combo cmbTop, New clsPlasma_ficha, 0, frmPlasma_Ficha_Detalle, " FABRICANTE_ID = " & cmbFabricante.getPK_SALIDA
        
    End If
End Sub

Private Sub cmbFichaDureza_change()
    Dim oPF As New clsPlasma_ficha
    If cmbFichaDureza.getTEXTO = "" Then
        cmbDureza.limpiar
        txtDurezaReq = ""
    Else
        With oPF
            .Carga cmbFichaDureza.getPK_SALIDA
            If cmbTest.getPK_SALIDA = PLASMA_TIPOS.PT_DUREZA_ROCKWELL Or cmbTest.getPK_SALIDA = PLASMA_TIPOS.PT_DUREZA_BRINELL Then
                cmbDureza.MostrarElemento .getMACRO_DUREZA
                txtDurezaReq = .getMACRO_DUREZA_REQ
            End If
            If cmbTest.getPK_SALIDA = PLASMA_TIPOS.PT_DUREZA_VICKERS Then
                cmbDureza.MostrarElemento .getMICRO_DUREZA
                txtDurezaReq = .getMICRO_DUREZA_REQ
            End If
        End With
    End If
    Set oPF = Nothing
End Sub

Private Sub cmbTest_change()
    Select Case cmbTest.getPK_SALIDA
        Case PLASMA_TIPOS.PT_PLASMA
            frmPlasma.visible = True
            frmDureza.visible = False
        Case Else
            frmPlasma.visible = False
            frmDureza.visible = True
    End Select
End Sub

Private Sub cmbTop_change()
    If cmbTop.getTEXTO <> "" Then
        Dim oPF As New clsPlasma_ficha
        With oPF
            .Carga cmbTop.getPK_SALIDA
            cmbMicroestructura(1).MostrarElemento .getMICROESTRUCTURA
            cmbTraccion(1).MostrarElemento .getTRACCION
            cmbMacro(1).MostrarElemento .getMACRO_DUREZA
            cmbmicro(1).MostrarElemento .getMICRO_DUREZA
            ' MICROESTRUCTURA
            If .getMICROESTRUCTURA = 0 Then
                chkMicroestructura(1).Enabled = False
                chkMicroestructura(1).Value = Unchecked
                cmbMicroestructura(1).desactivar
            Else
                chkMicroestructura(1).Enabled = True
                chkMicroestructura(1).Value = Checked
                cmbMicroestructura(1).activar
            End If
            ' TRACCION
            If .getTRACCION = 0 Then
                chkTraccion(1).Enabled = False
                chkTraccion(1).Value = Unchecked
                cmbTraccion(1).desactivar
            Else
                chkTraccion(1).Enabled = True
                chkTraccion(1).Value = Checked
                cmbTraccion(1).activar
            End If
            ' MACRO_DUREZA
            If .getMACRO_DUREZA = 0 Then
                chkMacroDureza(1).Enabled = False
                chkMacroDureza(1).Value = Unchecked
                cmbMacro(1).desactivar
            Else
                chkMacroDureza(1).Enabled = True
                chkMacroDureza(1).Value = Checked
                cmbMacro(1).activar
            End If
            ' MICRO_DUREZA
            If .getMICRO_DUREZA = 0 Then
                chkMicroDureza(1).Enabled = False
                chkMicroDureza(1).Value = Unchecked
                cmbmicro(1).desactivar
            Else
                chkMicroDureza(1).Enabled = True
                chkMicroDureza(1).Value = Checked
                cmbmicro(1).activar
            End If
            ' ESPESOR
            If .getESPESOR = 0 Then
                chkEspesor(1).Enabled = False
                chkEspesor(1).Value = Unchecked
                cmbEspesor(1).desactivar
            Else
                chkEspesor(1).Enabled = True
                chkEspesor(1).Value = Checked
                cmbEspesor(1).activar
            End If
        End With
        Set oPF = Nothing
    Else
        cmbMicroestructura(1).limpiar
        cmbTraccion(1).limpiar
        cmbMacro(1).limpiar
        cmbmicro(1).limpiar
        cmbEspesor(1).limpiar
    End If

End Sub

Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_PLASMA_ENSAYOS
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Tipo Ensayo Plasma " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    Dim i As Integer
   On Error GoTo cmdok_Click_Error

    If validar = True Then
        Dim oPP As New clsPlasma_procesos
        Dim PROCESO As Long
        With oPP
            .setNOMBRE = txtDatos(0)
            .setFABRICANTE_ID = cmbFabricante.getPK_SALIDA
            .setTIPO = cmbTest.getPK_SALIDA
            If cmbTest.getPK_SALIDA = PLASMA_TIPOS.PT_PLASMA Then
                .setBOND_COAT_FICHA_ID = cmbBond.getPK_SALIDA
                .setBOND_SPECIFICATION = txtBondSpecification.Text
                .setBOND_MICROESTRUCTURA = chkMicroestructura(0).Value
                .setBOND_TRACCION = chkTraccion(0).Value
                .setBOND_MACRO_DUREZA = chkMacroDureza(0).Value
                .setBOND_MICRO_DUREZA = chkMicroDureza(0).Value
                .setBOND_ESPESOR = chkEspesor(0).Value
                .setTOP_COAT_FICHA_ID = cmbTop.getPK_SALIDA
                .setTOP_SPECIFICATION = txtTopSpecification.Text
                .setTOP_MICROESTRUCTURA = chkMicroestructura(1).Value
                .setTOP_TRACCION = chkTraccion(1).Value
                .setTOP_MACRO_DUREZA = chkMacroDureza(1).Value
                .setTOP_MICRO_DUREZA = chkMicroDureza(1).Value
                .setTOP_ESPESOR = chkEspesor(1).Value
            Else
                .setBOND_COAT_FICHA_ID = cmbFichaDureza.getPK_SALIDA
                .setBOND_SPECIFICATION = txtDurezaEsp.Text
                .setBOND_MICROESTRUCTURA = 0
                .setBOND_TRACCION = 0
                If cmbTest.getPK_SALIDA = PLASMA_TIPOS.PT_DUREZA_ROCKWELL Or cmbTest.getPK_SALIDA = PLASMA_TIPOS.PT_DUREZA_BRINELL Then
                    .setBOND_MACRO_DUREZA = 1
                    .setBOND_MICRO_DUREZA = 0
                Else
                    .setBOND_MACRO_DUREZA = 0
                    .setBOND_MICRO_DUREZA = 1
                End If
                .setBOND_ESPESOR = 0
                .setTOP_COAT_FICHA_ID = cmbFichaDureza.getPK_SALIDA
                .setTOP_SPECIFICATION = txtDurezaEsp.Text
                .setTOP_MICROESTRUCTURA = 0
                .setTOP_TRACCION = 0
                .setTOP_MACRO_DUREZA = 0
                .setTOP_MICRO_DUREZA = 0
                .setTOP_ESPESOR = 0
                If cmbTest.getPK_SALIDA = PLASMA_TIPOS.PT_DUREZA_SHORE_A Then
                    .setBOND_ESPESOR = 1
                    .setTOP_ESPESOR = 1
                    .setBOND_MACRO_DUREZA = 0
                    .setBOND_MICRO_DUREZA = 0
                End If
            End If
        End With
        Dim ohc As New clsHistorial_cambios
        If PK = 0 Then
          If MsgBox("Va a introducir un nuevo proceso. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
              PROCESO = oPP.Insertar
              If PROCESO > 0 Then
                  With ohc
                      .setTIPO = HC_TIPOS.HC_PLASMA_PROCESOS
                      .setIDENTIFICADOR = PROCESO
                      .setIDENTIFICADOR_TEXTO = txtDatos(0)
                      .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                      .setMOTIVO = HC_CREACION
                      .Insertar
                  End With
              End If
          Else
              Exit Sub
          End If
      Else
        If MsgBox("Va a modificar el proceso. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación del proceso."
            frmMotivo.Show 1
            If Trim(MOTIVO) = "" Then
                MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                Exit Sub
            End If
            oPP.Modificar (PK)
            PROCESO = PK
            With ohc
                .setTIPO = HC_TIPOS.HC_PLASMA_PROCESOS
                .setIDENTIFICADOR = PK
                .setIDENTIFICADOR_TEXTO = txtDatos(0)
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .setMOTIVO = Trim(MOTIVO)
                .Insertar
            End With
        Else
            Exit Sub
        End If
      End If
      Set ohc = Nothing
      Me.MousePointer = 0
      If PK = 0 Then
          MsgBox "El proceso se ha introducido correctamente.", vbOKOnly + vbInformation, App.Title
          Unload Me
      Else
          MsgBox "El proceso se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
          Unload Me
      End If
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:
      Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmPlasma_Procesos_Detalle"
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combo
    If PK <> 0 Then
        lbltitulo = "Modificación de Ficha de Plasma"
        cargar_ficha
    Else
        lbltitulo = "Alta de Ficha de Plasma"
        cmbTest.MostrarElemento PLASMA_TIPOS.PT_PLASMA
    End If
End Sub
Private Sub cargar_combo()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_mi_combo cmbFabricante, DECODIFICADORA.DECODIFICADORA_PLASMA_FABRICANTES
    oDeco.cargar_mi_combo cmbTest, DECODIFICADORA.IBERIA_ENSAYOS_FISICOS
    Set oDeco = Nothing
    ' Bond
    llenar_combo cmbBond, New clsPlasma_ficha, 0, frmPlasma_Ficha_Detalle, ""
    llenar_combo cmbMicroestructura(0), New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 1 "
    llenar_combo cmbTraccion(0), New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 2 "
    llenar_combo cmbMacro(0), New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 3 "
    llenar_combo cmbmicro(0), New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 4 "
    llenar_combo cmbEspesor(0), New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 5 "
    ' Top
    llenar_combo cmbTop, New clsPlasma_ficha, 0, frmPlasma_Ficha_Detalle, ""
    llenar_combo cmbMicroestructura(1), New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 1 "
    llenar_combo cmbTraccion(1), New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 2 "
    llenar_combo cmbMacro(1), New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 3 "
    llenar_combo cmbmicro(1), New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 4 "
    llenar_combo cmbEspesor(1), New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 5 "
    ' Dureza
    llenar_combo cmbFichaDureza, New clsPlasma_ficha, 0, frmPlasma_Ficha_Detalle, ""
    llenar_combo cmbDureza, New clsPlasma_ensayos, 0, frmPlasma_Ensayos_Detalle, " TIPO_ID = 3 "
    cmbDureza.desactivar
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Private Sub cargar_ficha()
    Dim i As Integer
    Dim oPP As New clsPlasma_procesos
    If oPP.Carga(PK) = True Then
        With oPP
            txtDatos(0) = .getNOMBRE
            cmbTest.MostrarElemento .getTIPO
            cmbFabricante.MostrarElemento .getFABRICANTE_ID
            If .getTIPO = PLASMA_TIPOS.PT_PLASMA Then
                cmbBond.MostrarElemento .getBOND_COAT_FICHA_ID
                cmbTop.MostrarElemento .getTOP_COAT_FICHA_ID
                txtBondSpecification = .getBOND_SPECIFICATION
                txtTopSpecification = .getTOP_SPECIFICATION
                ' MIRAR CHECKS
                chkMicroestructura(0).Value = .getBOND_MICROESTRUCTURA
                chkTraccion(0).Value = .getBOND_TRACCION
                chkMacroDureza(0).Value = .getBOND_MACRO_DUREZA
                chkMicroDureza(0).Value = .getBOND_MICRO_DUREZA
                chkEspesor(0).Value = .getBOND_ESPESOR
                
                chkMicroestructura(1).Value = .getTOP_MICROESTRUCTURA
                chkTraccion(1).Value = .getTOP_TRACCION
                chkMacroDureza(1).Value = .getTOP_MACRO_DUREZA
                chkMicroDureza(1).Value = .getTOP_MICRO_DUREZA
                chkEspesor(1).Value = .getTOP_ESPESOR
            Else
                cmbFichaDureza.MostrarElemento .getBOND_COAT_FICHA_ID
                txtDurezaEsp = .getBOND_SPECIFICATION
            End If
        End With
    End If
    Set oPP = Nothing
End Sub
Private Function validar() As Boolean
    validar = True
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe indicar la descripción del Proceso.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If cmbTest.getTEXTO = "" Then
        MsgBox "Debe indicar el tipo de TEST del Proceso.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
    If cmbTest.getPK_SALIDA = PLASMA_TIPOS.PT_PLASMA Then
        If cmbBond.getTEXTO = "" Then
            MsgBox "Debe indicar la ficha del Bond Coat.", vbInformation, App.Title
            validar = False
            Exit Function
        End If
        If cmbTop.getTEXTO = "" Then
            MsgBox "Debe indicar la ficha del Top Coat.", vbInformation, App.Title
            validar = False
            Exit Function
        End If
    Else
        If cmbFichaDureza.getTEXTO = "" Then
            MsgBox "Debe indicar la ficha de la Dureza.", vbInformation, App.Title
            validar = False
            Exit Function
        End If
'        If cmbDureza.getTEXTO = "" Then
'            MsgBox "Debe indicar el ensayo de Dureza.", vbInformation, App.Title
'            validar = False
'            Exit Function
'        End If
        
    End If
End Function
