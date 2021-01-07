VERSION 5.00
Begin VB.Form frmUsuarios 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios y permisos"
   ClientHeight    =   5205
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   Icon            =   "frmUsuario.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5205
   ScaleWidth      =   7050
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   855
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   4230
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   855
      Left            =   5730
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4230
      Width           =   1245
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Personales"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1740
      Left            =   30
      TabIndex        =   1
      Top             =   390
      Width           =   6945
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataField       =   "PASSWORD"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   1440
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   1200
         Width           =   2535
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataField       =   "NOMBRE"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1440
         TabIndex        =   5
         Top             =   810
         Width           =   5160
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataField       =   "USUARIO"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1440
         TabIndex        =   3
         Top             =   405
         Width           =   2190
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   4200
         Picture         =   "frmUsuario.frx":08CA
         Top             =   1200
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   480
         Left            =   6300
         Picture         =   "frmUsuario.frx":0D0C
         Top             =   180
         Width           =   480
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   195
         TabIndex        =   4
         Top             =   855
         Width           =   765
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   195
         TabIndex        =   2
         Top             =   450
         Width           =   735
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Contraseña"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   195
         TabIndex        =   6
         Top             =   1230
         Width           =   1125
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Permisos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   30
      TabIndex        =   8
      Top             =   2220
      Width           =   3615
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Expedientes"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   6
         Left            =   495
         TabIndex        =   15
         Top             =   2475
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recalculo"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   5
         Left            =   495
         TabIndex        =   14
         Top             =   2115
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Altas / Bajas usuarios"
         DataField       =   "PER_ELIMINACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   4
         Left            =   495
         TabIndex        =   13
         Top             =   1755
         Width           =   2760
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Eliminacion"
         DataField       =   "PER_MODIFICACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   3
         Left            =   495
         TabIndex        =   12
         Top             =   1395
         Width           =   1635
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modificacion"
         DataField       =   "PER_MODIFICACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   2
         Left            =   495
         TabIndex        =   11
         Top             =   1050
         Width           =   1725
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Alta"
         DataField       =   "PER_FACTURACION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   1
         Left            =   495
         TabIndex        =   10
         Top             =   675
         Width           =   1500
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Impresion"
         DataField       =   "PER_IMPRESION"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Index           =   0
         Left            =   495
         TabIndex        =   9
         Top             =   315
         Width           =   1365
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   2700
         Picture         =   "frmUsuario.frx":15D6
         Top             =   495
         Width           =   480
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Mantenimiento de Usuarios"
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
      Height          =   300
      Index           =   0
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6960
   End
End
Attribute VB_Name = "frmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    Dim ousu As New ClsUsuario
    With ousu
            .setNOMBRE = datos(1)
            .setusuario = datos(0)
            .setPASSWORD = datos(3)
            ' Per. Impresion
            If Check1(0).Value = Checked Then
                .setPER_1 = 1
            Else
                .setPER_1 = 0
            End If
            ' Per. Facturacion
            If Check1(1).Value = Checked Then
                .setPER_2 = 1
            Else
                .setPER_2 = 0
            End If
            ' Per. Modificacion
            If Check1(2).Value = Checked Then
                .setPER_3 = 1
            Else
                .setPER_3 = 0
            End If
            ' Per. Eliminacion
            If Check1(3).Value = Checked Then
                .setPER_4 = 1
            Else
                .setPER_4 = 0
            End If
            ' Per. Usuario
            If Check1(4).Value = Checked Then
                .setPER_5 = 1
            Else
                .setPER_5 = 0
            End If
            ' Per. Recalculo
            If Check1(5).Value = Checked Then
                .setPER_6 = 1
            Else
                .setPER_6 = 0
            End If
            ' Per. Expedientes
            If Check1(6).Value = Checked Then
                .setPER_7 = 1
            Else
                .setPER_7 = 0
            End If
            If gusuario = 0 Then ' Nuevo
                If .Insertar <> 0 Then
                    MsgBox "El usuario se ha insertado correctamente", vbInformation, App.Title
                    Unload Me
                End If
            Else
                If .Modificar(CLng(gusuario)) <> 0 Then
                    MsgBox "El usuario se ha modificado correctamente", vbInformation, App.Title
                    Unload Me
                End If
            End If
    End With
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    If gusuario <> 0 Then
        cargar_usuario
    End If
End Sub

Public Sub cargar_usuario()
    Dim clsusu As New ClsUsuario
    If clsusu.Cargar(gusuario) = True Then
     With clsusu
        datos(0) = .getUSUARIO
        datos(1) = .getNOMBRE
        datos(3) = .getPASSWORD
        If .getPER_1 = 1 Then
            Check1(0).Value = Checked
        End If
        If .getPER_2 = 1 Then
            Check1(1).Value = Checked
        End If
        If .getPER_3 = 1 Then
            Check1(2).Value = Checked
        End If
        If .getPER_4 = 1 Then
            Check1(3).Value = Checked
        End If
        If .getPER_5 = 1 Then
            Check1(4).Value = Checked
        End If
        If .getPER_6 = 1 Then
            Check1(5).Value = Checked
        End If
        If .getPER_7 = 1 Then
            Check1(6).Value = Checked
        End If
        Label1(0).Caption = "Modificacion del usuario : " & .getUSUARIO
        Label1(0).BackColor = &HC0FFFF
     End With
    End If
End Sub
