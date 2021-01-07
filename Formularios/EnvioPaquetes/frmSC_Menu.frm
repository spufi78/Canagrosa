VERSION 5.00
Begin VB.Form frmSC_Menu 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   Icon            =   "frmSC_Menu.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6780
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerico 
      BackColor       =   &H00C0C0C0&
      Caption         =   "General"
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
      Left            =   4545
      Picture         =   "frmSC_Menu.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   405
      Width           =   2175
   End
   Begin VB.CommandButton cmdCE 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ensayos de Eficacia"
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
      Left            =   2295
      Picture         =   "frmSC_Menu.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   405
      Width           =   2175
   End
   Begin VB.CommandButton cmdEnsayos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Determinaciones"
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
      Left            =   45
      Picture         =   "frmSC_Menu.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   405
      Width           =   2175
   End
   Begin VB.Label lblsubtitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione el tipo de Subcontratación"
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
      Left            =   -45
      TabIndex        =   3
      Top             =   0
      Width           =   6810
   End
End
Attribute VB_Name = "frmSC_Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdEnsayos_Click()
    frmSC_Muestras_NoEnviadas_listado.Show 1
'JGM    Me.Visible = False
    Unload Me
End Sub

Private Sub cmdCE_Click()
    frmSC_Muestras_NoEnviadas_CE_Listado.Show 1
'JGM    Me.Visible = False
    Unload Me
End Sub

Private Sub cmdGenerico_Click()
'M1163-I
'    frmSC_Generico_NoEnviadas_Listado.Show 1
    frmSC_Paquete_Detalle_Generico.PK = 0
    frmSC_Paquete_Detalle_Generico.Show 1
'M1163-F
'JGM    Me.Visible = False
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    permisos
End Sub

Private Sub permisos()
    cmdEnsayos.Enabled = True
    cmdCE.Enabled = True
    If USUARIO.getPER_SCG = True Then
       cmdGenerico.Enabled = True
    Else
       cmdGenerico.Enabled = False
    End If
End Sub
