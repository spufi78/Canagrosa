VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Acerca de MiApl"
   ClientHeight    =   3240
   ClientLeft      =   2295
   ClientTop       =   1605
   ClientWidth     =   5655
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2236.306
   ScaleMode       =   0  'User
   ScaleWidth      =   5310.337
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   180
      Top             =   2340
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "5"
      Height          =   240
      Left            =   4680
      TabIndex        =   5
      Top             =   2520
      Width           =   330
   End
   Begin VB.Label lbltiempo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "seg."
      Height          =   240
      Left            =   4950
      TabIndex        =   4
      Top             =   2520
      Width           =   465
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Julio González Moreno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   225
      Left            =   2940
      TabIndex        =   3
      Top             =   150
      Width           =   2310
   End
   Begin VB.Image Image1 
      Height          =   1320
      Left            =   90
      Picture         =   "frmAbout.frx":030A
      Top             =   45
      Width           =   2385
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackColor       =   &H00E0E0E0&
      Caption         =   "mailto : julio.gonzalez@ixitec.net"
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
      Height          =   225
      Left            =   45
      TabIndex        =   0
      Top             =   2880
      Width           =   5550
   End
   Begin VB.Label lblTitle 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Título de la aplicación"
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
      Height          =   255
      Left            =   90
      TabIndex        =   1
      Top             =   1440
      Width           =   5505
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Aplicación para la gestión de laboratorios."
      Height          =   195
      Left            =   225
      TabIndex        =   2
      Top             =   1710
      Width           =   5370
      WordWrap        =   -1  'True
   End
   Begin VB.Shape Shape1 
      Height          =   3195
      Left            =   0
      Top             =   0
      Width           =   5655
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Load()
    Me.Caption = "Acerca de " & App.Title
    lblTitle.Caption = App.Title
End Sub

Private Sub Timer1_Timer()
    Label2.Caption = CInt(Label2.Caption) - 1
    If CInt(Label2.Caption) = 0 Then
        Unload Me
    End If
End Sub
