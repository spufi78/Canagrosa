VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCambios 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "Listado de ultimas modificaciones realizadas."
   ClientHeight    =   6705
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   DrawStyle       =   5  'Transparent
   Icon            =   "frmCambios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6705
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   Begin Geslab.ControlPanelXP ControlPanelXP1 
      Height          =   6585
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7000
      _ExtentX        =   12356
      _ExtentY        =   11615
      Caption         =   "Listado de últimas modificaciones realizadas."
      BackColor       =   12632256
      TextColor       =   0
      HeaderColor     =   8421504
      PanelOpen       =   0   'False
      Object.Height          =   6585
      Begin RichTextLib.RichTextBox texto 
         Height          =   6135
         Left            =   45
         TabIndex        =   1
         Top             =   405
         Width           =   6900
         _ExtentX        =   12171
         _ExtentY        =   10821
         _Version        =   393217
         BackColor       =   16777215
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmCambios.frx":0442
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = Screen.Width - Me.Width - frmMenu.ButtonBar.Width - 500
    Me.Top = 600
    On Error Resume Next
    texto.LoadFile ReadINI(App.Path + "\config.ini", "Documentos", "Cambios"), 0
End Sub
