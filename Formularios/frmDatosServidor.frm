VERSION 5.00
Begin VB.Form frmMantDatosServidor 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos Servidor"
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4050
   Icon            =   "frmDatosServidor.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2490
   ScaleWidth      =   4050
   Begin VB.CommandButton Command1 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   0
      Left            =   2475
      Picture         =   "frmDatosServidor.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1800
      Width           =   1365
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Index           =   1
      Left            =   1035
      Picture         =   "frmDatosServidor.frx":0BDC
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1800
      Width           =   1365
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      Height          =   330
      Left            =   2745
      TabIndex        =   3
      Text            =   "3306"
      Top             =   1125
      Width           =   690
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   330
      Left            =   2790
      TabIndex        =   2
      Text            =   "127.0.0.1"
      Top             =   495
      Width           =   1050
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   270
      Picture         =   "frmDatosServidor.frx":0E76
      Top             =   1035
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   270
      Picture         =   "frmDatosServidor.frx":1740
      Top             =   360
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PUERTO BASE DE LA DATOS"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   855
      TabIndex        =   1
      Top             =   1035
      Width           =   1635
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "IP DEL SERVIDOR"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   855
      TabIndex        =   0
      Top             =   540
      Width           =   1725
   End
End
Attribute VB_Name = "frmMantDatosServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click(Index As Integer)
Dim registro As String
Dim IpServidor As String
Dim puertoLectura As String
On Error GoTo fallo:
Select Case Index
    Case 0:
    'aviso de posible perdida datos en curso
    'guardar datos anteriores
    MsgBox "El cambio de informacion ,puede provocar una perdida de datos." & vbCrLf & _
    "O no establecer la conexion con la base de datos", vbCritical
    
    IpServidor = Trim(Text1)
    puertoLectura = Trim(Text2)
    
    Unload Me
 
    'probar conexion con nuevos datos
    Case 1:
    'restaurar datos anteriores
End Select

Exit Sub
fallo:
MsgBox "Error"
End Sub
