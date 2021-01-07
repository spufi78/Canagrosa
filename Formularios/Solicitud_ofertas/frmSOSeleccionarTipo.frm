VERSION 5.00
Begin VB.Form frmSOSeleccionarTipo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seleccionar Tipo de Solicitud de Oferta"
   ClientHeight    =   4770
   ClientLeft      =   5970
   ClientTop       =   2100
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4770
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   0
      TabIndex        =   3
      Top             =   360
      Width           =   4755
      Begin VB.OptionButton optTipoSO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "REACTIVOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   540
         TabIndex        =   12
         Top             =   1350
         Width           =   3105
      End
      Begin VB.OptionButton optTipoSO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "ESTRUCTURALES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   540
         TabIndex        =   11
         Top             =   2880
         Width           =   3105
      End
      Begin VB.OptionButton optTipoSO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "FUNGIBLES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   540
         TabIndex        =   10
         Top             =   2490
         Width           =   3105
      End
      Begin VB.OptionButton optTipoSO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "MATERIAL OFICINA"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   540
         TabIndex        =   9
         Top             =   2100
         Width           =   3105
      End
      Begin VB.OptionButton optTipoSO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PRODUCTOS CONTROLADOS"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   540
         TabIndex        =   8
         Top             =   1740
         Width           =   3765
      End
      Begin VB.OptionButton optTipoSO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "PATRONES"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   540
         TabIndex        =   7
         Top             =   990
         Width           =   3105
      End
      Begin VB.OptionButton optTipoSO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "CALIBRACIÓN"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   540
         TabIndex        =   6
         Top             =   630
         Width           =   3105
      End
      Begin VB.OptionButton optTipoSO 
         BackColor       =   &H00C0C0C0&
         Caption         =   "EQUIPO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   540
         TabIndex        =   5
         Top             =   270
         Value           =   -1  'True
         Width           =   3105
      End
      Begin VB.OptionButton optTipoSO 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   210
         Visible         =   0   'False
         Width           =   435
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   2580
      Picture         =   "frmSOSeleccionarTipo.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3840
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   3690
      Picture         =   "frmSOSeleccionarTipo.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3840
      Width           =   1050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Seleccione el Tipo de Solicitud A Crear"
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
      Height          =   330
      Index           =   2
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   4755
   End
End
Attribute VB_Name = "frmSOSeleccionarTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarblnResultado As Boolean
Private mvarlngIdTipoSO As Long


Public Property Get Resultado() As Boolean

    Resultado = mvarblnResultado

End Property

Public Property Let Resultado(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub cmdcancel_Click()
    mvarblnResultado = False
    Me.Hide
End Sub


Private Sub cmdok_Click()

Dim iCont As Long


    For iCont = 1 To 8
        If optTipoSO(iCont).value Then
            mvarlngIdTipoSO = iCont
        End If
    Next iCont

    mvarblnResultado = True
    Me.Hide
    
End Sub




Public Property Get IdTipoSO() As Long

    IdTipoSO = mvarlngIdTipoSO

End Property

Public Property Let IdTipoSO(ByVal lngIdTipoSO As Long)

    mvarlngIdTipoSO = lngIdTipoSO

End Property

