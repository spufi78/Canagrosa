VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Begin VB.Form frmGestorDocumentos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestor de Documentos"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14490
   Icon            =   "frmGestorDocumentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   14490
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkEliminar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Eliminar el Archivo despues de Adjuntar"
      Height          =   240
      Left            =   8685
      TabIndex        =   12
      Top             =   8685
      Width           =   3570
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntar"
      Height          =   870
      Left            =   12330
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8685
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Carpeta de almacenamiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1770
      Left            =   45
      TabIndex        =   4
      Top             =   0
      Width           =   4200
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Iberia"
         Height          =   195
         Index           =   68
         Left            =   2205
         TabIndex        =   15
         Top             =   765
         Width           =   1545
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Facturas"
         Height          =   195
         Index           =   67
         Left            =   2205
         TabIndex        =   14
         Top             =   540
         Width           =   1320
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Metrología"
         Height          =   195
         Index           =   66
         Left            =   2205
         TabIndex        =   13
         Top             =   315
         Width           =   1680
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Calidad"
         Height          =   195
         Index           =   65
         Left            =   180
         TabIndex        =   10
         Top             =   1440
         Width           =   2265
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Logística"
         Height          =   195
         Index           =   64
         Left            =   180
         TabIndex        =   9
         Top             =   1215
         Width           =   2265
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Administración"
         Height          =   195
         Index           =   63
         Left            =   180
         TabIndex        =   8
         Top             =   990
         Width           =   2265
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Recepción"
         Height          =   195
         Index           =   62
         Left            =   180
         TabIndex        =   7
         Top             =   765
         Width           =   2265
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Documentación"
         Height          =   195
         Index           =   61
         Left            =   180
         TabIndex        =   6
         Top             =   540
         Width           =   2265
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Análisis"
         Height          =   195
         Index           =   60
         Left            =   180
         TabIndex        =   5
         Top             =   315
         Value           =   -1  'True
         Width           =   2265
      End
   End
   Begin VB.FileListBox File1 
      Height          =   6915
      Left            =   0
      Pattern         =   "*.pdf"
      TabIndex        =   3
      Top             =   1800
      Width           =   4245
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   870
      Left            =   13410
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8685
      Width           =   1050
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   4950
      TabIndex        =   1
      Top             =   8820
      Visible         =   0   'False
      Width           =   2760
   End
   Begin AcroPDFLibCtl.AcroPDF pdf1 
      Height          =   8610
      Left            =   4275
      TabIndex        =   0
      Top             =   45
      Width           =   10185
      _cx             =   17965
      _cy             =   15187
   End
End
Attribute VB_Name = "frmGestorDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    If File1.ListCount > 0 Then
        documento_escaner = Dir1.Path & "\" & File1.List(File1.ListIndex)
        If chkEliminar.value = Checked Then
            documento_escaner_eliminar = True
        End If
        Unload Me
    End If
End Sub
Private Sub cmdSalir_Click()
    documento_escaner = ""
    Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub File1_Click()
    If File1.ListCount > 0 Then
        mostrar_pdf Dir1.Path & "\" & File1.List(File1.ListIndex)
    Else
        pdf1.LoadFile vbNullString
    End If
End Sub

Private Sub Form_Load()
    documento_escaner_eliminar = False
    log Me.Name
    cargar_botones Me
    opTipo_Click (60)
End Sub
Private Sub mostrar_pdf(DOC As String)
    If Dir(DOC) <> "" Then
        pdf1.LoadFile DOC
        pdf1.setShowToolbar False
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    pdf1.LoadFile vbNullString
End Sub

Private Sub opTipo_Click(Index As Integer)
    Dim op As New clsParametros
    op.Carga CLng(Index), ""
    Dir1.Path = Replace(op.getVALOR, "/", "\")
    pdf1.LoadFile vbNullString
End Sub
