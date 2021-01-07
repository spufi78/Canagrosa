VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDonde 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de utilización de tipo de determinación"
   ClientHeight    =   8550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmDonde.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   10380
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9300
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7650
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6795
      Left            =   45
      TabIndex        =   0
      Top             =   795
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   11986
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de los tipos de análisis y baños donde se utiliza el tipo de determinación"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   420
      Width           =   5580
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9765
      Picture         =   "frmDonde.frx":000C
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Analisis y Baños donde se encuentra :"
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
      TabIndex        =   2
      Top             =   90
      Width           =   3975
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   780
      Left            =   0
      Top             =   0
      Width           =   10470
   End
End
Attribute VB_Name = "frmDonde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tipo As Integer
' 0 - Equipos
' 1 - DOCUMENTOS
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Me.SetFocus
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Caption = lbltitulo
End Sub
Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub
Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        Select Case tipo
            Case 0
                frmEquipoEdicion.PK = lista.ListItems(lista.SelectedItem.Index).Text
                frmEquipoEdicion.Show 1
            Case 1
                frmCA_Documento.PK = lista.ListItems(lista.SelectedItem.Index).Text
                frmCA_Documento.Show 1
        End Select
    End If
End Sub
