VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmUsuarios_Conectados 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Usuarios conectados al Geslab"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4005
   ControlBox      =   0   'False
   Icon            =   "frmUsuarios_Conectados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   4005
   Begin VB.CommandButton cmdBorrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2835
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5895
      Width           =   1125
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   30000
      Left            =   330
      Top             =   2910
   End
   Begin MSComctlLib.ListView datos 
      Height          =   5520
      Left            =   30
      TabIndex        =   0
      Top             =   360
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   9737
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Usuarios en Geslab"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   30
      TabIndex        =   1
      Top             =   30
      Width           =   3915
   End
End
Attribute VB_Name = "frmUsuarios_Conectados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdborrar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 3200
    Me.Top = 1300
    cabecera
    cargar_lista
End Sub
Public Sub cabecera()
    ' Datos
    With datos.ColumnHeaders.Add(, , "Empleado", 1700, lvwColumnLeft)
        .Tag = "Empleado"
    End With
    With datos.ColumnHeaders.Add(, , "Máquina", 1800, lvwColumnLeft)
        .Tag = "Máquina"
    End With
End Sub
Public Sub cargar_lista()
    On Error GoTo fallo
    Dim oemp As New clsUsuarios
    Dim rs As ADODB.RecordSet
    Set rs = oemp.Listado_Logo
    datos.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With datos.ListItems.Add(, , rs("nombre"))
                 .SubItems(1) = rs("uso")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set ovb = Nothing
    Exit Sub
fallo:
    Exit Sub
End Sub

Private Sub Timer1_Timer()
    cargar_lista
End Sub
