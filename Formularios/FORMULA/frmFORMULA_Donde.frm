VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmFORMULA_Donde 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de utilización de fórmula"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8940
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmFORMULA_Donde.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   8940
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   7875
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6030
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4980
      Left            =   45
      TabIndex        =   0
      Top             =   1005
      Width           =   8900
      _ExtentX        =   15690
      _ExtentY        =   8784
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H0080C0FF&
      Caption         =   "Tipos de Determinación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   4
      Top             =   780
      Width           =   8925
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de los tipos de determinación donde se utiliza la fórmula"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   420
      Width           =   4455
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   8325
      Picture         =   "frmFORMULA_Donde.frx":000C
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Determinaciones dónde se encuentra la fórmula:"
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
      Width           =   5025
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
Attribute VB_Name = "frmFORMULA_Donde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 250
    Me.Top = 250
    cargar_botones Me
    cabecera
    If PK <> 0 Then
        Dim oFormula As New clsFormulas
        oFormula.cargar (PK)
        lbltitulo = "Determinaciones dónde se encuentra la fórmula: : " & oFormula.getNOMBRE
        cargar_lista
    End If
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Nombre", 4900, lvwColumnLeft
        .Add , , "PNT", 3000, lvwColumnLeft
        .Add , , "ID", 800, lvwColumnCenter
    End With
End Sub
Public Sub cargar_lista()
    Dim rs As ADODB.Recordset
    Dim oTD As New clsTipos_determinacion
    Set rs = oTD.Listado_por_Formula(PK)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs("NOMBRE"))
             .SubItems(1) = rs("PNT")
             .SubItems(2) = Format(rs("ID_TIPO_DETERMINACION"), "0000")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oTD = Nothing
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
        frmTD_Detalle.PK = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
        frmTD_Detalle.Show 1
    End If
End Sub
