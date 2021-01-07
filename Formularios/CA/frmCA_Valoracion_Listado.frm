VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCA_Valoracion_Listado 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Listado de Valoraciones del Documento"
   ClientHeight    =   8700
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   13440
   Icon            =   "frmCA_Valoracion_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frmCA_Valoracion_Listado.frx":030A
   ScaleHeight     =   8700
   ScaleWidth      =   13440
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12375
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7785
      Width           =   1050
   End
   Begin VB.Frame frmanalisis 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7080
      Left            =   45
      TabIndex        =   4
      Top             =   630
      Width           =   13365
      Begin MSComctlLib.ListView lista 
         Height          =   6735
         Left            =   90
         TabIndex        =   0
         Top             =   225
         Width           =   13185
         _ExtentX        =   23257
         _ExtentY        =   11880
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
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Valoraciones del Documento realizadas por los Usuarios"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   3
      Top             =   360
      Width           =   4740
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Valoraciones del Documento"
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
      Index           =   0
      Left            =   90
      TabIndex        =   2
      Top             =   45
      Width           =   4170
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   -45
      Width           =   13455
   End
End
Attribute VB_Name = "frmCA_Valoracion_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    If PK <> 0 Then
        cargarLista
    End If
End Sub
Private Sub cargarLista()
    Dim oCA As New clsCa_documentos
    oCA.Carga PK
    lbltitulo(0) = "Listado de Valoraciones de : " & oCA.getNOMBRE
    Dim oCAV As New clsCa_documentos_val
    Set rs = oCAV.Listado(PK)
    lista.ListItems.Clear
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = Format(rs(1), "dd/mm/yyyy") ' Fecha
                .SubItems(2) = rs(2) ' Usuario
                Select Case rs(3)
                    Case 1
                        .SubItems(3) = "Bueno"
                    Case 2
                        .SubItems(3) = "Regular"
                    Case 3
                        .SubItems(3) = "Malo"
                End Select
                Select Case rs(4)
                    Case 1
                        .SubItems(4) = "Bueno"
                    Case 2
                        .SubItems(4) = "Regular"
                    Case 3
                        .SubItems(4) = "Malo"
                End Select
                Select Case rs(5)
                    Case 1
                        .SubItems(5) = "Bueno"
                    Case 2
                        .SubItems(5) = "Regular"
                    Case 3
                        .SubItems(5) = "Malo"
                End Select
                .SubItems(6) = rs(6)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "USUARIO_ID", 1, lvwColumnLeft
        .Add , , "Fecha", 1100, lvwColumnCenter
        .Add , , "Usuario", 2200, lvwColumnCenter
        .Add , , "Claridad", 1300, lvwColumnCenter
        .Add , , "Estructura", 1300, lvwColumnCenter
        .Add , , "Orden", 1300, lvwColumnCenter
        .Add , , "Comentarios", 5500, lvwColumnLeft
    End With
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        frmCA_Valoracion.PK_DOCUMENTO_ID = PK
        frmCA_Valoracion.PK_USUARIO_ID = lista.ListItems(lista.selectedItem.Index).Text
        frmCA_Valoracion.Show 1
    End If
End Sub
