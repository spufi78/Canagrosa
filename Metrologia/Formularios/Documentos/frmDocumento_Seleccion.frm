VERSION 5.00
Begin VB.Form frmDocumento_Seleccion 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tipos de documentos"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1710
   ControlBox      =   0   'False
   Icon            =   "frmDocumento_Seleccion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   1710
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDoc 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1320
      Index           =   0
      Left            =   75
      Picture         =   "frmDocumento_Seleccion.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   1575
   End
End
Attribute VB_Name = "frmDocumento_Seleccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim primer_boton As Boolean

Private Sub cmdDoc_Click(Index As Integer)
    gTipo_Documento = Index + 1
    Unload Me
End Sub

Private Sub cmdDoc_GotFocus(Index As Integer)
    cmdDoc(Index).BackColor = &HC0E0FF
End Sub

Private Sub cmdDoc_LostFocus(Index As Integer)
    cmdDoc(Index).BackColor = &HE0E0E0
End Sub

Private Sub Form_Load()
    log (Me.Name)
    gTipo_Documento = 0
    Dim oDocumentos_tipos As New clsDocumentos_tipos
    Dim rs As ADODB.Recordset
    primer_boton = True
    Set rs = oDocumentos_tipos.Listado
    If rs.RecordCount <> 0 Then
        Do
            Crear_Boton "", rs("NOMBRE"), rs("ID_TIPO_DOCUMENTO")
            rs.MoveNext
        Loop Until rs.EOF
    End If
End Sub

Public Sub Crear_Boton(imagen As String, NOMBRE As String, ID As Integer)
    Dim Index As Integer
    On Error Resume Next
    If primer_boton Then
        Index = 0
        primer_boton = False
    Else
        Index = cmdDoc.Count
        Load cmdDoc(Index)
        cmdDoc(Index).Top = cmdDoc(Index - 1).Top
        cmdDoc(Index).Left = cmdDoc(Index - 1).Left + cmdDoc(0).Width + 50
        Me.Width = Me.Width + cmdDoc(0).Width + 50
    End If
'    cmdDoc(Index).Picture = Nothing
'    If imagen <> "" Then
'        Dim ruta As String
'        ruta = App.Path & "\imagenes\" & imagen
'        If Dir(ruta) <> "" Then
'            Set cmdDoc(Index).Picture = LoadPicture(ruta)
'        End If
'    End If
    cmdDoc(Index).Caption = NOMBRE
    cmdDoc(Index).Visible = True
End Sub
