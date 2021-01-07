VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDocumentos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de documentos"
   ClientHeight    =   8790
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11760
   Icon            =   "frmDocumentos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8790
   ScaleWidth      =   11760
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   1215
      Left            =   2070
      TabIndex        =   10
      Top             =   3375
      Visible         =   0   'False
      Width           =   3765
   End
   Begin VB.CommandButton cmdBuscar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6870
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8100
      Width           =   1065
   End
   Begin VB.TextBox txtnumero 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5520
      TabIndex        =   7
      Top             =   8100
      Width           =   1275
   End
   Begin VB.TextBox txtanno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   750
      TabIndex        =   5
      Top             =   8100
      Width           =   945
   End
   Begin MSComCtl2.UpDown cambiar 
      Height          =   435
      Left            =   1695
      TabIndex        =   4
      Top             =   8100
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   767
      _Version        =   393216
      Value           =   2004
      BuddyControl    =   "txtanno"
      BuddyDispid     =   196612
      OrigLeft        =   1590
      OrigTop         =   6570
      OrigRight       =   1830
      OrigBottom      =   6975
      Max             =   2015
      Min             =   2004
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   10620
      Picture         =   "frmDocumentos.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   7860
      Width           =   1050
   End
   Begin VB.FileListBox File 
      Height          =   1065
      Left            =   5895
      Pattern         =   "*.doc"
      TabIndex        =   2
      Top             =   3375
      Visible         =   0   'False
      Width           =   3645
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6360
      Top             =   7800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentos.frx":114C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDocumentos.frx":23CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tree 
      Height          =   7410
      Left            =   60
      TabIndex        =   0
      Top             =   345
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   13070
      _Version        =   393217
      HideSelection   =   0   'False
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "ImageList1"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Buscar número General"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2580
      TabIndex        =   9
      Top             =   8160
      Width           =   2865
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Top             =   8160
      Width           =   585
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Listado de Documentos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   300
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   11625
   End
End
Attribute VB_Name = "frmDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cambiar_Change()
    cargar_tree (txtAnno)
End Sub
Private Sub cmdBuscar_Click()
    Dim rs As New ADODB.RecordSet
'    Dim tipo_muestra As String
    Dim ref As String
    Dim oMuestra As New clsMuestra
    ' Tipo de la muestra
'        For i = 1 To Tree.Nodes.Count
'            If Tree.Nodes(i).Bold = True Then
'                Tree.Nodes(i).Bold = False
'            End If
'        Next
    
    Set rs = oMuestra.obtener_id_muestra(CLng(txtnumero), CInt(txtAnno))
    Dim c As Integer
    Dim Pos As Integer
    Dim nombre As Integer
    If rs.RecordCount <> 0 Then
        ' Referencia
        oMuestra.CargaMuestra (rs(0))
        ref = oMuestra.CodigoParticular(rs(0)) & " " & _
              oMuestra.getREFERENCIA_CLIENTE & " " & _
              Format(oMuestra.getFECHA_CIERRE, "dd-mm-yyyy") & " Ed_" & _
              oMuestra.getULT_EDICION_IMP & ".doc"
        ' Eliminar caracter invalidos
        ref = Replace(ref, """", "'")
        ref = Replace(ref, "/", "")
        ref = Replace(ref, ":", "")
        ref = Replace(ref, "*", "")
        For i = 1 To Tree.Nodes.Count
            If Tree.Nodes(i).Key = ref Then
                Tree.Nodes(i).Selected = True
'                Tree.Nodes(i).Bold = True
                Exit Sub
            End If
        Next
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 50
    Me.Top = 50
    txtAnno = Year(Date)
    cambiar.Max = Year(Date)
    cargar_tree (txtAnno)
End Sub

Private Sub Tree_DblClick()
    Dim aux As String
'    Dim i As Integer
    Dim RUTA As String
    On Error GoTo fallo
    Me.MousePointer = 11
    RUTA = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta")
    If UCase(USUARIO.getUSUARIO) = "PRUEBA" Then
        RUTA = RUTA & "\Prueba" & "\"
    Else
        RUTA = RUTA & "\" & txtAnno & "\"
    End If
    Dim Pos As Integer
    If Tree.Nodes(Tree.SelectedItem.Index).Bold = False Then
        If Right(Tree.SelectedItem.Text, 1) <> ")" Then
            Pos = InStr(1, Tree.SelectedItem.Parent.Text, "(", vbTextCompare)
            If Pos > 0 Then
               aux = Mid(Tree.SelectedItem.Parent, 1, Pos - 2) & "\" & Tree.SelectedItem
               ver_documento_word (RUTA & Tree.Nodes(Tree.SelectedItem.Index).Parent.Parent & "\" & aux)
            End If
        End If
    End If
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al abrir el documento.", vbInformation, App.Title
End Sub

Public Sub cargar_tree(ANNO As Integer)
    On Error GoTo fallo
    Dim nodX As Node
    Dim aux, nombre As String
    Dim i, j, c As Integer
    Dim Pos As Integer
    Dim RUTA As String
    Dim mianno As String
    Dim mimes As String
    Dim clave As String
    Tree.Nodes.Clear
    Dim rs As New ADODB.RecordSet
    RUTA = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta")
    If UCase(USUARIO.getUSUARIO) = "PRUEBA" Then
        RUTA = RUTA & "\Prueba" & "\"
    Else
        RUTA = RUTA & "\" & ANNO & "\"
    End If
    ' Muestra los nombres que representan directorios.
    Dir1.Path = RUTA
    For j = 0 To Dir1.ListCount - 1
'    mianno = Dir(ruta, vbDirectory)   ' Recupera la primera entrada.
     mianno = Dir1.List(j)
     Pos = Len(Dir1.List(j))
     While Pos > 0 And c = 0
        If InStr(Pos, Dir1.List(j), "\", vbTextCompare) > 0 Then
            c = Pos
        End If
        Pos = Pos - 1
     Wend
     clave = Right(Dir1.List(j), Len(Dir1.List(j)) - c)
     Set nodX = Tree.Nodes.Add(, , clave, clave, 1)
     Tree.Nodes(nodX.Index).Bold = True
'     Do While mianno <> ""   ' Inicia el bucle.
       ' Ignora el directorio actual y el que lo abarca.
'       If mianno <> "." And mianno <> ".." Then
          ' Realiza una comparaciónnivel de bit para asegurarse de que MiNombre es un directorio.
'          If (GetAttr(ruta & mianno) And vbDirectory) = vbDirectory Then
             mimes = Dir(mianno & "\", vbDirectory)
             Do While mimes <> ""
               If mimes <> "." And mimes <> ".." Then
                 If UCase(USUARIO.getUSUARIO) = "PRUEBA" Then
'                     File.Path = mianno & "\"
'                     For i = 0 To File.ListCount - 1
'                         Set nodX = Tree.Nodes.Add(clave, tvwChild, clave & File.List(i), File.List(i), 1)
    '                     Set nodX = Tree.Nodes.Add(clave & mimes, tvwChild, File.List(i), File.List(i), 2)
'                     Next
                 Else
                     consulta = "select nombre from tipos_muestra where codigo = '" & mimes & "'"
                     Set rs = datos_bd(consulta)
                     If rs.RecordCount <> 0 Then
                        Set nodX = Tree.Nodes.Add(clave, tvwChild, clave & mimes, mimes & " (" & rs("nombre") & ")", 1)
                        File.Path = mianno & "\" & mimes
                        For i = 0 To File.ListCount - 1
    '                      Set nodX = Tree.Nodes.Add(mimes, tvwChild, mimes & "-" & CStr(i), File.List(i), 2)
                          Set nodX = Tree.Nodes.Add(clave & mimes, tvwChild, File.List(i), File.List(i), 2)
                        Next
                     End If
                 End If
               End If
               mimes = Dir
             Loop
'          End If   ' solamente si representa un directorio.
'       End If
'       mianno = Dir   ' Obtiene siguiente entrada.
'    Loop
    Next
    Exit Sub
fallo:
    MsgBox Err.Description, vbCritical, App.Title & ": " & Err.Number
End Sub

Private Sub txtnumero_GotFocus()
    txtnumero.BackColor = &H80C0FF
    txtnumero.SelStart = 0
    txtnumero.SelLength = Len(txtnumero)
End Sub

Private Sub txtnumero_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdBuscar_Click
    End If
End Sub

Private Sub txtnumero_LostFocus()
    txtnumero.BackColor = vbWhite
End Sub
