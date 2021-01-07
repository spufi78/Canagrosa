VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmOferta_DatosAdicionales 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ficheros Adjuntos a la Oferta"
   ClientHeight    =   7215
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   9375
   Icon            =   "frmOferta_DatosAdicionales.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   9375
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documento adjunto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2130
      Index           =   0
      Left            =   45
      TabIndex        =   8
      Top             =   4140
      Width           =   9285
      Begin VB.CommandButton cmdAnular 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Anular"
         Height          =   870
         Left            =   8235
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1125
         Width           =   960
      End
      Begin VB.CommandButton cmdMostrar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Mostrar"
         Height          =   870
         Index           =   0
         Left            =   7245
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1125
         Width           =   960
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escanear"
         Height          =   870
         Left            =   6255
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   1125
         Width           =   960
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   1530
         TabIndex        =   13
         Top             =   1035
         Visible         =   0   'False
         Width           =   3540
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar"
         Height          =   870
         Index           =   0
         Left            =   5220
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1125
         Width           =   1005
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   330
         Index           =   0
         Left            =   8235
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   915
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         DataSource      =   "Adodc1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   810
         TabIndex        =   0
         Top             =   315
         Width           =   7380
      End
      Begin VB.TextBox datos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataSource      =   "Adodc1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   2790
      End
      Begin VB.TextBox datos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataSource      =   "Adodc1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   6390
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   1800
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   645
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1350
         Width           =   4995
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Archivo"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   405
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   780
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   5715
         TabIndex        =   10
         Top             =   810
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   1125
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   8325
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6345
      Width           =   1005
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2520
      Top             =   6480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3270
      Left            =   45
      TabIndex        =   7
      Top             =   855
      Width           =   9300
      _ExtentX        =   16404
      _ExtentY        =   5768
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
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficheros adjuntos: "
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
      TabIndex        =   15
      Top             =   135
      Width           =   1980
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   8730
      Picture         =   "frmOferta_DatosAdicionales.frx":3AFA
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique los ficheros a adjuntar a la Oferta"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   405
      Width           =   2895
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   9450
   End
End
Attribute VB_Name = "frmOferta_DatosAdicionales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_OFERTA As Long
Private Sub cmdEscaner_Click()
   On Error GoTo cmdEscaner_Click_Error

    frmEscaner.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = ""
        nombreNuevo = InputBox("Introduzca nombre para el archivo Escaneado, sin ninguna extensión (SOLO LETRAS Y NUMEROS).", "Escaneando Archivo", , Screen.Width / 3, (Screen.Height / 3))
        nombreNuevo = Eliminar_Caracteres_Archivo(nombreNuevo)
        If Trim(nombreNuevo) <> "" Then
            datos(4).Text = documento_escaner
            datos(0).Text = nombreNuevo & ".pdf"
            cmdAdjuntar_Click (Index)
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdEscaner_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEscaner_Click of Formulario frmOferta_DatosAdicionales"
End Sub
Private Sub cmdAdjuntar_Click(Index As Integer)
   On Error GoTo cmdAdjuntar_Click_Error

    Me.MousePointer = 11
    ' Validar
    If validar = False Then
        Me.MousePointer = 0
        Exit Sub
    End If
    adjuntar PK_OFERTA
    cargar_lista
    MsgBox "El archivo se ha adjuntado correctamente.", vbOKOnly + vbInformation, App.Title

   On Error GoTo 0
   Exit Sub

cmdAdjuntar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntar_Click of Formulario frmOferta_DatosAdicionales"
End Sub
Private Sub adjuntar(OFERTA As Long)
    Dim strNivel As String
    
    strNivel = "0"
    If copiar(OFERTA) = False Then
        Me.MousePointer = 0
        strNivel = "-1"
        Exit Sub
    End If
    
    strNivel = "1"
    Dim oMuestras_Adjuntos As New clsOfertas_Adjuntos
    
    With oMuestras_Adjuntos
        strNivel = "2"
        .setOFERTA_ID = OFERTA
        strNivel = "3"
        .setRUTA = datos(0)
        strNivel = "4"
        .setEMPLEADO_ID = usuario.getID_EMPLEADO
        strNivel = "5"
        .setFECHA = Format(Date, "yyyy-mm-dd")
        strNivel = "6"
        .setOBSERVACIONES = datos(3)
        strNivel = "7"
        .Insertar
        strNivel = "8"
        Me.MousePointer = 0
        strNivel = "9"
    End With
    strNivel = "10"
    Exit Sub
fallo:
    Me.MousePointer = 0
    error_grave "Error al adjuntar el archivo. En Funcion frmOferta_DatosAdicionales.adjuntar. Nivel codigo: " & strNivel & Err.Description
End Sub

Private Sub cmdAnular_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Seguro de anular el documento adjunto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oMuestras_Adjunto As New clsOfertas_Adjuntos
            oMuestras_Adjunto.Eliminar lista.ListItems(lista.selectedItem.Index).SubItems(3), lista.ListItems(lista.selectedItem.Index).SubItems(4)
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
'    cd.InitDir = "c:\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(4).Text = cd.FileName  ' cd.FileTitle
        datos(0).Text = cd.FileTitle
    End If
End Sub

Private Sub cmdMostrar_Click(Index As Integer)
    Dim destino As String
    destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\OFERTAS-ADJUNTOS\" & lista.ListItems(lista.selectedItem.Index).SubItems(3) & "\" & datos(0)
    On Error GoTo fallo
    If Dir(destino) <> "" Then
        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbMaximizedFocus)
    End If
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_lista
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Fichero", 5500, lvwColumnLeft
        .Add , , "Fecha", 1400, lvwColumnCenter
        .Add , , "Empleado", 1400, lvwColumnCenter
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "Orden", 1, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim m As Long
    Dim rs As ADODB.RecordSet
    m = PK_OFERTA
    If m > 0 Then
        Dim oMuestra_Adjunto As New clsOfertas_Adjuntos
        Set rs = oMuestra_Adjunto.Listado(m)
        lista.ListItems.Clear
        If rs.RecordCount > 0 Then
            Do
                With lista.ListItems.Add(, , rs(0))
                     .SubItems(1) = Format(rs(1), "dd-mm-yyyy")
                     .SubItems(2) = rs(2)
                     .SubItems(3) = rs(3)
                     .SubItems(4) = rs(4)
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        If lista.ListItems.Count > 0 Then
            lista_Click
        End If
    End If
End Sub
Private Function validar() As Boolean
    validar = False
    If datos(4) = "" Then
        MsgBox "Escriba una ruta.", vbInformation, App.Title
        Exit Function
    End If
    If Dir(datos(4)) = "" Then
        MsgBox "La ruta introducida no existe.", vbInformation, App.Title
        Exit Function
    End If
    validar = True
End Function

Private Function copiar(OFERTA As Long) As Boolean
    Dim origen As String
    Dim destino As String
    origen = datos(4)
    On Error Resume Next
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\OFERTAS-ADJUNTOS"
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\OFERTAS-ADJUNTOS\" & OFERTA
    On Error GoTo fallo
    destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\OFERTAS-ADJUNTOS\" & OFERTA & "\" & datos(0)
    
    If Trim(Dir(destino, vbArchive)) = "" Then
        ' si no existe el archivo en destino, lo copia
        FileCopy origen, destino
    End If
    
    copiar = True
    Exit Function
fallo:
    copiar = False
    error_grave "Error al adjuntar el archivo en funcion frmOferta_DatosAdicionales.Copiar " & vbCrLf & " Nº OFERTA: " & OFERTA & vbCrLf & "Origen: " & origen & vbCrLf & "Destino: " & destino & vbCrLf & Err.Description
End Function
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        Dim oMuestra_Adjunto As New clsOfertas_Adjuntos
        With oMuestra_Adjunto
        If .Carga(lista.ListItems(lista.selectedItem.Index).SubItems(3), lista.ListItems(lista.selectedItem.Index).SubItems(4)) Then
            datos(0) = .getRUTA
            Dim oempleado As New clsUsuarios
            oempleado.CARGAR (.getEMPLEADO_ID)
            datos(1) = oempleado.getUSUARIO
            datos(2) = Format(.getFECHA, "dd-mm-yyyy")
            datos(3) = .getOBSERVACIONES
            datos(4) = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\OFERTAS-ADJUNTOS\" & PK_OFERTA & "\" & .getRUTA
        End If
        End With
    End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdMostrar_Click (0)
    End If
End Sub
