VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmNC_Adjuntos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ficheros Adjuntos a la Incidencia"
   ClientHeight    =   7185
   ClientLeft      =   735
   ClientTop       =   1365
   ClientWidth     =   9180
   Icon            =   "frmNC_Adjuntos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdEscaner 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Escanear"
      Height          =   870
      Left            =   3210
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6270
      Width           =   960
   End
   Begin VB.CommandButton cmdAdjuntar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntar"
      Height          =   870
      Index           =   0
      Left            =   90
      Picture         =   "frmNC_Adjuntos.frx":3AFA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6255
      Width           =   960
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Anular"
      Height          =   870
      Left            =   1102
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6255
      Width           =   960
   End
   Begin VB.CommandButton cmdMostrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar"
      Height          =   870
      Index           =   0
      Left            =   2115
      Picture         =   "frmNC_Adjuntos.frx":43C4
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   6255
      Width           =   1005
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Documento adjunto"
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
      Height          =   2580
      Index           =   0
      Left            =   45
      TabIndex        =   8
      Top             =   3600
      Width           =   9060
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   4
         Left            =   3015
         TabIndex        =   13
         Top             =   630
         Visible         =   0   'False
         Width           =   2115
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   285
         Index           =   0
         Left            =   8055
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   915
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1035
         TabIndex        =   0
         Top             =   315
         Width           =   6930
      End
      Begin VB.TextBox datos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataSource      =   "Adodc1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1035
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   630
         Width           =   1890
      End
      Begin VB.TextBox datos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         DataSource      =   "Adodc1"
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   6165
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   630
         Width           =   1800
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   1230
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1215
         Width           =   8865
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fichero"
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
         Top             =   360
         Width           =   645
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
         Top             =   675
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
         Left            =   5580
         TabIndex        =   10
         Top             =   675
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
         Top             =   990
         Width           =   1260
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8100
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6255
      Width           =   1005
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2520
      Top             =   6210
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lista 
      Height          =   2685
      Left            =   45
      TabIndex        =   7
      Top             =   855
      Width           =   9075
      _ExtentX        =   16007
      _ExtentY        =   4736
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
      Caption         =   "Indique los ficheros a adjuntar a la incidencia"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   16
      Top             =   405
      Width           =   3180
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   8550
      Picture         =   "frmNC_Adjuntos.frx":44DA
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ficheros adjuntos a la Incidencia"
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
      Width           =   3435
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   9360
   End
End
Attribute VB_Name = "frmNC_Adjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdAdjuntar_Click(Index As Integer)
    On Error GoTo fallo
    Me.MousePointer = 11
    ' Validar
    If validar = False Then
        Me.MousePointer = 0
        Exit Sub
    End If
    If copiar = False Then
        Me.MousePointer = 0
        Exit Sub
    End If
    Dim oNC_Adjuntos As New clsNc_adjuntos
    With oNC_Adjuntos
        .setNC_ID = PK
        .setRUTA = datos(0)
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setFECHA = Format(Date, "yyyy-mm-dd")
        .setOBSERVACIONES = datos(3)
        .Insertar
        cargar_lista
        Me.MousePointer = 0
        MsgBox "El archivo se ha adjuntado correctamente.", vbOKOnly + vbInformation, App.Title
    End With
    Exit Sub
fallo:
    Me.MousePointer = 0
    error_grave "Error al adjuntar el archivo. " & Err.Description
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Seguro de anular el documento adjunto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oMuestras_Adjunto As New clsNc_adjuntos
            oMuestras_Adjunto.Eliminar lista.ListItems(lista.SelectedItem.Index).SubItems(3), lista.ListItems(lista.SelectedItem.Index).SubItems(4)
            cargar_lista
        End If
    End If
End Sub

Private Sub cmdEscaner_Click()
    On Error Resume Next
    
    Dim strArchivo As String
    
    strArchivo = EscanearATemp
    
    If Trim(strArchivo) = "" Then Exit Sub
        
    'datos(0).Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
    datos(0).Text = strArchivo
    datos(1).Text = USUARIO.getUSUARIO
    datos(3).Text = ""
    datos(4).Text = strArchivo
    
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

Private Sub CMDMOSTRAR_Click(Index As Integer)
    Dim destino As String
    destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\" & datos(0)
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
    If PK > 0 Then
        lbltitulo(0) = "Fichero adjuntos a la incidencia, número : " & PK
        Me.Caption = lbltitulo(0)
        Dim oNC As New clsNc
        oNC.Carga PK
        If oNC.getESTADO_ID = C_NC_ESTADOS.CERRADA Then
            cmdAdjuntar(0).Enabled = False
            cmdEliminar.Enabled = False
        End If
        Dim oNC_Adjunto As New clsNc_adjuntos
        Dim rs As ADOdb.RecordSet
        Set rs = oNC_Adjunto.Listado(PK)
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

Private Function copiar() As Boolean
    On Error GoTo fallo
    Dim origen As String
    Dim destino As String
    origen = datos(4)
    destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\" & datos(0)
    FileCopy origen, destino
    copiar = True
    Exit Function
fallo:
    copiar = False
    error_grave "Error al adjuntar el archivo." & Err.Description
End Function
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        Dim oMuestra_Adjunto As New clsNc_adjuntos
        With oMuestra_Adjunto
        If .Carga(lista.ListItems(lista.SelectedItem.Index).SubItems(3), lista.ListItems(lista.SelectedItem.Index).SubItems(4)) Then
            datos(0) = .getRUTA
            Dim oempleado As New clsUsuarios
            oempleado.CARGAR (.getEMPLEADO_ID)
            datos(1) = oempleado.getUSUARIO
            datos(2) = Format(.getFECHA, "dd-mm-yyyy")
            datos(3) = .getOBSERVACIONES
            datos(4) = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\" & .getRUTA
        End If
        End With
    End If
End Sub
