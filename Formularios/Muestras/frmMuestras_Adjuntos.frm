VERSION 5.00
Object = "{05BFD3F1-6319-4F30-B752-C7A22889BCC4}#1.0#0"; "AcroPDF.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMuestras_Adjuntos 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ficheros Adjuntos"
   ClientHeight    =   7635
   ClientLeft      =   60
   ClientTop       =   240
   ClientWidth     =   13980
   Icon            =   "frmMuestras_Adjuntos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7635
   ScaleWidth      =   13980
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdMostrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Abrir Acrobat"
      Height          =   870
      Index           =   0
      Left            =   9495
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   6615
      Width           =   1500
   End
   Begin VB.Frame Frame1 
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
      Height          =   3390
      Index           =   0
      Left            =   45
      TabIndex        =   8
      Top             =   4185
      Width           =   9375
      Begin VB.CommandButton cmd_OR_GuardarCorreo 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Correo"
         Height          =   870
         Left            =   8280
         Picture         =   "frmMuestras_Adjuntos.frx":3AFA
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2385
         Width           =   960
      End
      Begin VB.CommandButton cmdAnular 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eliminar"
         Height          =   870
         Left            =   90
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2385
         Width           =   960
      End
      Begin VB.CommandButton cmdGestor 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Gestor"
         Height          =   870
         Left            =   5310
         Picture         =   "frmMuestras_Adjuntos.frx":3E04
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2385
         Width           =   960
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Escanear"
         Height          =   870
         Left            =   7290
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   2385
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
         Left            =   2430
         TabIndex        =   13
         Top             =   1170
         Visible         =   0   'False
         Width           =   3540
      End
      Begin VB.CommandButton cmdAdjuntar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Adjuntar"
         Height          =   870
         Index           =   0
         Left            =   6300
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2385
         Width           =   960
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   330
         Index           =   0
         Left            =   8190
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   315
         Width           =   1005
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataSource      =   "Adodc1"
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   810
         TabIndex        =   0
         Top             =   315
         Width           =   7290
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
         Left            =   4500
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   1800
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         Height          =   915
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   1395
         Width           =   7965
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Archivo"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   12
         Top             =   405
         Width           =   555
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   780
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   195
         Index           =   1
         Left            =   3780
         TabIndex        =   10
         Top             =   810
         Width           =   450
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   9
         Top             =   1125
         Width           =   1065
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   12870
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6660
      Width           =   1005
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4455
      Top             =   6300
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3405
      Left            =   45
      TabIndex        =   7
      Top             =   720
      Width           =   9390
      _ExtentX        =   16563
      _ExtentY        =   6006
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
   Begin AcroPDFLibCtl.AcroPDF pdf1 
      Height          =   5775
      Left            =   9495
      TabIndex        =   19
      Top             =   720
      Width           =   4470
      _cx             =   5080
      _cy             =   5080
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
      Top             =   90
      Width           =   1980
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13365
      Picture         =   "frmMuestras_Adjuntos.frx":46CE
      Top             =   135
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique los ficheros a adjuntar"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   14
      Top             =   405
      Width           =   2115
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   13950
   End
End
Attribute VB_Name = "frmMuestras_Adjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_MUESTRA As Long
Public PK_MUESTRAS_GRUPO As String
Public PK_NUMERO_RECEPCION_CE As Long
Public PK_DECODIFICADORA As Long
Public PK_DECODIFICADORA_VALOR As Long
Public tipo As Integer
' 0 : MUESTRAS
' 3 : DECODIFICADORA

Public Sub Inicializar()
    PK_MUESTRA = 0
    PK_MUESTRAS_GRUPO = ""
    PK_NUMERO_RECEPCION_CE = 0
    PK_DECODIFICADORA = 0
    PK_DECODIFICADORA_VALOR = 0
End Sub

Private Sub cmd_OR_GuardarCorreo_Click()

    Dim strTempDir As String, strFinalDir As String
    
    Dim objGO As New Geslab_MSOLink.clsMSOOutlook

On Error Resume Next
    limpiar_datos
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "Ruta") & "\temp"
On Error GoTo error_outlook
    strTempDir = ReadINI(App.Path + "\config.ini", "documentos", "Ruta") & "\temp"
    strFinalDir = strTempDir
    If objGO.Guarda_mensaje_outlook(conn, usuario, strTempDir, strFinalDir, "") Then
        If objGO.nombreGenerado <> "" Then
            datos(0) = objGO.nombreGenerado & ".pdf"
            datos(4).Text = strFinalDir & "\" & objGO.nombreGenerado & ".pdf"
            datos(3) = objGO.asuntoGenerado
            MsgBox "Correo convertido a pdf correctamente. Ahora modifique la observacion si lo desea y adjuntelo.", vbInformation, App.Title
            datos(3).SetFocus
        End If
    Else
        If objGO.nombreGenerado <> "" Then
            MsgBox "Se ha producido un error al generar el correo adjunto.", vbCritical, App.Title
        End If
    End If
    Set oGo = Nothing
    
Exit Sub
error_outlook:

    If Err.Number = 440 Then
        MsgBox "No se ha permitido acceder a MS Outlook para adjuntar la Orden de Recarga", vbInformation, "Adjuntar Correo"
        Set oGo = Nothing
    Else
        MsgBox "Error : " & Err.Description, vbCritical, App.Title
    End If
End Sub
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEscaner_Click of Formulario frmMuestras_Adjuntos"
End Sub
Private Sub cmdAdjuntar_Click(Index As Integer)
   On Error GoTo cmdAdjuntar_Click_Error

    Me.MousePointer = 11
    ' Validar
    If validar = False Then
        Me.MousePointer = 0
        Exit Sub
    End If
    Select Case tipo
    Case ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_MUESTRA
        If PK_NUMERO_RECEPCION_CE <> 0 Then
            Dim rs As ADODB.RecordSet
            Dim oce_recepcion As New clsCe_recepcion
            Set rs = oce_recepcion.Listado_por_recepcion(PK_NUMERO_RECEPCION_CE)
            If rs.RecordCount > 0 Then
                Do
                    adjuntar rs("MUESTRA_ID")
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            Set oce_recepcion = Nothing
        Else
            If PK_MUESTRA <> 0 Then
                adjuntar PK_MUESTRA
            Else
                Dim s() As String
                Dim x As Integer
                s = Split(PK_MUESTRAS_GRUPO, ";")
                For x = LBound(s) To UBound(s)
                    If s(x) <> "" Then
                        adjuntar CLng(s(x))
                    End If
                Next
            End If
        End If
        cargar_muestra
    Case ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_DECODIFICADORA  ' DECODIFICADORA
        adjuntar PK_DECODIFICADORA_VALOR
        cargar_decodificadora
    End Select
    MsgBox "El archivo se ha adjuntado correctamente.", vbOKOnly + vbInformation, App.Title
   On Error GoTo 0
   Exit Sub

cmdAdjuntar_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAdjuntar_Click of Formulario frmMuestras_Adjuntos"
End Sub
Private Sub adjuntar(ID As Long)
    Dim strNivel As String
    
    strNivel = "0"
    If copiar(ID) = False Then
        Me.MousePointer = 0
        strNivel = "-1"
        Exit Sub
    End If
    
    strNivel = "1"
    Select Case tipo
    Case ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_MUESTRA ' Muestra
        Dim oM As New clsMuestras_adjuntos
        With oM
            strNivel = "2"
            .setMUESTRA_ID = ID
            strNivel = "3"
            .setRUTA = datos(0)
            strNivel = "4"
            .setEMPLEADO_ID = usuario.getID_EMPLEADO
            strNivel = "5"
            .setFECHA = Format(Date, "yyyy-mm-dd")
            strNivel = "6"
            .setOBSERVACIONES = Trim(datos(3))
            strNivel = "7"
            .Insertar
        End With
        Set oM = Nothing
    Case ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_DECODIFICADORA  ' Decodificadora
        Dim o3 As New clsDecodificadora_adjuntos
        With o3
            .setDECODIFICADORA_ID = PK_DECODIFICADORA
            .setVALOR = ID
            .setRUTA = datos(0)
            .setEMPLEADO_ID = usuario.getID_EMPLEADO
            .setFECHA = Format(Date, "yyyy-mm-dd")
            .setOBSERVACIONES = Trim(datos(3))
            .Insertar
        End With
        Set o3 = Nothing
    End Select
    Me.MousePointer = 0
    strNivel = "10"
    Exit Sub
fallo:
    Me.MousePointer = 0
    error_grave "Error al adjuntar el archivo. En Funcion frmMuestras_Adjuntos.adjuntar. Nivel codigo: " & strNivel & Err.Description
End Sub

Private Sub cmdAnular_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("¿Seguro de anular el documento adjunto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Select Case tipo
            Case ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_MUESTRA  ' Muestra
                Dim oM As New clsMuestras_adjuntos
                oM.Eliminar lista.ListItems(lista.selectedItem.Index).SubItems(3), lista.ListItems(lista.selectedItem.Index).SubItems(4)
                Set oM = Nothing
                cargar_muestra
            Case ADJUNTOS_TIPOS.ADJUNTOS_TIPOS_DECODIFICADORA  ' Decodificadora
                Dim o3 As New clsDecodificadora_adjuntos
                o3.Eliminar PK_DECODIFICADORA, lista.ListItems(lista.selectedItem.Index).SubItems(3), lista.ListItems(lista.selectedItem.Index).SubItems(4)
                Set o3 = Nothing
                cargar_decodificadora
            End Select
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

Private Sub cmdGestor_Click()
   On Error GoTo cmdGestor_Click_Error

    documento_escaner_eliminar = False
    frmGestorDocumentos.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = ""
        nombreNuevo = InputBox("Introduzca nombre para el archivo Escaneado, sin ninguna extensión (SOLO LETRAS Y NUMEROS).", "Escaneando Archivo", , Screen.Width / 3, (Screen.Height / 3))
        nombreNuevo = Eliminar_Caracteres_Archivo(nombreNuevo)
        If Trim(nombreNuevo) <> "" Then
            datos(4).Text = documento_escaner
            datos(0).Text = nombreNuevo & ".pdf"
            cmdAdjuntar_Click (Index)
            If documento_escaner_eliminar = True Then
                On Error Resume Next
                Kill documento_escaner
            End If
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdGestor_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdGestor_Click of Formulario frmMuestras_Adjuntos"
    
End Sub

Private Sub cmdMostrar_Click(Index As Integer)
    If lista.ListItems.Count = 0 Then
        MsgBox "Seleccione algún archivo de la lista.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim destino As String
    
    Select Case tipo
    Case 0 ' Muestras
        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & RUTA_MUESTRA & "\" & lista.ListItems(lista.selectedItem.Index).SubItems(3) & "\" & datos(0)
    Case 3 ' Decodificadora
        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & RUTA_DECODIFICADORA & "\" & PK_DECODIFICADORA & "\" & lista.ListItems(lista.selectedItem.Index).SubItems(3) & "\" & datos(0)
    End Select
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
    pdf1.setShowToolbar True
    Select Case tipo
    Case 0
        cargar_muestra
    Case 3
        cargar_decodificadora
    End Select
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Fichero", 6000, lvwColumnLeft)
        .Tag = "Fichero"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1400, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Empleado", 1400, lvwColumnCenter)
        .Tag = "Empleado"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Orden", 1, lvwColumnCenter)
        .Tag = "Orden"
    End With
End Sub
Public Sub cargar_decodificadora()
    Dim m As Long
    Dim rs As ADODB.RecordSet
    If PK_DECODIFICADORA > 0 Then
        Dim oDA As New clsDecodificadora_adjuntos
        Set rs = oDA.Listado(PK_DECODIFICADORA, PK_DECODIFICADORA_VALOR)
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

Private Sub cargar_muestra()
    limpiar_datos
    Dim m As Long
        Dim rs As ADODB.RecordSet
    If PK_MUESTRA > 0 Then
        m = PK_MUESTRA
    Else
        If PK_NUMERO_RECEPCION_CE > 0 Then
            Dim oce_recepcion As New clsCe_recepcion
            Set rs = oce_recepcion.Listado_por_recepcion(PK_NUMERO_RECEPCION_CE)
            If rs.RecordCount > 0 Then
                m = rs("MUESTRA_ID")
            End If
            Set oce_recepcion = Nothing
        Else
            Dim s() As String
            s = Split(PK_MUESTRAS_GRUPO, ";")
            m = s(0)
        End If
    End If
    If m > 0 Then
        Dim oMuestra_Adjunto As New clsMuestras_adjuntos
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

Private Function copiar(MUESTRA As Long) As Boolean
    Dim origen As String
    Dim destino As String
    origen = datos(4)
    On Error Resume Next
    copiar = False
    Select Case tipo
    Case 0 ' Muestra
        MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & RUTA_MUESTRA & "\" & MUESTRA
        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & RUTA_MUESTRA & "\" & MUESTRA & "\" & datos(0)
    Case 3 ' DECODIFICADORA
        MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & RUTA_DECODIFICADORA
        MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & RUTA_DECODIFICADORA & "\" & PK_DECODIFICADORA
        MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & RUTA_DECODIFICADORA & "\" & PK_DECODIFICADORA & "\" & PK_DECODIFICADORA_VALOR
        destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & RUTA_DECODIFICADORA & "\" & PK_DECODIFICADORA & "\" & PK_DECODIFICADORA_VALOR & "\" & datos(0)
    End Select
    
    On Error GoTo fallo
    If Not gFSO.FileExists(origen) Then
        error_grave "Error al adjuntar el archivo en funcion frmMuestras_Adjuntos(Copiar). NO EXISTE EL ORIGEN. " & vbCrLf & " Origen: " & origen & vbCrLf & "Destino: " & destino & vbCrLf
        copiar = False
        Exit Function
    End If
    
    If Trim(origen) <> Trim(destino) Then
        gFSO.CopyFile origen, destino
    End If
   
'    If Trim(Dir(destino, vbArchive)) = "" Then
'        ' si no existe el archivo en destino, lo copia
'        FileCopy origen, destino
'    End If
    
    If Not gFSO.FileExists(destino) Then
        error_grave "Error al adjuntar el archivo en funcion frmMuestras_Adjuntos(Copiar). NO EXISTE EL DESTINO. " & vbCrLf & " Origen: " & origen & vbCrLf & "Destino: " & destino & vbCrLf
        copiar = False
        Exit Function
    End If
    
    
    copiar = True
    Exit Function
fallo:
    copiar = False
    error_grave "Error al adjuntar el archivo en funcion frmMuestras_Adjuntos.Copiar " & vbCrLf & " Nº Muestra: " & MUESTRA & vbCrLf & "Origen: " & origen & vbCrLf & "Destino: " & destino & vbCrLf & Err.Description
End Function
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        Dim oempleado As New clsUsuarios
        Select Case tipo
        Case 0
            Dim oM As New clsMuestras_adjuntos
            With oM
            If .Carga(lista.ListItems(lista.selectedItem.Index).SubItems(3), lista.ListItems(lista.selectedItem.Index).SubItems(4)) Then
                datos(0) = .getRUTA
                oempleado.CARGAR (.getEMPLEADO_ID)
                datos(1) = oempleado.getUSUARIO
                datos(2) = Format(.getFECHA, "dd-mm-yyyy")
                datos(3) = .getOBSERVACIONES
                datos(4) = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & RUTA_MUESTRA & "\" & gmuestra & "\" & .getRUTA
            End If
            End With
        Case 3 ' DECODIFICADORA
            Dim o3 As New clsDecodificadora_adjuntos
            With o3
            If .Carga(PK_DECODIFICADORA, lista.ListItems(lista.selectedItem.Index).SubItems(3), lista.ListItems(lista.selectedItem.Index).SubItems(4)) Then
                datos(0) = .getRUTA
                oempleado.CARGAR (.getEMPLEADO_ID)
                datos(1) = oempleado.getUSUARIO
                datos(2) = Format(.getFECHA, "dd-mm-yyyy")
                datos(3) = .getOBSERVACIONES
                datos(4) = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & RUTA_PEDIDO_CLIENTE & "\" & PK_OFERTA & "\" & .getRUTA
            End If
            End With
        End Select
        If UCase(Right(datos(0), 3)) = "PDF" Then
            mostrar_pdf datos(4)
        Else
            pdf1.LoadFile vbNullString
            pdf1.Visible = False
            DoEvents
        End If
        
    Else
        pdf1.LoadFile vbNullString
            pdf1.Visible = False
        limpiar_datos
    End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdMostrar_Click (0)
    End If
End Sub

Private Sub limpiar_datos()
    datos(0) = ""
    datos(1) = ""
    datos(2) = ""
    datos(3) = ""
    datos(4) = ""
End Sub
Private Sub mostrar_pdf(DOC As String)
    If Dir(DOC) <> "" Then
        pdf1.Visible = True
        pdf1.LoadFile DOC
        pdf1.setShowToolbar False
    End If
End Sub
