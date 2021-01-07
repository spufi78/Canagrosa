VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmArchivosAdjuntos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Ficheros Adjuntos a la Incidencia"
   ClientHeight    =   7185
   ClientLeft      =   3045
   ClientTop       =   3495
   ClientWidth     =   9180
   Icon            =   "frmArchivosAdjuntos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   9180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdEscaner 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Escanear"
      Height          =   870
      Left            =   3180
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6255
      Width           =   960
   End
   Begin VB.CommandButton cmdAdjuntar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntar"
      Height          =   870
      Index           =   0
      Left            =   90
      Picture         =   "frmArchivosAdjuntos.frx":3AFA
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
      Picture         =   "frmArchivosAdjuntos.frx":43C4
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
         Index           =   5
         Left            =   6180
         TabIndex        =   19
         Top             =   930
         Visible         =   0   'False
         Width           =   1785
      End
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
      Left            =   5535
      Top             =   6345
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
      Picture         =   "frmArchivosAdjuntos.frx":44DA
      Top             =   135
      Width           =   480
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
Attribute VB_Name = "frmArchivosAdjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private mvarobjArchivosAdjuntos As clsGenericCollection
Private mvarobjArchivoSel As New clsArchivoAdjunto
Private mvarenumTipoEdicion As enumTipoEdicion

Private mvarstrSubRuta As String
Private mvarstrRutaElemento As String
Private mvarstrNombre As String
Private mvarblnCopiarAlInstante As Boolean

Public Event AdjuntarArchivo(ByVal prmRutaOriginal As String, ByVal prmNombreArchivo As String, ByVal prmObservaciones As String, ByRef prmDatosActualizados As Boolean)
Public Event EliminarArchivo(ByVal prmRutaOriginal As String, ByVal prmNombreArchivo As String, ByVal prmidAdjunto As Long, ByRef prmDatosActualizados As Boolean)
Private mvarobjrs_archivos As ADODB.RecordSet
Private mvarblnacceso_directo As Boolean

Public Property Get rs_archivos() As ADODB.RecordSet

    Set rs_archivos = mvarobjrs_archivos

End Property

Public Property Set rs_archivos(objrs_archivos As ADODB.RecordSet)

    Set mvarobjrs_archivos = objrs_archivos

End Property

Private Sub OpcionesEdicion()
    If mvarenumTipoEdicion = visualizar Then
        cmdAdjuntar(0).Enabled = False
        cmdEliminar.Enabled = False
    End If
    
End Sub

Public Property Let SubRuta(ByVal cadena As String)
    mvarstrSubRuta = cadena
End Property

Private Sub cmdEscaner_Click()
    On Error Resume Next
    
    Dim strArchivo As String
    
    strArchivo = EscanearATemp
    
    If Trim(strArchivo) = "" Then Exit Sub
        
    'datos(0).Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
    'datos(0).Text = strArchivo ' solo el nombre del archivo
    datos(0).Text = Split(strArchivo, "\")(UBound(Split(strArchivo, "\")))
    datos(1).Text = usuario.getUSUARIO
    datos(3).Text = ""
    datos(4).Text = strArchivo
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjArchivosAdjuntos = Nothing

    Set mvarobjrs_archivos = Nothing
End Sub


Public Property Get ArchivosAdjuntos() As clsGenericCollection

    Set ArchivosAdjuntos = mvarobjArchivosAdjuntos

End Property

Public Property Set ArchivosAdjuntos(objArchivosAdjuntos As clsGenericCollection)

    Set mvarobjArchivosAdjuntos = objArchivosAdjuntos

End Property

Private Sub PresentarDatos()
    lbltitulo(0) = "Fichero adjuntos " & mvarstrNombre
    Me.Caption = lbltitulo(0)
    
    Dim objArchivo As clsArchivoAdjunto
    
    lista.ListItems.Clear
    
    If mvarblnacceso_directo Then
        If mvarobjrs_archivos.RecordCount <> 0 Then
            mvarobjrs_archivos.MoveFirst
            While Not mvarobjrs_archivos.EOF
                With lista.ListItems.Add(, , mvarobjrs_archivos("ruta"))
                     .SubItems(1) = Format(mvarobjrs_archivos("fecha"), "dd-mm-yyyy")
                     .SubItems(2) = mvarobjrs_archivos("empleado")
                     .SubItems(3) = mvarobjrs_archivos("id_adjunto")
                     .SubItems(4) = mvarobjrs_archivos("orden")
                     .SubItems(5) = Replace(mvarobjrs_archivos("ruta_completa"), "/", "\")
                     .SubItems(6) = mvarobjrs_archivos("observaciones")
                End With
                mvarobjrs_archivos.MoveNext
            Wend
        End If
    Else
        If mvarobjArchivosAdjuntos.Count > 0 Then
            For Each objArchivo In mvarobjArchivosAdjuntos.Iterator
                If objArchivo.getID_AUX <> -1 Then
                    With lista.ListItems.Add(, , IIf(objArchivo.getNOMBRE_ARCHIVO_TEMP <> "", objArchivo.getNOMBRE_ARCHIVO_TEMP, objArchivo.getNOMBRE_ARCHIVO))
                         .SubItems(1) = Format(objArchivo.getFECHA, "dd-mm-yyyy")
                         .SubItems(2) = objArchivo.getEMPLEADO_NOMBRE_APELLIDOS
                         .SubItems(3) = objArchivo.getID_ADJUNTO
                         .SubItems(4) = objArchivo.getORDEN
                         .SubItems(6) = objArchivo.getOBSERVACIONES
                         If objArchivo.getRUTA_TEMPORAL <> "" Then
                            .SubItems(5) = objArchivo.getRUTA_TEMPORAL
                        Else
                            .SubItems(5) = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & mvarstrRutaElemento & "\" & mvarstrSubRuta & "\" & objArchivo.getNOMBRE_ARCHIVO
                        End If
                    End With
                End If
            Next objArchivo
        End If
    End If
    
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If

End Sub

Private Sub cmdAdjuntar_Click(Index As Integer)
    
' Añade el archivo a la colección

Dim objDoc As New clsArchivoAdjunto
Dim blnDatosActualizados As Boolean, strNombreArchivo As String

If Not mvarblnacceso_directo Then

    With objDoc
        .setEMPLEADO_ID = usuario.getID_EMPLEADO
        .setEMPLEADO_NOMBRE_APELLIDOS = usuario.getUSUARIO
        .setRUTA = datos(0).Text
        .setRUTA_TEMPORAL = datos(4).Text
        .setFECHA = Format(Date, "yyyy-mm-dd")
        .setOBSERVACIONES = datos(3).Text
        .setORDEN = lista.ListItems.Count
    End With

End If

blnDatosActualizados = False

RaiseEvent AdjuntarArchivo(datos(4).Text, datos(0).Text, datos(3).Text, blnDatosActualizados)

If Not mvarblnacceso_directo Then
    If Not blnDatosActualizados Then
        Call mvarobjArchivosAdjuntos.Add(objDoc)
    End If
End If

Call PresentarDatos


End Sub


Private Sub cmdcancel_Click()
    Me.Hide
End Sub

Private Sub cmdEliminar_Click()

Dim strNombreArchivo As String
Dim strRutaFinal As String, blnDatosActualizados As Boolean
Dim strId As String

strRutaFinal = datos(4).Text
strId = datos(5).Text

If Trim(Dir(strRutaFinal, vbArchive)) = "" Then Exit Sub

strNombreArchivo = Split(strRutaFinal, "\")(UBound(Split(strRutaFinal, "\")))
    
RaiseEvent EliminarArchivo(strRutaFinal, strNombreArchivo, CLng(strId), blnDatosActualizados)


If Not mvarblnacceso_directo Then
    If Not blnDatosActualizados Then
        Call mvarobjArchivosAdjuntos.Remove(mvarobjArchivoSel.getID_ADJUNTO)
    End If
End If

Call PresentarDatos

    
Exit Sub
'    If lista.ListItems.Count > 0 Then
'        If MsgBox("¿Seguro de anular el documento adjunto?", vbQuestion + vbYesNo, App.Title) = vbYes Then
'            Dim oMuestras_Adjunto As New clsNc_adjuntos
'            oMuestras_Adjunto.Eliminar lista.ListItems(lista.SelectedItem.Index).SubItems(3), lista.ListItems(lista.SelectedItem.Index).SubItems(4)
'            cargar_lista
'        End If
'    End If
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
'    cd.InitDir = "c:\"
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(0).Text = cd.FileTitle
        datos(1).Text = usuario.getUSUARIO
        datos(3).Text = ""
        datos(4).Text = cd.FileName
    End If
End Sub

Private Sub cmdMostrar_Click(Index As Integer)
    Dim destino As String
    
    'If Trim(mvarobjArchivoSel.getRUTA_TEMPORAL) <> "" Then
    '    destino = mvarobjArchivoSel.getRUTA_TEMPORAL
    'Else
   '     destino = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & mvarstrRutaElemento & "\" & mvarstrSubRuta & "\" & mvarobjArchivoSel.getNOMBRE_ARCHIVO
   ' End If
        
    destino = datos(4).Text
        
        
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
    ' Pone POR DEFECTO LA SUB RUTA
    mvarstrSubRuta = "" & mvarstrSubRuta & ""
    
    'cargar_lista
    PresentarDatos
    
    OpcionesEdicion
    
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Fichero", 5500, lvwColumnLeft
        .Add , , "Fecha", 1400, lvwColumnCenter
        .Add , , "Empleado", 1400, lvwColumnCenter
        .Add , , "ID", 1, lvwColumnCenter
        .Add , , "Orden", 1, lvwColumnCenter
        .Add , , "ruta_completa", 1, lvwColumnCenter
        .Add , , "observaciones", 1, lvwColumnCenter
    End With
End Sub



Private Sub lista_Click()
    
Dim objDoc As clsProcNcAdjuntos
Dim strId As String

If lista.ListItems.Count = 0 Then Exit Sub


If mvarblnacceso_directo Then
    datos(0).Text = lista.ListItems(lista.SelectedItem.Index).Text
    datos(1).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
    datos(2).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
    datos(3).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(6)
    datos(4).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(5)
    datos(5).Text = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
Else
    strId = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
    Set mvarobjArchivoSel = mvarobjArchivosAdjuntos.Item(strId)
    With mvarobjArchivoSel
        'MsgBox .getNOMBRE_ARCHIVO
        'datos(0).Text = .getRUTA
        datos(0).Text = .getNOMBRE_ARCHIVO
        
        datos(1).Text = .getEMPLEADO_NOMBRE_APELLIDOS
        datos(2).Text = Format(.getFECHA, "dd-mm-yyyy")
        datos(3).Text = .getOBSERVACIONES
        datos(5).Text = .getID_ADJUNTO
        If Trim(.getRUTA_TEMPORAL) <> "" Then
            datos(4).Text = .getRUTA_TEMPORAL
        Else
            datos(4).Text = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & mvarstrRutaElemento & "\" & mvarstrSubRuta & "\" & .getNOMBRE_ARCHIVO
        End If
    End With
End If

Exit Sub

'
'    If lista.ListItems.Count > 0 Then
'        Dim oMuestra_Adjunto As New clsNc_adjuntos
'        With oMuestra_Adjunto
'        If .Carga(lista.ListItems(lista.SelectedItem.Index).SubItems(3), lista.ListItems(lista.SelectedItem.Index).SubItems(4)) Then
'            datos(0) = .getRUTA
'            Dim oempleado As New clsUsuarios
'            oempleado.CARGAR (.getEMPLEADO_ID)
'            datos(1) = oempleado.getUSUARIO
'            datos(2) = Format(.getFECHA, "dd-mm-yyyy")
'            datos(3) = .getOBSERVACIONES
'            datos(4) = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\" & .getRUTA
'        End If
'        End With
'    End If
End Sub

Public Property Get TipoEdicion() As enumTipoEdicion

    TipoEdicion = mvarenumTipoEdicion

End Property

Public Property Let TipoEdicion(ByVal enumTipoEdicion As enumTipoEdicion)

    mvarenumTipoEdicion = enumTipoEdicion

End Property



Public Property Get RutaElemento() As String

    RutaElemento = mvarstrRutaElemento

End Property

Public Property Let RutaElemento(ByVal strRutaElemento As String)

    mvarstrRutaElemento = strRutaElemento

End Property

Public Property Get nombre() As String

    nombre = mvarstrNombre

End Property

Public Property Let nombre(ByVal strNOMBRE As String)

    mvarstrNombre = strNOMBRE

End Property

Public Property Get CopiarAlInstante() As Boolean

    CopiarAlInstante = mvarblnCopiarAlInstante

End Property

Public Property Let CopiarAlInstante(ByVal blnCopiarAlInstante As Boolean)

    mvarblnCopiarAlInstante = blnCopiarAlInstante

End Property

Public Property Get acceso_directo() As Boolean

    acceso_directo = mvarblnacceso_directo

End Property

Public Property Let acceso_directo(ByVal blnacceso_directo As Boolean)

    mvarblnacceso_directo = blnacceso_directo

End Property
