VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEmpleados_fexternos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Formadores Externos"
   ClientHeight    =   8100
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmEmpleados_fexternos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Adjuntos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Index           =   3
      Left            =   45
      TabIndex        =   31
      Top             =   4905
      Width           =   4470
      Begin VB.CommandButton cmdEliminar 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Index           =   3
         Left            =   3975
         Picture         =   "frmEmpleados_fexternos.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   33
         ToolTipText     =   "Eliminar accesorio"
         Top             =   2475
         Width           =   420
      End
      Begin VB.CommandButton cmdEscaner 
         BackColor       =   &H00E0E0E0&
         Height          =   465
         Left            =   3465
         Picture         =   "frmEmpleados_fexternos.frx":049E
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2475
         Width           =   465
      End
      Begin MSComctlLib.ListView lista 
         Height          =   2190
         Left            =   90
         TabIndex        =   34
         Top             =   225
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   3863
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   14609914
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
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   7155
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8175
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   7155
      Width           =   1050
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2160
      Index           =   13
      Left            =   4545
      TabIndex        =   26
      Top             =   4905
      Width           =   4695
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   1770
         Index           =   13
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   27
         Top             =   270
         Width           =   4530
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Otros Datos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   780
      Left            =   60
      TabIndex        =   20
      Top             =   4095
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   5625
         MaxLength       =   25
         TabIndex        =   11
         Top             =   315
         Width           =   3375
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1140
         MaxLength       =   25
         TabIndex        =   10
         Top             =   285
         Width           =   3465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pagina Web"
         Height          =   195
         Index           =   10
         Left            =   4680
         TabIndex        =   25
         Top             =   360
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   13
         Left            =   90
         TabIndex        =   21
         Top             =   375
         Width           =   930
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos del Formador"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3330
      Left            =   45
      TabIndex        =   12
      Top             =   720
      Width           =   9195
      Begin MSDataListLib.DataCombo cmbPais 
         Height          =   315
         Left            =   5280
         TabIndex        =   3
         Top             =   1140
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   5280
         MaxLength       =   30
         TabIndex        =   7
         Top             =   1980
         Width           =   2385
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1080
         MaxLength       =   30
         TabIndex        =   6
         Top             =   1980
         Width           =   2475
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   7
         Left            =   1080
         MaxLength       =   50
         TabIndex        =   9
         Top             =   2790
         Width           =   7980
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1080
         MaxLength       =   15
         TabIndex        =   8
         Top             =   2385
         Width           =   2520
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   2
         Top             =   1140
         Width           =   960
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1080
         MaxLength       =   75
         TabIndex        =   1
         Top             =   735
         Width           =   7980
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1080
         MaxLength       =   75
         TabIndex        =   0
         Top             =   330
         Width           =   7980
      End
      Begin MSDataListLib.DataCombo cmbProvincia 
         Height          =   315
         Left            =   1080
         TabIndex        =   4
         Top             =   1560
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbMunicipios 
         Height          =   315
         Left            =   5280
         TabIndex        =   5
         Top             =   1560
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
         Height          =   195
         Index           =   7
         Left            =   4380
         TabIndex        =   24
         Top             =   1620
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "FAX"
         Height          =   195
         Index           =   15
         Left            =   4380
         TabIndex        =   23
         Top             =   2040
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   195
         Index           =   11
         Left            =   240
         TabIndex        =   22
         Top             =   2040
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "E-mail"
         Height          =   195
         Index           =   6
         Left            =   225
         TabIndex        =   19
         Top             =   2835
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.I.F."
         Height          =   195
         Index           =   5
         Left            =   225
         TabIndex        =   18
         Top             =   2460
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
         Height          =   195
         Index           =   4
         Left            =   240
         TabIndex        =   17
         Top             =   1620
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pais"
         Height          =   195
         Index           =   3
         Left            =   4380
         TabIndex        =   16
         Top             =   1200
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
         Height          =   195
         Index           =   2
         Left            =   225
         TabIndex        =   15
         Top             =   1230
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
         Height          =   195
         Index           =   1
         Left            =   225
         TabIndex        =   14
         Top             =   780
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   13
         Top             =   375
         Width           =   555
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Mantenimiento de Formadores Externos"
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
      Left            =   135
      TabIndex        =   30
      Top             =   180
      Width           =   4125
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   8685
      Picture         =   "frmEmpleados_fexternos.frx":6CF0
      Top             =   45
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   9345
   End
End
Attribute VB_Name = "frmEmpleados_fexternos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long


Private Sub cmdEliminar_Click(Index As Integer)
    If lista.ListItems.Count > 0 Then
        lista.ListItems.Remove lista.selectedItem.Index
    End If
End Sub

Private Sub cmdEscaner_Click()
   On Error GoTo cmdEscaner_Click_Error

    If PK = 0 Then
        Dim c As String
        
        c = "Para añadir adjuntos, es necesario primero añadir al formador."
        c = c & vbNewLine & " Pulse aceptar, para insertar al formador en el sistema y "
        c = c & vbNewLine & " posteriormente añada los adjuntos que desee. "
        
        MsgBox c, vbInformation, App.Title
        Exit Sub
    End If
        
    frmEscaner.Show 1
    If documento_escaner <> "" Then
        Dim nombreNuevo As String
        nombreNuevo = ""
        nombreNuevo = InputBox("Introduzca nombre para el archivo Escaneado, sin ninguna extensión (SOLO LETRAS Y NUMEROS).", "Escaneando Archivo", , Screen.Width / 3, (Screen.Height / 3))
        nombreNuevo = Eliminar_Caracteres_Archivo(nombreNuevo)
        If Trim(nombreNuevo) <> "" Then
            If Dir(documento_escaner) = "" Then
                MsgBox "El documento que intenta vincular no existe en la ruta.", vbExclamation, App.Title
                Exit Sub
            End If
            On Error Resume Next
            Dim ruta As String
            ruta = ReadINI(App.Path + "\config.ini", "Documentos", "ca_fexternos")
            MkDir ruta
            MkDir ruta & "\" & CStr(PK)
            FileCopy documento_escaner, ruta & "\" & CStr(PK) & "\" & nombreNuevo & ".pdf"
            With lista.ListItems.Add(, , lista.ListItems.Count + 1)
                .SubItems(1) = nombreNuevo
                .SubItems(2) = nombreNuevo & ".pdf"
            End With
            MsgBox "Adjunto vinculado correctamente.", vbInformation, App.Title
        End If
    End If

   On Error GoTo 0
   Exit Sub

cmdEscaner_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEscaner_Click of Formulario frmEmpleados_Cualificaciones_Nueva"
End Sub
Private Sub cmbPais_LostFocus()
    cargar_provincias
End Sub
Private Sub cmbProvincia_LostFocus()
    cargar_municipios
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    If PK <> 0 Then
        Modificar
    Else
        Insertar
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_paises
    cabecera
    If PK <> 0 Then
        consulta
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEmpleados_fexternos = Nothing
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    On Error GoTo fallo
    Dim ruta As String
    ruta = ReadINI(App.Path + "\config.ini", "Documentos", "ca_fexternos")
    ruta = ruta & "\" & CStr(PK)
    ruta = ruta & "\" & lista.ListItems(lista.selectedItem.Index).SubItems(2)
    If ruta <> "" Then
        Dim r As Long
        If Dir(ruta) <> "" Then
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & ruta, vbMaximizedFocus)
        Else
            MsgBox "El documento vinculado no existe.", vbCritical, App.Title
        End If
    End If
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento.", vbCritical, App.Title
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40
       If Index = 15 Then
        txtDatos(1).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
     Case 38
       If Index = 1 Then
        txtDatos(15).SetFocus
       Else
        SendKeys "+{Tab}", True
       End If
     Case 27
        cmdcancel_Click
    End Select
End Sub

Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 16 Then
       If Index = 15 Then
        txtDatos(1).SetFocus
        KeyAscii = 0 ' Para evitar el "bip" del sistema
       Else
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
       End If
    End If
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = &HFFFFFF
End Sub

Private Sub borrar_campos()
    Dim i As Integer
    For i = 1 To 13
       If i < 9 Or i > 11 Then
        txtDatos(i) = ""
       End If
    Next
    cmbPais.Text = ""
    cmbProvincia.Text = ""
    cmbMunicipios.Text = ""
    txtDatos(1).SetFocus
End Sub

Private Sub bloquear_campos()
    Dim i As Integer
    For i = 1 To 13
       If i < 9 Or i > 11 Then
        txtDatos(i).Locked = True
       End If
    Next
    cmbMunicipios.Locked = True
    cmbProvincia.Locked = True
    cmbPais.Locked = True
End Sub

Private Sub Insertar()
    If valida_datos = False Then
        Exit Sub
    End If
    If MsgBox("Va a dar de alta al formador. ¿Esta seguro?", vbQuestion + vbYesNo, "Informacion") = vbYes Then
        Dim oF As New clsEmpleados_fexternos
        Set oF = mover_datos
        oF.Insertar
        borrar_campos
        Set oF = Nothing
    End If
End Sub

Private Sub Modificar()
    If valida_datos() = False Then
        Exit Sub
    End If
    Dim Pos As Integer
    Dim formador As Integer
    If MsgBox("Va a modificar los datos del formador. ¿Esta seguro?", vbQuestion + vbYesNo, "Informacion") = vbYes Then
        Dim oF As New clsEmpleados_fexternos
        Set oF = mover_datos
        oF.setID_FEXTERNO = PK
        If oF.Modificar = True Then
            Unload Me
        End If
        Set oF = Nothing
    End If

End Sub


Private Function valida_datos() As Boolean
    valida_datos = True
    If txtDatos(1) = "" Then
        MsgBox "El nombre no puede estar en blanco.", vbCritical, "Error"
        txtDatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtDatos(2) = "" Then
        MsgBox "La dirección no puede estar en blanco.", vbCritical, "Error"
        txtDatos(2).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtDatos(3) = "" Then
        MsgBox "El codigo postal no puede estar en blanco.", vbCritical, "Error"
        txtDatos(3).SetFocus
        valida_datos = False
        Exit Function
    Else
        If Not IsNumeric(txtDatos(3)) Then
            MsgBox "El codigo postal debe ser numérico.", vbCritical, "Error"
            txtDatos(3).SetFocus
            valida_datos = False
            Exit Function
        End If
    End If
    If txtDatos(6) = "" Then
        MsgBox "El CIF no puede estar en blanco.", vbCritical, "Error"
        txtDatos(6).SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbPais.Text = "" Then
        MsgBox "El pais no puede estar en blanco.", vbCritical, "Error"
        cmbPais.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbProvincia.Text = "" Then
        MsgBox "La provincia no puede estar en blanco.", vbCritical, "Error"
        cmbProvincia.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbMunicipios.Text = "" Then
        MsgBox "El municipio no puede estar en blanco.", vbCritical, "Error"
        cmbMunicipios.SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Private Sub consulta()
    On Error GoTo fallo
    Dim oF As New clsEmpleados_fexternos
    oF.Carga (PK)
    With oF
        txtDatos(1) = .getNOMBRE
        txtDatos(2) = .getDIRECCION
        txtDatos(3) = .getCOD_POSTAL
        txtDatos(6) = .getCIF
        txtDatos(4) = .getTELEFONO
        txtDatos(5) = .getFAX
        txtDatos(8) = .getRESPONSABLE
        txtDatos(7) = .getEMAIL
        txtDatos(13) = .getOBSERVACIONES
        txtDatos(12) = .getWEB
        ' Pais
        Dim opais As New clsPais
        opais.CargarPais (.getPAIS_ID)
        cmbPais.BoundText = opais.getNOMBRE
        cmbPais.Text = opais.getNOMBRE
        Set opais = Nothing
        ' Provincia
        Dim oprovincia As New clsProvincias
        oprovincia.CargarProvincia (.getPROVINCIA_ID)
        cmbProvincia.BoundText = oprovincia.getNOMBRE
        cmbProvincia.Text = oprovincia.getNOMBRE
        Set oprovincia = Nothing
        ' Municipio
        Dim oMunicipio As New clsMunicipios
        oMunicipio.CargarMunicipio (.getMUNICIPIO_ID)
        cmbMunicipios.BoundText = oMunicipio.getNOMBRE
        cmbMunicipios.Text = oMunicipio.getNOMBRE
        Set oMunicipio = Nothing
        ' Adjuntos
        Dim ADJUNTOS() As String
        Dim i As Integer
        If .getADJUNTOS <> "" Then
            ADJUNTOS = Split(.getADJUNTOS, ";")
            i = 0
            While i <= UBound(ADJUNTOS) - 1
                With lista.ListItems.Add(, , i)
                    .SubItems(1) = ADJUNTOS(i)
                End With
                i = i + 1
                lista.ListItems(lista.ListItems.Count).SubItems(2) = ADJUNTOS(i)
                i = i + 1
            Wend
        End If
    End With
    Set oF = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del formador.", vbCritical, Err.Description
End Sub

Private Sub desbloquear_controles()
    Dim i As Integer
    For i = 1 To 13
        txtDatos(i).Locked = False
    Next
    cmbMunicipios.Locked = False
    cmbProvincia.Locked = False
    cmbPais.Locked = False
End Sub

Private Function mover_datos() As clsEmpleados_fexternos
    On Error GoTo fallo
    Dim oF As New clsEmpleados_fexternos
    With oF
        .setNOMBRE = txtDatos(1)
        .setDIRECCION = txtDatos(2)
        If txtDatos(3) <> "" Then
            .setCOD_POSTAL = CLng(txtDatos(3))
        Else
            .setCOD_POSTAL = 0
        End If
        .setCIF = txtDatos(6)
        .setTELEFONO = txtDatos(4)
        .setFAX = txtDatos(5)
        .setRESPONSABLE = txtDatos(8)
        .setTIPO = "" ' Ojo
        .setTIPO = "0" ' JONATHAN.2010.05.13
        .setEMAIL = txtDatos(7)
        .setOBSERVACIONES = txtDatos(13)
        .setWEB = txtDatos(12)
        ' Pais
        If cmbPais.Text <> "" Then
            If IsNumeric(cmbPais.BoundText) Then
                .setPAIS_ID = cmbPais.BoundText
            Else
                Dim opais As New clsPais
                Dim pais As Long
                pais = opais.buscar(cmbPais.Text)
                If pais = 0 Then
                    opais.setNOMBRE = cmbPais.Text
                    .setPAIS_ID = opais.Insertar
                Else
                    .setPAIS_ID = pais
                End If
            End If
        End If
        ' Provincia
        If cmbProvincia.Text <> "" Then
            If IsNumeric(cmbProvincia.BoundText) Then
                .setPROVINCIA_ID = cmbProvincia.BoundText
            Else
                Dim oprov As New clsProvincias
                Dim PROVINCIA As Long
                PROVINCIA = oprov.buscar(cmbProvincia.Text)
                If PROVINCIA = 0 Then
                    oprov.setPAIS_ID = .getPAIS_ID
                    oprov.setNOMBRE = cmbProvincia.Text
                    .setPROVINCIA_ID = oprov.Insertar
                Else
                    .setPROVINCIA_ID = PROVINCIA
                End If
            End If
        End If
        ' Municipio
        If cmbMunicipios.Text <> "" Then
            If IsNumeric(cmbMunicipios.BoundText) Then
                .setMUNICIPIO_ID = cmbMunicipios.BoundText
            Else
                Dim omun As New clsMunicipios
                Dim municipio As Long
                municipio = omun.buscar(cmbMunicipios.Text)
                If municipio = 0 Then
                    omun.setPROVINCIA_ID = .getPROVINCIA_ID
                    omun.setNOMBRE = cmbMunicipios.Text
                    .setMUNICIPIO_ID = omun.Insertar
                Else
                    .setMUNICIPIO_ID = municipio
                End If
            End If
        End If
        ' Adjuntos
        Dim ADJUNTOS As String
        ADJUNTOS = ""
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            ADJUNTOS = ADJUNTOS & lista.ListItems(i).SubItems(1) & ";"
            ADJUNTOS = ADJUNTOS & lista.ListItems(i).SubItems(2) & ";"
        Next
        .setADJUNTOS = ADJUNTOS
    End With
    Set mover_datos = oF
    Set oF = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del formador.", vbCritical, Err.Description
End Function
Private Sub cargar_paises()
    Dim opais As New clsPais
    Set cmbPais.RowSource = opais.Listado  'recorset devuelto por la funcion
    cmbPais.ListField = "nombre" 'campo que veo
    cmbPais.DataField = "nombre" 'campo asociado
    cmbPais.BoundColumn = "id_pais" 'lo que realmente envia
    Set opais = Nothing
End Sub
Private Sub cargar_provincias()
'    cmbProvincia.Text = ""
    If cmbPais.Text <> "" Then
     If IsNumeric(cmbPais.BoundText) Then
        Dim oprovincia As New clsProvincias
        Set cmbProvincia.RowSource = oprovincia.Listado(CInt(cmbPais.BoundText))  'recorset devuelto por la funcion
        cmbProvincia.ListField = "nombre" 'campo que veo
        cmbProvincia.DataField = "nombre" 'campo asociado
        cmbProvincia.BoundColumn = "id_provincia" 'lo que realmente envia
        Set oprovincia = Nothing
     End If
    End If
End Sub
Private Sub cargar_municipios()
'    cmbMunicipios.Text = ""
    If cmbProvincia.Text <> "" Then
     If IsNumeric(cmbProvincia.BoundText) Then
        Dim omuni As New clsMunicipios
        Set cmbMunicipios.RowSource = omuni.Listado(CInt(cmbProvincia.BoundText))
        cmbMunicipios.ListField = "nombre" 'campo que veo
        cmbMunicipios.DataField = "nombre" 'campo asociado
        cmbMunicipios.BoundColumn = "id_municipio" 'lo que realmente envia
        Set omuni = Nothing
     End If
    End If
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
         .Add , , "ORDEN", 1, lvwColumnLeft
         .Add , , "Descripción", lista.Width, lvwColumnLeft
         .Add , , "Ruta", 1, lvwColumnLeft
    End With
End Sub

