VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEmpleados_Formadores 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Proveedores"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmEmpleados_Formadores.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7050
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   6570
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8175
      Style           =   1  'Graphical
      TabIndex        =   29
      Top             =   6570
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
      Height          =   1575
      Index           =   13
      Left            =   60
      TabIndex        =   26
      Top             =   4920
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   1230
         Index           =   13
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   8985
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
      Height          =   1095
      Left            =   60
      TabIndex        =   20
      Top             =   3780
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   12
         Left            =   1140
         MaxLength       =   25
         TabIndex        =   11
         Top             =   660
         Width           =   7875
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   8
         Left            =   1140
         MaxLength       =   25
         TabIndex        =   10
         Top             =   285
         Width           =   7875
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pagina Web"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   25
         Top             =   720
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
      Caption         =   " Datos del Proveedor"
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
      Left            =   60
      TabIndex        =   12
      Top             =   420
      Width           =   9195
      Begin VB.CheckBox chkSubcontrata 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Subcontrata"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   7605
         TabIndex        =   31
         Top             =   0
         Width           =   1275
      End
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
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nuevo Proveedor"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   60
      TabIndex        =   28
      Top             =   60
      Width           =   9180
   End
End
Attribute VB_Name = "frmEmpleados_Formadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'E0070-I
' se cambia gproveedores por PK
Option Explicit
Public PK As Long
'E0070-F

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
    'E0071-I
    'If gproveedor <> 0 Then
    If PK <> 0 Then
    'E0071-F
        modificar_proveedor
    Else
        insertar_proveedor
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_paises
    'E0072-I
    'If gproveedor <> 0 Then
    If PK <> 0 Then
    'E0072-F
        consulta_proveedor
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmEmpleados_Formadores = Nothing
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
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
       'E0066-I
       ' se comenta porque no encuentra la variable (ya está a 0 en el keypress)
       'KeyAscii = 0 ' Para evitar el "bip" del sistema
       'E0066-F
     Case 38
       If Index = 1 Then
        txtDatos(15).SetFocus
       Else
        SendKeys "+{Tab}", True
       End If
       'E0067-I
       'KeyAscii = 0 ' Para evitar el "bip" del sistema
       'E0067-F
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

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = &HFFFFFF
End Sub

Public Sub borrar_campos()
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

Public Sub bloquear_campos()
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

Public Sub insertar_proveedor()
    If valida_datos = False Then
        Exit Sub
    End If
    If MsgBox("Va a dar de alta el proveedor. ¿Esta seguro?", vbQuestion + vbYesNo, "Informacion") = vbYes Then
        'E0068-I
        'Se declara la variable porque el set de abajo no la encuentra
        Dim oProveedor As New clsProveedor
        'E0068-F
        Set oProveedor = mover_datos
        oProveedor.insertar
        borrar_campos
        Set oProveedor = Nothing
    End If
End Sub

Public Sub modificar_proveedor()
    If valida_datos() = False Then
        Exit Sub
    End If
    Dim Pos As Integer
    Dim proveedor As Integer
    If MsgBox("Va a modificar los datos del proveedor. ¿Esta seguro?", vbQuestion + vbYesNo, "Informacion") = vbYes Then
        'E0069-I
        'Se declara la variable porque el set de abajo no la encuentra
        Dim oProveedor As New clsProveedor
        'E0069-F
        Set oProveedor = mover_datos
        'E0073-I
        'oProveedor.setID_PROVEEDOR = gproveedor
        oProveedor.setID_PROVEEDOR = PK
        'E0073-F
        If oProveedor.Modificar = True Then
            Unload Me
        End If
        Set oProveedor = Nothing
    End If

End Sub


Public Function valida_datos() As Boolean
    valida_datos = True
    If txtDatos(1) = "" Then
        MsgBox "El nombre del proveedor no puede estar en blanco.", vbCritical, "Error"
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

Public Sub consulta_proveedor()
    On Error GoTo fallo
    Dim oProveedor As New clsProveedor
    lbltitulo.Caption = "Modificacion de Proveedor"
    lbltitulo.BackColor = &H80C0FF
    'E0073-I
    'oProveedor.Carga (gproveedor)
    oProveedor.Carga (PK)
    'E0073-F
    With oProveedor
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
        'E0200-I
        If .getES_SUBCONTRATA = 0 Then
            chkSubcontrata.value = Unchecked
        Else
            chkSubcontrata.value = Checked
        End If
        'E0200-F
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
    End With
    Set oProveedor = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del proveedor.", vbCritical, Err.Description
End Sub

Public Sub desbloquear_controles()
    Dim i As Integer
    For i = 1 To 13
        txtDatos(i).Locked = False
    Next
    cmbMunicipios.Locked = False
    cmbProvincia.Locked = False
    cmbPais.Locked = False
End Sub

Public Function mover_datos() As clsProveedor
    On Error GoTo fallo
    Dim oProveedor As New clsProveedor
    With oProveedor
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
        'E0200-I
        If chkSubcontrata.value = Unchecked Then
            .setES_SUBCONTRATA = 0
        Else
            .setES_SUBCONTRATA = 1
        End If
        'E0200-F
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
                    .setPAIS_ID = opais.insertar
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
                    .setPROVINCIA_ID = oprov.insertar
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
                    .setMUNICIPIO_ID = omun.insertar
                Else
                    .setMUNICIPIO_ID = municipio
                End If
            End If
        End If
    End With
    Set mover_datos = oProveedor
    Set oProveedor = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del proveedor.", vbCritical, Err.Description
End Function
Public Sub cargar_paises()
    Dim opais As New clsPais
    Set cmbPais.RowSource = opais.Listado  'recorset devuelto por la funcion
    cmbPais.ListField = "nombre" 'campo que veo
    cmbPais.DataField = "nombre" 'campo asociado
    cmbPais.BoundColumn = "id_pais" 'lo que realmente envia
    Set opais = Nothing
End Sub
Public Sub cargar_provincias()
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
Public Sub cargar_municipios()
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

