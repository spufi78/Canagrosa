VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmClientes_Direcciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Direcciones"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "frmClientes_Direcciones.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9330
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dirección envío de INFORMES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   60
      TabIndex        =   18
      Top             =   2670
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1065
         MaxLength       =   150
         TabIndex        =   5
         Top             =   300
         Width           =   7980
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   0
         Left            =   1065
         MaxLength       =   5
         TabIndex        =   6
         Top             =   705
         Width           =   1455
      End
      Begin MSDataListLib.DataCombo cmbPais_Informes 
         Height          =   315
         Left            =   5265
         TabIndex        =   7
         Top             =   705
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbProvincia_informes 
         Height          =   315
         Left            =   1065
         TabIndex        =   8
         Top             =   1140
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbMunicipios_informes 
         Height          =   315
         Left            =   5265
         TabIndex        =   9
         Top             =   1125
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   23
         Top             =   345
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   22
         Top             =   795
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pais"
         Height          =   195
         Index           =   6
         Left            =   4365
         TabIndex        =   21
         Top             =   765
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   20
         Top             =   1185
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
         Height          =   195
         Index           =   0
         Left            =   4365
         TabIndex        =   19
         Top             =   1185
         Width           =   675
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dirección envío Facturas/Ofertas/Materiales/Productos Controlados"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   60
      TabIndex        =   12
      Top             =   990
      Width           =   9195
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1065
         MaxLength       =   150
         TabIndex        =   0
         Top             =   300
         Width           =   7980
      End
      Begin MSDataListLib.DataCombo cmbPais_envio 
         Height          =   315
         Left            =   5265
         TabIndex        =   2
         Top             =   705
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbProvincia_envio 
         Height          =   315
         Left            =   1065
         TabIndex        =   3
         Top             =   1125
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbMunicipios_envio 
         Height          =   315
         Left            =   5265
         TabIndex        =   4
         Top             =   1125
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1065
         MaxLength       =   5
         TabIndex        =   1
         Top             =   705
         Width           =   1455
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Municipio"
         Height          =   195
         Index           =   7
         Left            =   4365
         TabIndex        =   17
         Top             =   1185
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Provincia"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   16
         Top             =   1185
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pais"
         Height          =   195
         Index           =   3
         Left            =   4365
         TabIndex        =   15
         Top             =   765
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "C.P."
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   795
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Direccion"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   345
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   7110
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4350
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8220
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   4350
      Width           =   1050
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   8700
      Picture         =   "frmClientes_Direcciones.frx":3AFA
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Dependencias de determinaciones"
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
      TabIndex        =   24
      Top             =   120
      Width           =   3645
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   915
      Left            =   0
      Top             =   0
      Width           =   9465
   End
End
Attribute VB_Name = "frmClientes_Direcciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmbPais_envio_Change()
    If cmbPais_envio.Text <> "" Then
     If IsNumeric(cmbPais_envio.BoundText) Then
        cmbProvincia_envio.Text = ""
        cmbMunicipios_envio.Text = ""
        Dim oProvincia As New clsProvincias
        Set cmbProvincia_envio.RowSource = oProvincia.Listado(CInt(cmbPais_envio.BoundText))  'recorset devuelto por la funcion
        cmbProvincia_envio.ListField = "nombre" 'campo que veo
        cmbProvincia_envio.DataField = "nombre" 'campo asociado
        cmbProvincia_envio.BoundColumn = "id_provincia" 'lo que realmente envia
        Set oProvincia = Nothing
     End If
    End If
End Sub
Private Sub cmbPais_Informes_Change()
    If cmbPais_Informes.Text <> "" Then
     If IsNumeric(cmbPais_Informes.BoundText) Then
        cmbProvincia_informes.Text = ""
        cmbMunicipios_informes.Text = ""
        Dim oProvincia As New clsProvincias
        Set cmbProvincia_informes.RowSource = oProvincia.Listado(CInt(cmbPais_Informes.BoundText))  'recorset devuelto por la funcion
        cmbProvincia_informes.ListField = "nombre" 'campo que veo
        cmbProvincia_informes.DataField = "nombre" 'campo asociado
        cmbProvincia_informes.BoundColumn = "id_provincia" 'lo que realmente envia
        Set oProvincia = Nothing
     End If
    End If
End Sub

Private Sub cmbProvincia_envio_Change()
    If cmbProvincia_envio.Text <> "" Then
     If IsNumeric(cmbProvincia_envio.BoundText) Then
        cmbMunicipios_envio.Text = ""
        Dim omuni As New clsMunicipios
        Set cmbMunicipios_envio.RowSource = omuni.Listado(CInt(cmbProvincia_envio.BoundText))
        cmbMunicipios_envio.ListField = "nombre" 'campo que veo
        cmbMunicipios_envio.DataField = "nombre" 'campo asociado
        cmbMunicipios_envio.BoundColumn = "id_municipio" 'lo que realmente envia
        Set omuni = Nothing
     End If
    End If

End Sub

Private Sub cmbProvincia_informes_Change()
    If cmbProvincia_informes.Text <> "" Then
     If IsNumeric(cmbProvincia_informes.BoundText) Then
        cmbMunicipios_informes.Text = ""
        Dim omuni As New clsMunicipios
        Set cmbMunicipios_informes.RowSource = omuni.Listado(CInt(cmbProvincia_informes.BoundText))
         cmbMunicipios_informes.ListField = "nombre" 'campo que veo
         cmbMunicipios_informes.DataField = "nombre" 'campo asociado
         cmbMunicipios_informes.BoundColumn = "id_municipio" 'lo que realmente envia
        Set omuni = Nothing
     End If
    End If
End Sub
Private Sub cmdOk_Click()
   On Error GoTo cmdok_Click_Error
    If USUARIO.getPER_MOD_CLIENTE = False Then
        MsgBox "Su usuario no tiene permisos para modificar los datos del cliente.", vbCritical, App.Title
        Exit Sub
    End If
    If valida_datos Then
        If MsgBox("¿Informar las direcciones del cliente?", vbYesNo + vbQuestion, App.Title) = vbYes Then
            Dim oCliente As New clsCliente
            With oCliente
                .setENVIO_DIRECCION = txtDatos(2)
                .setENVIO_COD_POSTAL = txtDatos(3)
                .setENVIO_PAIS_ID = cmbPais_envio.BoundText
                .setENVIO_PROVINCIA_ID = cmbProvincia_envio.BoundText
                .setENVIO_MUNICIPIO_ID = cmbMunicipios_envio.BoundText
                
                .setINFORMES_DIRECCION = txtDatos(1)
                .setINFORMES_COD_POSTAL = txtDatos(0)
                .setINFORMES_PAIS_ID = cmbPais_Informes.BoundText
                .setINFORMES_PROVINCIA_ID = cmbProvincia_informes.BoundText
                .setINFORMES_MUNICIPIO_ID = cmbMunicipios_informes.BoundText
                
                .modificar_direcciones (PK)
                Unload Me
            End With
        End If
    End If
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmClientes_Direcciones"
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_paises
    If PK <> 0 Then
        cargar_datos
    End If
End Sub
Private Sub cargar_paises()
    Dim opais As New clsPais
    Set cmbPais_envio.RowSource = opais.Listado  'recorset devuelto por la funcion
    cmbPais_envio.ListField = "nombre" 'campo que veo
    cmbPais_envio.DataField = "nombre" 'campo asociado
    cmbPais_envio.BoundColumn = "id_pais" 'lo que realmente envia
    Set cmbPais_Informes.RowSource = opais.Listado  'recorset devuelto por la funcion
    cmbPais_Informes.ListField = "nombre" 'campo que veo
    cmbPais_Informes.DataField = "nombre" 'campo asociado
    cmbPais_Informes.BoundColumn = "id_pais" 'lo que realmente envia
    Set opais = Nothing
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub cargar_datos()
    Dim oCliente As New clsCliente
    With oCliente
        If .CargaCliente(PK) = True Then
            lbltitulo = "Direcciones de : " & .getNOMBRE
            txtDatos(2) = .getENVIO_DIRECCION
            txtDatos(3) = .getENVIO_COD_POSTAL
            cmbPais_envio.BoundText = .getENVIO_PAIS_ID
            cmbProvincia_envio.BoundText = .getENVIO_PROVINCIA_ID
            cmbMunicipios_envio.BoundText = .getENVIO_MUNICIPIO_ID
            
            txtDatos(1) = .getINFORMES_DIRECCION
            txtDatos(0) = .getINFORMES_COD_POSTAL
            cmbPais_Informes.BoundText = .getINFORMES_PAIS_ID
            cmbProvincia_informes.BoundText = .getINFORMES_PROVINCIA_ID
            cmbMunicipios_informes.BoundText = .getINFORMES_MUNICIPIO_ID
            
        End If
    End With
    Set oCliente = Nothing
End Sub

Private Function valida_datos() As Boolean
    valida_datos = True
    ' ENVIOS
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
    If cmbPais_envio.Text = "" Then
        MsgBox "El pais no puede estar en blanco.", vbCritical, "Error"
        cmbPais_envio.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbProvincia_envio.Text = "" Then
        MsgBox "La provincia no puede estar en blanco.", vbCritical, "Error"
        cmbProvincia_envio.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbMunicipios_envio.Text = "" Then
        MsgBox "El municipio no puede estar en blanco.", vbCritical, "Error"
        cmbMunicipios_envio.SetFocus
        valida_datos = False
        Exit Function
    End If
    ' INFORMES
    If txtDatos(1) = "" Then
        MsgBox "La dirección no puede estar en blanco.", vbCritical, "Error"
        txtDatos(2).SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtDatos(0) = "" Then
        MsgBox "El codigo postal no puede estar en blanco.", vbCritical, "Error"
        txtDatos(0).SetFocus
        valida_datos = False
        Exit Function
    Else
        If Not IsNumeric(txtDatos(0)) Then
            MsgBox "El codigo postal debe ser numérico.", vbCritical, "Error"
            txtDatos(0).SetFocus
            valida_datos = False
            Exit Function
        End If
    End If
    If cmbPais_Informes.Text = "" Then
        MsgBox "El pais no puede estar en blanco.", vbCritical, "Error"
        cmbPais_Informes.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbProvincia_informes.Text = "" Then
        MsgBox "La provincia no puede estar en blanco.", vbCritical, "Error"
        cmbProvincia_informes.SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbMunicipios_informes.Text = "" Then
        MsgBox "El municipio no puede estar en blanco.", vbCritical, "Error"
        cmbMunicipios_informes.SetFocus
        valida_datos = False
        Exit Function
    End If
    
End Function
