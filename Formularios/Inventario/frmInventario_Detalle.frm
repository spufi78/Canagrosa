VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmInventario_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Inventario"
   ClientHeight    =   7815
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   12150
   Icon            =   "frmInventario_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   12150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Historial Cambios"
      Height          =   870
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6885
      Width           =   1365
   End
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   6885
      Width           =   1050
   End
   Begin VB.Frame frmQR 
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
      Height          =   4875
      Left            =   8010
      TabIndex        =   30
      Top             =   405
      Width           =   4095
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   135
         MaxLength       =   30
         TabIndex        =   33
         Top             =   270
         Width           =   3825
      End
      Begin VB.PictureBox picQR 
         AutoSize        =   -1  'True
         Height          =   3810
         Left            =   135
         Picture         =   "frmInventario_Detalle.frx":1272
         ScaleHeight     =   250
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   250
         TabIndex        =   31
         Top             =   855
         Width           =   3810
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Ubicación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   60
      TabIndex        =   26
      Top             =   2250
      Width           =   7890
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   1125
         MaxLength       =   30
         TabIndex        =   6
         Top             =   945
         Width           =   2070
      End
      Begin MSDataListLib.DataCombo cmbCentro 
         Height          =   315
         Left            =   1125
         TabIndex        =   4
         Top             =   225
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo cmbZona 
         Height          =   315
         Left            =   1125
         TabIndex        =   5
         Top             =   585
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin XtremeSuiteControls.PushButton cmdZonas 
         Height          =   300
         Left            =   7110
         TabIndex        =   35
         Top             =   585
         Width           =   645
         _Version        =   851970
         _ExtentX        =   1138
         _ExtentY        =   529
         _StockProps     =   79
         Appearance      =   5
         Picture         =   "frmInventario_Detalle.frx":2F114
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Centro"
         Height          =   195
         Index           =   3
         Left            =   135
         TabIndex        =   29
         Top             =   315
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Zona"
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   28
         Top             =   630
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Toma Red"
         Height          =   195
         Index           =   7
         Left            =   135
         TabIndex        =   27
         Top             =   1005
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11025
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6885
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   9945
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6885
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
      Height          =   1380
      Index           =   13
      Left            =   90
      TabIndex        =   22
      Top             =   5445
      Width           =   7845
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         Height          =   1050
         Index           =   7
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   225
         Width           =   7635
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos Técnicos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   60
      TabIndex        =   19
      Top             =   3675
      Width           =   7890
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   6
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   10
         Top             =   1290
         Width           =   3105
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   3
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   7
         Top             =   240
         Width           =   3105
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   5
         Left            =   1140
         MaxLength       =   100
         TabIndex        =   9
         Top             =   930
         Width           =   3105
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   4
         Left            =   1140
         MaxLength       =   50
         TabIndex        =   8
         Top             =   585
         Width           =   3105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Clave"
         Height          =   195
         Index           =   20
         Left            =   120
         TabIndex        =   25
         Top             =   1335
         Width           =   405
      End
      Begin VB.Label Centro 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "IP"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   24
         Top             =   300
         Width           =   150
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Usuario"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   21
         Top             =   990
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Gateway"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   20
         Top             =   645
         Width           =   630
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos Descriptivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   60
      TabIndex        =   16
      Top             =   405
      Width           =   7890
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   1125
         MaxLength       =   75
         TabIndex        =   1
         Top             =   675
         Width           =   5550
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   1
         Left            =   1125
         MaxLength       =   100
         TabIndex        =   3
         Top             =   1410
         Width           =   5535
      End
      Begin MSDataListLib.DataCombo cmbTipo 
         Height          =   315
         Left            =   1125
         TabIndex        =   0
         Top             =   315
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbResponsable 
         Height          =   330
         Left            =   1125
         TabIndex        =   2
         Top             =   1035
         Width           =   6600
         _ExtentX        =   11642
         _ExtentY        =   582
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   300
         Left            =   7065
         TabIndex        =   36
         Top             =   315
         Width           =   645
         _Version        =   851970
         _ExtentX        =   1138
         _ExtentY        =   529
         _StockProps     =   79
         Appearance      =   5
         Picture         =   "frmInventario_Detalle.frx":35976
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Responsable"
         Height          =   195
         Index           =   13
         Left            =   135
         TabIndex        =   34
         Top             =   1065
         Width           =   930
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   32
         Top             =   360
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Serie"
         Height          =   195
         Index           =   5
         Left            =   135
         TabIndex        =   18
         Top             =   1455
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   17
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Nuevo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   420
      Left            =   45
      TabIndex        =   23
      Top             =   -15
      Width           =   13590
   End
End
Attribute VB_Name = "frmInventario_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private Sub cmbTipo_change()
    txtCodigo = ""
    If cmbTipo.Text <> "" Then
        Dim oInventario As New clsInventario
        txtCodigo = oInventario.calcularCodigo(cmbTipo.BoundText)
        Set oInventario = Nothing
        
    End If
End Sub
Private Sub cmdHistorialCambios_Click()
    If PK <> 0 Then
        frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_INVENTARIO
        frmHistorialCambios.PK_ID = PK
        frmHistorialCambios.PK_TITULO = "Inventario " & txtDatos(0)
        frmHistorialCambios.Show 1
    End If
End Sub
Private Sub cmdAdjuntos_Click()
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_INVENTARIO
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
End Sub
Private Sub cmdok_Click()

    If Not valida_datos Then Exit Sub
    Dim oInventario As New clsInventario
    Dim ohc As New clsHistorial_cambios
   On Error GoTo cmdok_Click_Error

    With oInventario
        .setTIPO_ID = cmbTipo.BoundText
        .setNOMBRE = txtDatos(0)
        .setIP = txtDatos(3)
        .setGATEWAY = txtDatos(4)
        .setUSUARIO_ID = cmbResponsable.getPK_SALIDA
        .setCENTRO_ID = cmbCentro.BoundText
        .setTOMA_RED = txtDatos(2)
        If cmbZona.BoundText = "" Then
            .setZONA_ID = 0
        Else
            .setZONA_ID = cmbZona.BoundText
        End If
        .setSERIE = txtDatos(1)
        .setUSUARIO = txtDatos(5)
        .setPASS = txtDatos(6)
        .setOBSERVACIONES = txtDatos(7)
        
        If PK <> 0 Then
            If MsgBox("Va a modificar los datos del Inventario. ¿Esta seguro?", vbQuestion + vbYesNo, "Informacion") = vbYes Then
                 frmMotivo.lbltitulo = "Indique detalladamente el motivo de modificación."
                 frmMotivo.Show 1
                 If Trim(MOTIVO) = "" Then
                    MsgBox "Para modificar los datos es necesario introducir el motivo de la modificación.", vbInformation, App.Title
                    Exit Sub
                End If
                If .Modificar(PK) = True Then
                    With ohc
                        .setTIPO = HC_TIPOS.HC_INVENTARIO
                        .setIDENTIFICADOR = PK
                        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                        .setIDENTIFICADOR_TEXTO = txtDatos(0)
                        .setMOTIVO = Trim(MOTIVO)
                        .Insertar
                    End With
                    Set ohc = Nothing
                End If
            End If
        Else
            If MsgBox("Va a dar de alta el Inventario. ¿Esta seguro?", vbQuestion + vbYesNo, "Informacion") = vbYes Then
                Dim ID As Long
                ID = .Insertar
                If ID > 0 Then
                    With ohc
                        .setTIPO = HC_TIPOS.HC_INVENTARIO
                        .setIDENTIFICADOR = ID
                        .setIDENTIFICADOR_TEXTO = txtDatos(0)
                        .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                        .setMOTIVO = HC_CREACION
                        .Insertar
                    End With
                    Set ohc = Nothing
                End If
                PK = ID
            End If
        End If
    End With
    ' Grabar QR
    ' GetDirTemp
    Dim Conversor As Class1
    Set Conversor = New Class1
    Dim fichero As String
    fichero = DIRECTORIO_TEMPORAL & "QR_" & PK & ".jpg"
    Conversor.GrabarJpg picQR.Image, fichero, CByte(70)
    ' Subir a BD
    Dim oDoc As New clsDocumentacion
    oDoc.SubirQR PK, fichero, "QR_" & PK & ".jpg"
    Set Conversor = Nothing
    ' Cerrar
    MsgBox "Registro almacenado correctamente.", vbInformation, App.Title
    Unload Me
    
    Set oInventario = Nothing
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmInventario_Detalle"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdZonas_Click()
    Dim oform As New frmDecodificadoraModal
    oform.CODIGO = DECODIFICADORA.DECODIFICADORA_INVENTARIO_ZONAS
    oform.Show 1
    Set oform = Nothing
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbZona, DECODIFICADORA_INVENTARIO_ZONAS
    Set oDeco = Nothing
End Sub

Private Sub Form_Load()
    log (Me.Name)
'    Dim i As Integer
'    Dim rs As ADODB.Recordset
'    Dim oi As New clsInventario
'    Set rs = oi.Listado("", "", "", "")
'    If rs.RecordCount > 0 Then
'        Do
'            txtCodigo = oi.calcularCodigoID(rs(0))
'            Dim Conversor As Class1
'            Set Conversor = New Class1
'            Dim fichero As String
'            fichero = DIRECTORIO_TEMPORAL & "QR_" & rs(0) & ".jpg"
'            Conversor.GrabarJpg picQR.Image, fichero, CByte(70)
'            ' Subir a BD
'            Dim oDoc As New clsDocumentacion
'            oDoc.SubirQR rs(0), fichero, "QR_" & rs(0) & ".jpg"
'            Set Conversor = Nothing
'
'
'            rs.MoveNext
'        Loop Until rs.EOF
'    End If
    
    cargar_botones Me
    cargar_combos
    If PK <> 0 Then
        cmdAdjuntos.Enabled = True
        consulta
    Else
        cmdAdjuntos.Enabled = False
    End If
    permisos
End Sub
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, DECODIFICADORA_INVENTARIO_TIPOS
    oDeco.cargar_combo cmbZona, DECODIFICADORA_INVENTARIO_ZONAS
    Set oDeco = Nothing
    cargar_combo cmbCentro, New clsCentros
    llenar_combo cmbResponsable, New clsEmpleados, 0, frmEmpleados_Gestion, ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    PK = 0
    Set frmInventario_Detalle = Nothing
End Sub

Private Sub PushButton1_Click()
    Dim oform As New frmDecodificadoraModal
    oform.CODIGO = DECODIFICADORA.DECODIFICADORA_INVENTARIO_TIPOS
    oform.Show 1
    Set oform = Nothing
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbTipo, DECODIFICADORA_INVENTARIO_TIPOS
    Set oDeco = Nothing
    
End Sub

Private Sub txtCodigo_Change()
    Dim cQrCode As New ClsQrCode
    Set cQrCode = New ClsQrCode
    picQR.Picture = cQrCode.GetPictureQrCode(txtCodigo.Text, picQR.ScaleWidth, picQR.ScaleHeight)
    If picQR.Picture Is Nothing Then MsgBox "Error!"
    Set cQrCode = Nothing
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 And Index <> 13 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = &HFFFFFF
End Sub
Private Sub borrar_campos()
    Dim i As Integer
'    For i = 1 To 18
'        txtdatos(i) = ""
'    Next
    txtDatos(1).SetFocus
End Sub

Private Function valida_datos() As Boolean
    valida_datos = True
    If cmbTipo.Text = "" Then
        MsgBox "El Tipo no puede estar en blanco.", vbCritical, "Error"
        cmbTipo.SetFocus
        valida_datos = False
        Exit Function
    End If
    If txtDatos(0) = "" Then
        MsgBox "El nombre  no puede estar en blanco.", vbCritical, "Error"
        txtDatos(0).SetFocus
        valida_datos = False
        Exit Function
    End If
    If cmbCentro.Text = "" Then
        MsgBox "El Centro no puede estar en blanco.", vbCritical, "Error"
        cmbCentro.SetFocus
        valida_datos = False
        Exit Function
    End If
    If Trim(cmbResponsable.getTEXTO) = "" Then
        MsgBox "El Responsable no puede estar en blanco.", vbCritical, "Error"
        cmbResponsable.SetFocus
        valida_datos = False
        Exit Function
    End If
    
End Function

Private Sub consulta()
    On Error GoTo fallo
    Dim oInventario As New clsInventario
    lbltitulo.Caption = "Modificacion de Inventario"
    lbltitulo.BackColor = &H80C0FF
    With oInventario
        .Carga PK
        cmbTipo.BoundText = .getTIPO_ID
        txtDatos(0) = .getNOMBRE
        cmbResponsable.MostrarElemento .getUSUARIO_ID
        txtDatos(1) = .getSERIE
        cmbCentro.BoundText = .getCENTRO_ID
        cmbZona.BoundText = .getZONA_ID
        txtDatos(2) = .getTOMA_RED
        txtDatos(3) = .getIP
        txtDatos(4) = .getGATEWAY
        txtDatos(5) = .getUSUARIO
        txtDatos(6) = .getPASS
        txtDatos(7) = .getOBSERVACIONES
        
        txtCodigo = .calcularCodigoID(PK)
        cmbTipo.Enabled = False
    End With
    Set oInventario = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del oinventario.", vbCritical, Err.Description
End Sub
Private Sub permisos()
    If UCase(USUARIO.getUSUARIO) = "JULIO" Then
        txtDatos(6).PasswordChar = ""
    End If
End Sub
