VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmEP_Paquete_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Envío de paquetes - Detalle del paquete"
   ClientHeight    =   8760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10770
   Icon            =   "frmEP_Paquete_Detalle.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   10770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
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
      ForeColor       =   &H80000008&
      Height          =   825
      Left            =   45
      TabIndex        =   20
      Top             =   6930
      Width           =   10680
      Begin VB.CheckBox chkFechaRecepcion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   240
         Left            =   135
         TabIndex        =   23
         Top             =   360
         Width           =   240
      End
      Begin MSComCtl2.DTPicker fechaRecepcion 
         Height          =   315
         Left            =   2025
         TabIndex        =   21
         Top             =   315
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   16515073
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha de Recepción"
         Height          =   240
         Index           =   5
         Left            =   405
         TabIndex        =   22
         Top             =   375
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdAdjuntos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Adjuntos"
      Height          =   885
      Left            =   7020
      Style           =   1  'Graphical
      TabIndex        =   19
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   9510
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7800
      Width           =   1215
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   885
      Left            =   8287
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7800
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle"
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
      Height          =   4650
      Left            =   30
      TabIndex        =   12
      Top             =   2250
      Width           =   10680
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   4245
         Index           =   2
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   270
         Width           =   10455
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del paquete"
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
      Height          =   1500
      Left            =   30
      TabIndex        =   8
      Top             =   690
      Width           =   10680
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Proveedor"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   3000
         TabIndex        =   18
         Top             =   30
         Width           =   1395
      End
      Begin VB.OptionButton opTipo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   1890
         TabIndex        =   17
         Top             =   30
         Value           =   -1  'True
         Width           =   1125
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   9270
         TabIndex        =   7
         Top             =   660
         Width           =   1320
      End
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1215
         MaxLength       =   255
         TabIndex        =   2
         Top             =   1050
         Width           =   9375
      End
      Begin MSDataListLib.DataCombo cmbMensajeria 
         Height          =   315
         Left            =   1215
         TabIndex        =   1
         Top             =   690
         Width           =   6285
         _ExtentX        =   11086
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Text            =   ""
      End
      Begin pryCombo.miCombo cmbClientes 
         Height          =   330
         Left            =   1215
         TabIndex        =   0
         Top             =   330
         Width           =   7320
         _ExtentX        =   12912
         _ExtentY        =   582
      End
      Begin MSComCtl2.DTPicker datFEnvio 
         Height          =   315
         Left            =   9270
         TabIndex        =   6
         Top             =   300
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   16515073
         CurrentDate     =   2
         MinDate         =   2
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Hora"
         Height          =   240
         Index           =   0
         Left            =   8640
         TabIndex        =   14
         Top             =   705
         Width           =   555
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   240
         Left            =   8640
         TabIndex        =   13
         Top             =   345
         Width           =   510
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   240
         Index           =   4
         Left            =   135
         TabIndex        =   11
         Top             =   1095
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   240
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   375
         Width           =   915
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mensajería"
         Height          =   240
         Left            =   135
         TabIndex        =   9
         Top             =   735
         Width           =   915
      End
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Envío de Paquetes"
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
      TabIndex        =   16
      Top             =   0
      Width           =   1980
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Indique los datos para realizar el envío de paquetes"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   15
      Top             =   300
      Width           =   3660
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   10110
      Picture         =   "frmEP_Paquete_Detalle.frx":08CA
      Top             =   90
      Width           =   480
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   10710
   End
End
Attribute VB_Name = "frmEP_Paquete_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long



Private Sub chkFechaRecepcion_Click()
    If chkFechaRecepcion.Value = Checked Then
        fechaRecepcion.Enabled = True
        fechaRecepcion = Date
    Else
        fechaRecepcion.Enabled = False
    End If
End Sub

Private Sub cmdAdjuntos_Click()
    'M1138-I
    If PK = 0 Then Exit Sub
    With frmAdjuntos
        .TOBJETO = TOBJETO.TOBJETO_PAQUETE
        .COBJETO = PK
        .Show 1
    End With
    Set frmAdjuntos = Nothing
    'M1138-F
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.top = 1500
    Me.Left = 700
    cargar_botones Me
    
    Call cargar_combos
    
    Dim titulo As String
    If PK <> 0 Then
        lbltitulo = "Envío de paquetes - Modificación de envío"
        CARGAR
    Else
        'M1138-I
        cmdAdjuntos.Visible = False
        'M1138-F
        lbltitulo = "Envío de paquetes - Alta de envío"
        Me.Caption = lbltitulo
        datFEnvio = Date
        txtDatos(3) = Format(Time, "hh:nn:ss")
    End If
End Sub

Private Sub cmbMensajeria_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

' botones
Private Sub cmdok_Click()
    If datos_correctos Then
        Dim oPaquete As New clsEP_Paquetes
        Dim PAQUETE As Long
        With oPaquete
            .setASUNTO = txtDatos(1)
            .setDETALLE = txtDatos(2)
            If opTipo(0).Value = True Then
                .setTIPO = 0
            Else
                .setTIPO = 1
            End If
            If cmbClientes.getTEXTO = "" Then
                .setCLIENTE_ID = 0
            Else
                .setCLIENTE_ID = cmbClientes.getPK_SALIDA
            End If
            .setMENSAJERIA_ID = IIf(cmbMensajeria.BoundText = "", 0, cmbMensajeria.BoundText)
            .setFECHA_CREACION = Format(datFEnvio, "yyyy-mm-dd")
            .setHORA_CREACION = Format(txtDatos(3), "hh:nn:ss")
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setFECHA_RECEPCION = ""
            If chkFechaRecepcion.Value = Checked Then
                .setFECHA_RECEPCION = fechaRecepcion
            End If
        End With
      
        If PK = 0 Then
            If MsgBox("Va a crear un envío. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                PAQUETE = oPaquete.Insertar
            Else
                Exit Sub
            End If
        Else
            If MsgBox("Va a modificar el envío. ¿Está seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
                oPaquete.Modificar (PK)
                PAQUETE = PK
            Else
                Exit Sub
            End If
        End If
        If PK = 0 Then
            MsgBox "El envío se ha creado correctamente. ", vbOKOnly + vbInformation, App.Title
            PK = PAQUETE
        Else
            MsgBox "El envío se ha modificado correctamente.", vbOKOnly + vbInformation, App.Title
            Unload Me
        End If
        frmEP_Listado.cargar_lista
        Unload Me
    End If
End Sub

Private Sub cmdcancel_Click()
    PK = 0
    Unload Me
End Sub

' Funciones auxiliares del formulario
Private Sub cargar_combos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbMensajeria, DECODIFICADORA.EP_EMPRESAS_MENSAJERIA
    cargar_clientes
End Sub
Private Sub cargar_clientes()
    cmbClientes.Limpiar
    If opTipo(0).Value = True Then
        llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    Else
        llenar_combo cmbClientes, New clsProveedor, 0, frmProveedores_Detalle, ""
    End If
End Sub

Public Function datos_correctos() As Boolean
    datos_correctos = True
    If Trim(cmbClientes.getTEXTO) = "" Then ' cliente
        MsgBox "Debe indicar un cliente.", vbInformation, App.Title
        cmbClientes.SetFocus
        datos_correctos = False
        Exit Function
    End If
    If Trim(cmbMensajeria.Text) = "" Then ' mensajeria
        MsgBox "Debe indicar una empresa de mensajería.", vbInformation, App.Title
        cmbMensajeria.SetFocus
        datos_correctos = False
        Exit Function
    End If
    If Trim(txtDatos(1)) = "" Then ' descripción
        MsgBox "Debe indicar una descripción para el envío.", vbInformation, App.Title
        txtDatos(1).SetFocus
        datos_correctos = False
        Exit Function
    End If
    If Trim(txtDatos(2)) = "" Then ' detalle
        MsgBox "Debe indicar un detalle para el envío.", vbInformation, App.Title
        txtDatos(2).SetFocus
        datos_correctos = False
        Exit Function
    End If
    
End Function

Public Sub CARGAR()
    Dim oPaquete As New clsEP_Paquetes
    
    If oPaquete.Carga(PK) Then
        lbltitulo = "Envío de paquetes - Modificación de envío : " & Format(oPaquete.getID_PAQUETE, "0000")
        Me.Caption = lbltitulo
        With oPaquete
            If .getTIPO = 1 Then
                opTipo(1).Value = True
            End If
            txtDatos(1) = .getASUNTO
            txtDatos(2) = .getDETALLE
            txtDatos(3) = .getHORA_CREACION
            cmbClientes.MostrarElemento .getCLIENTE_ID
            cmbMensajeria.BoundText = .getMENSAJERIA_ID
            datFEnvio = .getFECHA_CREACION
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            If .getFECHA_RECEPCION <> "" Then
                chkFechaRecepcion.Value = Checked
                fechaRecepcion.Enabled = True
                fechaRecepcion = .getFECHA_RECEPCION
            End If
        End With
    End If
        
    Set oPaquete = Nothing
End Sub

Private Sub opTipo_Click(Index As Integer)
    If Index = 0 Then
        Label1(1).Caption = "Clientes"
    Else
        Label1(1).Caption = "Proveedores"
    End If
    cargar_clientes
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &HC0E0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
