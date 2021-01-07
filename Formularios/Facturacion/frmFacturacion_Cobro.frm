VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.ocx"
Begin VB.Form frmFacturacion_Cobro 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Datos del cobro"
   ClientHeight    =   4650
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   Icon            =   "frmFacturacion_Cobro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4650
   ScaleWidth      =   4560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   2340
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3735
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   3465
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3735
      Width           =   1050
   End
   Begin VB.TextBox txtDatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   3
      Left            =   45
      MaxLength       =   100
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2205
      Width           =   4440
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   990
      MaxLength       =   100
      TabIndex        =   0
      Top             =   1620
      Width           =   3495
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   990
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   3
      Top             =   855
      Width           =   3495
   End
   Begin VB.TextBox txtDatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   3105
      MaxLength       =   100
      TabIndex        =   2
      Top             =   450
      Width           =   1380
   End
   Begin MSComCtl2.DTPicker fecha 
      Height          =   330
      Left            =   990
      TabIndex        =   5
      Top             =   405
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
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
      CalendarTitleBackColor=   12632256
      Format          =   52756481
      CurrentDate     =   38002
   End
   Begin MSDataListLib.DataCombo cmbfp 
      Height          =   315
      Left            =   990
      TabIndex        =   4
      Top             =   1215
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      Text            =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos"
      Height          =   195
      Index           =   1
      Left            =   45
      TabIndex        =   14
      Top             =   1665
      Width           =   420
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones"
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   13
      Top             =   1980
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Forma Pago"
      Height          =   195
      Index           =   3
      Left            =   45
      TabIndex        =   12
      Top             =   1305
      Width           =   855
   End
   Begin VB.Label lblCampos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hora"
      Height          =   240
      Index           =   0
      Left            =   2610
      TabIndex        =   11
      Top             =   450
      Width           =   825
   End
   Begin VB.Label lblCampos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
      Height          =   240
      Index           =   2
      Left            =   45
      TabIndex        =   10
      Top             =   450
      Width           =   825
   End
   Begin VB.Label lblCampos 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Empleado"
      Height          =   240
      Index           =   4
      Left            =   45
      TabIndex        =   9
      Top             =   900
      Width           =   825
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      Caption         =   "Datos del cobro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   45
      TabIndex        =   8
      Top             =   45
      Width           =   4455
   End
End
Attribute VB_Name = "frmFacturacion_Cobro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdok_Click()
    If validar = False Then
        Exit Sub
    End If
    Dim ocobro As New clsDocs_pago_cobros
    With ocobro
        .setDOC_ID = gdoc
        .setFECHA = Format(fecha, "yyyy-mm-dd")
        .setHORA = Format(txtDatos(0), "hh:mm:ss")
        .setFP_ID = cmbfp.BoundText
        .setEMPLEADO_ID = usuario.getID_EMPLEADO
        .setDATOS = txtDatos(2)
        .setOBSERVACIONES = txtDatos(3)
        If .Insertar <> 0 Then
            Dim oDoc As New clsDocs_pago
            oDoc.Cobrar CLng(gdoc)
        End If
    End With
    MsgBox "Los datos del cobro se han insertado correctamente.", vbInformation, App.Title
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    cargar_combos
    fecha = Date
    txtDatos(0) = Format(Time, "hh:mm:ss")
    txtDatos(1) = usuario.getUSUARIO
    If gdoc > 0 Then
        cargar_datos
    End If
End Sub

Public Sub cargar_combos()
    cargar_combo cmbfp, New clsFP
End Sub

Public Sub cargar_datos()
    Dim ocobro As New clsDocs_pago_cobros
    If ocobro.Carga(CLng(gdoc)) = True Then
        With ocobro
            fecha = .getFECHA
            txtDatos(0) = Format(.getHORA, "hh:mm:ss")
            Dim oempleado As New clsUsuarios
            oempleado.CARGAR (.getEMPLEADO_ID)
            txtDatos(1) = oempleado.getNOMBRE
            txtDatos(2) = .getDATOS
            txtDatos(3) = .getOBSERVACIONES
            cmbfp.BoundText = .getFP_ID
        End With
        cmdok.Visible = False
        txtDatos(0).Locked = True
        txtDatos(3).Locked = True
        txtDatos(2).Locked = True
        fecha.Enabled = False
        cmbfp.Locked = True
    End If
End Sub

Public Function validar() As Boolean
    validar = True
    If cmbfp.Text = "" Then
        validar = False
        MsgBox "Introduzca la forma de pago.", vbExclamation, App.Title
    End If
End Function
