VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmDocumento_Anular 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imprimir"
   ClientHeight    =   5805
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4830
      Width           =   1275
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   915
      Left            =   3480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4830
      Width           =   1275
   End
   Begin VB.TextBox txtdatos 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   2265
      Index           =   1
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   2520
      Width           =   5985
   End
   Begin VB.TextBox txtdatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   870
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   660
      Width           =   2235
   End
   Begin MSComCtl2.DTPicker txtfecha 
      Height          =   390
      Left            =   4530
      TabIndex        =   2
      Top             =   660
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   688
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarTitleBackColor=   14737632
      Format          =   16384001
      CurrentDate     =   38002
   End
   Begin MSComCtl2.DTPicker fecha 
      Height          =   345
      Left            =   840
      TabIndex        =   9
      Top             =   1740
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      Format          =   16384001
      CurrentDate     =   40679
   End
   Begin MSDataListLib.DataCombo cmbUsuario 
      Height          =   315
      Left            =   840
      TabIndex        =   10
      Top             =   1350
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      Style           =   2
      ListField       =   ""
      BoundColumn     =   ""
      Text            =   ""
      Object.DataMember      =   ""
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
   Begin MSComCtl2.DTPicker hora 
      Height          =   345
      Left            =   3270
      TabIndex        =   13
      Top             =   1740
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   609
      _Version        =   393216
      Format          =   16384002
      CurrentDate     =   40679
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Hora"
      Height          =   195
      Index           =   1
      Left            =   2790
      TabIndex        =   14
      Top             =   1830
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
      Height          =   195
      Index           =   6
      Left            =   120
      TabIndex        =   12
      Top             =   1830
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Usuario"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1410
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Motivo"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2250
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   1
      X1              =   30
      X2              =   6105
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "FACTURA"
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
      Height          =   405
      Left            =   30
      TabIndex        =   5
      Top             =   30
      Width           =   6060
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3870
      TabIndex        =   4
      Top             =   720
      Width           =   570
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Numero"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   3
      Top             =   720
      Width           =   720
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      Index           =   0
      X1              =   15
      X2              =   6090
      Y1              =   1170
      Y2              =   1170
   End
End
Attribute VB_Name = "frmDocumento_Anular"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pk As Long
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If validar Then
        Dim oD As New clsDocumentos
        If oD.Carga(pk) Then
            If oD.getTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.factura Then
                MsgBox "Al anular la factura, sus albaranes se quedarán pendientes de facturar.", vbInformation, App.Title
            End If
        End If
        Dim oDA As New clsDocumentos_anulados
        With oDA
            .setDOCUMENTO_ID = pk
            .setUSUARIO_ID = cmbUsuario.BoundText
            .setFECHA = Format(fecha, "yyyy-mm-dd")
            .setHORA = Format(hora, "hh:mm:ss")
            .setMOTIVO = txtdatos(1)
            .Insertar
        End With
        Set oDA = Nothing
        MsgBox "Anulado correctamente.", vbInformation, App.Title
        Unload Me
    End If

   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmDocumento_Anular"
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    log (Me.Name)
    cargar_botones Me
    Cargar_Combo cmbUsuario, New ClsUsuario
    If pk > 0 Then
        cargar_datos
    End If

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmDocumento_Anular"
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &HC0E0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub
Private Sub txtdatos_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
       SendKeys "{Tab}", True
       KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtDatos_LostFocus(Index As Integer)
    txtdatos(Index).BackColor = &HFFFFFF
End Sub

Public Function validar() As Boolean
    If Trim(txtdatos(1)) = "" Then
        MsgBox "Introduzca el motivo de la anulación.", vbCritical, "Error"
        validar = False
        txtdatos(1).SetFocus
        Exit Function
    End If
    validar = True
End Function


Private Sub cargar_datos()
    Dim oDOCUMENTO As New clsDocumentos
    With oDOCUMENTO
        If .Carga(pk) = True Then
            Dim oDeco As New clsDecodificadora
            oDeco.Carga_valor DECODIFICADORA.D_TIPOS_FACTURACION, .getTIPO_DOCUMENTO_ID
            lbltitulo = "ANULAR " & oDeco.getDESCRIPCION
            Me.Caption = lbltitulo
            txtdatos(0) = .getNUMERO
            txtfecha = .getFECHA
            If .getANULADO = 0 Then
                cmbUsuario.BoundText = USUARIO.getID_EMPLEADO
                fecha = Date
                hora = Time
            Else
                Dim oDA As New clsDocumentos_anulados
                If oDA.Carga(pk) Then
                    cmbUsuario.BoundText = oDA.getUSUARIO_ID
                    fecha = oDA.getFECHA
                    hora = oDA.getHORA
                    txtdatos(1) = oDA.getMOTIVO
                End If
            End If
        End If
    End With
    Set oDOCUMENTO = Nothing
End Sub
