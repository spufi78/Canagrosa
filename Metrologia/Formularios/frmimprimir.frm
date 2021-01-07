VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmimprimir 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Imprimir"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkCopia 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Incluir ES COPIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   12
      Top             =   1935
      Value           =   1  'Checked
      Width           =   4425
   End
   Begin VB.CheckBox chkPrevisualizar 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Imprimir directamente en la impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   11
      Top             =   2355
      Value           =   1  'Checked
      Width           =   4425
   End
   Begin VB.TextBox txtdatos 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Index           =   1
      Left            =   1710
      TabIndex        =   4
      Top             =   2955
      Width           =   945
   End
   Begin VB.CheckBox chkLogo 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Imprimir con imagen de fondo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1470
      Value           =   1  'Checked
      Width           =   4425
   End
   Begin VB.CommandButton cmdBoton 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Correo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   2
      Left            =   1470
      Picture         =   "frmimprimir.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3990
      Width           =   1335
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
      TabIndex        =   5
      Top             =   660
      Width           =   2235
   End
   Begin VB.CommandButton cmdBoton 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   1
      Left            =   4800
      Picture         =   "frmimprimir.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3990
      Width           =   1305
   End
   Begin VB.CommandButton cmdBoton 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Index           =   0
      Left            =   90
      Picture         =   "frmimprimir.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3990
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker txtfecha 
      Height          =   390
      Left            =   4530
      TabIndex        =   6
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
      Format          =   51314689
      CurrentDate     =   38002
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Numero Copias"
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
      Left            =   240
      TabIndex        =   10
      Top             =   3045
      Width           =   1410
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
      BackColor       =   &H00C0E0FF&
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   7
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
Attribute VB_Name = "frmimprimir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public pk As Long
Private Sub cmdBoton_Click(Index As Integer)
    On Error GoTo fallo
    If validar Then
        Dim oDOCUMENTO As New clsDocumentos
        Select Case Index
        Case 0 ' Imprimimos
            oDOCUMENTO.imprimir pk, chkPrevisualizar.Value, chkLogo.Value, txtdatos(1), , chkCopia.Value
        Case 1 ' Salir
            Unload Me
        Case 2 ' correo
            oDOCUMENTO.Correo pk, False, chkLogo.Value, txtdatos(1), chkCopia.Value
        End Select
        Set oDOCUMENTO = Nothing
'        Unload Me
    End If
    Exit Sub
fallo:
    MsgBox "Error al imprimir el documento : " & Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Load()
   On Error GoTo Form_Load_Error

    log (Me.Name)
    Dim oDOCUMENTO As New clsDocumentos
    With oDOCUMENTO
        If .Carga(pk) = True Then
            chkLogo.Value = ReadINI(App.Path & "\config.ini", "parametros", "Empresa")
            If .getTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.ALBARAN Then
                lbltitulo.Caption = "ALBARAN"
                txtdatos(1) = ReadINI(App.Path & "\config.ini", "parametros", "Copias_albaran")
            Else
                lbltitulo.Caption = "FACTURA"
                Dim oObra As New clsObras
                oObra.Carga oDOCUMENTO.getOBRA_ID
                Dim ocliente As New clsCliente
                ocliente.CargaCliente oObra.getCLIENTE_ID
                txtdatos(1) = ocliente.getCOPIAS_FACTURA
                Set ocliente = Nothing
                Set oObra = Nothing
'                txtdatos(1) = ReadINI(App.Path & "\config.ini", "parametros", "Copias_facturas")
            End If
            chkPrevisualizar.Value = ReadINI(App.Path & "\config.ini", "parametros", "Previsualizar")
            txtdatos(0) = .getNUMERO
            txtfecha = .getFECHA
        End If
    End With

   On Error GoTo 0
   Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Form_Load of Formulario frmimprimir"
End Sub

Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &HC0FFFF
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
    If IsNumeric(txtdatos(1)) = False Then
        MsgBox "Introduzca el número de copias correctamente.", vbCritical, "Error"
        validar = False
        txtdatos(1).SetFocus
        Exit Function
    End If
    validar = True
End Function

