VERSION 5.00
Begin VB.Form frmSelMsg 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleccione Mensaje de Outlook"
   ClientHeight    =   3495
   ClientLeft      =   5310
   ClientTop       =   4230
   ClientWidth     =   8700
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSelMsg.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3495
   ScaleWidth      =   8700
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optHtml 
      BackColor       =   &H00C0C0C0&
      Caption         =   "HTML (Sin adjuntos ni imágenes, solo Texto del mensaje)"
      Enabled         =   0   'False
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
      Left            =   150
      TabIndex        =   10
      Top             =   2955
      Visible         =   0   'False
      Width           =   5715
   End
   Begin VB.OptionButton optMSG 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mensaje de Outlook (Guarda los archivos adjuntos)"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   150
      TabIndex        =   9
      Top             =   2565
      Value           =   -1  'True
      Visible         =   0   'False
      Width           =   5805
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancelar"
      Height          =   870
      Left            =   7515
      Picture         =   "frmSelMsg.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2505
      Width           =   1050
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   6390
      Picture         =   "frmSelMsg.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2505
      Width           =   1050
   End
   Begin VB.TextBox txtAsunto 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   2040
      Width           =   6765
   End
   Begin VB.TextBox txtFecha 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1680
      Width           =   6765
   End
   Begin VB.TextBox txtEnviado 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Height          =   330
      Left            =   1800
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1320
      Width           =   6765
   End
   Begin VB.Timer tm 
      Interval        =   1000
      Left            =   5985
      Top             =   2835
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Abra outlook si lo tiene cerrado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   12
      Top             =   375
      Width           =   2190
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   8055
      Picture         =   "frmSelMsg.frx":149E
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Seleccione en el Outlook el correo que desea vincular"
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
      Left            =   180
      TabIndex        =   11
      Top             =   75
      Width           =   5640
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Pulse [Aceptar] si éste es el seleccionado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   300
      Left            =   1710
      TabIndex        =   8
      Top             =   855
      Width           =   5010
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Asunto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   2100
      Width           =   720
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha Recep."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   240
      TabIndex        =   2
      Top             =   1740
      Width           =   1470
   End
   Begin VB.Label lblCap 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Enviado Por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   240
      TabIndex        =   0
      Top             =   1380
      Width           =   1290
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   8790
   End
End
Attribute VB_Name = "frmSelMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Event MensajeSeleccionado(ByRef Cancel As Boolean)

Private oO As New OUTLOOK.Application
Public correo As MailItem, oM As MailItem
Public Seleccionado As Boolean
Public Aceptar As Boolean

Private blnExisteImpresoraPDF As Boolean

Private Function comprobar_correo_seleccionado()
On Error GoTo error_comprobar_correo_seleccionado

    Dim oCol As OUTLOOK.Selection
    Dim strNombre As String, strDir As String
    Dim randomCont As Integer
    
    If oO.ActiveExplorer Is Nothing Then
        MsgBox "Para que pueda adjuntar correo, necestita tener MS Outlook abierto.", vbInformation, "Adjuntar Correo"
        cmdcancel_Click
        Exit Function
    End If
    
    Set oCol = oO.ActiveExplorer.Selection
    
    If oCol.Count = 0 Then
        lblMsg.Caption = ""
        txtAsunto.Text = ""
        txtEnviado.Text = ""
        txtFecha.Text = ""
        Seleccionado = False
    End If
    
    
    If oCol.Count > 1 Then
        lblMsg.Caption = "Debe señalar solo un mensaje, por favor."
        txtAsunto.Text = ""
        txtEnviado.Text = ""
        txtFecha.Text = ""
        Seleccionado = False
    Else
        Set oM = oCol.Item(1)
        lblMsg.Caption = "Pulse [Aceptar] si éste es el seleccionado"
        txtAsunto.Text = oM.Subject
        txtEnviado.Text = oM.SenderName & " (" & oM.SenderEmailAddress & ")"
        txtFecha.Text = oM.ReceivedTime
        Seleccionado = True
    End If

Exit Function
error_comprobar_correo_seleccionado:

    MsgBox "Para que pueda adjuntar correo, necestita permitir acceso a MS Outlook", vbInformation, "Adjuntar Correo"
    cmdcancel_Click
    Exit Function

End Function

Private Sub comprobar_impresora_pdf()
    
    On Error GoTo Error_No_ODF
    
    Dim oPDF As Object
    
    
    Set oPDF = CreateObject("PDFCreator.clsPDFCreator")
    
    
    blnExisteImpresoraPDF = True
    
    Set oPDF = Nothing

Exit Sub
Error_No_ODF:
    blnExisteImpresoraPDF = False
End Sub

Private Sub cmdcancel_Click()
    
    tm.Enabled = False
    Aceptar = False
    Me.Hide

End Sub

Private Sub cmdok_Click()

    If Not Seleccionado Then
        MsgBox "No ha seleccionado ningún mensaje de correo en Outlook", vbInformation, "Adjuntar Mensaje de Correo"
        Exit Sub
    End If
    
    Set correo = oM
    tm.Enabled = False
    Aceptar = True
    Me.Hide


End Sub



Private Sub Form_Load()
    blnExisteImpresoraPDF = True
'    comprobar_impresora_pdf

'    If Not blnExisteImpresoraPDF Then
'        MsgBox "Necesita tener la impresora virtual PDFCreator instalada"
        cmdok.Enabled = blnExisteImpresoraPDF
'    End If
    

    Seleccionado = False
    lblMsg.Caption = ""
    
    tm.Enabled = True
End Sub

Private Sub tm_Timer()

    comprobar_correo_seleccionado

End Sub


