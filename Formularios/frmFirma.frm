VERSION 5.00
Begin VB.Form frmFirma 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Captura de firma electrónica"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   14445
   Icon            =   "frmFirma.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9645
   ScaleWidth      =   14445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Limpiar"
      Height          =   870
      Left            =   12960
      Picture         =   "frmFirma.frx":164A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   90
      Width           =   1410
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Grabar firma"
      Height          =   1050
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7410
      Width           =   1410
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir sin firma"
      Height          =   1050
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8550
      Width           =   1410
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      DrawWidth       =   3
      Height          =   9510
      Left            =   45
      ScaleHeight     =   9450
      ScaleWidth      =   12735
      TabIndex        =   0
      Top             =   45
      Width           =   12795
   End
End
Attribute VB_Name = "frmFirma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public firmas As String
Dim X1 As Long      'This Is The X Position Of The Last Line Drawn
Dim y1 As Long      'This Is The Y Position Of The Last Line Drawn
Dim X2 As Long      'This Is The Start Mark Of The Box or Circle
Dim y2 As Long      'This Is The Satrt Mark Of The Box or Circle

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    On Error Resume Next
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\FIRMAS"
    MkDir ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\FIRMAS\PRUEBA"
    On Error GoTo Command2_Click_Error
    Dim Conversor As Class1
    Set Conversor = New Class1
    Dim oDoc As New clsDocumentacion
    Dim fichero As String
    fichero = DIRECTORIO_TEMPORAL & "\" & gmuestra & ".jpg"
    Conversor.GrabarJpg Picture1.Image, fichero, CByte(70)
    ' Subir a BD
'    oDoc.SubirFirma gmuestra, fichero, gmuestra & ".jpg"
    oDoc.SubirFirma gmuestra, fichero, CStr(gmuestra)
    Set Conversor = Nothing
    Dim oMuestra As New clsMuestra
    oMuestra.informar_firma (gmuestra)
    Unload Me
   On Error GoTo 0
   Exit Sub

Command2_Click_Error:

    error_grave ("Error " & Err.Number & " (" & Err.Description & ") in procedure Command2_Click of Formulario frmFirma")
End Sub

Private Sub Command3_Click()
    Picture1.Cls
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra gmuestra
    If Trim(oMuestra.getFIRMA) <> "" Then
        Dim oDoc As New clsDocumentacion
        Dim firma As String
        firma = oDoc.CargarFirma(CLng(Trim(Replace(oMuestra.getFIRMA, ".jpg", ""))), False)
        If firma <> "" Then
            Picture1.Picture = LoadPicture(firma)
        End If
    End If
End Sub

Private Sub Form_Resize()
'    Picture1.Left = 1
'    Picture1.Top = 1
'    Picture1.Width = Me.Width - Command3.Width - 300
'    Picture1.Height = Me.Height - 300
    
'    Command3.Left = Me.Width - Command3.Width - 200
'    cmdok.Left = Me.Width - cmdok.Width - 200
'    cmdcancel.Left = Me.Width - cmdcancel.Width - 200
    
End Sub

Private Sub Picture1_Click()
'    If Button = vbLeftButton Then
        'Draw Continuous Line...
'        Picture1.Line (X1, y1)-(x + 1, y + 1), vbBlack
        Picture1.Line (X2, y2)-(X2, y2), vbBlack
'        Picture1.Point X2, y2
'        X1 = x: y1 = y
'    End If

End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    X1 = x
    y1 = y
    X2 = x
    y2 = y
End Sub
Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbLeftButton Then
        'Draw Continuous Line...
        Picture1.Line (X1, y1)-(x, y), vbBlack
        X1 = x: y1 = y
    End If
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    X1 = 0
    y1 = 0
End Sub
