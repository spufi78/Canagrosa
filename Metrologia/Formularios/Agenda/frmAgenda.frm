VERSION 5.00
Begin VB.Form frmAgenda 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Agenda"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmAgenda.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   855
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2460
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   855
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2460
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos de Contacto"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   60
      TabIndex        =   4
      Top             =   375
      Width           =   9165
      Begin VB.TextBox txtdatos 
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
         Height          =   330
         Index           =   4
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1515
         Width           =   1755
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   330
         Index           =   3
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1110
         Width           =   1755
      End
      Begin VB.TextBox txtdatos 
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
         Height          =   330
         Index           =   2
         Left            =   1350
         MaxLength       =   30
         TabIndex        =   1
         Top             =   720
         Width           =   1755
      End
      Begin VB.TextBox txtdatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1350
         MaxLength       =   75
         TabIndex        =   0
         Top             =   330
         Width           =   7710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fax"
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   9
         Top             =   1575
         Width           =   255
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Movil"
         Height          =   195
         Index           =   15
         Left            =   210
         TabIndex        =   7
         Top             =   1170
         Width           =   375
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Telefono"
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   6
         Top             =   780
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   0
         Left            =   225
         TabIndex        =   5
         Top             =   375
         Width           =   555
      End
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nuevo Contacto"
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
      Height          =   315
      Left            =   60
      TabIndex        =   8
      Top             =   30
      Width           =   9180
   End
End
Attribute VB_Name = "frmAgenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    If gAgenda > 0 Then
        modificar_agenda
    Else
        insertar_agenda
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    If gAgenda > 0 Then
        consulta_agenda
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmAgenda = Nothing
End Sub
Private Sub txtDatos_GotFocus(Index As Integer)
    txtdatos(Index).BackColor = &H80C0FF
    txtdatos(Index).SelStart = 0
    txtdatos(Index).SelLength = Len(txtdatos(Index))
End Sub

Private Sub txtdatos_Keyup(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40 ' Abajo
       If Index = 4 Then
        txtdatos(1).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 1 Then
        txtdatos(4).SetFocus
       Else
        SendKeys "+{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 27
        cmdcancel_Click
     Case 121 ' F10
        cmdok_Click
    End Select
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

Public Sub borrar_campos()
    Dim i As Integer
    For i = 1 To 4
        txtdatos(i) = ""
    Next
    txtdatos(1).SetFocus
End Sub

Public Sub bloquear_campos()
    Dim i As Integer
    For i = 1 To 4
        txtdatos(i).Locked = True
    Next
End Sub

Public Sub insertar_agenda()
    If valida_datos = False Then
        Exit Sub
    End If
    pregunta = "Va a dar de alta el Contacto. ¿Esta seguro?"
    If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim aux As Long
        Set oagenda = mover_datos
        aux = gAgenda
        gAgenda = oagenda.Insertar
        If gAgenda > 0 Then
            MsgBox "El contacto se ha insertado correctamente.", vbInformation, App.Title
        End If
        If aux = -1 Then
            Unload Me
            Exit Sub
        End If
        borrar_campos
        Set oagenda = Nothing
    End If
End Sub

Public Sub modificar_agenda()
    If valida_datos() = False Then
        Exit Sub
    End If
    Dim pos As Integer
    Dim agenda As Integer
    pregunta = "Va a modificar los datos del contacto. ¿Esta seguro?"
    If MsgBox(pregunta, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Set oagenda = mover_datos
        oagenda.setID_AGENDA = gAgenda
        If oagenda.Modificar = True Then
            MsgBox "El contacto se ha modificado correctamente.", vbInformation, App.Title
            Unload Me
        End If
        Set oagenda = Nothing
    End If
End Sub
Public Function valida_datos() As Boolean
    valida_datos = True
    If txtdatos(1) = "" Then
        MsgBox "El nombre del contacto no puede estar en blanco.", vbCritical, "Error"
        txtdatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Public Sub consulta_agenda()
    On Error GoTo fallo
    Dim oagenda As New clsAgenda
    lbltitulo.Caption = "Modificacion de contacto"
    oagenda.Carga (gAgenda)
    With oagenda
        txtdatos(1) = .getNOMBRE
        txtdatos(2) = .getTELEFONO
        txtdatos(3) = .getMOVIL
        txtdatos(4) = .getFAX
    End With
    Set oagenda = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del agenda.", vbCritical, Err.Description
End Sub
Public Function mover_datos() As clsAgenda
    On Error GoTo fallo
    Dim oagenda As New clsAgenda
    With oagenda
        .setNOMBRE = UCase(txtdatos(1))
        .setTELEFONO = txtdatos(2)
        .setMOVIL = txtdatos(3)
        .setFAX = txtdatos(4)
    End With
    Set mover_datos = oagenda
    Set oagenda = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del contacto.", vbCritical, Err.Description
End Function
