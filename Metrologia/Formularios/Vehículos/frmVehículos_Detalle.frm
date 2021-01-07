VERSION 5.00
Begin VB.Form frmVehículos_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Agenda"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9300
   Icon            =   "frmVehículos_Detalle.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   9300
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   855
      Left            =   6750
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   2610
      Width           =   1215
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   855
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2610
      Width           =   1245
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   " Datos "
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
      Left            =   30
      TabIndex        =   4
      Top             =   570
      Width           =   9165
      Begin VB.TextBox txtDatos 
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
         Width           =   2175
      End
      Begin VB.TextBox txtDatos 
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
         Width           =   2175
      End
      Begin VB.TextBox txtDatos 
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
         Width           =   2175
      End
      Begin VB.TextBox txtDatos 
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
         Caption         =   "Remolque"
         Height          =   195
         Index           =   8
         Left            =   210
         TabIndex        =   8
         Top             =   1575
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "N.I.F."
         Height          =   195
         Index           =   15
         Left            =   210
         TabIndex        =   7
         Top             =   1170
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Matrícula"
         Height          =   195
         Index           =   11
         Left            =   210
         TabIndex        =   6
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
         TabIndex        =   5
         Top             =   375
         Width           =   555
      End
   End
   Begin VB.Label lbltitulo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenimiento de Vehículos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   11
      Top             =   90
      Width           =   4050
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   8700
      Picture         =   "frmVehículos_Detalle.frx":08CA
      Top             =   0
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   525
      Left            =   0
      Top             =   0
      Width           =   9345
   End
End
Attribute VB_Name = "frmVehículos_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pk As Long
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    If pk > 0 Then
        Modificar
    Else
        Insertar
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    If pk > 0 Then
        consulta
    End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Set frmVehículos_Detalle = Nothing
End Sub
Private Sub txtDatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub

Private Sub txtdatos_Keyup(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
     Case 40 ' Abajo
       If Index = 4 Then
        txtDatos(1).SetFocus
       Else
        SendKeys "{Tab}", True
       End If
       KeyAscii = 0 ' Para evitar el "bip" del sistema
     Case 38
       If Index = 1 Then
        txtDatos(4).SetFocus
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
    txtDatos(Index).BackColor = &HFFFFFF
End Sub

Public Sub borrar_campos()
    Dim i As Integer
    For i = 1 To 4
        txtDatos(i) = ""
    Next
    txtDatos(1).SetFocus
End Sub

Public Sub bloquear_campos()
    Dim i As Integer
    For i = 1 To 4
        txtDatos(i).Locked = True
    Next
End Sub

Public Sub Insertar()
    If valida_datos = False Then
        Exit Sub
    End If
    pregunta = "Va a dar de alta el Vehículo. ¿Esta seguro?"
    If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim aux As Long
        Set ov = mover_datos
        aux = pk
        pk = ov.Insertar
        If pk > 0 Then
            MsgBox "El vehículo se ha insertado correctamente.", vbInformation, App.Title
        End If
        Unload Me
    End If
End Sub

Public Sub Modificar()
    If valida_datos() = False Then
        Exit Sub
    End If
    Dim pos As Integer
    Dim agenda As Integer
    pregunta = "Va a modificar los datos del vehículo. ¿Esta seguro?"
    If MsgBox(pregunta, vbYesNo + vbQuestion, App.Title) = vbYes Then
        Set ov = mover_datos
        If ov.Modificar(pk) = True Then
            MsgBox "El vehículo se ha modificado correctamente.", vbInformation, App.Title
            Unload Me
        End If
        Set ov = Nothing
    End If
End Sub
Public Function valida_datos() As Boolean
    valida_datos = True
    If txtDatos(1) = "" Then
        MsgBox "El nombre del vehículo no puede estar en blanco.", vbCritical, "Error"
        txtDatos(1).SetFocus
        valida_datos = False
        Exit Function
    End If
End Function

Public Sub consulta()
    On Error GoTo fallo
    Dim ov As New clsVehiculos
    lbltitulo.Caption = "Modificacion de Vehículo"
    ov.Carga (pk)
    With ov
        txtDatos(1) = .getNOMBRE
        txtDatos(2) = .getMATRICULA
        txtDatos(3) = .getNIF
        txtDatos(4) = .getREMOLQUE
    End With
    Set ov = Nothing
    Exit Sub
fallo:
    MsgBox "Error al consultar los datos del agenda.", vbCritical, Err.Description
End Sub
Public Function mover_datos() As clsVehiculos
    On Error GoTo fallo
    Dim ov As New clsVehiculos
    With ov
        .setNOMBRE = UCase(txtDatos(1))
        .setMATRICULA = txtDatos(2)
        .setNIF = txtDatos(3)
        .setREMOLQUE = txtDatos(4)
    End With
    Set mover_datos = ov
    Set oagenda = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos del vehiculo.", vbCritical, Err.Description
End Function
