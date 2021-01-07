VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmMEN_Crear 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Creación de mensajes"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   11850
   Icon            =   "frmMEN_Crear.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   11850
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   9585
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5085
      Width           =   1050
   End
   Begin VB.CommandButton cmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10710
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5085
      Width           =   1050
   End
   Begin VB.CommandButton cmdNinguno 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ninguno"
      Height          =   735
      Left            =   10935
      Picture         =   "frmMEN_Crear.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2520
      Width           =   870
   End
   Begin VB.CommandButton cmdtodos 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Todos"
      Height          =   735
      Left            =   10935
      Picture         =   "frmMEN_Crear.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1710
      Width           =   870
   End
   Begin VB.TextBox txttexto 
      Appearance      =   0  'Flat
      Height          =   3255
      Index           =   0
      Left            =   90
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1755
      Width           =   6000
   End
   Begin VB.TextBox txttexto 
      Appearance      =   0  'Flat
      Height          =   330
      Index           =   1
      Left            =   90
      TabIndex        =   0
      Top             =   1350
      Width           =   6000
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Detalle del mensaje"
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
      Height          =   1005
      Left            =   90
      TabIndex        =   7
      Top             =   45
      Width           =   6000
      Begin VB.TextBox txttexto 
         Appearance      =   0  'Flat
         Height          =   330
         Index           =   2
         Left            =   765
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   225
         Width           =   5145
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   765
         TabIndex        =   14
         Top             =   585
         Width           =   1545
         _ExtentX        =   2725
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
         CalendarTitleBackColor=   14737632
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3195
         TabIndex        =   15
         Top             =   585
         Width           =   1545
         _ExtentX        =   2725
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
         CalendarTitleBackColor=   14737632
         Format          =   16515073
         CurrentDate     =   38002
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "De"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   270
         Width           =   210
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Válido"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   10
         Top             =   630
         Width           =   435
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   3
         Left            =   2655
         TabIndex        =   9
         Top             =   630
         Width           =   390
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4650
      Left            =   6210
      TabIndex        =   2
      Top             =   360
      Width           =   4620
      _ExtentX        =   8149
      _ExtentY        =   8202
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Listado de usuarios de destino"
      Height          =   195
      Left            =   6255
      TabIndex        =   13
      Top             =   135
      Width           =   4605
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "Datos del Mensaje"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   12
      Top             =   1125
      Width           =   6000
   End
End
Attribute VB_Name = "frmMEN_Crear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNinguno_Click()
    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            lista.ListItems(i).Checked = False
        Next
    End If

End Sub

Private Sub cmdok_Click()
   On Error GoTo cmdok_Click_Error

    If txttexto(1) = "" Then
        MsgBox "Introduzca el asunto del mensaje", vbExclamation, App.Title
        Exit Sub
    End If
    Dim oMensaje As New clsMensajes
    Dim men As Integer
    With oMensaje
        .setASUNTO = txttexto(1)
        .setTEXTO = txttexto(0)
        .setEMPLEADO_ID = usuario.getID_EMPLEADO
        .setFECHA_INICIO = Format(fdesde, "yyyy-mm-dd")
        .setHORA_INICIO = "00:00:00"
        .setFECHA_FIN = Format(fhasta, "yyyy-mm-dd")
        .setHORA_FIN = "00:00:00"
        men = .Insertar
        If men > 0 Then
            Dim omu As New clsMensajes_usuarios
            Dim i As Integer
            For i = 1 To lista.ListItems.Count
                If lista.ListItems(i).Checked = True Then
                    omu.setEMPLEADO_ID = lista.ListItems(i).SubItems(1)
                    omu.setMENSAJE_ID = men
                    omu.Insertar
                End If
            Next
        End If
    End With
    MsgBox "Mensaje generado correctamente.", vbInformation, App.Title
    Unload Me
    Exit Sub
   On Error GoTo 0
   Exit Sub

cmdok_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdok_Click of Formulario frmMEN_Crear"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdtodos_Click()
    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            lista.ListItems(i).Checked = True
        Next
    End If
End Sub

Private Sub Form_Load()
    cabecera
    cargar_botones Me
    cargar_usuarios
    txttexto(2) = usuario.getUSUARIO
    fdesde = Date
    fhasta = Date
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Usuario", 4350, lvwColumnLeft
        .Add , , "ID", 1, lvwColumnLeft
    End With
End Sub


Public Sub cargar_usuarios()
    Dim oempleado As New clsUsuarios
    Dim RS As ADODB.RecordSet
    Set RS = oempleado.Listado
    If RS.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , RS("APELLIDOS") & ", " & RS("NOMBRE") & " (" & RS("USUARIO") & ")")
              .SubItems(1) = RS("ID_EMPLEADO")
            End With
            RS.MoveNext
        Loop Until RS.EOF
    End If
    Set oempleado = Nothing
    Set RS = Nothing
End Sub
