VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmBsm 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Baños sin muestra"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   Icon            =   "frmBsm.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   4920
      Picture         =   "frmBsm.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3390
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   6045
      Picture         =   "frmBsm.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3390
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   2910
      Left            =   120
      TabIndex        =   3
      Top             =   420
      Width           =   6975
      Begin VB.TextBox txtDatos 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1620
         Index           =   0
         Left            =   1770
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1140
         Width           =   5115
      End
      Begin MSDataListLib.DataCombo cmbbanos 
         Height          =   360
         Left            =   1770
         TabIndex        =   0
         Top             =   330
         Width           =   5115
         _ExtentX        =   9022
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   390
         Left            =   1770
         TabIndex        =   1
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   688
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   50069505
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   13
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   675
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   420
         Width           =   585
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Observaciones"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   90
         TabIndex        =   4
         Top             =   1830
         Width           =   1635
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nuevo Baño sin Muestra"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Index           =   2
      Left            =   120
      TabIndex        =   6
      Top             =   60
      Width           =   6975
   End
End
Attribute VB_Name = "frmBsm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdok_Click()
    If validar = True Then
      Dim obsm As New clsBsm
      With obsm
          .setBANO_ID = cmbbanos.BoundText
          .setFECHA = Format(fecha, "yyyy-mm-dd")
          .setOBSERVACIONES = txtDatos(0)
      End With
      If gbsm = 0 Then
        If MsgBox("Va a introducir un nuevo baño sin muestra. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            If obsm.Insertar = False Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
      Else
        If MsgBox("Va a modificar el baño sin muestra. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            If obsm.Modificar(gbsm) = False Then
                Exit Sub
            End If
        Else
            Exit Sub
        End If
      End If
      If gbsm = 0 Then
          MsgBox "El baño se ha introducido correctamente.", vbOKOnly, App.Title
      Else
          MsgBox "El baño se ha modificado correctamente.", vbOKOnly, App.Title
      End If
      Unload Me
    End If
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdcancel_Click
    End If
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Call cargar_banos
    If gbsm <> 0 Then
        Label1(2) = "Modificación de Baño sin muestra"
        Label1(2).BackColor = &H80C0FF
        cargar_Bsm
    Else
        fecha.Value = Date
    End If
End Sub

Public Sub cargar_banos()
    Dim oanom As New clsBanos
    Set cmbbanos.RowSource = oanom.Listado
    cmbbanos.ListField = "bano"
    cmbbanos.BoundColumn = "id_bano"
End Sub
Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
End Sub
Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = vbWhite
End Sub
Public Sub cargar_Bsm()
    Dim oanom As New clsBsm
    Dim obano As New clsBanos
    With oanom
     .Cargar (CLng(gbsm))
     txtDatos(0) = .getOBSERVACIONES
     fecha.Value = .getFECHA
     obano.cargar_bano (.getBANO_ID)
     cmbbanos.Text = obano.getNOMBRE
    End With
    Set oanom = Nothing
    Set obano = Nothing
End Sub
Public Function validar() As Boolean
    validar = True
    If cmbbanos.Text = "" Then
        MsgBox "Debe seleccionar un baño", vbCritical, App.Title
        validar = False
        Exit Function
    End If
    If Trim(txtDatos(0)) = "" Then
        MsgBox "Debe darle observaciones al baño.", vbInformation, App.Title
        validar = False
        Exit Function
    End If
End Function

