VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#39.0#0"; "miCombo.ocx"
Begin VB.Form frmEquipoAccesorios 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Accesorios de Equipos"
   ClientHeight    =   2850
   ClientLeft      =   3645
   ClientTop       =   3585
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2850
   ScaleWidth      =   10470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   630
      Width           =   10455
      Begin VB.CheckBox chkDesmontarAccesorio 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Accesorio Desmontado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         TabIndex        =   11
         Top             =   840
         Width           =   2415
      End
      Begin pryCombo.miCombo cmbEquipo 
         Height          =   345
         Left            =   2340
         TabIndex        =   8
         Top             =   90
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   609
      End
      Begin VB.CommandButton cmdcancel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "ESC-Salir"
         Height          =   870
         Left            =   9540
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1290
         Width           =   870
      End
      Begin MSComCtl2.DTPicker txtFechaAlta 
         Height          =   315
         Left            =   2340
         TabIndex        =   3
         Top             =   450
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyy HH:mm"
         Format          =   60293123
         CurrentDate     =   40273
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker txtFechaBaja 
         Height          =   315
         Left            =   2340
         TabIndex        =   9
         Top             =   810
         Visible         =   0   'False
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   556
         _Version        =   393216
         CustomFormat    =   "dd/MM/yyy HH:mm"
         Format          =   60293123
         CurrentDate     =   40273
         MinDate         =   2
      End
      Begin VB.CommandButton cmdok 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aceptar"
         Height          =   870
         Left            =   8610
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1290
         Width           =   870
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Equipo Accesorio"
         Height          =   195
         Left            =   60
         TabIndex        =   5
         Top             =   120
         Width           =   1275
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha/Hora Montaje/Alta"
         Height          =   195
         Left            =   60
         TabIndex        =   4
         Top             =   540
         Width           =   2025
      End
      Begin VB.Label lblFechaBaja 
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha/Hora Desmontaje/Baja"
         Height          =   195
         Left            =   60
         TabIndex        =   10
         Top             =   900
         Visible         =   0   'False
         Width           =   2175
      End
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   9900
      Picture         =   "frmEquipoAccesorios.frx":0000
      Top             =   90
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Ventana de gestión de Accesorios de Equipo"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   7
      Top             =   300
      Width           =   3195
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Accesorios de Equipo"
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
      TabIndex        =   6
      Top             =   0
      Width           =   2310
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   10470
   End
End
Attribute VB_Name = "frmEquipoAccesorios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarobjEquipo As clsEquipos
Private mvarobjAccesorio As New clsEquipoAccesorios
Public PK As Long



Private Function comprobar_datos() As Boolean

    comprobar_datos = False
    Dim strCad As String
    
    If cmbEquipo.getPK_SALIDA = 0 Then
        MsgBox "Debe indicar el Equipo que Consta como Accesorio", vbInformation, "Accesorio de Equipo"
        Exit Function
    End If


    If chkDesmontarAccesorio.value = vbChecked Then
        If txtFechaAlta.value >= txtFechaBaja.value Then
            MsgBox "La fecha de Baja del Accesorio debe ser posterior a la fecha de Alta del Accesorio", vbInformation, "Accesorio de Equipo"
            Exit Function
        End If
    End If

    comprobar_datos = True

End Function

Private Sub guardar_datos()

If PK = 0 Then
    mvarobjAccesorio.Insertar
Else
    mvarobjAccesorio.Modificar
End If

End Sub

Private Sub recoger_datos()


With mvarobjAccesorio
    
    .setID_ACCESORIO = cmbEquipo.getPK_SALIDA
    .setEQUIPO_ID = mvarobjEquipo.getID_EQUIPO
    .setCUSERID = usuario.getID_EMPLEADO
    .setFECHA_ALTA = txtFechaAlta.value
    
    
    If chkDesmontarAccesorio.value = vbChecked Then
        .setEN_USO = 0
        .setFECHA_BAJA = txtFechaBaja.value
    Else
        .setEN_USO = 1
        .setFECHA_BAJA = CDate("01/01/1900 00:00")
    End If
    
    
    
End With
End Sub

Private Sub cmdcancel_Click()

    Unload Me
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjEquipo = Nothing

End Sub




Private Sub presentar_datos()


txtFechaAlta.value = Now

If PK = 0 Then Exit Sub

mvarobjAccesorio.Carga PK, mvarobjEquipo.getID_EQUIPO

With mvarobjAccesorio
    cmbEquipo.MostrarElemento PK
    txtFechaAlta.value = mvarobjAccesorio.getFECHA_ALTA
    If .getEN_USO = 1 Then
        chkDesmontarAccesorio.value = vbUnchecked
    Else
        chkDesmontarAccesorio.value = vbChecked
        txtFechaBaja.value = .getFECHA_BAJA
    End If
    chkDesmontarAccesorio_Click
End With

End Sub

Private Sub cmdok_Click()

    If Not comprobar_datos Then Exit Sub
    
    recoger_datos
    
    guardar_datos
    
    Unload Me
    
End Sub

Private Sub cargar_combos()

    If PK = 0 Then
        ' cuando es dar de alta
        
        ' JONATHAN.10.08.2010 -> SE ESTIPULA QUE LOS ACCESORIOS PUEDEN ESTAR EN MÁS DE UN EQUPO COMO ACCESORIO.
        'llenar_combo cmbEquipo, New clsEquipos, 0, frmEquipoEdicion, " AND ID_EQUIPO NOT IN (SELECT ID_ACCESORIO FROM EQUIPOS_ACCESORIOS) AND ID_EQUIPO <> " & CStr(mvarobjEquipo.getID_EQUIPO)
        llenar_combo cmbEquipo, New clsEquipos, 0, frmEquipoEdicion, " AND ES_ACCESORIO=1 AND ID_EQUIPO <> " & CStr(mvarobjEquipo.getID_EQUIPO)
    Else
        ' cuando NO es dar de alta
        llenar_combo cmbEquipo, New clsEquipos, 0, frmEquipoEdicion, " AND ID_EQUIPO = " & CStr(PK)
    End If

End Sub


Private Sub chkDesmontarAccesorio_Click()

If chkDesmontarAccesorio.value = vbChecked Then
    lblFechaBaja.Visible = True
    txtFechaBaja.Visible = True
Else
    lblFechaBaja.Visible = False
    txtFechaBaja.Visible = False
End If

End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cargar_combos
    
    
    presentar_datos

    opciones_edicion


End Sub



Private Sub opciones_edicion()

    
    If PK = 0 Then
        chkDesmontarAccesorio.Visible = False
        Exit Sub
    End If
    
    If mvarobjAccesorio.getEN_USO = 1 Then Exit Sub
    
       
    cmbEquipo.desactivar
    txtFechaAlta.Enabled = False
    txtFechaBaja.Enabled = False
    chkDesmontarAccesorio.Enabled = False
    
    cmdok.Visible = False
    
End Sub

Public Property Get EQUIPO() As clsEquipos

    Set EQUIPO = mvarobjEquipo

End Property

Public Property Set EQUIPO(objEquipo As clsEquipos)

    Set mvarobjEquipo = objEquipo

End Property

