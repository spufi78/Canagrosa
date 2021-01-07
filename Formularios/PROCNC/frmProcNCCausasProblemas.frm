VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmProcNCCausasProblemas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Causas del Problema"
   ClientHeight    =   9450
   ClientLeft      =   1800
   ClientTop       =   1665
   ClientWidth     =   11055
   Icon            =   "frmProcNCCausasProblemas.frx":0000
   LinkTopic       =   "frmProcNCCausasProblemas"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   11055
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Estudio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   45
      TabIndex        =   28
      Top             =   7740
      Width           =   10965
      Begin MSDataListLib.DataCombo cmbDESVIACION_ID 
         Height          =   315
         Left            =   90
         TabIndex        =   29
         Top             =   225
         Width           =   10485
         _ExtentX        =   18494
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
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   10035
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8460
      Width           =   1020
   End
   Begin VB.Frame fraCausaRaiz 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Causa Raiz"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   30
      TabIndex        =   13
      Top             =   6390
      Width           =   10965
      Begin VB.TextBox txtCausaRaiz 
         Appearance      =   0  'Flat
         Height          =   795
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         Top             =   480
         Width           =   10845
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "¿Cual ha sido la causa raíz?"
         Height          =   195
         Index           =   43
         Left            =   90
         TabIndex        =   16
         Top             =   240
         Width           =   2010
      End
   End
   Begin VB.Frame fraCausasContributibas 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Causas Contributivas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4965
      Left            =   30
      TabIndex        =   2
      Top             =   1380
      Width           =   10965
      Begin VB.TextBox txtCC 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   2970
         TabIndex        =   9
         Top             =   3120
         Width           =   7905
      End
      Begin VB.TextBox txtCC_Desc 
         Appearance      =   0  'Flat
         Height          =   525
         Index           =   4
         Left            =   1950
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   10
         Top             =   3420
         Width           =   8925
      End
      Begin VB.TextBox txtCC 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   2970
         TabIndex        =   7
         Top             =   2190
         Width           =   7905
      End
      Begin VB.TextBox txtCC_Desc 
         Appearance      =   0  'Flat
         Height          =   525
         Index           =   3
         Left            =   1950
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   2490
         Width           =   8925
      End
      Begin VB.TextBox txtCC 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   2970
         TabIndex        =   5
         Top             =   1230
         Width           =   7905
      End
      Begin VB.TextBox txtCC_Desc 
         Appearance      =   0  'Flat
         Height          =   525
         Index           =   2
         Left            =   1950
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   1530
         Width           =   8925
      End
      Begin VB.TextBox txtCC 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   2970
         TabIndex        =   11
         Top             =   4050
         Width           =   7905
      End
      Begin VB.TextBox txtCC_Desc 
         Appearance      =   0  'Flat
         Height          =   525
         Index           =   5
         Left            =   1950
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   4350
         Width           =   8925
      End
      Begin VB.TextBox txtCC_Desc 
         Appearance      =   0  'Flat
         Height          =   525
         Index           =   1
         Left            =   1950
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   570
         Width           =   8925
      End
      Begin VB.TextBox txtCC 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   2970
         TabIndex        =   3
         Top             =   270
         Width           =   7905
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "5.- Descripcion de la Causa Contributiva"
         Height          =   195
         Index           =   7
         Left            =   60
         TabIndex        =   25
         Top             =   4110
         Width           =   2835
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "¿Por qué ha ocurrido?"
         Height          =   195
         Index           =   6
         Left            =   210
         TabIndex        =   24
         Top             =   4380
         Width           =   1575
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "4.- Descripcion de la Causa Contributiva"
         Height          =   195
         Index           =   5
         Left            =   60
         TabIndex        =   27
         Top             =   3180
         Width           =   2835
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "¿Por qué ha ocurrido?"
         Height          =   195
         Index           =   4
         Left            =   210
         TabIndex        =   26
         Top             =   3450
         Width           =   1575
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "3.- Descripcion de la Causa Contributiva"
         Height          =   195
         Index           =   3
         Left            =   60
         TabIndex        =   22
         Top             =   2250
         Width           =   2835
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "¿Por qué ha ocurrido?"
         Height          =   195
         Index           =   2
         Left            =   210
         TabIndex        =   21
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "2.- Descripcion de la Causa Contributiva"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   1290
         Width           =   2835
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "¿Por qué ha ocurrido?"
         Height          =   195
         Index           =   0
         Left            =   210
         TabIndex        =   19
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "1.- Descripcion de la Causa Contributiva"
         Height          =   195
         Index           =   36
         Left            =   60
         TabIndex        =   18
         Top             =   330
         Width           =   2835
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "¿Por qué ha ocurrido?"
         Height          =   195
         Index           =   33
         Left            =   210
         TabIndex        =   17
         Top             =   600
         Width           =   1575
      End
   End
   Begin VB.Frame fraCausaDirecta 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Causa Directa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1365
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   10995
      Begin VB.TextBox txtCausaDirecta 
         Appearance      =   0  'Flat
         Height          =   795
         Left            =   60
         MaxLength       =   65000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Top             =   510
         Width           =   10875
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "¿Que ha ocurrido?"
         Height          =   195
         Index           =   32
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmProcNCCausasProblemas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long

Private mvarobjProcNC As New clsProcNc
Private rs As ADODB.Recordset
Private strSql As String
Private mvarblnEditable As Boolean



Private Function guardar_datos() As Boolean

On Error GoTo guardar_datos_Error
    guardar_datos = False
        
    mvarobjProcNC.guardar_datos_causas txtCausaRaiz.Text, txtCausaDirecta.Text, txtCC(1).Text, txtCC_Desc(1).Text, _
    txtCC(2).Text, txtCC_Desc(2).Text, _
    txtCC(3).Text, txtCC_Desc(3).Text, _
    txtCC(4).Text, txtCC_Desc(4).Text, _
    txtCC(5).Text, txtCC_Desc(5).Text, _
    IIf(cmbDESVIACION_ID.Text = "", 0, cmbDESVIACION_ID.BoundText)
        
    guardar_datos = True
    
On Error GoTo 0
    Exit Function
guardar_datos_Error:
    MsgBox Err.Description
    guardar_datos = False
End Function

Private Sub cmdcancel_Click()
    If Not mvarblnEditable Then Unload Me

    If Not guardar_datos Then Exit Sub
    
    Unload Me

End Sub

Private Sub Form_Activate()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2

End Sub

Private Sub Form_Load()
    
    cargar_botones Me
    
    cargar_datos

    opciones_edicion
End Sub



Private Sub opciones_edicion()

    txtCausaDirecta.Enabled = mvarblnEditable
    txtCausaRaiz.Enabled = mvarblnEditable
    txtCC(1).Enabled = mvarblnEditable
    txtCC_Desc(1).Enabled = mvarblnEditable
    txtCC(2).Enabled = mvarblnEditable
    txtCC_Desc(2).Enabled = mvarblnEditable
    txtCC(3).Enabled = mvarblnEditable
    txtCC_Desc(3).Enabled = mvarblnEditable
    txtCC(4).Enabled = mvarblnEditable
    txtCC_Desc(4).Enabled = mvarblnEditable
    txtCC(5).Enabled = mvarblnEditable
    txtCC_Desc(5).Enabled = mvarblnEditable
    cmbDESVIACION_ID.Enabled = mvarblnEditable

End Sub


Private Sub cargar_datos()
    Dim oDeco As New clsDecodificadora
    oDeco.cargar_combo cmbDESVIACION_ID, DECODIFICADORA.NC_DESVIACIONES
    
    mvarobjProcNC.Carga PK
    PresentarDatos_Otros

End Sub
Private Sub PresentarDatos_Otros()

    With mvarobjProcNC
        
        txtCausaDirecta.Text = .getCAUSA_DIRECTA
        
        txtCC(1).Text = .getCAUSA_CONTRIBUTIVA_1
        txtCC_Desc(1).Text = .getCAUSA_CONTRIBUTIVA_1_DESCRIPCION
        
        txtCC(2).Text = .getCAUSA_CONTRIBUTIVA_2
        txtCC_Desc(2).Text = .getCAUSA_CONTRIBUTIVA_2_DESCRIPCION
        
        txtCC(3).Text = .getCAUSA_CONTRIBUTIVA_3
        txtCC_Desc(3).Text = .getCAUSA_CONTRIBUTIVA_3_DESCRIPCION
        
        txtCC(4).Text = .getCAUSA_CONTRIBUTIVA_4
        txtCC_Desc(4).Text = .getCAUSA_CONTRIBUTIVA_4_DESCRIPCION
        
        txtCC(5).Text = .getCAUSA_CONTRIBUTIVA_5
        txtCC_Desc(5).Text = .getCAUSA_CONTRIBUTIVA_5_DESCRIPCION
        
        txtCausaRaiz.Text = .getCAUSA_RAIZ
        
        cmbDESVIACION_ID.BoundText = .getDESVIACION_ID
    End With
    
End Sub

Public Property Get Editable() As Boolean

    Editable = mvarblnEditable

End Property

Public Property Let Editable(ByVal blnEditable As Boolean)

    mvarblnEditable = blnEditable

End Property

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If UnloadMode = 0 Then cmdcancel_Click
End Sub
Private Sub txtCausaDirecta_GotFocus()
    txtCausaDirecta.BackColor = &H80C0FF
    txtCausaDirecta.SelStart = 0
    txtCausaDirecta.SelLength = Len(txtCausaDirecta)

End Sub
Private Sub txtCausaDirecta_LostFocus()
    txtCausaDirecta.BackColor = vbWhite
End Sub
Private Sub txtCausaRaiz_GotFocus()
    txtCausaRaiz.BackColor = &H80C0FF
    txtCausaRaiz.SelStart = 0
    txtCausaRaiz.SelLength = Len(txtCausaRaiz)
   
End Sub

Private Sub txtCausaRaiz_LostFocus()
    txtCausaRaiz.BackColor = vbWhite
End Sub

Private Sub txtCC_Desc_GotFocus(Index As Integer)
    txtCC_Desc(Index).BackColor = &H80C0FF
    txtCC_Desc(Index).SelStart = 0
    txtCC_Desc(Index).SelLength = Len(txtCC_Desc(Index))

End Sub

Private Sub txtCC_Desc_LostFocus(Index As Integer)
    txtCC_Desc(Index).BackColor = vbWhite
End Sub

Private Sub txtCC_GotFocus(Index As Integer)
    txtCC(Index).BackColor = &H80C0FF
    txtCC(Index).SelStart = 0
    txtCC(Index).SelLength = Len(txtCC(Index))
End Sub

Private Sub txtCC_LostFocus(Index As Integer)
    txtCC(Index).BackColor = vbWhite
End Sub
