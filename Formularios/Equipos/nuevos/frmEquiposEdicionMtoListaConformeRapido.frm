VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEquiposEdicionMtoListaConformeRapido 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Lista de Mantenimientos"
   ClientHeight    =   5895
   ClientLeft      =   3525
   ClientTop       =   2340
   ClientWidth     =   7650
   Icon            =   "frmEquiposEdicionMtoListaConformeRapido.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5895
   ScaleWidth      =   7650
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   5490
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5010
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5010
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4335
      Left            =   30
      TabIndex        =   1
      Top             =   630
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   7646
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
      BackColor       =   14609914
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
   Begin VB.Label lblCap 
      Caption         =   "La siguiente lista de mantenimientos (aquellos que estén chequeados) se cerrarán conforme de manera automática:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   7545
   End
End
Attribute VB_Name = "frmEquiposEdicionMtoListaConformeRapido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarobjrs As ADODB.Recordset
Private mvarstrExcepciones As String
Private mvarblnResultado As Boolean

Public Property Get RESULTADO() As Boolean

    RESULTADO = mvarblnResultado

End Property

Public Property Let RESULTADO(ByVal blnResultado As Boolean)

    mvarblnResultado = blnResultado

End Property

Private Sub Carga()


    If RS.RecordCount = 0 Then
        Me.Hide
        mvarstrExcepciones = ""
        Exit Sub
    End If
    
    
    RS.MoveFirst
    lista.ListItems.Clear
    
    While Not RS.EOF
        With lista.ListItems.Add(, , RS!EQUIPO_ID)
            .SubItems(1) = RS!EQUIPO
            .SubItems(2) = Format(RS!fecha_actual, "dd/mm/yyyy")
            .SubItems(3) = RS!id_mantenimiento
            .Checked = True
        End With
        RS.MoveNext
    Wend

End Sub

Private Sub cmdcancel_Click()
    mvarblnResultado = False
    Me.Hide
End Sub

Private Sub recoger_excepciones()
Dim lngCont As Long

    mvarstrExcepciones = ";"

    For lngCont = 1 To lista.ListItems.Count
        If Not lista.ListItems(lngCont).Checked Then
            mvarstrExcepciones = mvarstrExcepciones & lista.ListItems(lngCont).SubItems(3) & ";"
        End If
    Next lngCont

End Sub

Private Sub cmdOk_Click()
    recoger_excepciones
    mvarblnResultado = True
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = 0 Then Cancel = 1
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set mvarobjrs = Nothing

End Sub


Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Nº Equipo", 1200, lvwColumnLeft
        .Add , , "Equipo", 5000, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnCenter
        .Add , , "id_mantenimiento", 0, lvwColumnLeft
        
    End With

End Sub

Private Sub Form_Load()
    
    cargar_botones Me
    cabecera
    

    Carga
    
End Sub


Public Property Get RS() As ADODB.Recordset

    Set RS = mvarobjrs

End Property

Public Property Set RS(objrs As ADODB.Recordset)

    Set mvarobjrs = objrs

End Property

Public Property Get Excepciones() As String

    Excepciones = mvarstrExcepciones

End Property

Public Property Let Excepciones(ByVal strExcepciones As String)

    mvarstrExcepciones = strExcepciones

End Property
