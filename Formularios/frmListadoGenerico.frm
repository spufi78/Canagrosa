VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmListadoGenerico 
   Caption         =   "Seleccionar"
   ClientHeight    =   9165
   ClientLeft      =   4215
   ClientTop       =   1485
   ClientWidth     =   8820
   LinkTopic       =   "Form1"
   ScaleHeight     =   9165
   ScaleWidth      =   8820
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aceptar"
      Height          =   870
      Left            =   6630
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8280
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   7710
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8265
      Width           =   1050
   End
   Begin VB.TextBox txtFiltro 
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
      Height          =   285
      Left            =   750
      TabIndex        =   2
      Top             =   60
      Width           =   7995
   End
   Begin MSFlexGridLib.MSFlexGrid grdListado 
      Height          =   7815
      Left            =   30
      TabIndex        =   0
      Top             =   390
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   13785
      _Version        =   393216
      BackColor       =   12640511
      BackColorSel    =   -2147483636
      HighLight       =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
   End
   Begin VB.Label lbl 
      Caption         =   "Filtro:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   210
      TabIndex        =   1
      Top             =   90
      Width           =   495
   End
End
Attribute VB_Name = "frmListadoGenerico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private mvarstrConsulta As String
Private mvarstrCampoId As String
Private mvarstrCampoFiltro As String
Private mvarlngSelect_Id As Long
Private mvarstrFiltro As String
Private mvarstrSelectedText As String
Private mvarstrCampoTexto As String

Private mvarRs As ADODB.RecordSet

Private mvarstrCampos() As String
Private mvarintIndice As Integer
Private mvarintIndiceCampoTexto As Integer
Private Sub CargarDatos(Optional ByVal recargar As Boolean = True)

On Error GoTo CargarDatos_Error
Dim intCont As Integer
    If recargar Then
        mvarstrFiltro = txtfiltro.Text
        Set mvarRs = datos_bd(mvarstrConsulta & " AND " & mvarstrCampoFiltro & " like '%" & mvarstrFiltro & "%'")
    End If
    
    grdListado.Rows = 1
    
    If mvarRs Is Nothing Then Exit Sub
    If mvarRs.RecordCount = 0 Then Exit Sub
    
    With grdListado
        
        mvarRs.MoveFirst
        While Not mvarRs.EOF
        .Rows = grdListado.Rows + 1
            For intCont = 0 To mvarintIndice - 1
                .TextMatrix(.Rows - 1, intCont) = mvarRs(mvarstrCampos(intCont))
            Next intCont
            mvarRs.MoveNext
        Wend
    End With
    

On Error GoTo 0
    Exit Sub
CargarDatos_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure CargarDatos of Formulario frmListadoGenerico"
End Sub

Private Sub cmdcancel_Click()
    mvarlngSelect_Id = 0
    Unload Me
End Sub

Private Sub cmdok_Click()
Dim lngFila As Long

    With grdListado
        If .Rows < 1 Or .RowSel < 1 Then
            MsgBox "No ha seleccionado ningún Registro", vbInformation, "Filtrar"
            Exit Sub
        End If
        
        lngFila = .RowSel
        
        mvarlngSelect_Id = CLng(.TextMatrix(lngFila, 0))
        mvarstrSelectedText = .TextMatrix(lngFila, mvarintIndiceCampoTexto)
        Unload Me
    End With
End Sub


Private Sub Form_Load()

    If Trim(mvarstrFiltro) <> "" Then
        Set mvarRs = datos_bd(mvarstrConsulta & " AND lower(" & mvarstrCampoFiltro & ") like '%" & LCase(mvarstrFiltro) & "%'")
        txtfiltro.Text = mvarstrFiltro
    Else
        Set mvarRs = datos_bd(mvarstrConsulta)
    End If
    
    configurarLista
    
    CargarDatos (False)
    
End Sub


Public Property Get consulta() As String

    consulta = mvarstrConsulta

End Property

Public Property Let consulta(ByVal strConsulta As String)

    mvarstrConsulta = strConsulta

End Property

Public Property Get CampoId() As String

    CampoId = mvarstrCampoId

End Property

Public Property Let CampoId(ByVal strCampoId As String)

    mvarstrCampoId = strCampoId

End Property

Public Property Get CampoFiltro() As String

    CampoFiltro = mvarstrCampoFiltro

End Property

Public Property Let CampoFiltro(ByVal strCampoFiltro As String)

    mvarstrCampoFiltro = strCampoFiltro

End Property

Public Property Get Select_Id() As Long

    Select_Id = mvarlngSelect_Id

End Property

Public Property Let Select_Id(ByVal lngSelect_Id As Long)

    mvarlngSelect_Id = lngSelect_Id

End Property

Private Sub configurarLista()
On Error GoTo configurarLista_Error
Dim fld As ADODB.Field
Dim wd As Currency

    grdListado.FixedCols = 0
    grdListado.COLS = mvarRs.Fields.Count
    grdListado.ColWidth(0) = 0

    If mvarRs.Fields.Count < 1 Then
        wd = 0.99
    Else
        wd = (grdListado.Width * 0.98) / (mvarRs.Fields.Count - 1)
    End If
    
    mvarintIndice = 1
    ReDim Preserve mvarstrCampos(mvarintIndice)
    mvarstrCampos(0) = mvarstrCampoId
    
    
    For Each fld In mvarRs.Fields
        If LCase(Trim(fld.Name)) <> LCase(mvarstrCampoId) Then
            mvarintIndice = mvarintIndice + 1
            grdListado.ColWidth(mvarintIndice - 1) = wd
            ReDim Preserve mvarstrCampos(mvarintIndice)
            mvarstrCampos(mvarintIndice - 1) = fld.Name
            grdListado.TextMatrix(0, mvarintIndice - 1) = fld.Name
            If Trim(fld.Name) = mvarstrCampoTexto Then
                mvarintIndiceCampoTexto = mvarintIndice - 1
            End If
            
        End If
    Next fld
    
    grdListado.Rows = 1
    

On Error GoTo 0
    Exit Sub
configurarLista_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure configurarLista of Formulario frmListadoGenerico"
End Sub

Public Property Get FILTRO() As String

    FILTRO = mvarstrFiltro

End Property

Public Property Let FILTRO(ByVal strFiltro As String)

    mvarstrFiltro = strFiltro

End Property

Private Sub grdListado_DblClick()
Call cmdok_Click
End Sub

Private Sub txtfiltro_Change()
    Call CargarDatos
End Sub



Public Property Get SelectedText() As String

    SelectedText = mvarstrSelectedText

End Property

Public Property Let SelectedText(ByVal strSelectedText As String)

    mvarstrSelectedText = strSelectedText

End Property

Public Property Get CampoTexto() As String

    CampoTexto = mvarstrCampoTexto

End Property

Public Property Let CampoTexto(ByVal strCampoTexto As String)

    mvarstrCampoTexto = strCampoTexto

End Property
