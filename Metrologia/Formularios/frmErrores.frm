VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmErrores 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Envío de Consultas y Errores"
   ClientHeight    =   5400
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12750
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmErrores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   12750
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir Pendiente"
      Height          =   1035
      Left            =   10080
      Picture         =   "frmErrores.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4320
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.TextBox txterror 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   705
      Width           =   5850
   End
   Begin VB.TextBox txttexto 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   2955
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1290
      Width           =   5880
   End
   Begin VB.CheckBox chkcorregidos 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar corregidos"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6090
      TabIndex        =   6
      Top             =   4815
      Width           =   2715
   End
   Begin VB.CheckBox chkmis 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mostrar solo mis errores"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   6090
      TabIndex        =   5
      Top             =   4500
      Width           =   2445
   End
   Begin VB.CommandButton cmdNueva 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Limpiar Campos"
      Height          =   1035
      Left            =   60
      Picture         =   "frmErrores.frx":1194
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4305
      Width           =   1305
   End
   Begin VB.CommandButton cmdaceptar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Insertar"
      Height          =   1035
      Left            =   4680
      Picture         =   "frmErrores.frx":1A5E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4305
      Width           =   1260
   End
   Begin VB.CommandButton cmdBorrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Minimizar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   11430
      Picture         =   "frmErrores.frx":2328
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4320
      Width           =   1305
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3840
      Left            =   6090
      TabIndex        =   4
      Top             =   435
      Width           =   6630
      _ExtentX        =   11695
      _ExtentY        =   6773
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Alta de nueva consulta"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   60
      TabIndex        =   11
      Top             =   90
      Width           =   5880
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Listado de Consultas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   6105
      TabIndex        =   10
      Top             =   90
      Width           =   6600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción larga (Incluir datos precisos)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   1095
      Width           =   3495
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Descripción breve del error"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   45
      TabIndex        =   8
      Top             =   480
      Width           =   2325
   End
End
Attribute VB_Name = "frmErrores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkcorregidos_Click()
    cargar_lista
End Sub
Private Sub chkmis_Click()
    cargar_lista
End Sub

Private Sub cmdAceptar_Click()
    On Error GoTo fallo
    If txterror.Text = "" Then
        MsgBox "El error debe estar informado.", vbCritical, "Error"
        txterror.SetFocus
        Exit Sub
    End If
    If MsgBox("Va a dar de alta la incidencia. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
        Dim consulta As String
        Dim cod As Integer
        Dim rs As New ADODB.Recordset
        consulta = "select max(codigo) from errores"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            cod = 1
        Else
            cod = rs.Fields(0) + 1
        End If
        Set rs = Nothing
        consulta = "Insert into errores " & _
                   " values(" & _
                   cod & ",'" & txterror.Text & "','" & _
                   txttexto.Text & "'," & USUARIO.getID_EMPLEADO & ",'" & Format(Date, "yyyy/mm/dd") & "','" & Format(Time, "hh:mm") & "',0)"
        execute_bd consulta
        Me.MousePointer = 11
        cargar_lista
        enviar_mail
        ' Enviar a la web de proyectos
'        enviar_web (cod)
        Me.MousePointer = 0
        MsgBox "Incidencia generada correctamente.", vbInformation, App.Title
        cmdNueva_Click
    End If
    Exit Sub
fallo:
    Me.MousePointer = 0
    cmdNueva_Click
    MsgBox "Se ha producido un error al generar la incidencia.", vbCritical, Err.Description
End Sub

Private Sub cmdBorrar_Click()
    Me.WindowState = vbMinimized
End Sub


Private Sub cmdImprimir_Click()
    With frmReport
        .informe = "errores"
        .CRITERIO = "{ERRORES.CORREGIDO}=0"
        .imprimir = False
        .generar
        .Show 1
    End With
    Unload frmReport
End Sub

Private Sub cmdNueva_Click()
    txttexto = ""
    txterror = ""
    txterror.SetFocus
End Sub
Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 200
    Me.Top = 2200
    Call cabecera
    Call cargar_lista
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "ID", 400, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1200, lvwColumnCenter)
        .Tag = "Tipo"
    End With
    With lista.ColumnHeaders.Add(, , "Error", 3500, lvwColumnLeft)
        .Tag = "Error"
    End With
    With lista.ColumnHeaders.Add(, , "Descripcion", 1, lvwColumnLeft)
        .Tag = "Descripcion"
    End With
    With lista.ColumnHeaders.Add(, , "Usuario", 1200, lvwColumnCenter)
        .Tag = "Usuario"
    End With
End Sub
Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    Dim corregidos As String
    Dim mios As String
    If chkcorregidos.Value = Checked Then
        corregidos = ""
    Else
        corregidos = " and e.corregido = 0 "
    End If
    If chkmis.Value = Checked Then
        mios = " and e.empleado_id = " & USUARIO.getID_EMPLEADO
    Else
        mios = ""
    End If
    consulta = "select e.codigo,e.fecha,e.error,e.texto,em.usuario " & _
               "  from errores as e, usuarios as em " & _
               " where e.empleado_id = em.id_empleado " & _
               corregidos & mios & _
               " order by e.codigo desc"
    lista.ListItems.Clear
    Set rs = datos_bd(consulta)
    If rs.EOF = False Or rs.BOF = False Then
        Do
           With lista.ListItems.Add(, , rs.Fields(0))
            .SubItems(1) = Format(rs.Fields(1), "dd-mm-yyyy")
            .SubItems(2) = rs.Fields(2)
            .SubItems(3) = rs.Fields(3)
            .SubItems(4) = UCase(rs.Fields(4))
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
End Sub
Public Function enviar_web(incidencia As Long) As Boolean
    On Error GoTo falloConexion
    Dim consulta As String
    Dim cod As Integer
    Dim rs As New ADODB.Recordset
    Dim cw As ADODB.Connection
    Set cw = New ADODB.Connection
    cw.ConnectionString = "DRIVER={MySQL ODBC 3.51 Driver};" _
                         & "SERVER=ixitec.net;" _
                         & "DATABASE=ixiteca_dprj1;" _
                         & "UID=ixiteca_ixiteca;" _
                         & "PWD=bde40ea;" _
                         & "OPTION=" & 1 + 2 + 8 + 32 + 2048 + 16384
    cw.Open
    consulta = "select max(task_id) from tasks"
    rs.ActiveConnection = cw
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenForwardOnly
    rs.LockType = adLockReadOnly
    rs.Open consulta
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        cod = 1
    Else
        cod = rs.Fields(0) + 1
    End If
    Set rs = Nothing
    Dim proyecto As Integer
    proyecto = 7 ' Farmalab
    consulta = "Insert into tasks " & _
               " values(" & cod & ",'" & CStr(incidencia) & ". " & txterror.Text & "'," & cod & ",0," & proyecto & ",2,'" & _
               Format(Date, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss") & "',3,1,0,'" & Format(Date + 3, "yyyy-mm-dd") & " " & Format(Time, "hh:mm:ss") & "',0,0,0,'" & _
               txttexto.Text & " (Reportado por " & USUARIO.getNOMBRE & ")',0,'',2,0,0,0,0,1,'','','',0)"
    cw.Execute consulta
    consulta = "Insert into user_tasks " & _
               " values(2,0," & cod & ",100,0)"
    cw.Execute consulta
    cw.Close
    enviar_web = True
    Me.MousePointer = 0
    MsgBox "Incidencia generada correctamente.", vbInformation, App.Title
    Exit Function
falloConexion:
    Me.MousePointer = 0
    enviar_web = False
    MsgBox Err.Description
End Function

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        txterror = lista.ListItems(lista.SelectedItem.Index).SubItems(2)
        txttexto = lista.ListItems(lista.SelectedItem.Index).SubItems(3)
    End If
End Sub
Private Sub lista_KeyDown(KeyCode As Integer, Shift As Integer)
    lista_Click
End Sub
Public Sub enviar_mail()
   On Error Resume Next
   Ret = Enviar_Mail_CDO("julio.gonzalez@ixitec.net;david.escamez@ixitec.net", txterror.Text, txttexto.Text & "(Reportado por : " & USUARIO.getUSUARIO & ")", vbNullString)
End Sub

