VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEquiposTendencias 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gráficas de resultados de Verificación"
   ClientHeight    =   8610
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   13680
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   13680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtparametro 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   5895
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   17
      Top             =   630
      Width           =   7695
   End
   Begin VB.TextBox txtperiodicidad 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   330
      Left            =   990
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   15
      Top             =   630
      Width           =   3555
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12555
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   7650
      Width           =   1050
   End
   Begin VB.CommandButton cmdTodas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ninguna"
      Height          =   285
      Index           =   1
      Left            =   2925
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7020
      Width           =   1665
   End
   Begin VB.CommandButton cmdTodas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Todas"
      Height          =   285
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   7050
      Width           =   1665
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Tipo Gráfico"
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
      Height          =   780
      Left            =   5805
      TabIndex        =   5
      Top             =   7200
      Visible         =   0   'False
      Width           =   4155
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmEquiposTendencias.frx":0000
         Left            =   1665
         List            =   "frmEquiposTendencias.frx":000D
         TabIndex        =   8
         Top             =   270
         Width           =   2400
      End
      Begin VB.OptionButton opd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3D"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   495
         Width           =   690
      End
      Begin VB.OptionButton opd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2D"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   270
         Value           =   -1  'True
         Width           =   690
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Left            =   1215
         TabIndex        =   9
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.TextBox txtnum 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3060
      TabIndex        =   1
      Text            =   "1"
      Top             =   7380
      Visible         =   0   'False
      Width           =   930
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5895
      Left            =   30
      TabIndex        =   0
      Top             =   1125
      Width           =   4530
      _ExtentX        =   7990
      _ExtentY        =   10398
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
   Begin MSComCtl2.UpDown cambiar 
      Height          =   450
      Left            =   4005
      TabIndex        =   2
      Top             =   7380
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   794
      _Version        =   393216
      Value           =   1
      BuddyControl    =   "txtnum"
      BuddyDispid     =   196617
      OrigLeft        =   1590
      OrigTop         =   6570
      OrigRight       =   1830
      OrigBottom      =   6975
      Max             =   5000
      Min             =   1
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.DTPicker txtFecha 
      Height          =   345
      Left            =   3060
      TabIndex        =   20
      Top             =   7965
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   609
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
      Format          =   52232193
      CurrentDate     =   2
      MinDate         =   2
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Generar gráfico desde la fecha "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   19
      Top             =   8055
      Width           =   2820
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Parámetro"
      Height          =   195
      Index           =   0
      Left            =   4950
      TabIndex        =   18
      Top             =   705
      Width           =   720
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Periodicidad"
      Height          =   195
      Index           =   9
      Left            =   45
      TabIndex        =   16
      Top             =   705
      Width           =   870
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      Caption         =   "Gráficas de resultados de Verificación"
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
      TabIndex        =   14
      Top             =   120
      Width           =   3990
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13095
      Picture         =   "frmEquiposTendencias.frx":0025
      Top             =   0
      Width           =   480
   End
   Begin VB.OLE OLE2 
      Class           =   "MSGraph.Chart.8"
      Height          =   5865
      Left            =   4590
      OleObjectBlob   =   "frmEquiposTendencias.frx":032F
      TabIndex        =   13
      Top             =   1125
      Width           =   9015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "resultados."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4320
      TabIndex        =   4
      Top             =   7470
      Visible         =   0   'False
      Width           =   990
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Generar gráfico para los últimos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   45
      TabIndex        =   3
      Top             =   7470
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   510
      Left            =   0
      Top             =   0
      Width           =   13650
   End
End
Attribute VB_Name = "frmEquiposTendencias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_EQUIPO_ID As Long
Public PK_PERIODICIDAD As Long
Public PK_PARAMETRO As String
Public PK_RANGO_MIN As String
Public PK_RANGO_MAX As String
Public PK_TIPO As Integer

Private Sub cmbTipo_Click()
    tipo_grafico
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmdTodas_Click(Index As Integer)
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        If Index = 0 Then
            lista.ListItems(i).Checked = True
        Else
            lista.ListItems(i).Checked = False
        End If
    Next
    cargar_grafico
End Sub

Private Sub Form_Load()
    log (Me.Name)
    txtFecha = Date - 365
    With OLE2
      .Format = "CF_TEXT"
      .SizeMode = vbOLESizeStretch
    ' .CreateEmbed "", "MSGRAPH"
    End With
    cargar_botones Me
    cargar_lista
    
    Dim oEq As New clsEquipos
    oEq.Carga_Datos_Basicos PK_EQUIPO_ID
    If PK_TIPO = 2 Then
        lbltitulo = "Gráficas de resultados de Verificación : " & oEq.getDESCRIPCION
    ElseIf PK_TIPO = 1 Then
        lbltitulo = "Gráficas de resultados de Calibración : " & oEq.getDESCRIPCION
    End If
    Me.Caption = lbltitulo
    Dim op As New clsEquiposPeriodicidad
    op.Carga PK_PERIODICIDAD
    txtperiodicidad = op.getDESCRIPCION
    txtparametro = PK_PARAMETRO
End Sub

Private Sub cargar_lista()
    Dim oMuestra As New clsMuestra
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    lista.ColumnHeaders.Clear
    With lista.ColumnHeaders
        .Add , , "ID", 260, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnCenter
    End With
    If PK_TIPO = 2 Then ' VERIFICACION
    'M1319: Adición de la cláusula REALIZADO = 1 en verificaciones
        consulta = "SELECT DISTINCT B.DESCRIPCION " & _
                   "  FROM EQ_VERIFICACION_EQUIPOS A, EQ_VERIFICACION_PARAMETROS_RESULTADOS B " & _
                   " WHERE A.ID_VERIFICACION = B.VERIFICACION_ID " & _
                   " AND A.ESTADO <> 0 " & _
                   " AND B.REALIZADO = 1" & _
                   " AND A.EQUIPO_ID = " & PK_EQUIPO_ID & _
                   " AND A.PERIODICIDAD_ID = " & PK_PERIODICIDAD & _
                   " AND B.DESCRIPCION = '" & PK_PARAMETRO & "'"
    ElseIf PK_TIPO = 1 Then ' CALIBRACION
        consulta = "SELECT DISTINCT B.DESCRIPCION " & _
                   "  FROM eq_calibracion_equipos A, eq_calibracion_parametros_resultados B " & _
                   " WHERE A.ID_CALIBRACION = B.CALIBRACION_ID " & _
                   " AND A.ESTADO <> 0 " & _
                   " AND A.EQUIPO_ID = " & PK_EQUIPO_ID & _
                   " AND A.PERIODICIDAD_ID = " & PK_PERIODICIDAD & _
                   " AND B.DESCRIPCION = '" & PK_PARAMETRO & "'"
    Else
        Exit Sub
    End If
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        Do
            With lista.ColumnHeaders
                .Add , , rs(0), 1500, lvwColumnLeft
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    With lista.ColumnHeaders
       .Add , , "MIN", 0, lvwColumnLeft
       .Add , , "MAX", 0, lvwColumnLeft
    End With
    
    If PK_TIPO = 2 Then ' VERIFICACION
        'M1319: Adición de la cláusula REALIZADO = 1 en verificaciones
        consulta = "SELECT A.ID_VERIFICACION,A.FECHA_ACTUAL,B.descripcion,B.resultado,B.rango_min, B.rango_max  " & _
                   "  FROM EQ_VERIFICACION_EQUIPOS A, EQ_VERIFICACION_PARAMETROS_RESULTADOS B " & _
                   " WHERE A.ID_VERIFICACION = B.VERIFICACION_ID " & _
                   " AND A.ESTADO <> 0 " & _
                   " AND B.REALIZADO = 1" & _
                   " AND A.EQUIPO_ID = " & PK_EQUIPO_ID & _
                   " AND A.PERIODICIDAD_ID = " & PK_PERIODICIDAD & _
                   " AND B.DESCRIPCION = '" & PK_PARAMETRO & "'" & _
                   " AND A.FECHA_ACTUAL >= '" & Format(txtFecha, "yyyy-mm-dd") & "' " & _
                   " ORDER BY A.ID_VERIFICACION DESC, B.DESCRIPCION"
    ElseIf PK_TIPO = 1 Then ' CALIBRACION
        consulta = "SELECT A.ID_CALIBRACION,A.FECHA_ACTUAL,B.descripcion,B.resultado,B.rango_min, B.rango_max  " & _
                   "  FROM eq_calibracion_equipos A, eq_calibracion_parametros_resultados B " & _
                   " WHERE A.ID_CALIBRACION = B.CALIBRACION_ID " & _
                   " AND A.ESTADO <> 0 " & _
                   " AND A.EQUIPO_ID = " & PK_EQUIPO_ID & _
                   " AND A.PERIODICIDAD_ID = " & PK_PERIODICIDAD & _
                   " AND B.DESCRIPCION = '" & PK_PARAMETRO & "'" & _
                   " AND A.FECHA_ACTUAL >= '" & Format(txtFecha, "yyyy-mm-dd") & "' " & _
                   " ORDER BY A.ID_CALIBRACION DESC, B.DESCRIPCION"
    Else
        Exit Sub
    End If
    Set rs = datos_bd(consulta)
    lista.ListItems.Clear
    Dim registros As Integer
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(3) ' Resultado
                .SubItems(3) = rs(4) ' Rango. Min
                .SubItems(4) = rs(5) ' Rango. Max
            End With
            lista.ListItems(lista.ListItems.Count).Checked = True
            registros = registros + 1
            rs.MoveNext
        Loop Until rs.EOF
    
    End If
    On Error Resume Next
    txtnum.Text = CStr(registros)
    cambiar.min = 1
    cambiar.Max = registros
    
'    lbltitulo(3) = "Histórico de la determinación : " & rs(4)
'    cargar_grafico
    Set rs = Nothing
End Sub

Private Sub lista_Click()
    cargar_grafico
End Sub

Private Sub lista_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   If lista.ListItems.Count > 0 Then
     lista.SortKey = ColumnHeader.Index - 1
     If lista.SortOrder = 0 Then
        lista.SortOrder = 1
     Else
        lista.SortOrder = 0
     End If
     lista.Sorted = True
   End If
End Sub

Private Sub cargar_grafico()
    Dim i As Integer
   On Error GoTo cargar_grafico_Error
    If txtnum = "" Then
        Exit Sub
    End If
    i = 1
    Dim Msg, linea1, linea2, NewLine, Tabb
    
    Tabb = vbTab
    NewLine = vbNewLine
    linea1 = "Fecha" & Tabb
    linea2 = lista.ColumnHeaders(3).Text & Tabb
    linea3 = "Lim.Inf" & Tabb
    linea4 = "Lim.Sup" & Tabb
    For i = lista.ListItems.Count To 1 Step -1
        If lista.ListItems(i).Checked = True Then
            If i <= CInt(txtnum) Then
                If IsNumeric(lista.ListItems(i).SubItems(2)) Then
                   linea1 = linea1 & lista.ListItems(i).SubItems(1) & Tabb
                   linea2 = linea2 & Replace(lista.ListItems(i).SubItems(2), ".", ",") & Tabb
                   linea3 = linea3 & Replace(lista.ListItems(i).SubItems(3), ".", ",") & Tabb
                   linea4 = linea4 & Replace(lista.ListItems(i).SubItems(4), ".", ",") & Tabb
                End If
            End If
        End If
    Next
    Msg = linea1 & NewLine & linea2 & NewLine & linea3 & NewLine & linea4
    
      With OLE2
          .DataText = Msg
          .Update
'          .Refresh
      End With

   On Error GoTo 0
   Exit Sub

cargar_grafico_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_grafico of Formulario frmEquiposTendencias"
End Sub

Private Sub opd_Click(Index As Integer)
    tipo_grafico
End Sub

Private Sub txtFecha_Change()
'    cargar_grafico
    cargar_lista
End Sub

Private Sub txtnum_Change()
    cargar_grafico
End Sub

Private Sub tipo_grafico()
    If opd(0).value = True Then
        Select Case cmbTipo
        Case "Linea"
            grafico.ChartType = VtChChartType2dLine
        Case "Barra"
            grafico.ChartType = VtChChartType2dBar
        Case "Area"
            grafico.ChartType = VtChChartType2dArea
        Case Else
            grafico.ChartType = VtChChartType2dLine
        End Select
    Else
        Select Case cmbTipo
        Case "Linea"
            grafico.ChartType = VtChChartType3dLine
        Case "Barra"
            grafico.ChartType = VtChChartType3dBar
        Case "Area"
            grafico.ChartType = VtChChartType3dArea
        Case Else
            grafico.ChartType = VtChChartType3dLine
        End Select
    End If

End Sub
