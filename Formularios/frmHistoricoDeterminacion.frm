VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmHistoricoDeterminacion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Histórico de determinación"
   ClientHeight    =   7440
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   12630
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7440
   ScaleWidth      =   12630
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   11565
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6525
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   10485
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   6525
      Width           =   1050
   End
   Begin VB.CommandButton cmdTodas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ninguna"
      Height          =   285
      Index           =   1
      Left            =   1770
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   6330
      Width           =   1665
   End
   Begin VB.CommandButton cmdTodas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Todas"
      Height          =   285
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6330
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
      Left            =   6030
      TabIndex        =   7
      Top             =   6570
      Visible         =   0   'False
      Width           =   4155
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmHistoricoDeterminacion.frx":0000
         Left            =   1665
         List            =   "frmHistoricoDeterminacion.frx":000D
         TabIndex        =   10
         Top             =   270
         Width           =   2400
      End
      Begin VB.OptionButton opd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3D"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   9
         Top             =   495
         Width           =   690
      End
      Begin VB.OptionButton opd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2D"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   8
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
         TabIndex        =   11
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
      Left            =   6795
      TabIndex        =   3
      Top             =   6390
      Visible         =   0   'False
      Width           =   930
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5895
      Left            =   30
      TabIndex        =   0
      Top             =   405
      Width           =   3450
      _ExtentX        =   6085
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
      Left            =   7740
      TabIndex        =   4
      Top             =   6390
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   794
      _Version        =   393216
      Value           =   2004
      BuddyControl    =   "txtnum"
      BuddyDispid     =   196616
      OrigLeft        =   1590
      OrigTop         =   6570
      OrigRight       =   1830
      OrigBottom      =   6975
      Max             =   2015
      Min             =   2004
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin MSChart20Lib.MSChart grafico 
      Height          =   1155
      Left            =   9630
      OleObjectBlob   =   "frmHistoricoDeterminacion.frx":0025
      TabIndex        =   2
      Top             =   6165
      Visible         =   0   'False
      Width           =   1350
   End
   Begin MSComCtl2.DTPicker fdesde 
      Height          =   330
      Left            =   1125
      TabIndex        =   17
      Top             =   6840
      Width           =   1320
      _ExtentX        =   2328
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
      Format          =   60555265
      CurrentDate     =   38002
   End
   Begin MSComCtl2.DTPicker fhasta 
      Height          =   330
      Left            =   3060
      TabIndex        =   18
      Top             =   6840
      Width           =   1275
      _ExtentX        =   2249
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
      Format          =   60555265
      CurrentDate     =   38002
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Fecha desde"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   20
      Top             =   6885
      Width           =   930
   End
   Begin VB.Label lblCampos 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "hasta"
      Height          =   195
      Index           =   2
      Left            =   2565
      TabIndex        =   19
      Top             =   6885
      Width           =   405
   End
   Begin VB.OLE OLE2 
      Class           =   "MSGraph.Chart.8"
      Height          =   5910
      Left            =   3510
      OleObjectBlob   =   "frmHistoricoDeterminacion.frx":1D95
      TabIndex        =   16
      Top             =   405
      Width           =   9060
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
      Left            =   8100
      TabIndex        =   6
      Top             =   6480
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
      Left            =   3690
      TabIndex        =   5
      Top             =   6435
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Histórico de la determinación"
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
      Height          =   330
      Index           =   3
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12615
   End
End
Attribute VB_Name = "frmHistoricoDeterminacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LIMITE_INFERIOR As String
Public LIMITE_SUPERIOR As String

Private Sub cmbTipo_Click()
    tipo_grafico
End Sub

'Private Sub cmdcancel_Click()
'    Unload Me
'End Sub

Private Sub cmdImprimir_Click()
    If MsgBox("Va a imprimir. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
        On Error GoTo fallo
        grafico.EditCopy
        DoEvents   ' may be needed for large datasets
        Printer.Print " "
        Printer.FontBold = True
        Printer.Font = "Verdana"
        Printer.FontSize = 14
        Printer.Print lbltitulo(3).Caption
        Printer.Print " "
        Printer.PaintPicture Clipboard.GetData(), 1000, 1000
        Printer.EndDoc
    End If
    Exit Sub
fallo:
    MsgBox "Error al imprimir : " & Err.Description, vbCritical, App.Title
End Sub

Private Sub cmdSalir_Click()
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

Private Sub fdesde_Change()
    cargar_lista
End Sub

Private Sub fhasta_Change()
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    fdesde = Date - 180
    fhasta = Date
    With OLE2
      .Format = "CF_TEXT"
      .SizeMode = vbOLESizeStretch
    '  .CreateEmbed "", "MSGRAPH"
    End With
    
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnLeft)
        .Tag = "Muestra"
    End With
    With lista.ColumnHeaders.Add(, , "Muestra", 900, lvwColumnCenter)
        .Tag = "Muestra"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1200, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Resultado", 1000, lvwColumnRight)
        .Tag = "Resultado"
    End With
    'Tipos de graficos
    cargar_botones Me
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim oMuestra As New clsMuestra
    Dim rs As New ADODB.Recordset
'    Set rs = oMuestra.obtener_bano_anteriores(gmuestra, gdeterminacion)
    Set rs = oMuestra.obtener_muestras_anteriores(gmuestra, gdeterminacion, False, fdesde.value, fhasta.value)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        txtnum = rs.RecordCount
        cambiar.min = 1
        cambiar.Max = rs.RecordCount
        lbltitulo(3) = "Histórico de la determinación : " & rs(4)
        Do
           With lista.ListItems.Add(, , rs(0))
            .SubItems(1) = oMuestra.CodigoParticular(rs(6))
            .SubItems(2) = rs(5)
            ' En Resultado comprobar superindices
            If subindices(rs(3)) = True Then
                .SubItems(3) = subindice_formateado(rs(3))
            Else
                .SubItems(3) = rs(3)
            End If
           End With
           rs.MoveNext
           lista.ListItems(lista.ListItems.Count).Checked = True
        Loop Until rs.EOF
    Else
        txtnum = 0
    End If
    cargar_grafico
    Set oMuestra = Nothing
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

Public Sub cargar_grafico()
    Dim i As Integer
   On Error GoTo cargar_grafico_Error
    If txtnum = "" Then
        Exit Sub
    End If
    i = 1
    
'    Dim total As Integer
'    total = 0
'    With grafico
'       .ColumnCount = 1
'
'       For Row = 1 To lista.ListItems.Count
'            If lista.ListItems(Row).Checked = True Then
'                total = total + 1
'            End If
'       Next
''       .RowCount = txtnum
'       .RowCount = total
'       For Row = lista.ListItems.Count To 1 Step -1
'        If lista.ListItems(Row).Checked = True Then
'             If Row <= CInt(txtnum) Then
'                .Row = i
'                .RowLabel = lista.ListItems(Row).Text
'                If IsNumeric(lista.ListItems(Row).SubItems(3)) Then
'                   .Data = Replace(lista.ListItems(Row).SubItems(3), ".", ",")
'                Else
'                   .Data = 0
'                End If
'                i = i + 1
'             End If
'        End If
'       Next Row
'    End With


    Dim Msg, linea1, linea2, linea3, linea4, NewLine, Tabb
    
    Tabb = vbTab
    NewLine = vbNewLine
    linea1 = "Muestra" & Tabb
    linea2 = "Resultado" & Tabb
    linea3 = "Lim.Inf" & Tabb
    linea4 = "Lim.Sup" & Tabb
    For i = lista.ListItems.Count To 1 Step -1
        If lista.ListItems(i).Checked = True Then
            If i <= CInt(txtnum) Then
                If IsNumeric(lista.ListItems(i).SubItems(3)) Then
                   linea1 = linea1 & lista.ListItems(i).SubItems(1) & Tabb
                   linea2 = linea2 & Replace(lista.ListItems(i).SubItems(3), ".", ",") & Tabb
                   linea3 = linea3 & Replace(LIMITE_INFERIOR, ".", ",") & Tabb
                   linea4 = linea4 & Replace(LIMITE_SUPERIOR, ".", ",") & Tabb
                End If
'            Else
'                   linea1 = linea1 & "" & Tabb
'                   linea2 = linea2 & "" & Tabb
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_grafico of Formulario frmHistoricoDeterminacion"
End Sub

Private Sub opd_Click(Index As Integer)
    tipo_grafico
End Sub

Private Sub txtnum_Change()
'    cargar_grafico
End Sub

Public Sub tipo_grafico()
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
