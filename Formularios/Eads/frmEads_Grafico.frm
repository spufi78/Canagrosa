VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmEads_Grafico 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Histórico de determinaciones"
   ClientHeight    =   8460
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14490
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8460
   ScaleWidth      =   14490
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   960
      Left            =   13275
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   7470
      Width           =   1185
   End
   Begin VB.Frame Frame1 
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
      Height          =   780
      Left            =   45
      TabIndex        =   6
      Top             =   7515
      Width           =   4155
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmEads_Grafico.frx":0000
         Left            =   1665
         List            =   "frmEads_Grafico.frx":000D
         TabIndex        =   9
         Top             =   270
         Width           =   2400
      End
      Begin VB.OptionButton opd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3D"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   495
         Width           =   690
      End
      Begin VB.OptionButton opd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2D"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   7
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
         TabIndex        =   10
         Top             =   315
         Width           =   315
      End
   End
   Begin VB.TextBox txtnum 
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
      Left            =   3600
      TabIndex        =   2
      Top             =   5805
      Visible         =   0   'False
      Width           =   930
   End
   Begin MSComctlLib.ListView lista 
      Height          =   1305
      Left            =   6615
      TabIndex        =   0
      Top             =   90
      Visible         =   0   'False
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   2302
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
   Begin MSChart20Lib.MSChart grafico 
      Height          =   6555
      Left            =   45
      OleObjectBlob   =   "frmEads_Grafico.frx":0025
      TabIndex        =   1
      Top             =   855
      Width           =   10350
   End
   Begin MSComCtl2.UpDown cambiar 
      Height          =   450
      Left            =   4545
      TabIndex        =   3
      Top             =   5805
      Visible         =   0   'False
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   794
      _Version        =   393216
      Value           =   2004
      BuddyControl    =   "txtnum"
      BuddyDispid     =   196614
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
   Begin MSComctlLib.ListView deter 
      Height          =   6570
      Left            =   10395
      TabIndex        =   14
      Top             =   855
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   11589
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
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gráfico de resultados de determinaciones de los baños"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   13
      Top             =   420
      Width           =   3870
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13860
      Picture         =   "frmEads_Grafico.frx":1D95
      Top             =   180
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Gráfico de resultados de determinaciones de los baños"
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
      TabIndex        =   12
      Top             =   120
      Width           =   5775
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "resultados."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   4905
      TabIndex        =   5
      Top             =   5895
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Generar gráfico para los últimos "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   90
      TabIndex        =   4
      Top             =   5910
      Visible         =   0   'False
      Width           =   3405
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   825
      Left            =   0
      Top             =   0
      Width           =   14490
   End
End
Attribute VB_Name = "frmEads_Grafico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbTipo_Click()
    tipo_grafico
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub deter_Click()
    If deter.ListItems.Count > 0 Then
        cargar_grafico deter.ListItems(deter.selectedItem.Index).Text
    End If
End Sub

Private Sub Form_Activate()
    log (Me.Name)
    cargar_botones Me
    cabecera
    cargar_lista
End Sub

Public Sub cargar_lista()
    If lista.ListItems.Count <> 0 Then
        txtnum = lista.ListItems.Count
        cambiar.min = 1
        cambiar.Max = lista.ListItems.Count
    Else
        txtnum = 0
    End If
    Dim i As Integer
    For Col = 8 To lista.ColumnHeaders.Count
        With deter.ListItems.Add(, , Col)
            .SubItems(1) = lista.ColumnHeaders(Col).Text
        End With
    Next
    If deter.ListItems.Count > 0 Then
        cargar_grafico 8
    End If
End Sub
Public Sub cargar_grafico(columna As Integer)
    Dim i As Integer
    i = 1
    With grafico
'       .ColumnCount = lista.ColumnHeaders.Count - 7
       .columnCount = 1
       .RowCount = txtnum
'       For Col = 8 To lista.ColumnHeaders.Count
       For Col = columna To columna
'        .Column = Col - 7
        .Column = 1
        i = 1
        .ColumnLabel = lista.ColumnHeaders(Col).Text
        For Row = lista.ListItems.Count To 1 Step -1
              If Row <= CInt(txtnum) Then
                 .Row = i
                 .RowLabel = lista.ListItems(Row).SubItems(1)
                 If IsNumeric(lista.ListItems(Row).SubItems(Col - 1)) Then
                    .Data = Replace(lista.ListItems(Row).SubItems(Col - 1), ".", ",")
                 Else
                    .Data = 0
                 End If
                 i = i + 1
              End If
         Next Row
       Next
    End With
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

Private Sub cabecera()
    With deter.ColumnHeaders.Add(, , "Columna", 1, lvwColumnLeft)
        .Tag = "Columna"
    End With
    With deter.ColumnHeaders.Add(, , "Determinación", deter.Width, lvwColumnCenter)
        .Tag = "Determinación"
    End With

End Sub
