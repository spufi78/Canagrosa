VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmHistoricoDeterminacionCE 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Histórico del CE"
   ClientHeight    =   9570
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14895
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9570
   ScaleWidth      =   14895
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13770
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8595
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   12690
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   8595
      Width           =   1050
   End
   Begin VB.CommandButton cmdTodas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ninguna"
      Height          =   285
      Index           =   1
      Left            =   1770
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8445
      Width           =   1665
   End
   Begin VB.CommandButton cmdTodas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Todas"
      Height          =   285
      Index           =   0
      Left            =   60
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8445
      Width           =   1665
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
      TabIndex        =   3
      Top             =   8910
      Width           =   930
   End
   Begin MSComctlLib.ListView lista 
      Height          =   7920
      Left            =   30
      TabIndex        =   0
      Top             =   405
      Width           =   3450
      _ExtentX        =   6085
      _ExtentY        =   13970
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
      TabIndex        =   4
      Top             =   8910
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   794
      _Version        =   393216
      Value           =   2004
      BuddyControl    =   "txtnum"
      BuddyDispid     =   196612
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
      OleObjectBlob   =   "frmHistoricoDeterminacionCE.frx":0000
      TabIndex        =   2
      Top             =   8280
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.OLE OLE2 
      Class           =   "MSGraph.Chart.8"
      Height          =   7935
      Left            =   3510
      OleObjectBlob   =   "frmHistoricoDeterminacionCE.frx":1D70
      TabIndex        =   11
      Top             =   405
      Width           =   11355
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
      Left            =   4365
      TabIndex        =   6
      Top             =   9000
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
      Left            =   90
      TabIndex        =   5
      Top             =   9015
      Width           =   2895
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Histórico del CE"
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
      Width           =   14865
   End
End
Attribute VB_Name = "frmHistoricoDeterminacionCE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public ID_MUESTRA As Long
Public LIMITE_INFERIOR As String
Public LIMITE_SUPERIOR As String
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

Private Sub Form_Load()
    log (Me.Name)
    With OLE2
      .Format = "CF_TEXT"
      .SizeMode = vbOLESizeStretch
    '  .CreateEmbed "", "MSGRAPH"
    End With
    
    With lista.ColumnHeaders
        .Add , , "ID_MUESTRA", 1, lvwColumnLeft
        .Add , , "Designacion", 800, lvwColumnCenter
        .Add , , "Probeta", 300, lvwColumnCenter
        .Add , , "Area", 300, lvwColumnCenter
        .Add , , "Fecha", 1000, lvwColumnCenter
        .Add , , "Res.", 700, lvwColumnRight
    End With
    'Tipos de graficos
    cargar_botones Me
    cargar_lista
End Sub

Private Sub cargar_lista()
    Dim oCe_resultados As New clsCe_resultados
    Dim rs As New ADODB.Recordset
   On Error GoTo cargar_lista_Error

    Set rs = oCe_resultados.resultadosHistorico(ID_MUESTRA)
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        txtnum = rs.RecordCount
        cambiar.min = 1
        cambiar.Max = rs.RecordCount
        lbltitulo(3) = "Histórico del CE"
        Do
           With lista.ListItems.Add(, , rs(0)) ' ID
            .SubItems(1) = rs(1) ' DESIGNACION
            .SubItems(2) = rs(2) ' PROBETA
            .SubItems(3) = rs(3) ' AREA
            .SubItems(4) = Format(rs(4), "dd-mm-yyyy") ' FECHA
            
            If (rs(7) = TIPOS_MUESTRAS.ESPESOR Or _
               rs(7) = TIPOS_MUESTRAS.DUREZA_VICKERS Or _
               rs(7) = TIPOS_MUESTRAS.RUGOSIDAD) And rs(5) <> "" Then
               Dim res() As String
               res = Split(rs(5), "-")
               If UBound(res) > 0 Then
                   .SubItems(5) = res(1) ' RESULTADO
                End If
            Else
                If rs(5) <> "" Then
                    .SubItems(5) = rs(5) ' RESULTADO
                Else
                    .SubItems(5) = rs(6) ' CONFORME
                End If
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

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmHistoricoDeterminacionCE"
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
    Dim Msg, linea1, linea2, linea3, linea4, NewLine, Tabb
    
    Tabb = vbTab
    NewLine = vbNewLine
    linea1 = "Muestra" & Tabb
    linea2 = "Resultado" & Tabb
'    linea3 = "Lim.Inf" & Tabb
'    linea4 = "Lim.Sup" & Tabb
    For i = lista.ListItems.Count To 1 Step -1
        If lista.ListItems(i).Checked = True Then
            If i <= CInt(txtnum) Then
                If IsNumeric(lista.ListItems(i).SubItems(5)) Then
                   linea1 = linea1 & lista.ListItems(i).Text & Tabb
                   linea2 = linea2 & Replace(lista.ListItems(i).SubItems(5), ".", ",") & Tabb
'                   linea3 = linea3 & Replace(LIMITE_INFERIOR, ".", ",") & Tabb
'                   linea4 = linea4 & Replace(LIMITE_SUPERIOR, ".", ",") & Tabb
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

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_grafico of Formulario frmHistoricoDeterminacionCE"
End Sub
Private Sub txtnum_Change()
    cargar_grafico
End Sub

