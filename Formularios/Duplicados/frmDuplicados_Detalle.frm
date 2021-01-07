VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frmDuplicados_Detalle 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Control de duplicados"
   ClientHeight    =   9390
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   10065
   DrawWidth       =   10
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9390
   ScaleWidth      =   10065
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRevisada 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Marcar como revisadas"
      Height          =   870
      Left            =   45
      Picture         =   "frmDuplicados_Detalle.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8505
      Width           =   2445
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   45
      TabIndex        =   10
      Top             =   360
      Width           =   9960
      Begin VB.TextBox txtdeter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   1305
         TabIndex        =   18
         Top             =   225
         Width           =   8580
      End
      Begin VB.TextBox txtdif 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   330
         Left            =   8730
         TabIndex        =   16
         Top             =   630
         Width           =   1110
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   360
         Left            =   1305
         TabIndex        =   14
         Top             =   630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   61734913
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   360
         Left            =   4950
         TabIndex        =   15
         Top             =   630
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   14737632
         Format          =   61734913
         CurrentDate     =   38000
         MinDate         =   2
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "% Dif. Duplicados"
         Height          =   195
         Index           =   23
         Left            =   7425
         TabIndex        =   17
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Determinación"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha hasta"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   3
         Left            =   3915
         TabIndex        =   12
         Top             =   720
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha desde"
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   11
         Top             =   735
         Width           =   930
      End
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   8910
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   8505
      Width           =   1050
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   7830
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8505
      Width           =   1050
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
      Left            =   3510
      TabIndex        =   3
      Top             =   8550
      Width           =   4155
      Begin VB.ComboBox cmbTipo 
         Height          =   315
         ItemData        =   "frmDuplicados_Detalle.frx":08CA
         Left            =   1665
         List            =   "frmDuplicados_Detalle.frx":08D7
         TabIndex        =   6
         Top             =   270
         Width           =   2400
      End
      Begin VB.OptionButton opd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "3D"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   495
         Width           =   690
      End
      Begin VB.OptionButton opd 
         BackColor       =   &H00C0C0C0&
         Caption         =   "2D"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
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
         TabIndex        =   7
         Top             =   315
         Width           =   315
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   3015
      Left            =   45
      TabIndex        =   0
      Top             =   5490
      Width           =   9930
      _ExtentX        =   17515
      _ExtentY        =   5318
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
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
      Height          =   3945
      Left            =   90
      OleObjectBlob   =   "frmDuplicados_Detalle.frx":08EF
      TabIndex        =   2
      Top             =   1485
      Width           =   9900
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDuplicados_Detalle.frx":2677
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Control de duplicados"
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
      Width           =   10095
   End
End
Attribute VB_Name = "frmDuplicados_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_ID_TIPO_DETERMINACION As Long
Public PK_DIF As String
Public fecha_desde As String
Public fecha_hasta As String
Private Sub cmbTipo_Click()
    tipo_grafico
End Sub

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

Private Sub cmdRevisada_Click()
    If lista.ListItems.Count > 0 Then
        Dim oDET As New clsDeterminaciones
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            oDET.marcar_como_revisada CLng(lista.ListItems(i).Text)
        Next
        MsgBox "Determinaciones revisadas correctamente.", vbOKOnly + vbInformation, App.Title
        cargar_lista
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fdesde_Change()
     cargar_lista
End Sub
Private Sub fhasta_Change()
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cabecera
    'Tipos de graficos
    If fecha_desde <> "" Then
        fdesde = fecha_desde
    Else
        fdesde = Date - 30
    End If
    If fecha_hasta <> "" Then
        fhasta = fecha_hasta
    Else
        fhasta = Date
    End If
    cargar_botones Me
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim oMuestra As New clsMuestra
    Dim rs As New ADODB.Recordset
    If PK_DIF <> "" Then
        txtdif = PK_DIF
    Else
        txtdif = "0"
    End If
    Set rs = oMuestra.historico_determinaciones(PK_ID_TIPO_DETERMINACION, fdesde, fhasta)
    Dim objSI As ListSubItem
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        txtdeter = rs(6)
        lbltitulo(3) = "Control de duplicados de la determinación : " & rs(6)
        Do
            b = False
           If IsNumeric(rs(5)) Then
            If CSng(Replace(rs(5), ".", ",")) < 0 Or CSng(Format(Replace(rs(5), ".", ","), "#.0")) > CSng(Format(Replace(txtdif, ".", ","), "#.0")) Then
                 b = True
            Else
                 b = False
            End If
           End If
           With lista.ListItems.Add(, , rs(1))
            Set objSI = .ListSubItems.Add(, , rs(2))
            If b Then
                objSI.bold = True
                objSI.ForeColor = RGB(255, 0, 0)
            End If
            Set objSI = .ListSubItems.Add(, , Format(rs(7), "dd-mm-yyyy"))
            If b Then
                objSI.bold = True
                objSI.ForeColor = RGB(255, 0, 0)
            End If
            Set objSI = .ListSubItems.Add(, , rs(4))
            If b Then
                objSI.bold = True
                objSI.ForeColor = RGB(255, 0, 0)
            End If
            Set objSI = .ListSubItems.Add(, , rs(5))
            If b Then
                objSI.bold = True
                objSI.ForeColor = RGB(255, 0, 0)
            End If
            Set objSI = .ListSubItems.Add(, , rs(8))
            If b Then
                objSI.bold = True
                objSI.ForeColor = RGB(255, 0, 0)
            End If
            
'            If rs(9) <> 0 Then
                Set objSI = .ListSubItems.Add(, , rs(10))
                If b Then
                    objSI.bold = True
                    objSI.ForeColor = RGB(255, 0, 0)
                End If
                Set objSI = .ListSubItems.Add(, , Format(rs(11), "dd-mm-yyyy"))
                If b Then
                    objSI.bold = True
                    objSI.ForeColor = RGB(255, 0, 0)
                End If
'            End If
            Set objSI = .ListSubItems.Add(, , rs(0))
           End With
           If b Then
              lista.ListItems(lista.ListItems.Count).SmallIcon = 1
           End If
                
           rs.MoveNext
        Loop Until rs.EOF
    End If
    cargar_grafico
    Set oMuestra = Nothing
    Set rs = Nothing
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
    i = 1
    With grafico
       .columnCount = 1
       .RowCount = lista.ListItems.Count
       For Row = lista.ListItems.Count To 1 Step -1
            .Row = i
            .RowLabel = lista.ListItems(Row).SubItems(1)
            If IsNumeric(lista.ListItems(Row).SubItems(4)) Then
               .Data = Replace(lista.ListItems(Row).SubItems(4), ".", ",")
               
            Else
               .Data = 0
            End If
            i = i + 1
       Next Row
    End With
   On Error GoTo 0
   Exit Sub

cargar_grafico_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_grafico of Formulario frmDuplicados_Detalle"
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).SubItems(8)
        frmVerMuestra.Show 1
    End If
    
End Sub

Private Sub opd_Click(Index As Integer)
    tipo_grafico
End Sub

Private Sub txtnum_Change()
    cargar_grafico
End Sub

Public Sub tipo_grafico()
    If opd(0).Value = True Then
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
    With lista.ColumnHeaders
        .Add , , "ID_DETERMINACION", 240, lvwColumnLeft
        .Add , , "Muestra", 900, lvwColumnCenter
        .Add , , "Fecha", 1150, lvwColumnCenter
        .Add , , "Resultado", 900, lvwColumnRight
        .Add , , "%Dif.Duplicados", 900, lvwColumnRight
        .Add , , "Analista", 2000, lvwColumnCenter
        .Add , , "Revisado por", 2000, lvwColumnCenter
        .Add , , "F.Revisión", 1150, lvwColumnCenter
        .Add , , "ID_MUESTRA", 1, lvwColumnLeft
    End With
End Sub
