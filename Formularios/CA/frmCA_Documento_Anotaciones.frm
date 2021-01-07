VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#13.2#0"; "Codejock.ReportControl.v13.2.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmCA_Documento_Anotaciones 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Anotaciones del documento"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   Icon            =   "frmCA_Documento_Anotaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   7200
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   11865
      _Version        =   851970
      _ExtentX        =   20929
      _ExtentY        =   12700
      _StockProps     =   64
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      FullColumnScrolling=   -1  'True
      InitialSelectionEnable=   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Anotación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1845
      Left            =   60
      TabIndex        =   3
      Top             =   7230
      Width           =   9645
      Begin VB.TextBox txttexto 
         BackColor       =   &H00FFFFFF&
         Height          =   1545
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   210
         Width           =   7905
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirCalibracion 
         Height          =   435
         Left            =   8220
         TabIndex        =   5
         Top             =   240
         Width           =   1305
         _Version        =   851970
         _ExtentX        =   2302
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmCA_Documento_Anotaciones.frx":08CA
      End
      Begin XtremeSuiteControls.PushButton cmdEliminarCalibracion 
         Height          =   435
         Left            =   8220
         TabIndex        =   6
         Top             =   1320
         Width           =   1305
         _Version        =   851970
         _ExtentX        =   2302
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmCA_Documento_Anotaciones.frx":712C
      End
      Begin XtremeSuiteControls.PushButton PushButton1 
         Height          =   435
         Left            =   8220
         TabIndex        =   7
         Top             =   780
         Width           =   1305
         _Version        =   851970
         _ExtentX        =   2302
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Modificar"
         Appearance      =   5
         Picture         =   "frmCA_Documento_Anotaciones.frx":D98E
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   9795
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8205
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   10875
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8205
      Width           =   1050
   End
End
Attribute VB_Name = "frmCA_Documento_Anotaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_ID As Long

Private Sub cmdAnadirCalibracion_Click()
    If Trim(txttexto) <> "" Then
        Dim oDA As New clsCa_documentos_anotaciones
        With oDA
            .setDOCUMENTO_ID = PK_ID
            .setTEXTO = txttexto
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setFECHA = Format(Date, "yyyy-mm-dd")
            'M1380-I
            '.setORDEN = wndReportControl.Records.Count
            .valorMaximoOrden (PK_ID)
            'M1380-f
            .Insertar
        End With
        Set oDA = Nothing
        cargar_lista
        txttexto = ""
        txttexto.SetFocus
    End If
End Sub

Private Sub cmdEliminarCalibracion_Click()
    If wndReportControl.SelectedRows.Count = 0 Then
        MsgBox "No hay nada seleccionado.", vbCritical, App.Title
    Else
        Dim oCA As New clsCa_documentos_anotaciones
        oCA.Eliminar PK_ID, wndReportControl.SelectedRows(0).Record.Item(3).Caption
        Set oCA = Nothing
        cargar_lista
        txttexto = ""
        txttexto.SetFocus
    End If
End Sub

Private Sub cmdImprimir_Click()
    If wndReportControl.Records.Count = 0 Then
        Exit Sub
    End If
    wndReportControl.PrintOptions.Header.TextCenter = True
    wndReportControl.PrintOptions.Header.Font.bold = True
    wndReportControl.PrintOptions.Header.Font.Size = 11
    wndReportControl.PrintOptions.Header.FormatString = Me.Caption
    wndReportControl.PrintPreviewOptions.Title = "Impresión de anotaciones..."
    wndReportControl.PrintOptions.Footer.FormatString = Now
    wndReportControl.PrintPreview True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargar_lista
End Sub
Private Sub cabecera()
    wndReportControl.AllowColumnRemove = False
    
    Dim Column As ReportColumn, RecordControl As ReportRecord, RecordHeader As ReportRecord, RecordPaintManager As ReportRecord
    
    Set Column = wndReportControl.Columns.Add(0, "Fecha", 250, True)
    Column.Alignment = xtpAlignmentCenter
    Set Column = wndReportControl.Columns.Add(1, "Usuario", 250, True)
    Column.Alignment = xtpAlignmentCenter
    Set Column = wndReportControl.Columns.Add(2, "Texto", 1000, True)
    Column.Alignment = xtpAlignmentIconLeft
    Set Column = wndReportControl.Columns.Add(3, "Orden", 0, False)
    Column.Alignment = xtpAlignmentIconCenter

    Dim oCA As New clsCa_documentos
    oCA.Carga PK_ID
    Me.Caption = "Anotaciones del documento : " & oCA.getNOMBRE
    Set oCA = Nothing
End Sub
Private Sub cargar_lista()
    wndReportControl.Records.DeleteAll
    Dim ohc As New clsCa_documentos_anotaciones
    Dim RS As ADODB.Recordset
    Set RS = ohc.Listado(PK_ID)
    Dim c As Integer
    If RS.RecordCount > 0 Then
        c = RS.RecordCount
        Do
            AddGroupRecord wndReportControl.Records, RS(0), RS(1), RS(2), RS(3)
            RS.MoveNext
            c = c - 1
        Loop Until RS.EOF
    End If
    Set RS = Nothing
    Set ohc = Nothing
    
    wndReportControl.PaintManager.FixedRowHeight = False
    wndReportControl.Populate
End Sub

Private Sub Form_Resize()
'    On Error Resume Next
'    wndReportControl.Move 0, 0, ScaleWidth, ScaleHeight - 900
'    cmdsalir.Left = Me.Width - 1200
'    cmdsalir.Top = Me.Height - 1280
'
'    cmdImprimir.Left = Me.Width - 2300
'    cmdImprimir.Top = Me.Height - 1280
End Sub

Private Sub AddGroupRecord(Records As ReportRecords, fecha As String, USUARIO As String, texto As String, ORDEN As Integer)
    Dim Record As ReportRecord
    Dim Item As ReportRecordItem
    
    Set Record = wndReportControl.Records.Add()
    Record.AddItem fecha
    Record.AddItem USUARIO
    Record.AddItem texto
    Record.AddItem ORDEN
'    Set Item = Record.AddItem("")
'    Item.Focusable = False
End Sub

Private Sub PushButton1_Click()
    If wndReportControl.SelectedRows.Count = 0 Then
        MsgBox "No hay nada seleccionado.", vbCritical, App.Title
    Else
        Dim oCA As New clsCa_documentos_anotaciones
        oCA.setUSUARIO_ID = USUARIO.getID_EMPLEADO
        oCA.setFECHA = Format(Date, "yyyy-mm-dd")
        oCA.setTEXTO = txttexto
        oCA.Modificar PK_ID, wndReportControl.SelectedRows(0).Record.Item(3).Caption
        Set oCA = Nothing
        cargar_lista
        txttexto = ""
        txttexto.SetFocus
    End If

End Sub

Private Sub txttexto_GotFocus()
    txttexto.BackColor = &HC0E0FF
End Sub

Private Sub txttexto_LostFocus()
    txttexto.BackColor = vbWhite
End Sub

Private Sub wndReportControl_SelectionChanged()
    txttexto = wndReportControl.SelectedRows(0).Record.Item(2).Caption
End Sub
