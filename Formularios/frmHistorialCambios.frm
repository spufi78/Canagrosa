VERSION 5.00
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#13.2#0"; "Codejock.ReportControl.v13.2.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmHistorialCambios 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Historial de Cambios"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10935
   Icon            =   "frmHistorialCambios.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   10935
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   5850
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   7875
      _Version        =   851970
      _ExtentX        =   13891
      _ExtentY        =   10319
      _StockProps     =   64
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      FullColumnScrolling=   -1  'True
      InitialSelectionEnable=   0   'False
   End
   Begin VB.Frame frmSeguimiento 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seguimiento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   45
      TabIndex        =   3
      Top             =   6300
      Visible         =   0   'False
      Width           =   7890
      Begin VB.TextBox txttexto 
         BackColor       =   &H00FFFFFF&
         Height          =   870
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   210
         Width           =   6285
      End
      Begin XtremeSuiteControls.PushButton cmdAnadirCalibracion 
         Height          =   435
         Left            =   6480
         TabIndex        =   5
         Top             =   180
         Width           =   1305
         _Version        =   851970
         _ExtentX        =   2302
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Añadir"
         Appearance      =   5
         Picture         =   "frmHistorialCambios.frx":08CA
      End
      Begin XtremeSuiteControls.PushButton cmdEliminarCalibracion 
         Height          =   435
         Left            =   6480
         TabIndex        =   6
         Top             =   645
         Visible         =   0   'False
         Width           =   1305
         _Version        =   851970
         _ExtentX        =   2302
         _ExtentY        =   767
         _StockProps     =   79
         Caption         =   "Eliminar"
         Appearance      =   5
         Picture         =   "frmHistorialCambios.frx":712C
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   8730
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6480
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9810
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6480
      Width           =   1050
   End
End
Attribute VB_Name = "frmHistorialCambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_TIPO As Integer
Public PK_ID As Long
Public PK_TITULO As String
Private Sub cmdAnadirCalibracion_Click()
    'M1108-I
    If txttexto <> "" Then
        Dim ohc As New clsHistorial_cambios
        With ohc
            .setTIPO = PK_TIPO
            .setIDENTIFICADOR = PK_ID
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setMOTIVO = "Seguimiento : " & txttexto
            .Insertar
        End With
        Set ohc = Nothing
        txttexto = ""
        Form_Load
    End If
    'M1108-F
End Sub

Private Sub cmdEliminarCalibracion_Click()
    'M1108-I
    If wndReportControl.SelectedRows.Count = 0 Then
        MsgBox "No hay nada seleccionado.", vbCritical, App.Title
    Else
        Dim ohc As New clsHistorial_cambios
        ohc.EliminarIdentificadorTS CLng(PK_TIPO), PK_ID, Format(Left(wndReportControl.SelectedRows(0).Record.Item(1).Caption, 10), "yyyy-mm-dd") & " " & Right(wndReportControl.SelectedRows(0).Record.Item(1).Caption, 8)
        Set ohc = Nothing
        txttexto = ""
        Form_Load
    End If
    'M1108-F
End Sub

Private Sub cmdImprimir_Click()
    wndReportControl.PrintOptions.Header.TextCenter = True
    wndReportControl.PrintOptions.Header.Font.bold = True
    wndReportControl.PrintOptions.Header.Font.Size = 11
    wndReportControl.PrintOptions.Header.FormatString = "Histórico cambios: " & PK_TITULO
    wndReportControl.PrintPreviewOptions.Title = "Impresión de historial de cambios..."
    wndReportControl.PrintOptions.Footer.FormatString = Now
    wndReportControl.PrintPreview True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    'M1108-I
    wndReportControl.Records.DeleteAll
    wndReportControl.Columns.DeleteAll
    'M1108-F
    wndReportControl.AllowColumnRemove = False
    'M1108-I
    If PK_TIPO = HC_TIPOS.HC_OFERTAS Or PK_TIPO = HC_TIPOS.HC_PEDIDOS Then
        frmSeguimiento.visible = True
    End If
    'M1108-F
    Dim Column As ReportColumn, RecordControl As ReportRecord, RecordHeader As ReportRecord, RecordPaintManager As ReportRecord
    
    Set Column = wndReportControl.Columns.Add(0, "Versión", 80, True)
    Column.Alignment = xtpAlignmentCenter
    Set Column = wndReportControl.Columns.Add(1, "Fecha", 250, True)
    Column.Alignment = xtpAlignmentCenter
'    Column.TreeColumn = True
'    Column.Editable = False
    Set Column = wndReportControl.Columns.Add(2, "Usuario", 250, True)
    Column.Alignment = xtpAlignmentCenter
'    Column.Editable = True
'    Column.Icon = 1
'    Column.Sortable = False
    Set Column = wndReportControl.Columns.Add(3, "Motivo", 1000, True)
    Column.Alignment = xtpAlignmentIconLeft
'    Column.Editable = True
    
'    Set RecordControl =
    Dim ohc As New clsHistorial_cambios
    Dim rs As ADODB.Recordset
    Set rs = ohc.Listado(PK_TIPO, PK_ID)
    Dim c As Integer
    If rs.RecordCount > 0 Then
        c = rs.RecordCount
        Do
            AddGroupRecord wndReportControl.Records, rs(0), rs(1), rs(2), c
            rs.MoveNext
            c = c - 1
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set ohc = Nothing
    ' Set other options
'    wndReportControl.PaintManager.ColumnStyle = xtpColumnFlat
'    wndReportControl.AllowEdit = False
'    wndReportControl.EditOnClick = False
'    wndReportControl.MultipleSelection = False
'    wndReportControl.AllowColumnSort = False
    wndReportControl.PaintManager.FixedRowHeight = False
'    wndReportControl.PaintManager.TreeIndent = 10
    ' Apply changes
    wndReportControl.Populate
    If PK_TIPO = HC_TIPOS.HC_OFERTAS Or PK_TIPO = HC_TIPOS.HC_PEDIDOS Then
        cmdEliminarCalibracion.visible = True
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    wndReportControl.Move 0, 0, ScaleWidth, ScaleHeight - 1150
    cmdSalir.Left = Me.Width - 1200
    cmdSalir.top = Me.Height - 1280
    
    cmdImprimir.Left = Me.Width - 2300
    cmdImprimir.top = Me.Height - 1280
End Sub

Private Sub AddGroupRecord(Records As ReportRecords, fecha As String, USUARIO As String, MOTIVO As String, version As Integer)
    Dim Record As ReportRecord
    Dim Item As ReportRecordItem
    
    Set Record = wndReportControl.Records.Add()
    Record.AddItem version
    Record.AddItem fecha
    Record.AddItem USUARIO
    Record.AddItem MOTIVO
'    Set Item = Record.AddItem("")
'    Item.Focusable = False
End Sub

