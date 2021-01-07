VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#13.2#0"; "Codejock.ReportControl.v13.2.1.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmMuestras_Ediciones 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Ediciones de la Muestra"
   ClientHeight    =   9120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11925
   Icon            =   "frmMuestras_Ediciones.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin XtremeReportControl.ReportControl wndReportControl 
      Height          =   6930
      Left            =   45
      TabIndex        =   0
      Top             =   -15
      Width           =   11865
      _Version        =   851970
      _ExtentX        =   20929
      _ExtentY        =   12224
      _StockProps     =   64
      MultipleSelection=   0   'False
      EditOnClick     =   0   'False
      FullColumnScrolling=   -1  'True
      InitialSelectionEnable=   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
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
      Begin VB.TextBox txtedicion 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   915
         TabIndex        =   8
         Top             =   180
         Width           =   585
      End
      Begin VB.TextBox txttexto 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   1185
         Left            =   915
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   570
         Width           =   7095
      End
      Begin XtremeSuiteControls.PushButton cmdAnadir 
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
         Picture         =   "frmMuestras_Ediciones.frx":08CA
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
         Picture         =   "frmMuestras_Ediciones.frx":712C
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
         Picture         =   "frmMuestras_Ediciones.frx":D98E
      End
      Begin MSComCtl2.UpDown btnEdiciones 
         Height          =   330
         Left            =   1501
         TabIndex        =   9
         Top             =   180
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   582
         _Version        =   393216
         Value           =   99
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtedicion"
         BuddyDispid     =   196610
         OrigLeft        =   1740
         OrigTop         =   180
         OrigRight       =   1980
         OrigBottom      =   510
         Max             =   99
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.DTPicker fecha 
         Height          =   330
         Left            =   2700
         TabIndex        =   13
         Top             =   180
         Width           =   1350
         _ExtentX        =   2381
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
         CalendarTitleBackColor=   12632256
         Format          =   52428801
         CurrentDate     =   38002
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha"
         Height          =   225
         Index           =   2
         Left            =   2115
         TabIndex        =   14
         Top             =   225
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Motivo"
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   11
         Top             =   990
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edición"
         Height          =   225
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   240
         Width           =   615
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
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "* Introduzca en cada edición, todos los motivos de su generación. Motivo 1, pulse añadir, Motivo 2, pulse añadir...."
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
      Height          =   195
      Left            =   135
      TabIndex        =   12
      Top             =   6975
      Width           =   11715
   End
End
Attribute VB_Name = "frmMuestras_Ediciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public PK As Long

Private Sub cmdAnadir_Click()
   On Error GoTo cmdAnadir_Click_Error

    If Len(txttexto) < 3 Then
        MsgBox "Indique un texto bien descrito con el cambio realizado.", vbExclamation, App.Title
    Else
        Dim oME As New clsMuestras_ediciones
        With oME
            .setMUESTRA_ID = PK
            .setEDICION = txtedicion
            .setOBSERVACIONES = Trim(txttexto)
            .setFECHA = fecha
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            If .Insertar(False) > 0 Then
                cargar_lista
                txttexto = ""
                txttexto.SetFocus
            Else
                MsgBox "Error al insertar el registro.", vbCritical, App.Title
            End If
        End With
    End If

   On Error GoTo 0
   Exit Sub

cmdAnadir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAnadir_Click of Formulario frmMuestras_Ediciones"
End Sub

Private Sub cmdEliminarCalibracion_Click()
   On Error GoTo cmdEliminarCalibracion_Click_Error

    If wndReportControl.SelectedRows.Count = 0 Then
        MsgBox "No hay ningún registro seleccionado.", vbCritical, App.Title
    Else
        Dim oME As New clsMuestras_ediciones
        oME.Eliminar wndReportControl.SelectedRows(0).Record.Item(4).Caption
        Set oME = Nothing
        cargar_lista
        txttexto = ""
        txttexto.SetFocus
    End If

   On Error GoTo 0
   Exit Sub

cmdEliminarCalibracion_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEliminarCalibracion_Click of Formulario frmMuestras_Ediciones"
End Sub

Private Sub cmdImprimir_Click()
    If wndReportControl.Records.Count = 0 Then
        MsgBox "No existen registros.", vbExclamation, App.Title
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
    fecha = Date
    cargar_lista
End Sub
Private Sub cabecera()
    wndReportControl.AllowColumnRemove = False
    
    Dim Column As ReportColumn, RecordControl As ReportRecord, RecordHeader As ReportRecord, RecordPaintManager As ReportRecord
    
    Set Column = wndReportControl.Columns.Add(0, "Edición", 100, True)
    Column.Alignment = xtpAlignmentCenter
    Set Column = wndReportControl.Columns.Add(1, "Fecha", 150, True)
    Column.Alignment = xtpAlignmentCenter
    Set Column = wndReportControl.Columns.Add(2, "Usuario", 250, True)
    Column.Alignment = xtpAlignmentCenter
    Set Column = wndReportControl.Columns.Add(3, "Motivo", 1000, True)
    Column.Alignment = xtpAlignmentIconLeft
    Set Column = wndReportControl.Columns.Add(4, "ID", 0, True)
    Column.Alignment = xtpAlignmentIconCenter

    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra PK
    Me.Caption = "Ediciones de la muestra : " & oMuestra.getTITULO_MUESTRA
    
    btnEdiciones.min = 2
    If oMuestra.getCERRADA <> 1 Then
        txtedicion.Text = oMuestra.getULT_EDICION_IMP + 1
        btnEdiciones.Max = oMuestra.getULT_EDICION_IMP + 1
    Else
        txtedicion.Text = oMuestra.getULT_EDICION_IMP
        btnEdiciones.Max = oMuestra.getULT_EDICION_IMP
    End If
    btnEdiciones.Value = txtedicion.Text
    Set oMuestra = Nothing
End Sub
Private Sub cargar_lista()
    wndReportControl.Records.DeleteAll
    Dim oME As New clsMuestras_ediciones
    Dim rs As ADODB.Recordset
    Set rs = oME.Listado(PK)
    Dim c As Integer
    If rs.RecordCount > 0 Then
        c = rs.RecordCount
        Do
            AddGroupRecord wndReportControl.Records, rs("EDICION"), rs("FECHA"), rs("USUARIO_ID"), rs("OBSERVACIONES"), rs("ID")
            rs.MoveNext
            c = c - 1
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oME = Nothing
    
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

Private Sub AddGroupRecord(Records As ReportRecords, EDICION As Integer, fecha As String, USUARIO As String, texto As String, ID As Long)
    Dim Record As ReportRecord
    Dim Item As ReportRecordItem
    
    Set Record = wndReportControl.Records.Add()
    Record.AddItem EDICION
    Record.AddItem fecha
    Record.AddItem USUARIO
    Record.AddItem texto
    Record.AddItem ID
'    Set Item = Record.AddItem("")
'    Item.Focusable = False
End Sub

Private Sub PushButton1_Click()
   On Error GoTo PushButton1_Click_Error

    If wndReportControl.SelectedRows.Count = 0 Then
        MsgBox "No hay ningún registro seleccionado.", vbCritical, App.Title
    Else
        If Len(txttexto) < 3 Then
            MsgBox "Indique un texto bien descrito con el cambio realizado.", vbExclamation, App.Title
        Else
            Dim oME As New clsMuestras_ediciones
            With oME
                .setMUESTRA_ID = PK
                .setEDICION = txtedicion
                .setOBSERVACIONES = Trim(txttexto)
                .setFECHA = fecha
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                If .Modificar(wndReportControl.SelectedRows(0).Record.Item(4).Caption) = True Then
                    cargar_lista
                    txttexto = ""
                    txttexto.SetFocus
                Else
                    MsgBox "Error al modificar el registro.", vbCritical, App.Title
                End If
            End With
        End If
    End If

   On Error GoTo 0
   Exit Sub

PushButton1_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure PushButton1_Click of Formulario frmMuestras_Ediciones"
End Sub

Private Sub txttexto_GotFocus()
    txttexto.BackColor = &HC0E0FF
End Sub

Private Sub txttexto_LostFocus()
    txttexto.BackColor = vbWhite
End Sub

Private Sub wndReportControl_SelectionChanged()
    txtedicion = wndReportControl.SelectedRows(0).Record.Item(0).Caption
    fecha = wndReportControl.SelectedRows(0).Record.Item(1).Caption
    txttexto = wndReportControl.SelectedRows(0).Record.Item(3).Caption
End Sub
