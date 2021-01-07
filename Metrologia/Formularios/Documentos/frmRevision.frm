VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRevision 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado Comparativo de Revisión de Facturación"
   ClientHeight    =   8865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10050
   Icon            =   "frmRevision.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   10050
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Estadísticas"
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
      Left            =   90
      TabIndex        =   9
      Top             =   8010
      Width           =   8745
      Begin VB.TextBox txttotal 
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   3
         Left            =   7830
         Locked          =   -1  'True
         TabIndex        =   23
         Top             =   225
         Width           =   825
      End
      Begin VB.TextBox txttotal 
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
         ForeColor       =   &H000000FF&
         Height          =   375
         Index           =   2
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   225
         Width           =   825
      End
      Begin VB.TextBox txttotal 
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
         Height          =   375
         Index           =   1
         Left            =   3555
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   225
         Width           =   825
      End
      Begin VB.TextBox txttotal 
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
         Height          =   375
         Index           =   0
         Left            =   1530
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   225
         Width           =   825
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "No en Excel"
         Height          =   285
         Index           =   3
         Left            =   6750
         TabIndex        =   24
         Top             =   315
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "O.T. Erróneas"
         Height          =   285
         Index           =   2
         Left            =   4545
         TabIndex        =   14
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "O.T. Correctas"
         Height          =   285
         Index           =   1
         Left            =   2430
         TabIndex        =   12
         Top             =   315
         Width           =   1320
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "O.T. Registradas"
         Height          =   285
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   315
         Width           =   1320
      End
   End
   Begin VB.TextBox txtwork 
      Height          =   330
      Left            =   6345
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   8280
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seleccione documento a comparar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1500
      Index           =   0
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   10005
      Begin VB.CommandButton cmdComparar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Comparar"
         Height          =   825
         Left            =   9000
         Picture         =   "frmRevision.frx":3AFA
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   405
         Width           =   915
      End
      Begin VB.CommandButton cmdEXplorar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Explorar"
         Height          =   330
         Index           =   0
         Left            =   8100
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         Width           =   780
      End
      Begin VB.TextBox datos 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   810
         Locked          =   -1  'True
         TabIndex        =   0
         Top             =   1080
         Width           =   7200
      End
      Begin MSDataListLib.DataCombo cmbclientes 
         Height          =   315
         Left            =   810
         TabIndex        =   16
         Top             =   270
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         BackColor       =   12640511
         ListField       =   ""
         BoundColumn     =   ""
         Text            =   ""
         Object.DataMember      =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker desde 
         Height          =   345
         Left            =   810
         TabIndex        =   18
         Top             =   630
         Width           =   1320
         _ExtentX        =   2328
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
         CalendarBackColor=   12640511
         CalendarTitleBackColor=   12640511
         Format          =   16842753
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker hasta 
         Height          =   345
         Left            =   2745
         TabIndex        =   21
         Top             =   630
         Width           =   1320
         _ExtentX        =   2328
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
         CalendarBackColor=   12640511
         CalendarTitleBackColor=   12640511
         Format          =   16842753
         CurrentDate     =   38002
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   5
         Left            =   2250
         TabIndex        =   20
         Top             =   720
         Width           =   390
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Desde"
         Height          =   195
         Index           =   6
         Left            =   135
         TabIndex        =   19
         Top             =   720
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   8
         Left            =   135
         TabIndex        =   17
         Top             =   315
         Width           =   480
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ruta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   135
         TabIndex        =   5
         Top             =   1125
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   825
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8010
      Width           =   1005
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8460
      Top             =   8145
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      Filter          =   "Hojas Excel (*.xls) | *.xls"
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5700
      Left            =   45
      TabIndex        =   3
      Top             =   2250
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   10054
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
      BackColor       =   14609914
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7740
      Top             =   8145
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":43C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":4C9E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":5578
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRevision.frx":5E52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblsub 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Maersk y Safmarine"
      Height          =   240
      Left            =   45
      TabIndex        =   7
      Top             =   360
      Width           =   8115
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   9450
      Picture         =   "frmRevision.frx":74AC
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      BackColor       =   &H00FFFFFF&
      Caption         =   " Listado Comparativo de Revisión de Facturación"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   -45
      TabIndex        =   6
      Top             =   0
      Width           =   10050
   End
End
Attribute VB_Name = "frmRevision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdComparar_Click()
    If validar Then
        cargar_lista
    End If
End Sub

Private Sub cmdEXplorar_Click(Index As Integer)
    On Error Resume Next
    cd.DialogTitle = "Abrir hoja Excel"
    cd.InitDir = App.Path
    cd.ShowOpen
    If cd.FileName <> "" Then
        datos(0).Text = cd.FileName
        datos(1) = ""
        datos(2) = ""
        datos(3) = ""
        datos(3).SetFocus
        
    End If
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.Left = 50
    Me.Top = 50
    cargar_botones Me
    cabecera
    desde = Date
    hasta = Date
    cargar_combo cmbclientes, New clsCliente
    cargar_lista
End Sub

Public Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Icono", 300, lvwColumnLeft
        .Add , , "WorkOrder", 1500, lvwColumnCenter
        .Add , , "Booking", 1500, lvwColumnCenter
        .Add , , "C.Size", 1200, lvwColumnCenter
        .Add , , "C.Number", 1500, lvwColumnCenter
        .Add , , "Fecha", 1200, lvwColumnCenter
        .Add , , "Importe", 1200, lvwColumnRight
        .Add , , "I.Registrado", 1200, lvwColumnRight
        .Add , , "ID_OT", 1, lvwColumnLeft
    End With
End Sub
Public Sub cargar_lista()
    If Trim(datos(0)) <> "" Then
     If Dir(datos(0)) <> "" Then
        lista.ListItems.Clear
        txttotal(0) = "0"
        txttotal(1) = "0"
        txttotal(2) = "0"
        txttotal(3) = "0"
        Me.MousePointer = 11
        On Error GoTo fallo
        Dim fila As Integer
        Dim XLA As Excel.Application
        Dim XLW As Excel.Workbook
        Dim XLS As Excel.Worksheet
        Set XLA = New Excel.Application
        Set XLW = XLA.Workbooks.Open(datos(0))
        Set XLS = XLW.Worksheets(1)
        Dim oOT As New clsOt
        Dim total As Currency
        fila = 1
        Do
            If XLS.Cells(fila, 1) <> "" Then
                If XLS.Cells(fila, 8) <> "" And IsNumeric(XLS.Cells(fila, 8)) Then
                    If fila > 3 And XLS.Cells(fila, 8) = XLS.Cells(fila - 1, 8) Then
                        total = lista.ListItems(lista.ListItems.Count).SubItems(6)
                        lista.ListItems(lista.ListItems.Count).SubItems(6) = moneda(total + CCur(XLS.Cells(fila, 4)))
                    Else
                        With lista.ListItems.Add(, , "0")
                             .SubItems(1) = XLS.Cells(fila, 8) ' Workorder
                             .SubItems(2) = XLS.Cells(fila, 10) ' Booking
                             .SubItems(3) = XLS.Cells(fila, 5) ' Size
                             .SubItems(4) = XLS.Cells(fila, 7) ' Number
                             .SubItems(6) = moneda(XLS.Cells(fila, 4))
                             If oOT.Carga_Workorder(XLS.Cells(fila, 8)) Then
                                 .SubItems(5) = Format(oOT.getFECHA_SERVICIO, "dd-mm-yyyy") ' Fecha
                                 .SubItems(7) = moneda(oOT.Total_OT(oOT.getID_OT))
                                 .SubItems(8) = oOT.getID_OT
                             Else
                                 .SubItems(7) = ""
                                 .SubItems(8) = "0"
                             End If
                        End With
                    End If
                    If lista.ListItems(lista.ListItems.Count).SubItems(6) = lista.ListItems(lista.ListItems.Count).SubItems(7) Then
                         lista.ListItems(lista.ListItems.Count).SmallIcon = 2
                         lista.ListItems(lista.ListItems.Count).Text = "2"
'                         txttotal(1) = CInt(txttotal(1)) + 1
                    Else
                         lista.ListItems(lista.ListItems.Count).SmallIcon = 1
                         lista.ListItems(lista.ListItems.Count).Text = "1"
'                         txttotal(2) = CInt(txttotal(2)) + 1
                    End If
                    lista.ListItems(lista.ListItems.Count).EnsureVisible
                    DoEvents
                End If
            End If
            fila = fila + 1
'            txttotal(0) = CInt(txttotal(0)) + 1
        Loop Until XLS.Cells(fila, 1) = ""
        XLA.Quit
        ' Comparamos las que no estan en excel
        ' Sacamos las ot en la base de datos por cliente y fechas
        ' Miramos si esta en la lista, sino, la añadimos y le ponemos una exclamacion
        Dim rs As ADODB.Recordset
        Set rs = oOT.Listado_Para_Comparativo(cmbclientes.BoundText, desde.Value, hasta.Value)
        Dim h As Integer
        Dim encontrado As Boolean
        If rs.RecordCount > 0 Then
            Do
                encontrado = False
                For h = 1 To lista.ListItems.Count
                    If lista.ListItems(h).SubItems(8) = rs(0) Then
                        encontrado = True
                    End If
                Next
                If Not encontrado Then
                    If oOT.Carga(rs(0)) Then
                        With lista.ListItems.Add(, , "4")
                             .SubItems(1) = oOT.getWORKORDER  ' Workorder
                             .SubItems(2) = oOT.getBOOKING_REFERENCE  ' Booking
                             .SubItems(6) = " "
                             .SubItems(7) = moneda(oOT.Total_OT(oOT.getID_OT))
                             .SubItems(8) = oOT.getID_OT
                        End With
                        lista.ListItems(lista.ListItems.Count).SmallIcon = 4
                        lista.ListItems(lista.ListItems.Count).EnsureVisible
                    End If
                End If
                rs.MoveNext
            Loop Until rs.EOF
        End If
        estadisticas
        Set XLW = Nothing
        Set XLA = Nothing
        MsgBox "Comparativo generado correctamente.", vbInformation, App.Title
        Me.MousePointer = 0
     End If
    End If
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al generar el listado comparativo." & Err.Description, vbCritical, App.Title
End Sub
Private Function validar() As Boolean
    validar = False
    If cmbclientes.Text = "" Then
        MsgBox "Seleccione un cliente.", vbExclamation, App.Title
        Exit Function
    End If
    If datos(0) = "" Then
        MsgBox "Escriba una ruta.", vbInformation, App.Title
        Exit Function
    End If
    If Dir(datos(0)) = "" Then
        MsgBox "La ruta introducida no existe.", vbInformation, App.Title
        Exit Function
    End If
    validar = True
End Function

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index).SubItems(8) <> "" Then
            frmOt.PK = lista.ListItems(lista.SelectedItem.Index).SubItems(8)
            frmOt.Show 1
        End If
    End If
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
Private Sub estadisticas()
    Dim i As Integer
    txttotal(1) = "0"
    txttotal(2) = "0"
    txttotal(3) = "0"
    For i = 1 To lista.ListItems.Count
      Select Case lista.ListItems(i).SmallIcon
        Case 2
         txttotal(1) = CInt(txttotal(1)) + 1
        Case 1
         txttotal(2) = CInt(txttotal(2)) + 1
        Case 4
         txttotal(3) = CInt(txttotal(3)) + 1
      End Select
    Next
    txttotal(0) = lista.ListItems.Count
End Sub
