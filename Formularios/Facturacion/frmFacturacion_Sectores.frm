VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmFacturacion_Sectores 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturación por familias y sectores"
   ClientHeight    =   12930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10830
   Icon            =   "frmFacturacion_Sectores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12930
   ScaleWidth      =   10830
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExcelFacturas 
      BackColor       =   &H00E0E0E0&
      Caption         =   "FACTURAS"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   12015
      Width           =   1080
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criterio de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2760
      Left            =   45
      TabIndex        =   7
      Top             =   360
      Width           =   10680
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de documento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   7470
         TabIndex        =   34
         Top             =   1350
         Width           =   1905
         Begin VB.CheckBox chkTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Albarán"
            Height          =   285
            Index           =   1
            Left            =   180
            TabIndex        =   37
            Top             =   720
            Visible         =   0   'False
            Width           =   780
         End
         Begin VB.CheckBox chkTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Factura"
            Height          =   285
            Index           =   2
            Left            =   180
            TabIndex        =   36
            Top             =   315
            Width           =   915
         End
         Begin VB.CheckBox chkTipo 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cargo / Abono"
            Height          =   285
            Index           =   3
            Left            =   180
            TabIndex        =   35
            Top             =   720
            Width           =   1365
         End
      End
      Begin VB.Frame frmMensual 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fechas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1320
         Left            =   90
         TabIndex        =   17
         Top             =   1350
         Width           =   7350
         Begin VB.OptionButton optFecha 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   195
            Index           =   1
            Left            =   4185
            TabIndex        =   19
            Top             =   225
            Width           =   240
         End
         Begin VB.OptionButton optFecha 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Option1"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   18
            Top             =   180
            Width           =   240
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00C0C0C0&
            Height          =   1050
            Left            =   45
            TabIndex        =   29
            Top             =   180
            Width           =   4020
            Begin MSComCtl2.DTPicker fechaI 
               Height          =   330
               Left            =   630
               TabIndex        =   30
               Top             =   405
               Width           =   1395
               _ExtentX        =   2461
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
               Format          =   16449537
               CurrentDate     =   38002
            End
            Begin MSComCtl2.DTPicker fechaF 
               Height          =   330
               Left            =   2475
               TabIndex        =   31
               Top             =   405
               Width           =   1395
               _ExtentX        =   2461
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
               Format          =   16449537
               CurrentDate     =   38002
            End
            Begin VB.Label Label1 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Fin"
               Height          =   315
               Index           =   12
               Left            =   2160
               TabIndex        =   33
               Top             =   450
               Width           =   315
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               BackColor       =   &H00C0C0C0&
               Caption         =   "Inicio"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   32
               Top             =   450
               Width           =   375
            End
         End
         Begin VB.Frame Frame5 
            BackColor       =   &H00C0C0C0&
            Height          =   1050
            Left            =   4095
            TabIndex        =   20
            Top             =   180
            Width           =   3165
            Begin VB.CheckBox chkAnyo 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Año completo"
               Height          =   240
               Left            =   1710
               TabIndex        =   26
               Top             =   270
               Width           =   1365
            End
            Begin VB.OptionButton optTrimestre 
               BackColor       =   &H00C0C0C0&
               Caption         =   "1T"
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   25
               Top             =   720
               Width           =   600
            End
            Begin VB.OptionButton optTrimestre 
               BackColor       =   &H00C0C0C0&
               Caption         =   "2T"
               Height          =   195
               Index           =   2
               Left            =   765
               TabIndex        =   24
               Top             =   720
               Width           =   645
            End
            Begin VB.OptionButton optTrimestre 
               BackColor       =   &H00C0C0C0&
               Caption         =   "3T"
               Height          =   195
               Index           =   3
               Left            =   1440
               TabIndex        =   23
               Top             =   720
               Width           =   690
            End
            Begin VB.OptionButton optTrimestre 
               BackColor       =   &H00C0C0C0&
               Caption         =   "4T"
               Height          =   195
               Index           =   4
               Left            =   2115
               TabIndex        =   22
               Top             =   720
               Width           =   510
            End
            Begin VB.TextBox txtanno 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Verdana"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   360
               Left            =   540
               TabIndex        =   21
               Top             =   225
               Width           =   720
            End
            Begin MSComCtl2.UpDown cambiar 
               Height          =   360
               Left            =   1260
               TabIndex        =   27
               Top             =   225
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   635
               _Version        =   393216
               Value           =   2004
               BuddyControl    =   "chkAnyo"
               BuddyDispid     =   196617
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
            Begin VB.Label Label1 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Año"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   28
               Top             =   270
               Width           =   585
            End
         End
      End
      Begin VB.CommandButton cmdGenerar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   870
         Left            =   9495
         Picture         =   "frmFacturacion_Sectores.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   945
         Width           =   1080
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8550
         TabIndex        =   9
         Top             =   270
         Value           =   1  'Checked
         Width           =   780
      End
      Begin VB.CheckBox chkAirbus 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Airbus"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   8550
         TabIndex        =   8
         Top             =   630
         Width           =   780
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   765
         TabIndex        =   10
         Top             =   270
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   609
      End
      Begin pryCombo.miCombo cmbCC 
         Height          =   345
         Left            =   765
         TabIndex        =   13
         Top             =   630
         Width           =   7665
         _ExtentX        =   13520
         _ExtentY        =   609
      End
      Begin MSDataListLib.DataCombo cmbsectores 
         Bindings        =   "frmFacturacion_Sectores.frx":1194
         Height          =   315
         Left            =   765
         TabIndex        =   15
         Top             =   990
         Width           =   6615
         _ExtentX        =   11668
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sector"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   16
         Top             =   1035
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Familia"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   14
         Top             =   675
         Width           =   495
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   11
         Top             =   315
         Width           =   495
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   1620
      TabIndex        =   5
      Top             =   12015
      Visible         =   0   'False
      Width           =   6450
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Generando documento EXCEL. Por favor, espere."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   330
         Index           =   1
         Left            =   675
         TabIndex        =   6
         Top             =   225
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdVerExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "SECTORES"
      Height          =   870
      Left            =   8550
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   12015
      Width           =   1080
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9675
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   12015
      Width           =   1080
   End
   Begin MSComctlLib.ListView lista 
      Height          =   4890
      Left            =   45
      TabIndex        =   1
      Top             =   3150
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   8625
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
   Begin MSComctlLib.ListView listaFacturas 
      Height          =   3240
      Left            =   45
      TabIndex        =   38
      Top             =   8685
      Width           =   10680
      _ExtentX        =   18838
      _ExtentY        =   5715
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   14737632
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
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Facturas asociadas al Alodine"
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
      Index           =   0
      Left            =   45
      TabIndex        =   39
      Top             =   8415
      Width           =   10680
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
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
      Height          =   375
      Left            =   45
      TabIndex        =   3
      Top             =   8010
      Width           =   10695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Facturación por familias y sectores"
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
      Height          =   315
      Index           =   3
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   10920
   End
End
Attribute VB_Name = "frmFacturacion_Sectores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------------------------------------------------
'MANTIS: 1278
'------------------------------------------------------------------------------------------------------

Private Trimestres(4, 2) As String

Private Sub chkAnyo_Click()
    If chkAnyo.Value = Checked Then
        Dim i As Integer
        For i = 1 To 4
            optTrimestre(i).Value = False
        Next
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.Value = Checked Then
        cmbClientes.Limpiar
        cmbClientes.desactivar
        chkAirbus.Enabled = True
    Else
        cmbClientes.activar
        chkAirbus.Enabled = False
    End If
End Sub

Private Sub cmdExcelFacturas_Click()
   
       Dim cadena As String
   On Error GoTo cmdExcelFacturas_Click_Error

       Frame6.Visible = True
       Me.MousePointer = vbHourglass
       Dim rs As New ADODB.Recordset
       rs.Fields.Append "c1", adChar, 250, adFldUpdatable
       rs.Fields.Append "c2", adChar, 250, adFldUpdatable
       rs.Fields.Append "c3", adChar, 250, adFldUpdatable
       rs.Fields.Append "c4", adChar, 250, adFldUpdatable
       rs.Fields.Append "c5", adChar, 250, adFldUpdatable
       rs.Fields.Append "c6", adChar, 250, adFldUpdatable
       rs.Fields.Append "c7", adChar, 250, adFldUpdatable
       rs.Fields.Append "c8", adChar, 250, adFldUpdatable
       rs.Fields.Append "c9", adChar, 250, adFldUpdatable
       rs.Fields.Append "c10", adChar, 250, adFldUpdatable
       rs.Open
       
       Dim i As Integer

       For i = 1 To listaFacturas.ListItems.Count
            rs.AddNew
            rs("c1") = listaFacturas.ListItems(i).SubItems(1)
            rs("c2") = listaFacturas.ListItems(i).SubItems(2)
            rs("c3") = listaFacturas.ListItems(i).SubItems(3)
            rs("c4") = listaFacturas.ListItems(i).SubItems(4)
            rs("c5") = listaFacturas.ListItems(i).SubItems(5)
            rs("c6") = listaFacturas.ListItems(i).SubItems(6)
            rs("c7") = listaFacturas.ListItems(i).SubItems(7)
            rs("c8") = listaFacturas.ListItems(i).SubItems(8)
            rs("c9") = listaFacturas.ListItems(i).SubItems(9)
            rs("c10") = listaFacturas.ListItems(i).SubItems(10)
            rs.Update
        Next i
        
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Facturacion"

        'Cabecera
        With XLS.Range("A1:J1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With XLS.Range("A1:J1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:J1").Borders
            .LineStyle = vbSolid
        End With
        
        XLS.Range("A1:A1").ColumnWidth = 10
        XLS.Range("B1:B1").ColumnWidth = 40
        XLS.Range("C1:C1").ColumnWidth = 15
        XLS.Range("D1:D1").ColumnWidth = 20
        XLS.Range("E1:E1").ColumnWidth = 15
        XLS.Range("F1:F1").ColumnWidth = 15
        XLS.Range("G1:G1").ColumnWidth = 15
        XLS.Range("H1:H1").ColumnWidth = 15
        XLS.Range("I1:I1").ColumnWidth = 15
        XLS.Range("J1:J1").ColumnWidth = 15
        XLS.Cells(1, 1) = "NºDOC"
        XLS.Cells(1, 2) = "Cliente"
        XLS.Cells(1, 3) = "Fecha"
        XLS.Cells(1, 4) = "Pedido"
        XLS.Cells(1, 5) = "Importe"
        XLS.Cells(1, 6) = "Dto%"
        XLS.Cells(1, 7) = "Base"
        XLS.Cells(1, 8) = "IVA%"
        XLS.Cells(1, 9) = "Cuota I.V.A."
        XLS.Cells(1, 10) = "Total"

        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Cells(i, 1) = ClrStr(rs("c1"), False, True, True)
            XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True)
            XLS.Cells(i, 3) = Format(rs("c3"), "yyyy-mm-dd")
            XLS.Cells(i, 4) = "'" & ClrStr(rs("c4"), False, True, True)
            XLS.Cells(i, 5) = moneda_bd(rs("C5"))
            XLS.Cells(i, 6) = moneda_bd(rs("C6"))
            XLS.Cells(i, 7) = moneda_bd(rs("C7"))
            XLS.Cells(i, 8) = moneda_bd(rs("C8"))
            XLS.Cells(i, 9) = moneda_bd(rs("C9"))
            XLS.Cells(i, 10) = moneda_bd(rs("C10"))
            i = i + 1
             
            XLS.Range("A" & i).EntireRow.Insert
            rs.MoveNext
          Loop Until rs.EOF
        End If
        Frame6.Visible = False
        Me.MousePointer = vbNormal
        XLA.Visible = True
        
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cmdExcelFacturas_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExcelFacturas_Click of Formulario frmFacturacion_Sectores"

End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    Me.Left = 50
    Me.top = 50
    cabecera
    cargar_trimestres
    cargar_fechas
    cargar_checks
    cargar_combos
    
    cmbClientes.desactivar
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
'        .Add , , "Tipo", 600, lvwColumnLeft
        .Add , , "Sector", 4000, lvwColumnLeft
        .Add , , "Familia", 4000, lvwColumnLeft
        .Add , , "Total", 1800, lvwColumnRight
        .Add , , "DOCS", 0, lvwColumnLeft
    End With
    With listaFacturas.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "NºDOC", 800, lvwColumnCenter
        .Add , , "Cliente", 3200, lvwColumnLeft
        .Add , , "Fecha", 1200, lvwColumnCenter
        .Add , , "Pedido", 2700, lvwColumnCenter
        .Add , , "Importe", 1100, lvwColumnRight
        .Add , , "Dto%", 500, lvwColumnCenter
        .Add , , "Base", 1100, lvwColumnRight
        .Add , , "IVA%", 500, lvwColumnRight
        .Add , , "Cuota I.V.A.", 1100, lvwColumnRight
        .Add , , "Total", 1200, lvwColumnRight
    End With

End Sub
Private Sub cargar_trimestres()
    Trimestres(1, 1) = "-01-01"
    Trimestres(1, 2) = "-03-31"
    Trimestres(2, 1) = "-04-01"
    Trimestres(2, 2) = "-06-30"
    Trimestres(3, 1) = "-07-01"
    Trimestres(3, 2) = "-09-30"
    Trimestres(4, 1) = "-10-01"
    Trimestres(4, 2) = "-12-31"
End Sub
Private Sub cargar_fechas()
    fechaI = "01/01/" & Year(Date)
    fechaF = Date
    txtanno = Year(Date)
End Sub
Private Sub cargar_checks()
    Dim i As Integer
      'albaranes, facturas, cargo-abonos
    chkTipo(1).Value = 0
    chkTipo(2).Value = 1
    chkTipo(3).Value = 1
    optFecha(0).Value = True 'Ningún trimestre seleccionado
    ANYO = Date              'Año en curso
    
End Sub
Private Sub cargar_combos()
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    cargar_combo cmbsectores, New clsSectores
    llenar_combo cmbCC, New clsFamilias, 0, Me, ""
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oMuestra As New clsMuestra
    Dim fecha_inicio As String
    Dim fecha_fin As String
    Dim SECTOR As Long
    Dim strTipo As String
    Dim Suma As Double
    
    Dim i As Integer
    Dim encontrado As Boolean
   
    'Tipos: se usa el cero como comodín
    '       La cadena se envía formateada a la clase.(ENTVAGJ)
   On Error GoTo cargar_lista_Error

    strTipo = "(0"
    For i = 1 To 3
        If chkTipo(i).Value = 1 Then
            strTipo = strTipo & "," & i
        End If
    Next i
    strTipo = strTipo & ")"
    
   'Sectores
    If cmbsectores.BoundText = "" Then
        SECTOR = 0
    Else
        SECTOR = cmbsectores.BoundText
    End If
    
    'Fechas
    If optFecha(0).Value = True Then 'Fechas libres
        fecha_inicio = Format(fechaI, "yyyy-mm-dd")
        fecha_fin = Format(fechaF, "yyyy-mm-dd")
    Else
        If chkAnyo.Value = 1 Then
           fecha_inicio = Trim(txtanno.Text) & "-01-01"
           fecha_fin = Trim(txtanno.Text) & "-12-31"
        Else
            i = 1
            encontrado = False
            Do
                If optTrimestre(i).Value = True Then
                    encontrado = True
                    fecha_inicio = Trim(txtanno.Text) & Trimestres(i, 1)
                    fecha_fin = Trim(txtanno.Text) & Trimestres(i, 2)
                End If
                i = i + 1
            Loop Until encontrado = True Or i > 4 'Cuatro trimestres. solo se permite consultarlo de uno en uno, todo el año o usando intervalos personalizados
        End If
    End If
    Dim cliente As Long
    cliente = 0
    If chkTodos.Value = Unchecked And cmbClientes.getTEXTO <> "" Then
        cliente = cmbClientes.getPK_SALIDA
    End If
    
    
    lista.ListItems.Clear
    Me.MousePointer = vbHourglass
    Set rs = oMuestra.listadoFacturacionSectores(cmbCC.getPK_SALIDA, SECTOR, strTipo, fecha_inicio, fecha_fin, cliente, chkAirbus.Value)
    If rs.RecordCount > 0 Then
        Suma = 0
        Do
            With lista.ListItems.Add(, , rs("SECTOR"))
             .SubItems(1) = rs("FAMILIA")
             .SubItems(2) = moneda(rs("SUMA"))
             .SubItems(3) = rs("DOCS")
             Suma = Suma + CDbl(rs("SUMA"))
            End With
            rs.MoveNext
        Loop Until rs.EOF
        
    End If
    Me.MousePointer = vbNormal
    lbltotal = "Total:  " & moneda(CStr(Suma)) & "  " 'Total (sumatorio de cada subtotal)
    Set oMuestra = Nothing
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cargar_lista_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_lista of Formulario frmFacturacion_Sectores"
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdGenerar_Click()
    cargar_lista
End Sub
Private Sub optFecha_Click(Index As Integer)
    If Index = 0 Then 'Búsqueda por intervalos
        Frame4.Enabled = True
        Frame5.Enabled = False
    Else              'Búsqueda por trimestres o ejercicio completo
        Frame4.Enabled = False
        Frame5.Enabled = True
    End If
End Sub
Private Sub optTrimestre_Click(Index As Integer)
    chkAnyo.Value = 0
End Sub

Private Sub cmdVerExcel_Click()
   Dim cadena As String
   On Error GoTo cmdVerExcel_Click_Error

       Frame6.Visible = True
       Me.MousePointer = vbHourglass
       Dim rs As New ADODB.Recordset
       rs.Fields.Append "c1", adChar, 350, adFldUpdatable
       rs.Fields.Append "c2", adChar, 350, adFldUpdatable
       rs.Fields.Append "c3", adChar, 50, adFldUpdatable
       rs.Open
       
       Dim i As Integer

       For i = 1 To lista.ListItems.Count
            rs.AddNew
            rs("c1") = lista.ListItems(i).Text
            rs("c2") = lista.ListItems(i).SubItems(1)
            rs("c3") = lista.ListItems(i).SubItems(2)
'            RS("c4") = lista.ListItems(i).SubItems(3)
            rs.Update
        Next i
        
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Facturacion"

        'Cabecera
        With XLS.Range("A1:C1")
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
            .WrapText = True
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        With XLS.Range("A1:C1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:C1").Borders
            .LineStyle = vbSolid
        End With
        XLS.Range("A1:A1").ColumnWidth = 20
        XLS.Range("B1:B1").ColumnWidth = 40
        XLS.Range("C1:C1").ColumnWidth = 15
        XLS.Cells(1, 1) = "Sector"
        XLS.Cells(1, 2) = "Familia"
        XLS.Cells(1, 3) = "Total"

        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
'            XLS.Cells(i, 1) = RS("c1")
            XLS.Cells(i, 1) = ClrStr(rs("c1"), False, True, True)
            XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True)
            XLS.Cells(i, 3) = moneda_bd(rs("C3"))
            i = i + 1
             
            XLS.Range("A" & i).EntireRow.Insert
            rs.MoveNext
          Loop Until rs.EOF
        End If
        Frame6.Visible = False
        Me.MousePointer = vbNormal
        XLA.Visible = True
        
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cmdVerExcel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdVerExcel_Click of Formulario frmFacturacion_Sectores"
End Sub
Private Sub lista_Click()
   On Error GoTo lista_Click_Error

    listaFacturas.ListItems.Clear
    If lista.ListItems.Count > 0 Then
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT A.ID_DOC,A.NUMERO, B.NOMBRE,A.FECHA_FACTURA,C.CODIGO,A.TOTAL,A.DESCUENTO,A.IVA,A.TIPO " & _
                   "  FROM DOCS_PAGO A " & _
                   "  LEFT JOIN CLIENTES B ON A.CLIENTE_ID = B.ID_CLIENTE " & _
                   "  LEFT JOIN CLIENTES_PEDIDOS C ON A.PEDIDO_ID = C.ID_PEDIDO " & _
                   " WHERE A.ID_DOC IN (" & lista.ListItems(lista.selectedItem.Index).SubItems(3) & ")" & _
                   " ORDER BY A.ID_DOC "
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            Do
                Dim IMPORTE As Currency
                Dim BASE As Currency
                Dim IVA As Currency
                Dim NUMERO As String
                
'        .Add , , "NºDOC", 1000, lvwColumnCenter
'        .Add , , "Cliente", 3600, lvwColumnLeft
'        .Add , , "Fecha", 1200, lvwColumnCenter
'        .Add , , "Pedido", 3000, lvwColumnCenter
'        .Add , , "Importe", 1100, lvwColumnRight
 '       .Add , , "Dto%", 500, lvwColumnCenter
'        .Add , , "Base", 1100, lvwColumnRight
'        .Add , , "IVA%", 500, lvwColumnRight
'        .Add , , "Cuota I.V.A.", 1100, lvwColumnRight
'        .Add , , "Total", 1200, lvwColumnRight
                
                With listaFacturas.ListItems.Add(, , rs(0)) ' id_doc
                    Select Case rs(8)
                    Case 1
                        NUMERO = "A-" & Format(rs(1), "0000")
                    Case 2
                        NUMERO = "F-" & Format(rs(1), "0000")
                    Case 3
                        NUMERO = "B-" & Format(rs(1), "0000")
                    Case Else
                        NUMERO = Format(rs(1), "0000")
                    End Select
                 .SubItems(1) = NUMERO ' nºdoc
                 .SubItems(2) = rs(2) ' cliente
                 .SubItems(3) = Format(rs(3), "dd-mm-yyyy") ' Fecha
                 .SubItems(4) = rs(4) ' Pedido
'                 .SubItems(5) = moneda(rs(5)) ' importe
                 IMPORTE = Format(rs(5), "currency")
                 If IsNull(rs.Fields("descuento")) Or rs.Fields("descuento") = "0" Then
                    BASE = Format(IMPORTE, "0.00")
                 Else
                    BASE = Format(IMPORTE - ((IMPORTE * rs.Fields("descuento")) / 100), "0.00")
                 End If
                 IVA = Format((BASE * rs.Fields("iva")) / 100, "0.00")
                 .SubItems(5) = Format(IMPORTE, "currency")
                 .SubItems(6) = Format(rs.Fields("descuento"), "Standard")
                 .SubItems(7) = Format(BASE, "currency")
                 .SubItems(8) = rs.Fields("iva")
                 .SubItems(9) = Format(IVA, "currency")
                 .SubItems(10) = Format(BASE + IVA, "currency")
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = Nothing
    End If

   On Error GoTo 0
   Exit Sub

lista_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure lista_Click of Formulario frmFacturacion_Sectores"
End Sub

Private Sub listaFacturas_DblClick()
    If listaFacturas.ListItems.Count = 0 Then Exit Sub
    If USUARIO.getPER_FACTURACION = 0 Then
        MsgBox "No tiene permisos para ver la facturacion.", vbExclamation, App.Title
    Else
        gdoc = listaFacturas.ListItems(listaFacturas.selectedItem.Index).Text
        frmListadoDocPago.Show
    End If
End Sub


