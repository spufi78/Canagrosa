VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmOferta_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Ofertas"
   ClientHeight    =   9690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14595
   Icon            =   "frmOferta_Listado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9690
   ScaleWidth      =   14595
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Excel"
      Height          =   870
      Left            =   4455
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdProforma 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Proforma"
      Height          =   870
      Left            =   6615
      Picture         =   "frmOferta_Listado.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   37
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdHistorialCambios 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Seguimiento"
      Height          =   870
      Left            =   8775
      Style           =   1  'Graphical
      TabIndex        =   35
      Top             =   8730
      Width           =   1365
   End
   Begin VB.CommandButton cmdduplicar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Duplicar"
      Height          =   870
      Left            =   7695
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdEmail 
      BackColor       =   &H00E0E0E0&
      Caption         =   "E-mail"
      Height          =   870
      Left            =   5520
      Picture         =   "frmOferta_Listado.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8730
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por "
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
      Height          =   2040
      Left            =   45
      TabIndex        =   21
      Top             =   360
      Width           =   14505
      Begin VB.TextBox txtConcepto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   9090
         TabIndex        =   12
         Top             =   1620
         Width           =   2370
      End
      Begin VB.TextBox txtNumero 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6165
         TabIndex        =   11
         Top             =   1620
         Width           =   1830
      End
      Begin VB.CheckBox chkFechaAceptacion 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Check1"
         Height          =   285
         Left            =   135
         TabIndex        =   6
         Top             =   1305
         Width           =   195
      End
      Begin VB.TextBox txtFiltro 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   6165
         TabIndex        =   9
         Top             =   1260
         Width           =   5295
      End
      Begin VB.CheckBox chkTodas 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Mostrar todas las ediciones"
         Height          =   195
         Left            =   135
         TabIndex        =   10
         Top             =   1665
         Width           =   2715
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   960
         Left            =   12825
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   630
         Width           =   1095
      End
      Begin pryCombo.miCombo cmbclientes 
         Height          =   345
         Left            =   810
         TabIndex        =   0
         Top             =   225
         Width           =   11685
         _ExtentX        =   20611
         _ExtentY        =   609
      End
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   1
         Left            =   6165
         TabIndex        =   5
         Top             =   930
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   0
         Left            =   810
         TabIndex        =   1
         Top             =   585
         Width           =   3885
         _ExtentX        =   6853
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin MSDataListLib.DataCombo cmbDatos 
         Height          =   315
         Index           =   2
         Left            =   6165
         TabIndex        =   2
         Top             =   585
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         Text            =   ""
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
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Top             =   945
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
         Format          =   60096513
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   3420
         TabIndex        =   4
         Top             =   945
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
         Format          =   60096513
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fechaAceptacionDesde 
         Height          =   330
         Left            =   1440
         TabIndex        =   7
         Top             =   1305
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   60096513
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fechaAceptacionHasta 
         Height          =   330
         Left            =   3420
         TabIndex        =   8
         Top             =   1305
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Enabled         =   0   'False
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
         Format          =   60096513
         CurrentDate     =   38002
      End
      Begin VB.Label Label8 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Concepto"
         Height          =   195
         Left            =   8280
         TabIndex        =   34
         Top             =   1665
         Width           =   870
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Oferta"
         Height          =   195
         Left            =   5175
         TabIndex        =   33
         Top             =   1665
         Width           =   870
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descripción"
         Height          =   195
         Left            =   5175
         TabIndex        =   32
         Top             =   1305
         Width           =   960
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   3
         Left            =   2835
         TabIndex        =   31
         Top             =   1350
         Width           =   405
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "F.Aceptación"
         Height          =   195
         Index           =   0
         Left            =   405
         TabIndex        =   30
         Top             =   1350
         Width           =   945
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Oferta"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   29
         Top             =   990
         Width           =   930
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         Height          =   195
         Index           =   2
         Left            =   2835
         TabIndex        =   28
         Top             =   990
         Width           =   405
      End
      Begin VB.Label Label5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "SubTipo"
         Height          =   195
         Left            =   5175
         TabIndex        =   27
         Top             =   630
         Width           =   645
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Left            =   135
         TabIndex        =   24
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         Height          =   195
         Left            =   5175
         TabIndex        =   23
         Top             =   990
         Width           =   690
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   270
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   870
      Left            =   3357
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13455
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1179
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   8730
      Width           =   1050
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2268
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8730
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5925
      Left            =   45
      TabIndex        =   14
      Top             =   2430
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   10451
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
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      DataSource      =   "Adodc1"
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
      Height          =   315
      Left            =   8325
      Locked          =   -1  'True
      TabIndex        =   38
      Top             =   8340
      Width           =   1250
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      DataSource      =   "Adodc1"
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
      Height          =   315
      Left            =   9540
      Locked          =   -1  'True
      TabIndex        =   40
      Top             =   8340
      Width           =   1250
   End
   Begin VB.Label lblImporte 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   45
      TabIndex        =   39
      Top             =   8325
      Width           =   14460
   End
   Begin VB.Label lblsubtitulo 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   330
      Left            =   8910
      TabIndex        =   36
      Top             =   0
      Width           =   4380
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de Ofertas"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   345
      Index           =   3
      Left            =   15
      TabIndex        =   15
      Top             =   0
      Width           =   14535
   End
End
Attribute VB_Name = "frmOferta_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pk_CLIENTE As Long
Private Sub cmdExcel_Click()
   On Error GoTo cmdExcel_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
    Dim cadena As String
    Me.MousePointer = vbHourglass
    Dim i As Integer

     Dim XLA As excel.Application
     Dim XLW As excel.Workbook
     Dim XLS As excel.Worksheet
     
     Set XLA = New excel.Application
     Set XLW = XLA.Workbooks.Add
     Set XLS = XLW.Worksheets(1)
     XLW.Worksheets(3).Delete
     XLW.Worksheets(2).Delete
     XLW.Worksheets(1).Name = "Listado de Ofertas"

     'Cabecera
     With XLS.Range("A1:K1")
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
     With XLS.Range("A1:K1").Interior
         .Pattern = xlSolid
         .PatternColorIndex = xlAutomatic
         .color = &HC0C0FF
     End With
     With XLS.Range("A1:K1").Borders
         .LineStyle = vbSolid
     End With
     
     XLS.Range("A1:A1").ColumnWidth = 15
     XLS.Range("B1:B1").ColumnWidth = 10
     XLS.Range("C1:C1").ColumnWidth = 15
     XLS.Range("D1:D1").ColumnWidth = 40
     XLS.Range("E1:E1").ColumnWidth = 30
     XLS.Range("F1:F1").ColumnWidth = 30
     XLS.Range("G1:G1").ColumnWidth = 15
     XLS.Range("H1:H1").ColumnWidth = 15
     XLS.Range("I1:I1").ColumnWidth = 15
     XLS.Range("J1:J1").ColumnWidth = 15
     XLS.Range("K1:K1").ColumnWidth = 15
     XLS.Cells(1, 1) = "Número"
     XLS.Cells(1, 2) = "Ed."
     XLS.Cells(1, 3) = "Fecha"
     XLS.Cells(1, 4) = "Cliente"
     XLS.Cells(1, 5) = "Tipo"
     XLS.Cells(1, 6) = "SubTipo"
     XLS.Cells(1, 7) = "Importe"
     XLS.Cells(1, 8) = "Imp.Pedido"
     XLS.Cells(1, 9) = "Estado"
     XLS.Cells(1, 10) = "F.Aceptación"
     XLS.Cells(1, 11) = "Usuario"

     For i = 1 To lista.ListItems.Count
         XLS.Cells(i + 1, 1) = ClrStr(lista.ListItems(i).SubItems(1), False, True, True)
         XLS.Cells(i + 1, 2) = ClrStr(lista.ListItems(i).SubItems(2), False, True, True)
         XLS.Cells(i + 1, 3) = Format(lista.ListItems(i).SubItems(3), "mm/dd/yyyy") ' Fecha
         XLS.Cells(i + 1, 4) = ClrStr(lista.ListItems(i).SubItems(4), False, True, True)
         XLS.Cells(i + 1, 5) = ClrStr(lista.ListItems(i).SubItems(5), False, True, True)
         XLS.Cells(i + 1, 6) = ClrStr(lista.ListItems(i).SubItems(6), False, True, True)
         XLS.Cells(i + 1, 7) = moneda_bd(lista.ListItems(i).SubItems(7))
         XLS.Cells(i + 1, 8) = moneda_bd(lista.ListItems(i).SubItems(8))
         XLS.Cells(i + 1, 9) = ClrStr(lista.ListItems(i).SubItems(9), False, True, True)
         XLS.Cells(i + 1, 10) = Format(lista.ListItems(i).SubItems(10), "mm/dd/yyyy") ' Fecha
         XLS.Cells(i + 1, 11) = ClrStr(lista.ListItems(i).SubItems(11), False, True, True)
'         XLS.Range("A" & i).EntireRow.Insert
     Next
     Me.MousePointer = vbNormal
     XLA.visible = True

   On Error GoTo 0
   Exit Sub

cmdExcel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExcel_Click of Formulario frmOferta_Listado"

End Sub

Private Sub cmdHistorialCambios_Click()
    'M1108-I
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    frmHistorialCambios.PK_TIPO = HC_TIPOS.HC_OFERTAS
    frmHistorialCambios.PK_ID = lista.ListItems(lista.selectedItem.Index).SubItems(1)
    frmHistorialCambios.PK_TITULO = "Seguimiento de Oferta Nº" & lista.ListItems(lista.selectedItem.Index).SubItems(1)
    frmHistorialCambios.Show 1
    'M1108-F
End Sub
Private Sub chkFechaAceptacion_Click()
    If chkFechaAceptacion.Value = Checked Then
        fechaAceptacionDesde.Enabled = True
        fechaAceptacionHasta.Enabled = True
    Else
        fechaAceptacionDesde.Enabled = False
        fechaAceptacionHasta.Enabled = False
    End If
    cargar_lista
End Sub

Private Sub chkTodas_Click()
    cargar_lista
End Sub

Private Sub cmbClientes_change()
    cargar_lista
End Sub

Private Sub cmbDatos_Change(Index As Integer)
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
'    If UCase(usuario.getUSUARIO) <> "JULIO" Then
'        frmOferta_Nueva.PK = 0
'        frmOferta_Nueva.Show 1
'    Else
        frmOferta_Nueva2.PK = 0
        frmOferta_Nueva2.Show 1
'    End If
    cargar_lista
End Sub

Private Sub cmdduplicar_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If MsgBox("Va a duplicar la oferta. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
      Dim OFERTA As Long
      Dim oOferta As New clsOfertas
      Dim oOfertaD As New clsOfertas
      Dim rs As ADODB.Recordset
      If oOferta.Carga(lista.ListItems(lista.selectedItem.Index), lista.ListItems(lista.selectedItem.Index).SubItems(2)) = True Then
          With oOfertaD
            .setEDICION = 1
            .setULTIMA = 1
            .setCLIENTE_ID = oOferta.getCLIENTE_ID
            .setESTADO_OFERTA = 0
            .setFECHA = Format(Date, "dd-mm-yyyy")
            .setLOGO_ENAC = oOferta.getLOGO_ENAC
            .setLOGO_ENACM = oOferta.getLOGO_ENACM
            .setLOGO_EQUA = oOferta.getLOGO_EQUA
            .setLOGO_NADCAP = oOferta.getLOGO_NADCAP
            .setNUMERO = oOferta.Calcular_Numero
            .setOBSERVACIONES = oOferta.getOBSERVACIONES
            .setPLAZO_ENTREGA = oOferta.getPLAZO_ENTREGA
            .setSELLO = oOferta.getSELLO
            .setTIPO_OFERTA = oOferta.getTIPO_OFERTA
            .setSUBTIPO_OFERTA = oOferta.getSUBTIPO_OFERTA
            .setTOTAL = oOferta.getTOTAL
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            
            .setDESCRIPCION = oOferta.getDESCRIPCION
            
            .setFECHA_ACEPTACION = "1900-01-01"
            .setFECHA_ANULACION = "1900-01-01"
'            .setFECHA_ACEPTACION = Format(oOferta.getFECHA_ACEPTACION, "yyyy-mm-dd")
            .setIPEDIDO = moneda_bd(oOferta.getIPEDIDO)
            OFERTA = .Insertar
            If OFERTA = 0 Then
                MsgBox "Error al insertar la oferta duplicada.", vbCritical, App.Title
                Exit Sub
            End If
          End With
          ' Detalle de la oferta
          Dim oFD As New clsOfertas_detalle
          Dim OFDD As New clsOfertas_detalle
          Set rs = oFD.Listado(lista.ListItems(lista.selectedItem.Index), lista.ListItems(lista.selectedItem.Index).SubItems(2))
          Do While Not rs.EOF
             With OFDD
                .setOFERTA_ID = OFERTA
                .setEDICION = 1
                .setORDEN = rs("ORDEN")
                .setBANO = rs("BANO")
                .setDETERMINACION = rs("DETERMINACION")
                .setRANGO = rs("RANGO")
                .setPRECIO = rs("PRECIO")
                If .Insertar = False Then
                    MsgBox "Error al insertar el detalle de la oferta duplicada.", vbCritical, App.Title
                    Exit Sub
                End If
             End With
            rs.MoveNext
          Loop
        'M1108-I
        Dim ohc As New clsHistorial_cambios
        With ohc
            .setTIPO = HC_TIPOS.HC_OFERTAS
            .setIDENTIFICADOR = oOfertaD.getNUMERO
            .setIDENTIFICADOR_TEXTO = "Oferta Nº" & oOfertaD.getNUMERO
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setMOTIVO = HC_CREACION & " (Duplicando la Nº" & oOferta.getNUMERO & ")"
            .Insertar
        End With
        Set ohc = Nothing
        'M1108-F
          MsgBox "La oferta se ha duplicado correctamente.", vbOKOnly + vbInformation, App.Title
          cargar_lista
      End If
    End If
    Exit Sub
fallo:
    MsgBox "Error en el proceso de duplicación.", vbCritical, App.Title
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.selectedItem.Index).SubItems(7) = "ENVIADA" Then
            MsgBox "No se puede eliminar una oferta ENVIADA.", vbExclamation, App.Title
            Exit Sub
        End If
        If MsgBox("Va a eliminar la oferta número : " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & "/Ed." & lista.ListItems(lista.selectedItem.Index).SubItems(2) & " ¿Estas seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oOferta As New clsOfertas
            If oOferta.Eliminar(lista.ListItems(lista.selectedItem.Index), lista.ListItems(lista.selectedItem.Index).SubItems(2)) = True Then
                oOferta.Quitar_Ultima lista.ListItems(lista.selectedItem.Index).SubItems(1)
                cargar_lista
            End If
        End If
    End If
End Sub

Private Sub cmdEmail_Click()
   On Error GoTo cmdEmail_Click_Error
    Dim vinculo As String
    Dim ASUNTO As String
    Dim oOferta As New clsOfertas
    Dim marcado As Boolean
    Dim primera_oferta As Long
    Dim cliente As String
    Dim Clientes_Distintos As Boolean
    marcado = False
    Clientes_Distintos = False
    On Error Resume Next
    MkDir App.Path & "\Ofertas"
   On Error GoTo cmdEmail_Click_Error
    If lista.ListItems.Count > 0 Then
        ' Generamos el pdf
        Dim i As Integer
        ' Verificar que sean del mismo cliente
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                primera_oferta = lista.ListItems(i).Text
                If cliente = "" Then
                    cliente = lista.ListItems(i).SubItems(4)
                End If
                If cliente <> lista.ListItems(i).SubItems(4) Then
                    Clientes_Distintos = True
                End If
            End If
        Next
        If Clientes_Distintos Then
            MsgBox "Marque para enviar sólo Ofertas del mismo Cliente.", vbCritical, App.Title
            Exit Sub
        End If
        For i = 1 To lista.ListItems.Count
            If lista.ListItems(i).Checked = True Then
                marcado = True
                oOferta.generar_oferta CLng(lista.ListItems(i)), App.Path & "\Ofertas\Oferta número " & lista.ListItems(i).SubItems(1) & "-" & Year(lista.ListItems(i).SubItems(3)) & ".pdf"
                vinculo = vinculo & App.Path & "\Ofertas\Oferta número " & lista.ListItems(i).SubItems(1) & "-" & Year(lista.ListItems(i).SubItems(3)) & ".pdf" & ";"
                ASUNTO = ASUNTO & lista.ListItems(i).SubItems(1) & "-" & Year(lista.ListItems(i).SubItems(3)) & " , "
            End If
        Next
        If Not marcado Then
            MsgBox "Marque las ofertas que desea enviar al cliente.", vbExclamation, App.Title
            Exit Sub
        End If
        ASUNTO = Left(ASUNTO, Len(ASUNTO) - 3)
        ' Enviar correo
        Dim oCliente As New clsCliente
        oOferta.Carga primera_oferta, lista.ListItems(lista.selectedItem.Index).SubItems(2)
        oCliente.CargaCliente oOferta.getCLIENTE_ID
        Dim ref As String
        ref = "Adjunto oferta número : " & ASUNTO
        ' Copia de correo
        Dim opar As New clsParametros
        opar.Carga PARAM_OFERTAS_COPIA_CORREO, ""
        genera_correo oCliente.getEMAIL2, ref, "", vinculo, Me.hdc, , opar.getVALOR
        Set opar = Nothing
        'M1108-I
        Dim ohc As New clsHistorial_cambios
        With ohc
            .setTIPO = HC_TIPOS.HC_OFERTAS
            .setIDENTIFICADOR = lista.ListItems(lista.selectedItem.Index).SubItems(1)
            .setIDENTIFICADOR_TEXTO = "Oferta Nº" & lista.ListItems(lista.selectedItem.Index).SubItems(1)
            .setUSUARIO_ID = USUARIO.getID_EMPLEADO
            .setMOTIVO = "Envío por correo"
            .Insertar
        End With
        Set ohc = Nothing
        'M1108-F
        
    End If
    Set oOferta = Nothing

   On Error GoTo 0
   Exit Sub

cmdEmail_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdEmail_Click of Formulario frmOferta_Listado"
End Sub

Private Sub cmdImprimir_Click()
    If lista.ListItems.Count > 0 Then
        Dim oOferta As New clsOfertas
        oOferta.imprimir (lista.ListItems(lista.selectedItem.Index))
        Set oOferta = Nothing
    End If
End Sub

Private Sub cmdLimpiar_Click()
    cmbClientes.limpiar
    cmbDatos(0).Text = ""
    cmbDatos(1).Text = ""
    chkTodas.Value = Unchecked
    txtFiltro = ""
    cargar_lista
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
'        If UCase(usuario.getUSUARIO) <> "JULIO" Then
'            frmOferta_Nueva.PK = lista.ListItems(lista.selectedItem.Index)
'            frmOferta_Nueva.PK_EDICION = lista.ListItems(lista.selectedItem.Index).SubItems(2)
'            frmOferta_Nueva.Show 1
'            If frmOferta_Nueva.Nueva_Edicion = True Then
'                cargar_lista
'            Else
'                actualizar_lista
'            End If
'        Else
            frmOferta_Nueva2.PK = lista.ListItems(lista.selectedItem.Index)
            frmOferta_Nueva2.PK_EDICION = lista.ListItems(lista.selectedItem.Index).SubItems(2)
            frmOferta_Nueva2.Show 1
            If frmOferta_Nueva2.Nueva_Edicion = True Then
                cargar_lista
            Else
                actualizar_lista
            End If
'        End If
    End If
End Sub

Private Sub cmdProforma_Click()
    If lista.ListItems.Count > 0 Then
        Dim oOferta As New clsOfertas
        oOferta.imprimirProforma (lista.ListItems(lista.selectedItem.Index))
        Set oOferta = Nothing
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fdesde_Change()
    cargar_lista
End Sub

Private Sub fechaAceptacionDesde_Change()
    cargar_lista

End Sub

Private Sub fechaAceptacionHasta_Change()
    cargar_lista

End Sub

Private Sub fhasta_Change()
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_combos
    llenar_combo cmbClientes, New clsCliente, 0, frmClientes, ""
    If pk_CLIENTE <> 0 Then
        cmbClientes.MostrarElemento pk_CLIENTE
    End If
    fhasta = Date
    fdesde = Date - 180
    fechaAceptacionHasta = Date
    fechaAceptacionDesde = Date - 180
    cargar_lista
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 300, lvwColumnLeft
        .Add , , "Número", 800, lvwColumnCenter
        .Add , , "Ed.", 400, lvwColumnCenter
        .Add , , "Fecha", 1050, lvwColumnCenter
        .Add , , "Cliente", 3000, lvwColumnLeft
        .Add , , "Tipo", 1500, lvwColumnCenter
        .Add , , "SubTipo", 1200, lvwColumnCenter
        .Add , , "Importe", 1250, lvwColumnRight
        .Add , , "Imp.Pedido", 1250, lvwColumnRight
        .Add , , "Estado", 1200, lvwColumnCenter
        .Add , , "F.Acep/Rech.", 1050, lvwColumnCenter
        .Add , , "Usuario", 1100, lvwColumnCenter
    End With
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oOferta As New clsOfertas
    lista.ListItems.Clear
    Dim cliente As Long
    If cmbClientes.getTEXTO = "" Then
        cliente = 0
    Else
        cliente = cmbClientes.getPK_SALIDA
    End If
    Dim i1 As Currency
    Dim i2 As Currency
    Set rs = oOferta.Listado(cliente, cmbDatos(1).BoundText, cmbDatos(0).BoundText, cmbDatos(2).BoundText, chkTodas.Value, fdesde.Value, fhasta.Value, chkFechaAceptacion.Value, fechaAceptacionDesde, fechaAceptacionHasta, txtFiltro, txtNumero, txtConcepto)
    lblsubtitulo = "Se muestran : " & rs.RecordCount & " registros"
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = Format(rs(3), "dd-mm-yyyy")
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             .SubItems(6) = rs(8)
'             .SubItems(7) = moneda(rs(10))  ' Importe
             .SubItems(7) = rs(10)  ' Importe
             If Not IsNull(rs(11)) Then
                 .SubItems(8) = moneda(rs(11))  ' I.PEDIDO
             End If
             .SubItems(9) = rs(6)
             'ACEPTACION
             If Format(rs(9), "yyyy-mm-dd") <> "1900-01-01" Then
                 .SubItems(10) = Format(rs(9), "dd-mm-yyyy")
             End If
             'RECHAZO
             If Format(rs(12), "yyyy-mm-dd") <> "1900-01-01" Then
                 .SubItems(10) = Format(rs(12), "dd-mm-yyyy")
             End If
             .SubItems(11) = rs(7)
            End With
            If Not IsNull(rs(10)) Then
                If IsNumeric(rs(10)) Then
                    i1 = i1 + rs(10)
                End If
            End If
            If Not IsNull(rs(11)) Then
                If IsNumeric(rs(11)) Then
                    i2 = i2 + rs(11)
                End If
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oOferta = Nothing
    Text1 = moneda(CStr(i1))
    Text2 = moneda(CStr(i2))
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_Click()
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
      cmdModificar.Enabled = True
      cmdEliminar.Enabled = True
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
Public Sub actualizar_lista()
    Dim oOferta As New clsOfertas
    Dim rs As ADODB.Recordset
    Set rs = oOferta.Listado_PK(lista.ListItems(lista.selectedItem.Index).Text)
    If rs.RecordCount > 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = Format(rs(3), "dd-mm-yyyy")
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs(4)
        lista.ListItems(lista.selectedItem.Index).SubItems(5) = rs(5)
        lista.ListItems(lista.selectedItem.Index).SubItems(6) = rs(8)
        lista.ListItems(lista.selectedItem.Index).SubItems(7) = rs(10) ' Importe
       
        If Not IsNull(rs(11)) Then
            lista.ListItems(lista.selectedItem.Index).SubItems(8) = moneda(rs(11))  ' I.PEDIDO
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(8) = ""
        End If
        
        lista.ListItems(lista.selectedItem.Index).SubItems(9) = rs(6)
        If Format(rs(9), "yyyy-mm-dd") <> "1900-01-01" Then
            lista.ListItems(lista.selectedItem.Index).SubItems(10) = Format(rs(9), "dd-mm-yyyy")
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(10) = ""
        End If
        lista.ListItems(lista.selectedItem.Index).SubItems(11) = rs(7)
    End If
    Set oOferta = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Public Sub cargar_combos()
    Dim oDecodificadora As New clsDecodificadora
    oDecodificadora.cargar_combo cmbDatos(1), DECODIFICADORA.ESTADOS_OFERTAS
    oDecodificadora.cargar_combo cmbDatos(0), DECODIFICADORA.TIPOS_DE_OFERTAS
    oDecodificadora.cargar_combo cmbDatos(2), DECODIFICADORA.SUBTIPOS_DE_OFERTAS
End Sub

Private Sub txtConcepto_Change()
    cargar_lista
End Sub

Private Sub txtfiltro_Change()
    cargar_lista
End Sub

Private Sub txtNumero_Change()
    cargar_lista
End Sub
