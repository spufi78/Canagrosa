VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmContabilidad_Asientos 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Importación y verificación de asientos de Contaplus"
   ClientHeight    =   8070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13305
   Icon            =   "frmContabilidad_Asientos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8070
   ScaleWidth      =   13305
   Begin VB.CommandButton cmdListado 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado en Excel"
      Height          =   915
      Left            =   8640
      Picture         =   "frmContabilidad_Asientos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   7020
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Leyenda Revisión"
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   4590
      TabIndex        =   9
      Top             =   6885
      Width           =   3810
      Begin VB.OptionButton opLeyenda 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Morado: No coinciden importe o fecha "
         ForeColor       =   &H00C000C0&
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   12
         Top             =   765
         Width           =   3075
      End
      Begin VB.OptionButton opLeyenda 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Rojo: No encuentro numero factura Contaplus"
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   11
         Top             =   270
         Width           =   3660
      End
      Begin VB.OptionButton opLeyenda 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Azul: No encuentro la factura en Geslab"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   10
         Top             =   510
         Width           =   3150
      End
   End
   Begin VB.CommandButton cmdFactura 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Factura Geslab"
      Height          =   915
      Left            =   10395
      Picture         =   "frmContabilidad_Asientos.frx":0BD4
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   7020
      Width           =   1680
   End
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   12150
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7020
      Width           =   1050
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Left            =   45
      TabIndex        =   3
      Top             =   6840
      Width           =   4350
      Begin VB.CommandButton cmdAsignar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Asignar Asientos"
         Height          =   870
         Left            =   2745
         Picture         =   "frmContabilidad_Asientos.frx":149E
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   225
         Width           =   1545
      End
      Begin VB.CommandButton cmdDesmarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Desmarcar Todas"
         Height          =   870
         Left            =   1305
         Picture         =   "frmContabilidad_Asientos.frx":1D68
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   225
         Width           =   1410
      End
      Begin VB.CommandButton cmdMarcar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Marcar Todas"
         Height          =   870
         Left            =   90
         Picture         =   "frmContabilidad_Asientos.frx":2072
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   225
         Width           =   1185
      End
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
      Height          =   1050
      Left            =   45
      TabIndex        =   1
      Top             =   405
      Width           =   13230
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Carga de Archivo de Contaplus"
         Height          =   780
         Left            =   90
         Picture         =   "frmContabilidad_Asientos.frx":24B4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   2490
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5310
      Left            =   45
      TabIndex        =   7
      Top             =   1485
      Width           =   13200
      _ExtentX        =   23283
      _ExtentY        =   9366
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
      NumItems        =   0
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   8145
      Top             =   6795
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   12735
      Picture         =   "frmContabilidad_Asientos.frx":2D7E
      Top             =   -45
      Width           =   480
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "  Importación y verificación de asientos de Contaplus"
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
      Height          =   390
      Index           =   4
      Left            =   15
      TabIndex        =   0
      Top             =   15
      Width           =   13245
   End
End
Attribute VB_Name = "frmContabilidad_Asientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAsignar_Click()
    Dim i As Integer
    Dim oDoc As New clsDocs_pago
   On Error GoTo cmdAsignar_Click_Error
    Me.MousePointer = 11
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            If IsNumeric(lista.ListItems(i).SubItems(8)) Then
                oDoc.informar_asiento lista.ListItems(i).SubItems(8), lista.ListItems(i).Text
            End If
        End If
    Next
    Me.MousePointer = 0
    MsgBox "Asientos informados correctamente.", vbInformation, App.Title

   On Error GoTo 0
   Exit Sub

cmdAsignar_Click_Error:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdAsignar_Click of Formulario frmContabilidad_Asientos"
End Sub

Private Sub cmdBuscar_Click()
    On Error Resume Next
    cd.DialogTitle = "Abrir fichero"
'    cd.InitDir = "c:\"
    cd.Filter = "Archivos de texto|*.txt"
    cd.ShowOpen
    If cd.FileName <> "" Then
        cargar_asientos cd.FileName
'        datos(0).Text = cd.FileTitle
'        datos(1).Text = usuario.getUSUARIO
'        datos(3).Text = ""
'        datos(4).Text = cd.FileName
    End If
'    cargar_asientos ""
End Sub

Private Sub cmdDesmarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = False
    Next
End Sub

Private Sub cmdFactura_Click()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.SelectedItem.Index).SubItems(8) <> "" Then
            gdoc = CLng(lista.ListItems(lista.SelectedItem.Index).SubItems(8))
            frmListadoDocPago.Show
        End If
    End If
End Sub

Private Sub cmdListado_Click()
        Dim XLA As Excel.Application
        Dim XLW As Excel.Workbook
        Dim XLS As Excel.Worksheet
        
   On Error GoTo cmdListado_Click_Error

        Set XLA = New Excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Listado de Asientos"
        XLA.Visible = True
        XLS.Range("1:1").HorizontalAlignment = xlCenter
        XLS.Range("1:1").VerticalAlignment = xlCenter
        XLS.Range("1:1").RowHeight = 30
        XLS.Range("1:1").WrapText = True
        'Cabecera
        XLS.Cells(1, 1) = "C.Asiento"
        XLS.Cells(1, 2) = "C.Número"
        XLS.Cells(1, 3) = "C.Texto"
        XLS.Cells(1, 4) = "C.Importe"
        XLS.Cells(1, 5) = "C.Fecha"
        XLS.Cells(1, 6) = "G.Número"
        XLS.Cells(1, 7) = "G.Importe"
        XLS.Cells(1, 8) = "G.Fecha"
        XLS.Cells(1, 9) = "Error"
        fila = 2
        ' Datos
        For i = 1 To lista.ListItems.Count
            XLS.Range(XLS.Cells(fila, 4), XLS.Cells(fila, 4)).NumberFormat = "0.00"
            XLS.Range(XLS.Cells(fila, 7), XLS.Cells(fila, 7)).NumberFormat = "0.00"
            XLS.Cells(fila, 1) = lista.ListItems(i).Text ' Asiento
            XLS.Cells(fila, 2) = lista.ListItems(i).SubItems(1) ' Num
            XLS.Cells(fila, 3) = lista.ListItems(i).SubItems(2) ' Des
            XLS.Cells(fila, 4) = moneda_bd(lista.ListItems(i).SubItems(3)) ' Importe
            XLS.Cells(fila, 5) = Format(lista.ListItems(i).SubItems(4), "dd-mm-yyyy") ' Fecha
            XLS.Cells(fila, 6) = lista.ListItems(i).SubItems(5) ' Nunmero
            XLS.Cells(fila, 7) = moneda_bd(lista.ListItems(i).SubItems(6)) ' Importe
            XLS.Cells(fila, 8) = Format(lista.ListItems(i).SubItems(7), "dd-mm-yyyy") ' Fecha
            XLS.Cells(fila, 9) = lista.ListItems(i).SubItems(9) ' Error
            fila = fila + 1
        Next

   On Error GoTo 0
   Exit Sub

cmdListado_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdListado_Click of Formulario frmContabilidad_Asientos"

End Sub

Private Sub cmdMarcar_Click()
    Dim i As Integer
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).Checked = True
    Next
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    Me.SetFocus
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 50
    Me.Top = 50
    fdesde = Date
    fhasta = Date
    cabecera
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "C.Asiento", 1200, lvwColumnLeft
        .Add , , "C.Numero", 1200, lvwColumnLeft
        .Add , , "C.Descripcion", 3500, lvwColumnLeft
        .Add , , "C.Importe", 1200, lvwColumnRight
        .Add , , "C.Fecha", 1200, lvwColumnCenter
        .Add , , "G.Numero", 1200, lvwColumnCenter
        .Add , , "G.Importe", 1200, lvwColumnRight
        .Add , , "G.Fecha", 1200, lvwColumnRight
        .Add , , "ID", 1, lvwColumnCenter ' 8
        .Add , , "Error", 1000, lvwColumnCenter ' 9
    End With
End Sub
Private Sub cargar_asientos(fichero As String)
   On Error GoTo cargar_asientos_Error

    Open fichero For Input As #10
    Dim linea As String
    Dim concepto As String
    Dim Pos As Integer
    Dim NUMERO As String
    Dim i As Integer
    Dim asiento As Long
    asiento = 0
    Me.MousePointer = 11
    Line Input #10, linea
    While Not EOF(10)
        If Mid(linea, 15, 2) = 70 Then
            NUMERO = ""
            If asiento = Mid(linea, 1, 6) Then
                lista.ListItems(lista.ListItems.Count).SubItems(3) = moneda(CCur(lista.ListItems(lista.ListItems.Count).SubItems(3)) + CCur(Replace(Mid(linea, 255, 16), ".", ",")))
            Else
                With lista.ListItems.Add(, , Mid(linea, 1, 6))
                    .SubItems(2) = Mid(linea, 55, 37)
                    .SubItems(3) = moneda(Mid(linea, 255, 16))
                    .SubItems(4) = Mid(linea, 13, 2) & "-" & Mid(linea, 11, 2) & "-" & Mid(linea, 7, 4)
                    ' Buscar numero de factura
                    concepto = Mid(linea, 55, 37)
                    Pos = InStr(1, Replace(concepto, "-", "/"), "/")
                    If Pos > 0 Then
                        For i = Pos - 1 To 1 Step -1
                            If IsNumeric(Mid(concepto, i, 1)) Then
                                NUMERO = Mid(concepto, i, 1) & NUMERO
                            Else
                                Exit For
                            End If
                        Next
                    Else
                        Pos = InStr(1, concepto, "º")
                        If Pos > 0 Then
                            For i = Pos + 1 To Len(concepto)
                                If IsNumeric(Mid(concepto, i, 1)) Then
                                    NUMERO = NUMERO & Mid(concepto, i, 1)
                                Else
                                    Exit For
                                End If
                            Next
                        End If
                    End If
                    .SubItems(1) = NUMERO
                    ' Buscar factura geslab
                    If NUMERO <> "" Then
                        Dim rs As ADODB.RecordSet
                        Set rs = datos_bd("SELECT * FROM DOCS_PAGO WHERE NUMERO = " & NUMERO & " AND FECHA_FACTURA = '" & Format(Mid(linea, 13, 2) & "-" & Mid(linea, 11, 2) & "-" & Mid(linea, 7, 4), "YYYY-MM-DD") & "'")
                        If rs.RecordCount > 0 Then
                            .SubItems(5) = rs("NUMERO")
                            .SubItems(6) = moneda(rs("TOTAL") - (rs("TOTAL") * rs("DESCUENTO") / 100))
                            .SubItems(7) = Format(rs("FECHA_FACTURA"), "dd-mm-yyyy")
                            .SubItems(8) = rs("ID_DOC")
                        End If
                    End If
                End With
            End If
            If lista.ListItems.Count Mod 50 = 0 Then
                DoEvents
            End If
            lista.ListItems(lista.ListItems.Count).EnsureVisible
            asiento = Mid(linea, 1, 6)
'            If lista.ListItems.Count = 10 Then Exit Sub
        End If
        Line Input #10, linea
    Wend
    Close #10
    ' VALIDACIONES
    For i = 1 To lista.ListItems.Count
        lista.ListItems(i).SubItems(9) = ""
        ' VAL 1. No decodifica bien el numero de contaplus
        If lista.ListItems(i).SubItems(1) = "" Then
            lista.ListItems(i).SubItems(9) = 1
        Else
            ' VAL 2. No existe la factura en geslab
            If lista.ListItems(i).SubItems(5) = "" Then
                lista.ListItems(i).SubItems(9) = 2
            Else
                ' VAL 3. No coinciden los importes
                If lista.ListItems(i).SubItems(3) <> lista.ListItems(i).SubItems(6) Then
                    lista.ListItems(i).SubItems(9) = 3
                End If
            End If
        End If
        
        If lista.ListItems(i).SubItems(9) <> "" Then
            colorear_linea i, lista.ListItems(i).SubItems(9)
            DoEvents
        End If
        lista.ListItems(i).EnsureVisible
    Next

    Me.MousePointer = 0
    MsgBox "Proceso Finalizado.", vbInformation, App.Title
   On Error GoTo 0
   Exit Sub

cargar_asientos_Error:
    Me.MousePointer = 0

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cargar_asientos of Formulario frmContabilidad_Asientos"
End Sub
Public Sub colorear_linea(fila As Integer, validacion As Integer)
    Dim i As Integer
    Dim color As Long
    Select Case validacion
        Case 1
            color = vbRed
        Case 2
            color = vbBlue
        Case 3
            color = vbMagenta
    End Select
    lista.ListItems(fila).ForeColor = color
    For i = 1 To lista.ColumnHeaders.Count - 1
        lista.ListItems(fila).ListSubItems(i).ForeColor = color
    Next
End Sub

Private Function contar_marcados() As Integer
    Dim i As Integer
    Dim cont As Integer
    cont = 0
    For i = 1 To lista.ListItems.Count
        If lista.ListItems(i).Checked = True Then
            cont = cont + 1
        End If
    Next
    contar_marcados = cont
End Function

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

Private Sub lista_DblClick()
    cmdFactura_Click
End Sub

