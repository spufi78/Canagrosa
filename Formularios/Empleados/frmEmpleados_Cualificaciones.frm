VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{CEA3B7F6-2847-4E5E-A551-DB7A62489D44}#46.0#0"; "miCombo.ocx"
Begin VB.Form frmEmpleados_Cualificaciones 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cualificaciones de Empleados"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12675
   Icon            =   "frmEmpleados_Cualificaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12675
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000001&
      Height          =   780
      Left            =   2970
      TabIndex        =   19
      Top             =   3960
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
         TabIndex        =   20
         Top             =   225
         Width           =   5415
      End
   End
   Begin VB.CommandButton cmdVerExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   885
      Left            =   6030
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   7785
      Width           =   1155
   End
   Begin VB.CommandButton cmdPNT 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver P.N.T."
      Height          =   885
      Left            =   4860
      Picture         =   "frmEmpleados_Cualificaciones.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   7785
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Imprimir"
      Height          =   885
      Left            =   3669
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   7785
      Width           =   1155
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   885
      Left            =   2481
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   7785
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   885
      Left            =   1293
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   7785
      Width           =   1155
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   885
      Left            =   105
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   7785
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   45
      TabIndex        =   3
      Top             =   720
      Width           =   12570
      Begin VB.CheckBox chkmodalidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Externa"
         Height          =   330
         Index           =   1
         Left            =   4365
         TabIndex        =   13
         Top             =   495
         Value           =   1  'Checked
         Width           =   870
      End
      Begin VB.CheckBox chkmodalidad 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Interna"
         Height          =   330
         Index           =   0
         Left            =   4365
         TabIndex        =   12
         Top             =   225
         Value           =   1  'Checked
         Width           =   825
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   810
         MaxLength       =   75
         TabIndex        =   9
         Top             =   360
         Width           =   2355
      End
      Begin VB.CommandButton cmdLimpiarCampos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar Campos"
         Height          =   780
         Left            =   11115
         Picture         =   "frmEmpleados_Cualificaciones.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1410
      End
      Begin pryCombo.miCombo cmbFormador 
         Height          =   330
         Left            =   6390
         TabIndex        =   14
         Top             =   360
         Width           =   4620
         _ExtentX        =   8149
         _ExtentY        =   582
      End
      Begin VB.Shape Shape1 
         Height          =   690
         Left            =   3420
         Top             =   180
         Width           =   1950
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Formador"
         Height          =   195
         Index           =   6
         Left            =   5580
         TabIndex        =   15
         Top             =   405
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Modalidad"
         Height          =   195
         Index           =   2
         Left            =   3465
         TabIndex        =   11
         Top             =   405
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "P.N.T."
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   450
         Width           =   465
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   885
      Left            =   11475
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7785
      Width           =   1155
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   630
      Top             =   7470
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4410
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Cualificaciones.frx":79E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Cualificaciones.frx":82C0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEmpleados_Cualificaciones.frx":8B9A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5970
      Left            =   75
      TabIndex        =   7
      Top             =   1740
      Width           =   12570
      _ExtentX        =   22172
      _ExtentY        =   10530
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
      Caption         =   "Listado de cualificaciones del empleado"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   2820
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cualificaciones del Empleado"
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
      TabIndex        =   0
      Top             =   45
      Width           =   3120
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   690
      Left            =   0
      Top             =   0
      Width           =   12825
   End
End
Attribute VB_Name = "frmEmpleados_Cualificaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK As Long


Private Sub cmdVerExcel_Click()
   On Error GoTo cmdVerExcel_Click_Error

    If lista.ListItems.Count = 0 Then Exit Sub
       Frame3.Visible = True
       Me.MousePointer = vbHourglass
       Dim rs As New ADODB.Recordset
       rs.Fields.Append "c1", adChar, 250, adFldUpdatable
       rs.Fields.Append "c2", adChar, 250, adFldUpdatable
       rs.Fields.Append "c3", adChar, 250, adFldUpdatable
       rs.Fields.Append "c4", adChar, 20, adFldUpdatable
       rs.Fields.Append "c5", adChar, 20, adFldUpdatable
       rs.Fields.Append "c6", adChar, 20, adFldUpdatable
       rs.Open
       
       Dim i As Integer

       For i = 1 To lista.ListItems.Count
            rs.AddNew
            rs("c1") = lista.ListItems(i).SubItems(1)
            rs("c2") = lista.ListItems(i).SubItems(2)
            rs("c3") = lista.ListItems(i).SubItems(3)
            rs("c4") = lista.ListItems(i).SubItems(4)
            rs("c5") = lista.ListItems(i).SubItems(5)
            rs("c6") = lista.ListItems(i).SubItems(6)
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
        XLW.Worksheets(1).Name = "Cualificaciones del Empleado"
 
        'Cabecera
        With XLS.Range("A1:F1")
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
        With XLS.Range("A1:F1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:F1").Borders
            .LineStyle = vbSolid
        End With
        
        XLS.Range("A1:A1").ColumnWidth = 60
        XLS.Range("B1:B1").ColumnWidth = 60
        XLS.Range("C1:C1").ColumnWidth = 60
        XLS.Range("D1:D1").ColumnWidth = 15
        XLS.Range("E1:E1").ColumnWidth = 15
        XLS.Range("F1:F1").ColumnWidth = 15
        
        XLS.Cells(1, 1) = "P.N.T."
        XLS.Cells(1, 2) = "Modalidad"
        XLS.Cells(1, 3) = "Formador"
        XLS.Cells(1, 4) = "F.Formación"
        XLS.Cells(1, 5) = "F.Obtención"
        XLS.Cells(1, 6) = "F.Recualificación"
 
        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Cells(i, 1) = ClrStr(rs("c1"), False, True, True)
            XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True)
            XLS.Cells(i, 3) = ClrStr(rs("c3"), False, True, True)
            XLS.Cells(i, 4) = rs("c4")
            XLS.Cells(i, 5) = rs("c5")
            XLS.Cells(i, 6) = rs("c6")
            i = i + 1
             
            XLS.Range("A" & i).EntireRow.Insert
            rs.MoveNext
          Loop Until rs.EOF
        End If
        Frame3.Visible = False
        Me.MousePointer = vbNormal
        XLA.Visible = True
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cmdVerExcel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdVerExcel_Click of Formulario frmEmpleados_Cualificaciones"
End Sub
Private Sub chkfiltro_Click()
    cargar_lista
End Sub

Private Sub cmbestados_Change()
    cargar_lista
End Sub


Private Sub chkmodalidad_Click(Index As Integer)
    cargar_lista

End Sub

Private Sub cmbFormador_change()
    cargar_lista
End Sub

Private Sub cmdAnadir_Click()
    frmEmpleados_Cualificaciones_Nueva.EMPLEADO_ID = PK
    frmEmpleados_Cualificaciones_Nueva.ID_CUALIFICACION = 0
    frmEmpleados_Cualificaciones_Nueva.Show 1
    cargar_lista
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR la cualificacion : " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
            Dim oEC As New clsEmpleados_cualificaciones
            If oEC.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
                lista.ListItems.Remove lista.selectedItem.Index
            End If
            Set oEC = Nothing
        End If
        lista.SetFocus
    End If
End Sub

Private Sub cmdImprimir_Click()
    frmReport.iniciar
    frmReport.informe = "\Empleados\rptEmpleados_Cualificaciones"
    frmReport.criterio = "{empleados.ID_EMPLEADO} =" & PK
    frmReport.imprimir = False
    frmReport.generar
    frmReport.Show 1
    Unload frmReport
End Sub

Private Sub cmdLimpiarCampos_Click()
    txtDatos(0) = ""
    chkmodalidad(0).Value = Checked
    chkmodalidad(1).Value = Checked
    cmbFormador.Limpiar
    cargar_lista
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        frmEmpleados_Cualificaciones_Nueva.EMPLEADO_ID = PK
        frmEmpleados_Cualificaciones_Nueva.ID_CUALIFICACION = lista.ListItems(lista.selectedItem.Index).Text
        frmEmpleados_Cualificaciones_Nueva.Show 1
        actualizar_lista
    End If
End Sub

Private Sub cmdPNT_Click()
    If lista.ListItems.Count > 0 Then
        frmCA_Documento.PK = lista.ListItems(lista.selectedItem.Index).SubItems(7)
        frmCA_Documento.Show 1
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
Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       cmdModificar_Click
    End If
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    llenar_combo cmbFormador, New clsEmpleados, 0, frmEmpleados_Gestion, ""
    If PK > 0 Then
        cargar_lista
    End If
End Sub

Private Sub cargar_lista()
    Dim oempleado As New clsEmpleados
    If (oempleado.CARGAR(PK)) Then
        lbltitulo = "Cualificaciones del Empleado : " & oempleado.getNOMBRE
        Me.Caption = lbltitulo
        Dim oEC As New clsEmpleados_cualificaciones
        Dim rs As ADODB.Recordset
        lista.ListItems.Clear
        If cmbFormador.getTEXTO = "" Then
            Set rs = oEC.Listado(PK, txtDatos(0), chkmodalidad(0).Value, chkmodalidad(1).Value, 0)
        Else
            Set rs = oEC.Listado(PK, txtDatos(0), chkmodalidad(0).Value, chkmodalidad(1).Value, cmbFormador.getPK_SALIDA)
        End If
        If rs.RecordCount > 0 Then
            Do
                With lista.ListItems.Add(, , rs(0))
                    .SubItems(1) = rs(1)
                    If rs(2) = 0 Then
                        .SubItems(2) = "Interna"
                        If Not IsNull(rs(3)) Then
                            .SubItems(3) = rs(3)
                        End If
                    Else
                        .SubItems(2) = "Externa"
                        If Not IsNull(rs(7)) Then
                            .SubItems(3) = rs(7)
                        End If
                    End If
                    .SubItems(4) = Format(rs(4), "dd-mm-yyyy")
                    If Format(rs(5), "yyyy-mm-dd") <> "1900-01-01" Then
                        .SubItems(5) = Format(rs(5), "dd-mm-yyyy")
                    End If
                    If rs(8) = 1 Then
                        .SubItems(6) = "En histórico"
                    Else
                        If Format(rs(6), "yyyy-mm-dd") <> "1900-01-01" Then
                            .SubItems(6) = Format(rs(6), "dd-mm-yyyy")
                        End If
                    End If
                    .SubItems(7) = rs(9)
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oEC = Nothing
        Set rs = Nothing
    End If
    Set oempleado = Nothing
End Sub
Private Sub actualizar_lista()
    Dim oEC As New clsEmpleados_cualificaciones
    Dim rs As ADODB.Recordset
    Set rs = oEC.Listado_Cualificacion(lista.ListItems(lista.selectedItem.Index).Text)
    If rs.RecordCount > 0 Then
        Do
           lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
                If rs(2) = 0 Then
                    lista.ListItems(lista.selectedItem.Index).SubItems(2) = "Interna"
                    lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs(3)
                Else
                    Dim oEF As New clsProveedor
                    oEF.Carga rs(7)
                    lista.ListItems(lista.selectedItem.Index).SubItems(2) = "Externa"
                    lista.ListItems(lista.selectedItem.Index).SubItems(3) = oEF.getNOMBRE
                    Set oEF = Nothing
                End If
                lista.ListItems(lista.selectedItem.Index).SubItems(4) = Format(rs(4), "dd-mm-yyyy")
                If rs(5) <> "1900-01-01" Then
                    lista.ListItems(lista.selectedItem.Index).SubItems(5) = Format(rs(5), "dd-mm-yyyy")
                End If
                If rs(8) = 1 Then
                    lista.ListItems(lista.selectedItem.Index).SubItems(6) = "En histórico"
                Else
                    If Format(rs(6), "yyyy-mm-dd") <> "1900-01-01" Then
                        lista.ListItems(lista.selectedItem.Index).SubItems(6) = Format(rs(6), "dd-mm-yyyy")
                    End If
                End If
                lista.ListItems(lista.selectedItem.Index).SubItems(7) = rs(9)
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oEC = Nothing
    Set rs = Nothing
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "P.N.T.", 3500, lvwColumnLeft
        .Add , , "Modalidad", 1400, lvwColumnCenter
        .Add , , "Formador", 3500, lvwColumnCenter
        .Add , , "F. Formación", 1200, lvwColumnCenter
        .Add , , "F. Obtención", 1200, lvwColumnCenter
        .Add , , "F. Recualificación", 1200, lvwColumnCenter
        .Add , , "DOCUMENTO_ID", 1, lvwColumnCenter
    End With
End Sub

Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub
