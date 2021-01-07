VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#13.2#0"; "Codejock.Controls.v13.2.1.ocx"
Begin VB.Form frmInformeRegistro 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de registro"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10965
   Icon            =   "frmInformeRegistro.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6555
   ScaleWidth      =   10965
   Begin VB.CommandButton cmdsalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   9765
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5625
      Width           =   1140
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos del Informe"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4725
      Left            =   45
      TabIndex        =   0
      Top             =   450
      Width           =   10845
      Begin VB.CommandButton cmdExcel 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informe Resultado"
         Height          =   870
         Left            =   90
         Picture         =   "frmInformeRegistro.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1710
         Width           =   1725
      End
      Begin VB.CheckBox chkimprimir 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Imprimir los informes de registro. Si no esta marcada, los informes solo se generán."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6120
         TabIndex        =   16
         Top             =   1125
         Value           =   1  'Checked
         Visible         =   0   'False
         Width           =   4650
      End
      Begin VB.CommandButton cmdInforme 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Informe Registro"
         Height          =   870
         Left            =   90
         Picture         =   "frmInformeRegistro.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   2655
         Width           =   1725
      End
      Begin VB.CheckBox chkTodas 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todas"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9570
         TabIndex        =   11
         Top             =   765
         Width           =   1020
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Todos"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   9570
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin MSDataListLib.DataCombo cmbClientes 
         Height          =   315
         Left            =   1785
         TabIndex        =   2
         Top             =   300
         Width           =   7650
         _ExtentX        =   13494
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
      Begin MSComCtl2.DTPicker fhasta 
         Height          =   330
         Left            =   4185
         TabIndex        =   4
         Top             =   1170
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   52363265
         CurrentDate     =   38002
      End
      Begin MSComCtl2.DTPicker fdesde 
         Height          =   330
         Left            =   1800
         TabIndex        =   3
         Top             =   1170
         Width           =   1545
         _ExtentX        =   2725
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
         Format          =   52363265
         CurrentDate     =   38002
      End
      Begin MSDataListLib.DataCombo cmbMuestras 
         Height          =   315
         Left            =   1785
         TabIndex        =   10
         Top             =   735
         Width           =   7650
         _ExtentX        =   13494
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
      Begin VB.CommandButton cmdListado 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Listado Resumen"
         Height          =   870
         Left            =   90
         Picture         =   "frmInformeRegistro.frx":0EDE
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   3645
         Width           =   1725
      End
      Begin VB.Label lblCampos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2EAF0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Genera una hoja excel, con los resultados de las determinaciones analizadas para el tipo de muestras seleccionado."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   840
         Index           =   6
         Left            =   2025
         TabIndex        =   18
         Top             =   1710
         Width           =   8580
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2EAF0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Genera un documento word, con un listado resumen del informe de registro de las muestras seleccionadas."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Index           =   4
         Left            =   2025
         TabIndex        =   13
         Top             =   3645
         Width           =   8595
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00D2EAF0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Genera los documentos de registro para el cuaderno de ensayo."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   810
         Index           =   5
         Left            =   2025
         TabIndex        =   15
         Top             =   2655
         Width           =   8595
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo de Muestra"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   90
         TabIndex        =   9
         Top             =   795
         Width           =   1545
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "hasta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   2
         Left            =   3510
         TabIndex        =   7
         Top             =   1215
         Width           =   555
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Recepcion"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1230
         Width           =   1575
      End
      Begin VB.Label lblCampos 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   90
         TabIndex        =   5
         Top             =   390
         Width           =   675
      End
   End
   Begin XtremeSuiteControls.ProgressBar pb 
      Height          =   285
      Left            =   90
      TabIndex        =   19
      Top             =   5220
      Width           =   10815
      _Version        =   851970
      _ExtentX        =   19076
      _ExtentY        =   503
      _StockProps     =   93
      Appearance      =   6
      UseVisualStyle  =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Caption         =   "Informes de registro"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Index           =   4
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   10920
   End
End
Attribute VB_Name = "frmInformeRegistro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkTodas_Click()
    If chkTodas.value = Checked Then
        cmbMuestras.Text = ""
        cmbMuestras.Enabled = False
    Else
        cmbMuestras.Enabled = True
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.value = Checked Then
        cmbClientes.Text = ""
        cmbClientes.Enabled = False
    Else
        cmbClientes.Enabled = True
    End If
End Sub

Private Sub cmdExcel_Click()
    Dim rs As ADODB.RecordSet
   On Error GoTo cmdExcel_Click_Error

    Set rs = listado_muestras()
    If rs.RecordCount = 0 Then
        MsgBox "No existen registros para la seleccion.", vbInformation, App.Title
        Exit Sub
    Else
        Me.MousePointer = 11
        pb.min = 0
        pb.Max = rs.RecordCount
        pb.Text = "Generando...."
        Dim i As Integer
        Dim alveo As Long
        Dim rs_aux As ADODB.RecordSet
        Dim rs_deter As ADODB.RecordSet
        Dim rs_ad As ADODB.RecordSet
        Dim numero_deter As Integer
        Dim XLA As Excel.Application
        Dim XLW As Excel.Workbook
        Dim XLS As Excel.Worksheet
        Dim oalveo As New clsAlveogramas
        Set XLA = New Excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
'        XLW.Name = rs(8)
'        XLS.Name = rs(8)
        Dim tipo_analisis_anterior As Integer
        tipo_analisis_anterior = rs(3)
        XLW.Worksheets(1).Name = Left(Replace(rs(8), ":", " "), 31)
        XLA.Visible = False
'        XLS.Range("1:1").HorizontalAlignment = xlCenter
'        XLS.Range("1:1").VerticalAlignment = xlCenter
'        XLS.Range("1:1").RowHeight = 70
'        XLS.Range("1:1").WrapText = True
        'Cabecera
        i = 1
        cabecera i, rs(3), XLS, rs(13) ' TIPO_ANALISIS_ID , BANO_ID
        i = i + 1
        ' Datos
'        XLS.Range(XLS.Cells(i, 6), XLS.Cells(i + 1000, 8)) = "0.00"
        Do
            XLS.Cells(i, 1) = rs(9)
            XLS.Cells(i, 2) = rs(1)
            XLS.Cells(i, 3) = rs(2)
            XLS.Cells(i, 4) = rs(11)
            XLS.Cells(i, 5) = rs(8)
            XLS.Cells(i, 6) = Format(rs(5), "yyyy-mm-dd")
            XLS.Cells(i, 7) = Format(rs(12), "yyyy-mm-dd")
            XLS.Cells(i, 8) = Format(rs(10), "yyyy-mm-dd")
            XLS.Range(XLS.Cells(i, 6), XLS.Cells(i, 6)).HorizontalAlignment = xlRight
            XLS.Range(XLS.Cells(i, 7), XLS.Cells(i, 7)).HorizontalAlignment = xlRight
            XLS.Range(XLS.Cells(i, 8), XLS.Cells(i, 8)).HorizontalAlignment = xlRight
            XLS.Cells(i, 9) = rs(4)
            ' Determinaciones
            consulta = "select * from determinaciones where muestra_id = " & rs(6) & " order by orden"
            Set rs_deter = datos_bd(consulta)
            Col = 10
            If rs_deter.RecordCount > 0 Then
                Do
                    If IsNumeric(rs_deter("resultado")) Then
                        XLS.Range(XLS.Cells(i, Col), XLS.Cells(i, Col)).NumberFormat = "0.00"
                        XLS.Cells(i, Col) = CSng(Replace(rs_deter("resultado"), ".", ","))
                    Else
                        XLS.Cells(i, Col) = rs_deter("resultado")
                    End If
                    Col = Col + 1
                    ' Alveograma
                    If rs_deter("tipo_determinacion_id") = CInt(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma")) Then
                        alveo = oalveo.ComprobarAlveograma(rs(6), rs_deter("id_determinacion"))
                        If alveo <> 0 Then
                            ' Alveograma_valores
                            Dim oalveo_val As New clsAlveograma_valores
                            With oalveo_val
                             If .CargarAlveogramaValores(alveo, 0) = True Then
                                XLS.Cells(i, Col + 0) = CSng(Replace(.getTENACIDAD, ".", ","))
                                XLS.Cells(i, Col + 1) = CSng(Replace(.getEXTENSIBILIDAD, ".", ","))
                                XLS.Cells(i, Col + 2) = CSng(Replace(.getTENACIDAD / .getEXTENSIBILIDAD, ".", ","))
                                XLS.Cells(i, Col + 3) = CSng(Replace(.getS, ".", ","))
                                XLS.Cells(i, Col + 4) = CSng(Replace(.getS, ".", ",") * 6.54)
                                XLS.Cells(i, Col + 5) = CSng(Replace(.getEXTENSIBILIDAD, ".", ",") * 2.22)
                             End If
                             If .CargarAlveogramaValores(alveo, 1) = True Then
                                XLS.Cells(i, Col + 6) = CSng(Replace(.getTENACIDAD, ".", ","))
                                XLS.Cells(i, Col + 7) = CSng(Replace(.getEXTENSIBILIDAD, ".", ","))
                                XLS.Cells(i, Col + 8) = CSng(Replace(.getTENACIDAD / .getEXTENSIBILIDAD, ".", ","))
                                XLS.Cells(i, Col + 9) = CSng(Replace(.getS, ".", ","))
                                XLS.Cells(i, Col + 10) = CSng(Replace(.getS, ".", ",") * 6.54)
                                XLS.Cells(i, Col + 11) = CSng(Replace(.getEXTENSIBILIDAD, ".", ",") * 2.22)
                             End If
                            End With
                            Col = Col + 12
                        End If
                    End If
                    rs_deter.MoveNext
                Loop Until rs_deter.EOF
            End If
            i = i + 1
            rs.MoveNext
            If rs.EOF = False Then
              If rs(3) <> tipo_analisis_anterior Then
                XLW.Worksheets.Add
                Set XLS = XLW.Worksheets(1)
                XLW.Worksheets(1).Name = Left(Replace(rs(8), ":", " "), 31)
                i = 1
                cabecera i, rs(3), XLS, rs(13)
                i = i + 1
                tipo_analisis_anterior = rs(3)
              End If
            End If
            If pb.value < pb.Max Then
                pb.value = pb.value + 1
            End If
        Loop Until rs.EOF
    End If
    Set rs = Nothing
        Me.MousePointer = 0
    pb.Text = "Finalizado...."
    XLA.Visible = True
'    MsgBox "Proceso terminado.", vbInformation, App.Title

   On Error GoTo 0
   Exit Sub

cmdExcel_Click_Error:

        Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExcel_Click of Formulario frmInformeRegistro"
End Sub

Private Sub cmdInforme_Click()
    Dim rs As ADODB.RecordSet

    Set rs = listado_muestras()
    If rs.RecordCount >= 1 Then
        Dim oMuestra As New clsMuestra
        If MsgBox("¿Va a imprimir " & rs.RecordCount & " ensayos. ¿Esta seguro?", vbQuestion + vbYesNo, App.Title) = vbYes Then
            Do
                oMuestra.Informe_Recepcion rs(6), True
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set oMuestra = Nothing
    End If
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        cmdSalir_Click
    End If
End Sub
Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 20
    Me.Top = 20
    cargar_clientes
    cargar_muestras
    fdesde = Date
    fhasta = Date
End Sub
Public Sub cargar_clientes()
    Dim ocliente As New clsCliente
    Set cmbClientes.RowSource = ocliente.Listado("", "", "") 'recorset devuelto por la funcion
    cmbClientes.ListField = "nombre" 'campo que veo
    cmbClientes.DataField = "id_cliente" 'campo asociado
    cmbClientes.BoundColumn = "id_cliente" 'lo que realmente envia
    Set ocliente = Nothing
End Sub
Public Sub cargar_muestras()
    Dim oMuestra As New clsTipos_muestra
    Set cmbMuestras.RowSource = oMuestra.Listado_todas
    cmbMuestras.ListField = "nombre" 'lo que enseña
    cmbMuestras.DataField = "id_tipo_muestra" 'campo asociado
    cmbMuestras.BoundColumn = "id_tipo_muestra" 'lo que realmente envia
    Set oMuestra = Nothing
End Sub
Private Function listado_muestras() As ADODB.RecordSet
    On Error GoTo fallo
    Dim consulta As String
    Dim strMuestra As String
    Dim strClientes As String
    Dim f_desde As String
    Dim f_hasta As String
    f_desde = Format(fdesde, "yyyy-mm-dd")
    f_hasta = Format(fhasta, "yyyy-mm-dd")
    ' Tipo de muestra
    strMuestra = ""
    If chkTodas.value = Unchecked Then
        If cmbMuestras.Text = "" Then
            MsgBox "Debe seleccionar un tipo de muestras.", vbExclamation, App.Title
            Exit Function
        End If
        strMuestra = " AND mu.tipo_muestra_id=" & cmbMuestras.BoundText
    End If
    ' Clientes
    strClientes = ""
    If chkTodos.value = Unchecked Then
        If cmbClientes.Text = "" Then
            MsgBox "Debe seleccionar un cliente.", vbExclamation, App.Title
            Exit Function
        End If
        strClientes = " AND mu.cliente_id = " & cmbClientes.BoundText
    End If
    ' Fechas
    Dim FECHA_DESDE As String
    FECHA_DESDE = " AND mu.fecha_recepcion>='" & f_desde & "'"
    Dim FECHA_HASTA As String
    FECHA_HASTA = " AND mu.fecha_recepcion<='" & f_hasta & "'"
    consulta = "SELECT cl.id_cliente, " & _
               "concat(tm.codigo,'-',CAST(mu.id_particular AS CHAR)), " & _
               "cl.nombre, " & _
               "mu.tipo_analisis_id, " & _
               "mu.referencia_cliente, " & _
               "mu.fecha_recepcion, " & _
               "mu.id_muestra, " & _
               "mu.precio, " & _
               "ta.nombre, " & _
               "mu.id_general, " & _
               "mu.fecha_cierre, " & _
               "tm.nombre, " & _
               "mu.fecha_muestreo,mu.bano_id " & _
               "FROM clientes as cl, " & _
                     "tipos_muestra as tm, " & _
                     "tipos_analisis as ta, " & _
                     "muestras as mu " & _
               "WHERE mu.cliente_id=cl.id_cliente AND " & _
                      "mu.tipo_muestra_id=tm.id_tipo_muestra AND " & _
                      "mu.tipo_analisis_id=ta.id_tipo_analisis " & _
                      FECHA_DESDE & FECHA_HASTA & _
                      strMuestra & _
                      strClientes & _
                      " order by mu.tipo_analisis_id,mu.id_muestra asc"
    Set listado_muestras = datos_bd(consulta)
    Exit Function
fallo:
    Me.MousePointer = 0
    MsgBox "Se ha producido un error al generar el listado.", vbCritical, Err.Description
End Function
Private Sub cmdListado_Click()
    Dim rs As ADODB.RecordSet
   On Error GoTo cmdListado_Click_Error

    Set rs = listado_muestras()
    If rs.RecordCount >= 1 Then
        Dim rs_aux As ADODB.RecordSet
        Dim rs_deter As ADODB.RecordSet
        Dim appword As Word.Application
        Dim docword As Word.Document
        ' Crear copia para su uso
        Set appword = CreateObject("word.application")
        Set docword = appword.Documents.Open(copiar_plantilla("Informe", 0, 0))
        appword.Visible = True
        Do
            With docword.Sections(1).Range
              ' Datos generales
              .InsertAfter "------------------------------------"
'              .InsertAfter "Datos generales" & vbCrLf
              .InsertAfter "Ensayo: " & rs(1)
              .InsertAfter "------------------------------------" & vbCrLf
              .InsertAfter "Cliente: " & rs(2) & vbCrLf
              .InsertAfter "Análisis: " & rs(8) & vbCrLf
              .InsertAfter "Referencia Cliente: " & rs(4) & vbCrLf
              ' Determinaciones
              consulta = "select tipo_determinacion_id from determinaciones where muestra_id = " & rs(6)
              Set rs_deter = datos_bd(consulta)
              If rs_deter.RecordCount > 0 Then
                  .InsertAfter "---- Listado de Determinaciones ----" & vbCrLf
                  Do
                    consulta = "select nombre from tipos_determinacion where id_tipo_determinacion = " & rs_deter(0)
                    Set rs_aux = datos_bd(consulta)
                    If rs_aux.RecordCount > 0 Then
                      .InsertAfter rs_aux(0) & vbCrLf
                    End If
                    rs_deter.MoveNext
                  Loop Until rs_deter.EOF
              End If
            End With
            rs.MoveNext
        Loop Until rs.EOF
        docword.Save
        Set docword = Nothing
        Set appword = Nothing
    End If
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cmdListado_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdListado_Click of Formulario frmInformeRegistro"
End Sub

Public Sub cabecera(fila As Integer, TIPO_ANALISIS As Integer, XLS As Excel.Worksheet, BANO_ID As Long)
    Dim consulta As String
    If BANO_ID <> 0 Then
        c = " and a.bano_id = " & BANO_ID
    Else
        c = " and a.tipo_analisis_id = " & TIPO_ANALISIS
    End If
    consulta = "select b.nombre,b.id_tipo_determinacion from determinaciones_analisis a, tipos_determinacion b " & _
               " where a.tipo_determinacion_id = b.id_tipo_determinacion " & _
               c & _
               " order by a.orden"
               
    XLS.Cells(fila, 1) = "General"
    XLS.Cells(fila, 2) = "Particular"
    XLS.Cells(fila, 3) = "Cliente"
    XLS.Cells(fila, 4) = "T.Muestra"
    XLS.Cells(fila, 5) = "T.Análisis"
    XLS.Cells(fila, 6) = "F.Recepcion"
    XLS.Cells(fila, 7) = "F.Comienzo"
    XLS.Cells(fila, 8) = "F.Cierre"
    XLS.Cells(fila, 9) = "Referencia"
        
    XLS.Range("1:1").HorizontalAlignment = xlCenter
    XLS.Range("1:1").VerticalAlignment = xlCenter
    XLS.Range("1:1").RowHeight = 70
    XLS.Range("1:1").WrapText = True
    
    Dim rs_deter As ADODB.RecordSet
    Set rs_deter = datos_bd(consulta)
    Col = 10
    If rs_deter.RecordCount > 0 Then
        Do
            XLS.Cells(fila, Col) = rs_deter(0)
            Col = Col + 1
            ' Alveograma
            If rs_deter(1) = CInt(ReadINI(App.Path + "\config.ini", "Alveograma", "Alveograma")) Then
                XLS.Cells(fila, Col + 0) = "TENACIDAD (Normal)"
                XLS.Cells(fila, Col + 1) = "EXTENSIBILIDAD (Normal)"
                XLS.Cells(fila, Col + 2) = "P/L (Normal)"
                XLS.Cells(fila, Col + 3) = "S (Normal)"
                XLS.Cells(fila, Col + 4) = "W (Normal)"
                XLS.Cells(fila, Col + 5) = "G (Normal)"
                XLS.Cells(fila, Col + 6) = "TENACIDAD (Reposo)"
                XLS.Cells(fila, Col + 7) = "EXTENSIBILIDAD (Reposo)"
                XLS.Cells(fila, Col + 8) = "P/L (Reposo)"
                XLS.Cells(fila, Col + 9) = "S (Reposo)"
                XLS.Cells(fila, Col + 10) = "W (Reposo)"
                XLS.Cells(fila, Col + 11) = "G (Reposo)"
                Col = Col + 12
            End If
            rs_deter.MoveNext
        Loop Until rs_deter.EOF
    End If
End Sub


