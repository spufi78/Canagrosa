VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmEmpleados_Listado 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Empleados"
   ClientHeight    =   9645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14250
   Icon            =   "frmEmpleados_Listado.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9645
   ScaleWidth      =   14250
   Begin VB.CommandButton cmdVerExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "EXCEL"
      Height          =   915
      Left            =   4860
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   8685
      Width           =   1185
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "LI - 26 Formadores"
      Height          =   915
      Index           =   1
      Left            =   5985
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8685
      Width           =   1185
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
      Height          =   1410
      Left            =   45
      TabIndex        =   8
      Top             =   585
      Width           =   14190
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   3
         Left            =   8865
         MaxLength       =   75
         TabIndex        =   23
         Top             =   810
         Width           =   2400
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   2
         Left            =   5220
         MaxLength       =   75
         TabIndex        =   20
         Top             =   810
         Width           =   2400
      End
      Begin VB.CheckBox chkexterno 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Externo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   11520
         TabIndex        =   19
         Top             =   855
         Width           =   1005
      End
      Begin VB.CommandButton cmdLimpiarCampos 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar Campos"
         Height          =   735
         Left            =   12645
         Picture         =   "frmEmpleados_Listado.frx":09EA
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   315
         Width           =   1410
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   0
         Left            =   5220
         MaxLength       =   75
         TabIndex        =   11
         Top             =   360
         Width           =   2400
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   330
         Index           =   1
         Left            =   1170
         MaxLength       =   75
         TabIndex        =   9
         Top             =   360
         Width           =   2265
      End
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   315
         Left            =   8865
         TabIndex        =   13
         Top             =   360
         Width           =   2385
         _ExtentX        =   4207
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
      Begin MSDataListLib.DataCombo cmbempresas 
         Height          =   315
         Left            =   1170
         TabIndex        =   17
         Top             =   810
         Width           =   2250
         _ExtentX        =   3969
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "D.N.I."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   8010
         TabIndex        =   24
         Top             =   855
         Width           =   480
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Departamento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   3690
         TabIndex        =   21
         Top             =   855
         Width           =   1290
      End
      Begin VB.Label lblEmpresa 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Empresa"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   18
         Top             =   810
         Width           =   825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Estado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   8010
         TabIndex        =   14
         Top             =   405
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Cod. Interno"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   3825
         TabIndex        =   12
         Top             =   405
         Width           =   1065
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   10
         Top             =   405
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdanadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   915
      Left            =   45
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   8685
      Width           =   1185
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   885
      Left            =   13065
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   8670
      Width           =   1155
   End
   Begin VB.CommandButton cmdmodificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   915
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8685
      Width           =   1185
   End
   Begin VB.CommandButton cmdeliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   915
      Left            =   2475
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8685
      Width           =   1185
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado"
      Height          =   915
      Index           =   0
      Left            =   3645
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   8685
      Width           =   1185
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6600
      Left            =   30
      TabIndex        =   0
      Top             =   2010
      Width           =   14190
      _ExtentX        =   25030
      _ExtentY        =   11642
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
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "FICHA DE PERSONAL"
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
      TabIndex        =   7
      Top             =   45
      Width           =   2325
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13635
      Picture         =   "frmEmpleados_Listado.frx":723C
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Detalle de ficha de Personal"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   315
      Width           =   1995
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   600
      Left            =   0
      Top             =   0
      Width           =   14175
   End
End
Attribute VB_Name = "frmEmpleados_Listado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub chkexterno_Click()
    cargar_lista
End Sub

Private Sub cmbempresas_Change()
    cargar_lista
End Sub

Private Sub cmbestados_Change()
    cargar_lista
End Sub
Private Sub cmdAnadir_Click()
    frmEmpleados_Gestion.PK = 0
    frmEmpleados_Gestion.Show 1
    cargar_lista
    lista.SetFocus
End Sub
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a ELIMINAR al Empleado: " & lista.ListItems(lista.selectedItem.Index).SubItems(1) & ". ¿Esta seguro?", vbExclamation + vbYesNo, "Informacion") = vbYes Then
            Dim oempleado As New clsEmpleados
            If oempleado.Eliminar(lista.ListItems(lista.selectedItem.Index)) = True Then
                cargar_lista
            End If
            Set oempleado = Nothing
        End If
        lista.SetFocus
    End If
End Sub
Private Sub cmdImprimir_Click(Index As Integer)
    Select Case Index
    Case 0
 'M1002-I
        Dim strCad As String
        Dim arrNom() As String
        Dim arrVal() As String
        Dim objfrm As New frmReport
        
        With objfrm
            .iniciar
            .informe = "Empleados\rptEmpleados"
            
            ReDim arrNom(5)
            ReDim arrVal(5)
            
            arrNom(1) = "nombre"
            arrVal(1) = Trim(txtdatos(1).Text)
            
            arrNom(2) = "codigo"
            arrVal(2) = Trim(txtdatos(0).Text)
            
            arrNom(3) = "estado"
            arrVal(3) = cmbestados.BoundText
            
            arrNom(4) = "external"
            arrVal(4) = chkexterno.Value
            
            arrNom(5) = "empresa"
            If Trim(cmbempresas.BoundText) = "" Then
                arrVal(5) = "1 to 3"
            Else
                arrVal(5) = cmbempresas.BoundText & " to " & cmbempresas.BoundText
            End If
                    
            .ParametrosNombre = arrNom
            .ParametrosValores = arrVal
            
            .criterio = "{empleados.NOMBRE} like '*" & Trim(txtdatos(1).Text) & "*' and {empleados.CODIGO_INTERNO} like '*" & Trim(txtdatos(0).Text) & "*' and {empleados.ESTADO_ID} = " & cmbestados.BoundText & " and {empleados.EXTERNAL} = " & chkexterno.Value & " and {empleados.EMPRESA_ID} IN " & arrVal(5)
            .imprimir = False
            .generar
            .Visible = True
            
        End With
'M01002-F
    Case 1
        frmReport.iniciar
        frmReport.informe = "\Empleados\rptEmpleados_LI26"
        frmReport.criterio = ""
        frmReport.imprimir = False
        frmReport.generar
        frmReport.Show 1
        Unload frmReport
    
    End Select
End Sub

Private Sub cmdLimpiarCampos_Click()
    txtdatos(0) = ""
    txtdatos(1) = ""
    cmbestados.BoundText = ""
    cmbestados.Text = ""
'M0970-I
    cmbempresas.BoundText = ""
    cmbempresas.Text = ""
'M0970-F
    cargar_lista
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        If USUARIO.getPER_MODIFICACION = 0 Then
            Exit Sub
        End If
        frmEmpleados_Gestion.PK = lista.ListItems(lista.selectedItem.Index)
        frmEmpleados_Gestion.Show 1
        actualizar_lista
        lista.SetFocus
    End If
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Codigo", 520, lvwColumnLeft
        .Add , , "Nombre", 3200, lvwColumnLeft
        .Add , , "Direccion", 3200, lvwColumnLeft
        .Add , , "D.N.I.", 1200, lvwColumnCenter
        .Add , , "Telefono", 1350, lvwColumnCenter
        .Add , , "Movil", 1350, lvwColumnCenter
        .Add , , "F.Alta", 1000, lvwColumnCenter
        .Add , , "F.Baja", 1000, lvwColumnCenter
'M0970-I
        .Add , , "Empresa", 1000, lvwColumnCenter
        .Add , , "Departamentos", 1, lvwColumnCenter
        .Add , , "ID_EMPRESA", 1, lvwColumnCenter
    End With
End Sub

Private Sub cmdVerExcel_Click()
        Dim rsEmple As New ADODB.Recordset
        Dim oEmpresas As New clsEmpleados_Empresas
        Dim oExpediente As New clsEmpleados_Expediente
        Dim total As Integer
        Dim rango As String
        Dim fechaI As String
        Dim fechaF As String
        Dim cadena As String
        
        With rsEmple
        .Fields.Append "c1", adChar, 10, adFldUpdatable
        .Fields.Append "c2", adChar, 100, adFldUpdatable
        .Fields.Append "c3", adChar, 320, adFldUpdatable
        .Fields.Append "c4", adChar, 100, adFldUpdatable
        .Fields.Append "c5", adChar, 100, adFldUpdatable
        .Fields.Append "c6", adChar, 30, adFldUpdatable
        .Fields.Append "c7", adChar, 30, adFldUpdatable
        .Fields.Append "c8", adChar, 35, adFldUpdatable
        .Fields.Append "c9", adChar, 200, adFldUpdatable
        .Fields.Append "c10", adChar, 200, adFldUpdatable
        .Fields.Append "c11", adChar, 200, adFldUpdatable
        .Open
        
        Dim i As Integer
        For i = 1 To lista.ListItems.Count
            .AddNew
            rsEmple("c1") = lista.ListItems(i).Text
            rsEmple("c2") = lista.ListItems(i).SubItems(1)
            rsEmple("c3") = lista.ListItems(i).SubItems(2)
            rsEmple("c4") = lista.ListItems(i).SubItems(3)
            rsEmple("c5") = lista.ListItems(i).SubItems(4)
            rsEmple("c6") = lista.ListItems(i).SubItems(5)
            rsEmple("c7") = lista.ListItems(i).SubItems(6)
            rsEmple("c8") = lista.ListItems(i).SubItems(7)
            rsEmple("c9") = lista.ListItems(i).SubItems(8)
            rsEmple("c10") = lista.ListItems(i).SubItems(10)
            rsEmple("c11") = lista.ListItems(i).SubItems(9)
            rsEmple.Update
        Next i
        End With
        
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Listado de empleados"
        XLA.Visible = True
        total = lista.ListItems.Count + 1
        rango = "1:" & total
        XLS.Range(rango).HorizontalAlignment = xlCenter
        XLS.Range(rango).VerticalAlignment = xlCenter
        XLS.Range(rango).RowHeight = 25
        XLS.Range(rango).WrapText = True

        'Cabecera
        XLS.Cells(1, 1) = "Código"
        XLS.Cells(1, 2) = "Nombre"
        XLS.Cells(1, 3) = "Dirección"
        XLS.Cells(1, 4) = "DNI"
        XLS.Cells(1, 5) = "Tipo de contrato"
        XLS.Cells(1, 6) = "Antigüedad"
        XLS.Cells(1, 7) = "Departamento"
        XLS.Cells(1, 8) = "Empresa"
        i = 2
        If rsEmple.RecordCount > 0 Then
          rsEmple.MoveFirst
          Do
            XLS.Cells(i, 1) = rsEmple("c1")
            XLS.Cells(i, 2) = rsEmple("c2")
            XLS.Cells(i, 3) = rsEmple("c3")
            XLS.Cells(i, 4) = rsEmple("c4")
            oExpediente.Cargar_Ultimo CLng(rsEmple("c1"))
            XLS.Cells(i, 5) = Trim(oExpediente.getTIPO_CONTRATO)
            fechaI = Format(Trim(rsEmple("c7")), "yyyy-mm-dd")
            fechaF = Format(Trim(rsEmple("c8")), "yyyy-mm-dd")
            If fechaF <> "" Then
               cadena = DateDiff("yyyy", fechaI, fechaF) & " años "
            Else
               fechaF = Format(Date, "yyyy-mm-dd")
               cadena = DateDiff("yyyy", fechaI, fechaF) & " años "
            End If
            XLS.Cells(i, 6) = cadena

            oEmpresas.CARGAR CInt(Trim(rsEmple("c10")))
            XLS.Cells(i, 8) = oEmpresas.getDESCRIPCION
            XLS.Cells(i, 7) = DEPARTAMENTOS(rsEmple("c11"))
            i = i + 1
            rsEmple.MoveNext
          Loop Until rsEmple.EOF
        End If
        
    Set rsEmple = Nothing

End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27 ' esc
            Unload Me
    End Select
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 100
    Me.top = 100
    cabecera
    cargar_estados
'M0970-I
    cargar_empresas
'    cargar_lista
'M0970-F
    permisos
End Sub

Private Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim ocli As New clsEmpleados
    Dim oEmpresa As New clsEmpleados_Empresas
    Dim empresa As Long
    Dim departamento As Long
    
'M0970-I
'    Set rs = ocli.Listado(txtdatos(1), txtdatos(0), cmbestados.BoundText)
    empresa = 0
    If cmbempresas.Text <> "" Then
        empresa = cmbempresas.BoundText
    End If
'    If (cmbempresas.selectedItem = Null Or Not IsNumeric(cmbempresas.selectedItem)) Then
'       empresa = 1
'    Else
'       empresa = cmbempresas.selectedItem
'    End If
     
    Set rs = ocli.Listado(txtdatos(1), txtdatos(0), cmbestados.BoundText, chkexterno.Value, empresa, departamento, txtdatos(3))
    lblsubtitulo = "Se han localizado " & rs.RecordCount & " con los filtros aplicados."
'M0970-F

    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , Format(rs("id_empleado"), "000"))
            .SubItems(1) = rs("nombre")
            If IsNull(rs("direccion")) = False Then
                .SubItems(2) = rs("direccion")
            End If
            If IsNull(rs("cif")) = False Then
                .SubItems(3) = rs("cif")
            End If
            
            If IsNull(rs("telefono")) = False Then
                .SubItems(4) = rs("telefono")
            End If
            If IsNull(rs("movil")) = False Then
                .SubItems(5) = rs("movil")
            Else
                .SubItems(5) = ""
            End If
            .SubItems(6) = Format(rs("fecha_incorporacion"), "dd-mm-yyyy")
            If rs("estado_id") <> 0 Then
                .SubItems(7) = Format(rs("fecha_baja"), "dd-mm-yyyy")
            Else
                .SubItems(7) = ""
            End If
            .SubItems(8) = rs("descripcion")
            .SubItems(9) = rs("departamentos")
            .SubItems(10) = rs("empresa_id")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set ocli = Nothing
    Set oEmpresa = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
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
Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        If lista.ListItems(lista.selectedItem.Index) <> "" Then
          cmdmodificar.Enabled = True
          cmdeliminar.Enabled = True
        End If
        permisos
    End If
End Sub

Private Sub lista_DblClick()
    cmdModificar_Click
End Sub

Private Sub actualizar_lista()
    Dim ocli As New clsEmpleados
    If ocli.CARGAR(lista.ListItems(lista.selectedItem.Index)) = True Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = ocli.getNOMBRE
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = ocli.getDIRECCION
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = ocli.getCIF
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = ocli.getTELEFONO
        lista.ListItems(lista.selectedItem.Index).SubItems(5) = ocli.getMOVIL
        lista.ListItems(lista.selectedItem.Index).SubItems(6) = Format(ocli.getFECHA_INCORPORACION, "dd-mm-yyyy")
        If ocli.getESTADO_ID = 0 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(7) = " "
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(7) = Format(ocli.getFECHA_BAJA, "dd-mm-yyyy")
        End If
    End If
    Set ocli = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
       cmdModificar_Click
    End If
End Sub
Private Sub permisos()
    If USUARIO.getPER_EMPLEADOS = False Then
        cmdanadir.Enabled = False
        cmdmodificar.Enabled = False
        cmdeliminar.Enabled = False
    End If
End Sub

Private Sub cargar_estados()
    Dim ooe As New clsEmpleados_Estados
    Set cmbestados.RowSource = ooe.Listado
    cmbestados.ListField = "nombre"
    cmbestados.BoundColumn = "id_estado"
    Set ooe = Nothing
    cmbestados.BoundText = 0
End Sub

'M0970-I
Private Sub cargar_empresas()
    Dim ooe As New clsEmpleados_Empresas
    Set cmbempresas.RowSource = ooe.Listado
    cmbempresas.ListField = "DESCRIPCION"
    cmbempresas.BoundColumn = "ID_EMPRESA"
    Set ooe = Nothing
    cmbestados.BoundText = 0
End Sub
'M0970-F
Private Sub txtDatos_Change(Index As Integer)
    cargar_lista
End Sub
Private Function DEPARTAMENTOS(texto As String) As String
      'Obtención de la lista de departamentos (DESCRIPCIONES)
                                
        Dim strDepartamentos() As String
        Dim intCount As Integer
        Dim VALOR As Long
        Dim oDepart As New clsDecodificadora
        DEPARTAMENTOS = ""
        strDepartamentos = Split(Trim(texto), ";")
                
        For intCount = LBound(strDepartamentos) To UBound(strDepartamentos)

            If strDepartamentos(intCount) <> "" Then 'Para prevenir el caso de encontrar un ; al final de la línea de parámetros
                VALOR = CLng(Solo_Numeros(strDepartamentos(intCount)))
                oDepart.Carga_valor 50, VALOR
                'intcount: número de parámetros
                If intCount > LBound(strDepartamentos) Then
                    DEPARTAMENTOS = DEPARTAMENTOS & ", "
                End If
                DEPARTAMENTOS = DEPARTAMENTOS & oDepart.getDESCRIPCION
            End If
        Next intCount
End Function

'Public Function Solo_Numeros(ByRef sText As String) As String
'    Dim sActualChar                 As String * 1
'    Dim lTotalChar                  As Long
'    Dim x                           As Long
    
'    lTotalChar = LenB(sText) \ 2
'
'    If CBool(lTotalChar) Then
'        For x = 1 To lTotalChar
'            sActualChar = Mid$(sText, x, 1)
'            If IsNumeric(sActualChar) Then Solo_Numeros = Solo_Numeros & sActualChar
'        Next
'    End If
    
'End Function

