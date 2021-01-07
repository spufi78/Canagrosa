VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEmpleados_Nominas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control nominas de Empleados"
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9960
   Icon            =   "frmEmpleados_Nominas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7335
   ScaleWidth      =   9960
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   915
      Left            =   8790
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   6330
      Width           =   1125
   End
   Begin VB.CommandButton cmdok 
      BackColor       =   &H00E0E0E0&
      Caption         =   "F10-Aceptar"
      Height          =   915
      Left            =   7620
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   6330
      Width           =   1125
   End
   Begin VB.TextBox txtdatos 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   5
      Left            =   8010
      MaxLength       =   30
      TabIndex        =   5
      Top             =   5940
      Width           =   1830
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe"
      Height          =   915
      Left            =   6420
      Picture         =   "frmEmpleados_Nominas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   6330
      Width           =   1125
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Datos nómina"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   6210
      TabIndex        =   16
      Top             =   1935
      Width           =   3660
      Begin VB.TextBox txtdatos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   1710
         Locked          =   -1  'True
         MaxLength       =   30
         TabIndex        =   3
         Top             =   1665
         Width           =   1830
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   2
         Top             =   1260
         Width           =   1830
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   1
         Top             =   855
         Width           =   1830
      End
      Begin VB.TextBox txtdatos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   1710
         MaxLength       =   30
         TabIndex        =   0
         Top             =   450
         Width           =   1830
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T. Líquido"
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
         Index           =   5
         Left            =   135
         TabIndex        =   20
         Top             =   1710
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "IRPF"
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
         Left            =   135
         TabIndex        =   19
         Top             =   1305
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T. Aportaciones"
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
         Left            =   135
         TabIndex        =   18
         Top             =   900
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "T. Devengado"
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
         Left            =   135
         TabIndex        =   17
         Top             =   495
         Width           =   1305
      End
   End
   Begin VB.ComboBox cmbMes 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmEmpleados_Nominas.frx":0BD4
      Left            =   6795
      List            =   "frmEmpleados_Nominas.frx":0BFC
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   540
      Width           =   1815
   End
   Begin VB.TextBox txtanno 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6795
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   990
      Width           =   825
   End
   Begin VB.TextBox txtdatos 
      Appearance      =   0  'Flat
      Height          =   1515
      Index           =   4
      Left            =   6210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   4365
      Width           =   3750
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6570
      Left            =   45
      TabIndex        =   7
      Top             =   450
      Width           =   6075
      _ExtentX        =   10716
      _ExtentY        =   11589
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
   Begin MSComCtl2.UpDown cambiar 
      Height          =   360
      Left            =   7620
      TabIndex        =   10
      Top             =   990
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   635
      _Version        =   393216
      Value           =   2004
      BuddyControl    =   "txtanno"
      BuddyDispid     =   196616
      OrigLeft        =   1590
      OrigTop         =   6570
      OrigRight       =   1830
      OrigBottom      =   6975
      Max             =   2099
      Min             =   1990
      SyncBuddy       =   -1  'True
      BuddyProperty   =   0
      Enabled         =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Seguridad Social"
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
      Index           =   6
      Left            =   6210
      TabIndex        =   21
      Top             =   5985
      Width           =   1560
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Año"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   12
      Left            =   6210
      TabIndex        =   15
      Top             =   1050
      Width           =   585
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   1350
      Left            =   8685
      Stretch         =   -1  'True
      Top             =   495
      Width           =   1185
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Observaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   2
      Left            =   6210
      TabIndex        =   14
      Top             =   4095
      Width           =   1590
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   6210
      TabIndex        =   13
      Top             =   630
      Width           =   450
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Haga doble click para ver el detalle del expediente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   765
      TabIndex        =   12
      Top             =   7065
      Width           =   4365
   End
   Begin VB.Label lbltitulo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      Caption         =   "Control nominas de Empleados"
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
      Left            =   45
      TabIndex        =   11
      Top             =   45
      Width           =   9870
   End
End
Attribute VB_Name = "frmEmpleados_Nominas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbMes_Click()
    lista_Click
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub
Private Sub cmdok_Click()
    If lista.ListItems.Count > 0 Then
        insertar_control
    End If
End Sub

Private Sub Command1_Click()
    Dim rs As ADODB.Recordset
    Dim fila As Integer
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Open(App.Path & "\informes\nominas.xls")
    Set XLS = XLW.Worksheets(1)
    fila = 5
    On Error Resume Next
    XLA.Application.ErrorCheckingOptions.TextDate = False
    XLA.Application.ErrorCheckingOptions.NumberAsText = False
    On Error GoTo fallo
    Dim i As Integer
    Dim op As New clsEmpleados
    Dim opexp As New clsEmpleados_Expediente
    Dim opcon As New clsEmpleados_Control
    Dim oon As New clsEmpleados_Nominas
    Dim ooa As New clsEmpleados_Anticipo
    Set rs = op.Listado_Nominas
    ' Mes
    XLS.Cells(1, 4) = "MES DE " & cmbMes.Text & " DEL " & txtanno
    ' Precio kilometro
    XLS.Cells(2, 2) = ReadINI(App.Path & "\config.ini", "parametros", "Precio_kilometro")
    If rs.RecordCount > 0 Then
        Do
            ' Copia de fila
            If fila > 5 Then
                  XLS.Range(XLS.Cells(fila - 1, 7), XLS.Cells(fila - 1, 7)).Copy
                  XLS.Range(XLS.Cells(fila, 7), XLS.Cells(fila, 7)).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                  XLS.Range(XLS.Cells(fila - 1, 10), XLS.Cells(fila - 1, 10)).Copy
                  XLS.Range(XLS.Cells(fila, 10), XLS.Cells(fila, 10)).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                  XLS.Range(XLS.Cells(fila - 1, 11), XLS.Cells(fila - 1, 11)).Copy
                  XLS.Range(XLS.Cells(fila, 11), XLS.Cells(fila, 11)).PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
                  XLS.Range(XLS.Cells(fila - 1, 13), XLS.Cells(fila - 1, 16)).Copy
                  XLS.Range(XLS.Cells(fila, 13), XLS.Cells(fila, 16)).PasteSpecial Paste:=xlPasteFormulasAndNumberFormats, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
            End If
            ' Datos operario
            XLS.Cells(fila, 2) = rs("codigo_interno")
            XLS.Cells(fila, 3) = rs("nombre")
            ' Datos Expediente
            If opexp.Cargar_Ultimo(rs("id_operario")) = True Then
                XLS.Cells(fila, 4) = Format(opexp.getPRECIO_HORA_EXTRA, "currency")
                XLS.Cells(fila, 8) = Format(opexp.getPLUS, "currency")
            End If
            ' Datos Control
            XLS.Cells(fila, 5) = opcon.Totales_mes(rs("id_operario"), 2, cmbMes.ListIndex + 1, txtanno)
            XLS.Cells(fila, 6) = opcon.Totales_mes(rs("id_operario"), 3, cmbMes.ListIndex + 1, txtanno)
            ' Anticipos
            XLS.Cells(fila, 11) = ooa.Totales_mes(rs("id_operario"), cmbMes.ListIndex + 1, txtanno)
            ' Datos nomina
            XLS.Cells(fila, 14) = oon.Liquido(rs("id_operario"), cmbMes.ListIndex + 1, txtanno)
            fila = fila + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    On Error Resume Next
    Kill App.Path & "\informes\Nominas_Generado.xls"
    XLS.SaveAs App.Path & "\informes\Nominas_Generado.xls"
    XLA.Visible = True
    Set XLW = Nothing
    Set XLA = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    MsgBox "Se ha producido un error al generar el informe de nóminas." & Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Load()
    Me.Left = 20
    Me.top = 20
    cabecera_lista
    cargar_botones Me
    ' Fecha
    txtanno = Year(Date)
    cmbMes.ListIndex = Month(Date) - 2
    cargar_lista
End Sub
Public Sub cabecera_lista()
    With lista.ColumnHeaders.Add(, , "C.Interno", 1000, lvwColumnLeft)
        .Tag = "C.Interno"
    End With
    With lista.ColumnHeaders.Add(, , "Nombre", 3000, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Apodo", 1500, lvwColumnLeft)
        .Tag = "Apodo"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 1, lvwColumnLeft)
        .Tag = "ID"
    End With
End Sub
Public Sub cargar_lista()
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim oPer As New clsEmpleados
    Set rs = oPer.Listado_Nominas
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs("codigo_interno"))
            .SubItems(1) = rs("nombre")
            .SubItems(2) = rs("apodo")
            .SubItems(3) = rs("id_operario")
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox "Error al cargar los datos." & Err.Description, vbCritical, App.Title
End Sub

Private Sub lista_Click()
    If lista.ListItems.Count > 0 Then
        consulta_Operario
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

Public Sub consulta_Operario()
   Dim oPer As New clsEmpleados
   oPer.CARGAR (lista.ListItems(lista.selectedItem.Index).SubItems(3))
   On Error Resume Next
   ' Foto
   Set img.Picture = Nothing
   borrar_campos
   If oPer.getfoto <> "" Then
     If Dir(oPer.getfoto) <> "" Then
       Set img.Picture = LoadPicture(oPer.getfoto)
     End If
   End If
   ' Datos
   Dim oon As New clsEmpleados_Nominas
   Dim rs As ADODB.Recordset
   Set rs = oon.Datos_Nomina(lista.ListItems(lista.selectedItem.Index).SubItems(3), cmbMes.ListIndex + 1, txtanno)
   If rs.RecordCount <> 0 Then
        txtDatos(0) = rs("total_devengado")
        txtDatos(1) = rs("aportaciones")
        txtDatos(2) = rs("retencion")
        txtDatos(3) = rs("total_liquido")
        txtDatos(4) = rs("comentario")
        txtDatos(5) = rs("ss")
   End If
   txtDatos(0).SetFocus
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gOperario = lista.ListItems(lista.selectedItem.Index).SubItems(3)
        frmEmpleados.Show
    End If
End Sub

Private Sub lista_KeyUp(KeyCode As Integer, Shift As Integer)
    lista_Click
End Sub

Private Sub txtanno_Change()
    lista_Click
End Sub

Private Sub txtdatos_GotFocus(Index As Integer)
    txtDatos(Index).BackColor = &H80C0FF
    txtDatos(Index).SelStart = 0
    txtDatos(Index).SelLength = Len(txtDatos(Index))
End Sub
Private Sub txtDatos_KeyPress(Index As Integer, KeyAscii As Integer)
    ' Escribir ',' al pulsar '.'
    If Index <> 4 And KeyAscii = 46 Then
         KeyAscii = 44
    End If
    If KeyAscii = 13 And Index <> 4 Then
        SendKeys "{Tab}", True
        KeyAscii = 0 ' Para evitar el "bip" del sistema
    End If
End Sub

Private Sub txtdatos_Keyup(Index As Integer, KeyCode As Integer, Shift As Integer)
    calcular_liquido
End Sub

Private Sub txtdatos_LostFocus(Index As Integer)
    txtDatos(Index).BackColor = &HFFFFFF
End Sub
Public Sub borrar_campos()
    Dim i As Integer
    For i = 0 To 5
        txtDatos(i) = ""
    Next
End Sub
Public Sub insertar_control()
    If valida_datos = False Then
        Exit Sub
    End If
    On Error GoTo fallo
'    pregunta = "Va a dar de alta una nueva nómina. ¿Esta seguro?"
'    If MsgBox(pregunta, vbQuestion + vbYesNo, App.Title) = vbYes Then
        Set oNomina = mover_datos(lista.ListItems(lista.selectedItem.Index).SubItems(3))
        oNomina.Insertar
'        MsgBox "El control de empleado se ha insertado correctamente.", vbInformation, App.Title
        borrar_campos
        txtDatos(0).SetFocus
'    End If
    ' Pasar al siguiente campo
    If lista.ListItems.Count > lista.selectedItem.Index Then
         Set lista.selectedItem = lista.ListItems(lista.selectedItem.Index + 1)
         lista_Click
    Else
         MsgBox "Ya se encuentra en el último registro.", vbInformation, App.Title
    End If
    Exit Sub
fallo:
    MsgBox "Se ha producido un error al insertar el control.", vbCritical, Err.Description
End Sub
Public Function valida_datos() As Boolean
    valida_datos = True
    Dim i As Integer
    For i = 0 To 3
        If IsNumeric(txtDatos(i)) = False Then
            MsgBox "El campo tiene que ser numérico.", vbCritical, "Error"
            txtDatos(i).SetFocus
            valida_datos = False
            Exit Function
        End If
    Next
End Function
Public Function mover_datos(indice As Integer) As clsEmpleados_Nominas
    On Error GoTo fallo
    Dim oNomina As New clsEmpleados_Nominas
    With oNomina
        .setMES = cmbMes.ListIndex + 1
        .setANNO = txtanno
        .setEMPLEADO_ID = indice
        .setTOTAL_DEVENGADO = txtDatos(0)
        .setAPORTACIONES = txtDatos(1)
        .setRETENCION = txtDatos(2)
        .setTOTAL_LIQUIDO = txtDatos(3)
        .setCOMENTARIO = txtDatos(4)
        .setSS = txtDatos(5)
    End With
    Set mover_datos = oNomina
    Set oNomina = Nothing
    Exit Function
fallo:
    MsgBox "Error al mover los datos de la nómina.", vbCritical, Err.Description
End Function

Public Sub calcular_liquido()
    If IsNumeric(txtDatos(0)) And IsNumeric(txtDatos(1)) And IsNumeric(txtDatos(2)) Then
        txtDatos(3) = CCur(txtDatos(0)) - (CCur(txtDatos(1)) + CCur(txtDatos(2)))
    End If
End Sub
