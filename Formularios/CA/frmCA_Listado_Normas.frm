VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmCA_Listado_Normas 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de NORMAS CONTROLADAS"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14040
   Icon            =   "frmCA_Listado_Normas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9300
   ScaleWidth      =   14040
   Begin VB.CommandButton cmdExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado Excel"
      Height          =   870
      Left            =   8190
      Style           =   1  'Graphical
      TabIndex        =   32
      Top             =   8370
      Width           =   1545
   End
   Begin VB.CommandButton cmdVincular 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Vincular"
      Height          =   870
      Left            =   11700
      Picture         =   "frmCA_Listado_Normas.frx":08CA
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8370
      Width           =   1230
   End
   Begin VB.CommandButton cmdListadoAuditoria 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado Auditoria"
      Height          =   870
      Left            =   6345
      Picture         =   "frmCA_Listado_Normas.frx":0F08
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   8370
      Width           =   1830
   End
   Begin VB.CommandButton cmdImprimir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Listado de Normas"
      Height          =   870
      Left            =   4470
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   8370
      Width           =   1830
   End
   Begin VB.CommandButton cmdMostrar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Mostrar"
      Height          =   870
      Index           =   0
      Left            =   3375
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   8370
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Filtrar por"
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
      Height          =   1320
      Left            =   45
      TabIndex        =   17
      Top             =   630
      Width           =   13965
      Begin VB.CheckBox chkMTL 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "MTL"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10215
         TabIndex        =   31
         Top             =   585
         Width           =   960
      End
      Begin VB.CheckBox chkConsulta 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Solo las no controladas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   11745
         TabIndex        =   28
         Top             =   1035
         Width           =   2010
      End
      Begin VB.CheckBox chkFTP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "FTP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10215
         TabIndex        =   27
         Top             =   1035
         Width           =   750
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   6255
         MaxLength       =   255
         TabIndex        =   7
         Top             =   960
         Width           =   3750
      End
      Begin VB.CheckBox chkEQA 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "EQA"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10215
         TabIndex        =   5
         Top             =   810
         Width           =   750
      End
      Begin VB.CheckBox chkENAC 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "ENAC"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10215
         TabIndex        =   3
         Top             =   135
         Width           =   810
      End
      Begin VB.CheckBox chkNADCAP 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "NADCAP"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   10215
         TabIndex        =   4
         Top             =   360
         Width           =   960
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   870
         Left            =   12825
         Picture         =   "frmCA_Listado_Normas.frx":17D2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   135
         Width           =   1050
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   870
         Left            =   11745
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   135
         Width           =   1050
      End
      Begin VB.TextBox txtDatos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1140
         MaxLength       =   255
         TabIndex        =   6
         Top             =   960
         Width           =   4020
      End
      Begin MSDataListLib.DataCombo cmbtipos 
         Height          =   315
         Left            =   1140
         TabIndex        =   0
         Top             =   225
         Width           =   4050
         _ExtentX        =   7144
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
      Begin MSDataListLib.DataCombo cmbestados 
         Height          =   315
         Left            =   1140
         TabIndex        =   2
         Top             =   600
         Width           =   4035
         _ExtentX        =   7117
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
      Begin MSDataListLib.DataCombo cmbSectores 
         Height          =   315
         Left            =   6255
         TabIndex        =   1
         Top             =   585
         Width           =   3780
         _ExtentX        =   6668
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
      Begin MSDataListLib.DataCombo cmbSubtipo 
         Height          =   315
         Left            =   6255
         TabIndex        =   23
         Top             =   225
         Width           =   3780
         _ExtentX        =   6668
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
         Caption         =   "Subtipo"
         Height          =   195
         Index           =   5
         Left            =   5535
         TabIndex        =   24
         Top             =   285
         Width           =   540
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         Height          =   195
         Index           =   4
         Left            =   5535
         TabIndex        =   22
         Top             =   990
         Width           =   495
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sector"
         Height          =   195
         Index           =   3
         Left            =   5535
         TabIndex        =   21
         Top             =   645
         Width           =   465
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Tipo"
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   20
         Top             =   285
         Width           =   315
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Pdte.Estado"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   19
         Top             =   645
         Width           =   870
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nombre"
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   18
         Top             =   960
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdEliminar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Eliminar"
      Height          =   870
      Left            =   2280
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8370
      Width           =   1050
   End
   Begin VB.CommandButton cmdModificar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Modificar"
      Height          =   870
      Left            =   1185
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8370
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnadir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Añadir"
      Height          =   870
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8370
      Width           =   1050
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   12960
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   8370
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6345
      Left            =   45
      TabIndex        =   10
      Top             =   1965
      Width           =   13965
      _ExtentX        =   24633
      _ExtentY        =   11192
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
      Caption         =   "Rellene todos los campos para la creación/modificación  de una norma"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   30
      Top             =   360
      Width           =   5025
   End
   Begin VB.Image imagen 
      Height          =   480
      Left            =   13455
      Picture         =   "frmCA_Listado_Normas.frx":209C
      Top             =   45
      Width           =   480
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Listado de NORMAS CONTROLADAS"
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
      TabIndex        =   29
      Top             =   45
      Width           =   3900
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      Height          =   645
      Left            =   0
      Top             =   0
      Width           =   14265
   End
End
Attribute VB_Name = "frmCA_Listado_Normas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'E0039-I
'Public FK_EQUIPO As Long
Public VINCULAR As Boolean
'E0039-F
Private Sub cmdExcel_Click()
   On Error GoTo cmdExcel_Click_Error

       Me.MousePointer = vbHourglass
       Dim rs As New ADODB.Recordset
       Dim fecha As String
      
       rs.Fields.Append "c1", adChar, 250, adFldUpdatable    'NORMA
       rs.Fields.Append "c2", adChar, 150, adFldUpdatable   'TIPO
       rs.Fields.Append "c3", adChar, 150, adFldUpdatable   'SUBTIPO
       rs.Fields.Append "c4", adChar, 150, adFldUpdatable    'CODIGO
       rs.Fields.Append "c5", adChar, 50, adFldUpdatable    'EDICION
       rs.Fields.Append "c6", adChar, 100, adFldUpdatable   'FECHA
       rs.Fields.Append "c7", adChar, 50, adFldUpdatable   'FTP
       rs.Fields.Append "c8", adChar, 50, adFldUpdatable   'MANTENIMIENTO
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
                rs("c7") = lista.ListItems(i).SubItems(7)
                rs("c8") = lista.ListItems(i).SubItems(8)
                rs.Update
'            End If
        Next i
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Add
        Set XLS = XLW.Worksheets(1)
        XLW.Worksheets(3).Delete
        XLW.Worksheets(2).Delete
        XLW.Worksheets(1).Name = "Listado de Normas"
        'Cabecera
        With XLS.Range("A1:H1")
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
        With XLS.Range("A1:H1").Interior
            .Pattern = xlSolid
            .PatternColorIndex = xlAutomatic
            .color = &HC0C0FF
        End With
        With XLS.Range("A1:P1").Borders
            .LineStyle = vbSolid
        End With
        XLS.Range("A1:A1").ColumnWidth = 70
        XLS.Range("B1:B1").ColumnWidth = 25
        XLS.Range("C1:C1").ColumnWidth = 25
        XLS.Range("D1:D1").ColumnWidth = 30
        XLS.Range("E1:E1").ColumnWidth = 10
        XLS.Range("F1:F1").ColumnWidth = 20
        XLS.Range("G1:G1").ColumnWidth = 10
        XLS.Range("H1:H1").ColumnWidth = 20

        XLS.Cells(1, 1) = "Norma"
        XLS.Cells(1, 2) = "Tipo"
        XLS.Cells(1, 3) = "Subtipo"
        XLS.Cells(1, 4) = "Codigo"
        XLS.Cells(1, 5) = "Edición"
        XLS.Cells(1, 6) = "Fecha"
        XLS.Cells(1, 7) = "Ftp"
        XLS.Cells(1, 8) = "Mantenimiento"
        i = 2
        If rs.RecordCount > 0 Then
          rs.MoveFirst
          Do
            XLS.Cells(i, 1) = rs("c1")
            XLS.Cells(i, 2) = ClrStr(rs("c2"), False, True, True)
            XLS.Cells(i, 3) = ClrStr(rs("c3"), False, True, True)
            XLS.Cells(i, 4) = ClrStr(rs("c4"), False, True, True)
            XLS.Cells(i, 5) = ClrStr(rs("c5"), False, True, True)
'            If IsDate(Trim(rs("c6"))) Then
'                XLS.Cells(i, 6) = CDate(Trim(rs("c6"))) ' Fecha
'            Else
                XLS.Cells(i, 6) = Trim(rs("c6")) ' Fecha
'            End If
            XLS.Cells(i, 7) = ClrStr(rs("c7"), False, True, True)
            XLS.Cells(i, 8) = ClrStr(rs("c8"), False, True, True)
            i = i + 1
            XLS.Range("A" & i).EntireRow.Insert
            rs.MoveNext
          Loop Until rs.EOF
        End If
        Me.MousePointer = vbNormal
        XLA.visible = True
        Set rs = Nothing

   On Error GoTo 0
   Exit Sub

cmdExcel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExcel_Click of Formulario frmCA_Listado_Normas, i : " & i
End Sub
Private Sub chkConsulta_Click()
    cargar_lista
End Sub

Private Sub chkMTL_Click()
    cargar_lista
End Sub

Private Sub cmdImprimir_Click()
    
    On Error GoTo fallo
    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim i As Integer
    Dim FECHA_REVISION As Date
    ' Generamos los datos del listado
    Dim rs As New ADODB.Recordset
    rs.Fields.Append "c1", adChar, 250, adFldUpdatable
    rs.Fields.Append "c2", adChar, 50, adFldUpdatable
    rs.Fields.Append "c3", adChar, 20, adFldUpdatable
    rs.Fields.Append "c4", adChar, 20, adFldUpdatable
    rs.Open
'    Dim existe As Boolean
'    existe = False
    For i = 1 To lista.ListItems.Count
'        If Trim(lista.ListItems(i).SubItems(6)) = "En vigor" Then
            rs.AddNew
'            existe = True
            rs("c1") = Trim(Left(lista.ListItems(i).SubItems(1), 250))
            rs("c2") = Left(lista.ListItems(i).SubItems(4), 50)
            rs("c3") = Left(lista.ListItems(i).SubItems(5), 20)
            rs("c4") = Left(lista.ListItems(i).SubItems(6), 20) ' Fecha
'            rs("c4") = Left(lista.ListItems(i).SubItems(9), 20) ' Fecha revision
            rs.Update
'            If IsDate(lista.ListItems(i).SubItems(5)) Then
'                If Format(FECHA_REVISION, "yyyy-mm-dd") < Format(lista.ListItems(i).SubItems(5), "yyyy-mm-dd") Then
'                    FECHA_REVISION = Format(lista.ListItems(i).SubItems(5), "yyyy-mm-dd")
'                End If
'            End If
            If IsDate(lista.ListItems(i).SubItems(9)) Then
                If Format(FECHA_REVISION, "yyyy-mm-dd") < Format(lista.ListItems(i).SubItems(9), "yyyy-mm-dd") Then
                    FECHA_REVISION = Format(lista.ListItems(i).SubItems(9), "yyyy-mm-dd")
                End If
            End If
'        End If
    Next

'    If Not existe Then
'        MsgBox "No existen en la lista documentos en vigor.", vbInformation, App.Title
'        Exit Sub
'    End If
    
    ' Generar Listado
    Dim Listado As New rptListadoModal
'    Listado.Orientation = rptOrientLandscape
    ' Cabecera
    With Listado.Sections("cabecera")
        .Controls("titulo").Caption = "Listado de Normas Controladas (LI-02)"
        .Controls("etiqueta4").Left = 170
        .Controls("etiqueta4").Width = 5800
        .Controls("etiqueta4").Caption = "Norma"
        .Controls("etiqueta5").Left = 6000
        .Controls("etiqueta5").Width = 1500
        .Controls("etiqueta5").Caption = "Código"
        .Controls("etiqueta10").Left = 7800
        .Controls("etiqueta10").Width = 1000
        .Controls("etiqueta10").Caption = "Edición"
        .Controls("etiqueta11").Left = 8400
        .Controls("etiqueta11").Width = 2500
        .Controls("etiqueta11").Caption = "Fecha"
    End With
    Set Listado.Sections("cabecera").Controls("logoc").Picture = LoadPicture(ReadINI(App.Path + "\config.ini", "logo", "logo"))
    'Detalle
    With Listado.Sections("detalle")
 '       .Controls("linea1").Visible = True
        .Controls("d1").Left = 170
        .Controls("d1").Width = 5800
        .Controls("d1").CanGrow = True
        .Controls("d1").Alignment = 0
        .Controls("d1").DataField = rs.Fields("c1").Name
        .Controls("d2").Left = 6000
        .Controls("d2").Width = 1500
        .Controls("d2").Alignment = 2
        .Controls("d2").DataField = rs.Fields("c2").Name
        .Controls("d3").Left = 7800
        .Controls("d3").Width = 1000
        .Controls("d3").Alignment = 2
        .Controls("d3").DataField = rs.Fields("c3").Name
        .Controls("d4").Left = 8400
        .Controls("d4").Width = 2500
        .Controls("d4").Alignment = 2
        .Controls("d4").DataField = rs.Fields("c4").Name
    End With
    ' Pie de Pagina
    With Listado.Sections("pie")
        .Controls("pie1").Caption = "Fecha : " & Format(Date, "dd-mm-yyyy")
        .Controls("pie2").Caption = "Fecha Ult.Revisión : " & Format(FECHA_REVISION, "dd-mm-yyyy")
        .Controls("pie3").Caption = "Firmado, Margarita Halcón"
        .Controls("pie3").visible = True
        .Controls("firma").visible = True
    End With
    Set Listado.DataSource = rs
    Listado.Caption = "Listado de Documentos de calidad."
'    Listado.WindowState = vbMaximized
    Listado.Show 1
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox "Error al generar el listado.", vbCritical, Err.Description
End Sub
Private Sub cmbestados_Change()
    cmdBuscar_Click
End Sub

Private Sub cmbSubtipo_Change()
    cmdBuscar_Click
End Sub
Private Sub cmbtipos_Change()
    cmdBuscar_Click
End Sub
Private Sub cmbsectores_Change()
    cmdBuscar_Click
End Sub

Private Sub cmdAnadir_Click()
    frmCA_Normas.PK = 0
    frmCA_Normas.Show 1
    cargar_lista
End Sub

Private Sub cmdBuscar_Click()
    cargar_lista
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdEliminar_Click()
    If lista.ListItems.Count > 0 Then
        If MsgBox("Va a eliminar la NORMA : " & lista.ListItems(lista.selectedItem.Index).SubItems(1), vbQuestion + vbYesNo, App.Title) = vbYes Then
            Dim oNorma As New clsCa_normas
            If oNorma.Eliminar(lista.ListItems(lista.selectedItem.Index).Text) = True Then
                cargar_lista
            End If
        End If
    End If
End Sub


Private Sub cmdLimpiar_Click()
    txtDatos(1) = ""
    txtDatos(0) = ""
    cmbtipos.Text = ""
    cmbSectores.Text = ""
    cmbestados.Text = ""
    cmbSubtipo.Text = ""
    cmdBuscar_Click
End Sub

Private Sub cmdListadoAuditoria_Click()
    frmCa_Filtro_Listado.TIPO_LLAMADA = 2
    frmCa_Filtro_Listado.Show 1
End Sub

Private Sub cmdModificar_Click()
    If lista.ListItems.Count > 0 Then
        If Not USUARIO.getPER_DOCUMENTACION_CALIDAD Then
            Exit Sub
        End If
        frmCA_Normas.PK = lista.ListItems(lista.selectedItem.Index).Text
        frmCA_Normas.Show 1
        actualizar_lista
    End If
End Sub

Private Sub cmdMostrar_Click(Index As Integer)
   On Error GoTo CMDMOSTRAR_Click_Error

    If lista.ListItems.Count = 0 Then
        Exit Sub
    End If
    Dim oNorma As New clsCa_normas
    oNorma.mostrar lista.ListItems(lista.selectedItem.Index).Text, False

   On Error GoTo 0
   Exit Sub

CMDMOSTRAR_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdMostrar_Click of Formulario frmCA_Listado_Normas"
End Sub

'E0038-I
Private Sub cmdVincular_Click()
'    Dim oEquipos_Normas As New clsEquipos_Normas

    If lista.ListItems.Count = 0 Then
        MsgBox "Debe seleccionar una norma de la lista.", vbInformation, App.Title
        Exit Sub
    Else
'        oEquipos_Normas.setNORMA_ID = lista.ListItems(lista.SelectedItem.Index).Text ' Este es el ID_NORMA
'        oEquipos_Normas.setEQUIPO_ID = FK_EQUIPO
'        Call oEquipos_Normas.Insertar    ' Se vincula la norma al equipo
        gID = lista.ListItems(lista.selectedItem.Index).Text
        Unload Me
    End If
End Sub
'E0038-F

Private Sub chkENAC_Click()
    cargar_lista
End Sub

Private Sub chkEQA_Click()
    cargar_lista
End Sub

Private Sub chkFTP_Click()
    cargar_lista
End Sub

Private Sub chkNADCAP_Click()
    cargar_lista
End Sub

Private Sub Form_Load()
    log (Me.Name)
    Me.top = 100
    Me.Left = 100
    cabecera
    cargar_botones Me
    cargar_combos
    cargar_lista
    permisos
    'E0042-I
    'El botón vincular sólo aparecerá si se ha abierto el formulario desde la gestión de equipos.
    If VINCULAR Then
        cmdVincular.visible = True
        cmdAnadir.visible = False
        cmdModificar.visible = False
        cmdEliminar.visible = False
        cmdMostrar(0).visible = False
        cmdImprimir.visible = False
        cmdListadoAuditoria.visible = False
    Else
        cmdVincular.visible = False
    End If
    'E0042-F
    'M1377-I
    If USUARIO.getPER_NORMAS_NO_CONTROLADAS = True Then
'JGM-I
         chkConsulta.visible = True
'        chkConsulta.Enabled = True
'        chkConsulta.value = vbChecked
'JGM-F
    Else
'JGM-I
         chkConsulta.visible = False
'        chkConsulta.value = vbUnchecked
'        chkConsulta.Enabled = False
'JGM-F
    End If
    'M1377-F
End Sub
Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "ID", 1, lvwColumnLeft
        .Add , , "Norma", 3700, lvwColumnLeft
        .Add , , "Tipo", 2000, lvwColumnLeft
        .Add , , "SubTipo", 2000, lvwColumnLeft
        .Add , , "Código", 2000, lvwColumnCenter
        .Add , , "Edición", 1000, lvwColumnCenter
        .Add , , "Fecha", 1000, lvwColumnCenter
        .Add , , "FTP", 800, lvwColumnCenter
        .Add , , "Mantenimiento", 1200, lvwColumnCenter
        .Add , , "FechaRevision", 0, lvwColumnCenter
    End With
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oca_normas As New clsCa_normas
    lista.ListItems.Clear
    Dim tipo As String
    Dim SECTOR As String
    Dim ESTADO As String
    Dim nombre As String
    Dim CODIGO As String
    Dim subtipo As String
    If cmbtipos.Text = "" Then
        tipo = 0
    Else
        tipo = cmbtipos.BoundText
    End If
    If cmbSectores.Text = "" Then
        SECTOR = 0
    Else
        SECTOR = cmbSectores.BoundText
    End If
    If cmbestados.Text = "" Then
        ESTADO = 0
    Else
        ESTADO = cmbestados.BoundText
    End If
    If cmbSubtipo.Text = "" Then
        subtipo = 0
    Else
        subtipo = cmbSubtipo.BoundText
    End If
    nombre = txtDatos(1)
    CODIGO = txtDatos(0)
    Set rs = oca_normas.Listado(tipo, SECTOR, ESTADO, nombre, CODIGO, subtipo, chkENAC.Value, chkNADCAP.Value, chkMTL.Value, chkEQA.Value, chkFTP.Value, chkConsulta.Value)
    lblsubtitulo = "Normas listadas : " & rs.RecordCount
    If rs.RecordCount <> 0 Then
        Do
            With lista.ListItems.Add(, , Format(rs(0), "0000"))
             .SubItems(1) = rs(1)
             .SubItems(2) = rs(2)
             .SubItems(3) = rs(3)
             .SubItems(4) = rs(4)
             .SubItems(5) = rs(5)
             If IsDate(rs(6)) Then
                 .SubItems(6) = Format(rs(6), "dd-mm-yyyy")
             Else
                .SubItems(6) = rs(6)
             End If
             If rs(8) = 1 Then
                .SubItems(7) = "Si"
             Else
                .SubItems(7) = "No"
             End If
             .SubItems(8) = rs(7)
             .SubItems(9) = Format(rs(9), "dd-mm-yyyy")
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oca_documentos = Nothing
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

Public Sub actualizar_lista()
    Dim rs As ADODB.Recordset
    Dim oNorma As New clsCa_normas
    Set rs = oNorma.Listado_por_Codigo(lista.ListItems(lista.selectedItem.Index))
    If rs.RecordCount <> 0 Then
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = rs(3)
        lista.ListItems(lista.selectedItem.Index).SubItems(4) = rs(4)
        lista.ListItems(lista.selectedItem.Index).SubItems(5) = rs(5)
        lista.ListItems(lista.selectedItem.Index).SubItems(6) = Format(rs(6), "dd-mm-yyyy")
        lista.ListItems(lista.selectedItem.Index).SubItems(8) = rs(7)
        If rs(8) = 1 Then
            lista.ListItems(lista.selectedItem.Index).SubItems(7) = "Si"
        Else
            lista.ListItems(lista.selectedItem.Index).SubItems(7) = "No"
        End If
        lista.ListItems(lista.selectedItem.Index).SubItems(9) = Format(rs(9), "dd-mm-yyyy")
    End If
    Set oNorma = Nothing
End Sub

Public Sub cargar_combos()
    Dim oDecodificadora As New clsDecodificadora
    oDecodificadora.cargar_combo cmbtipos, DECODIFICADORA.CA_NORMAS_TIPOS
    oDecodificadora.cargar_combo cmbSectores, DECODIFICADORA.CA_NORMAS_SECTORES
    oDecodificadora.cargar_combo cmbestados, DECODIFICADORA.CA_NORMAS_ESTADOS
    oDecodificadora.cargar_combo cmbSubtipo, DECODIFICADORA.CA_NORMAS_SUBTIPOS
End Sub
Public Sub permisos()
    If Not USUARIO.getPER_DOCUMENTACION_CALIDAD Then
'        cmdanadir.Enabled = False
        cmdModificar.Enabled = False
'        cmdEliminar.Enabled = False
        cmdImprimir.Enabled = False
        cmdListadoAuditoria.Enabled = False
    End If
    If Not USUARIO.getPER_ADMIN_PNT Then
        cmdAnadir.Enabled = False
        cmdEliminar.Enabled = False
    End If
End Sub
Private Sub txtDatos_Change(Index As Integer)
    cmdBuscar_Click
End Sub

'E0041-I
Private Sub Form_Unload(Cancel As Integer)
    'Se borra el dato que pudiera haber en FK_EQUIPO
    FK_EQUIPO = 0
End Sub
'E0041-F
