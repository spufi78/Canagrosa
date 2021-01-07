VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmIndicadores_Gestion 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generación de Indicadores"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11250
   Icon            =   "frmIndicadores_Gestion.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6015
   ScaleWidth      =   11250
   Begin VB.Frame frmAnual 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Anual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8100
      TabIndex        =   11
      Top             =   1755
      Visible         =   0   'False
      Width           =   3120
      Begin VB.TextBox txtanno2 
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
         Left            =   1350
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   1035
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   360
         Left            =   2385
         TabIndex        =   13
         Top             =   225
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno2"
         BuddyDispid     =   196610
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
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   810
         TabIndex        =   14
         Top             =   270
         Width           =   585
      End
   End
   Begin VB.Frame frmMensual 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Mensual"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1275
      Left            =   8100
      TabIndex        =   5
      Top             =   405
      Visible         =   0   'False
      Width           =   3120
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
         Left            =   855
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   720
         Width           =   990
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
         ItemData        =   "frmIndicadores_Gestion.frx":0442
         Left            =   855
         List            =   "frmIndicadores_Gestion.frx":046A
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   270
         Width           =   2085
      End
      Begin MSComCtl2.UpDown cambiar 
         Height          =   360
         Left            =   1845
         TabIndex        =   8
         Top             =   720
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   635
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196613
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
         Caption         =   "Mes"
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
         Left            =   270
         TabIndex        =   10
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   12
         Left            =   270
         TabIndex        =   9
         Top             =   780
         Width           =   585
      End
   End
   Begin VB.CommandButton cmdGenerar 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Generar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   8055
      Picture         =   "frmIndicadores_Gestion.frx":04D3
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5085
      Width           =   1080
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   10170
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5085
      Width           =   1050
   End
   Begin MSComctlLib.ListView lista 
      Height          =   5565
      Left            =   60
      TabIndex        =   0
      Top             =   390
      Width           =   7950
      _ExtentX        =   14023
      _ExtentY        =   9816
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
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Deter."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   6
      Left            =   4050
      TabIndex        =   4
      Top             =   3480
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Generación de Indicadores"
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
      Left            =   60
      TabIndex        =   1
      Top             =   30
      Width           =   11145
   End
End
Attribute VB_Name = "frmIndicadores_Gestion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdGenerar_Click()
    On Error GoTo fallo
    Dim oIndicadores As New clsIndicadores
    Dim rs As ADODB.Recordset
    Dim rs_resultado As ADODB.Recordset
    Set rs = oIndicadores.generar(lista.ListItems(lista.selectedItem.Index).SubItems(3))
    If rs.RecordCount = 0 Then
        MsgBox "No existen campos para la seleccion.", vbInformation, App.Title
        Exit Sub
    Else
        Me.MousePointer = 11
        On Error Resume Next
        Kill ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & lista.ListItems(lista.selectedItem.Index).Text & " " & Format(Date, "dd-mm-yyyy") & ".xls"
        On Error GoTo fallo
        Dim XLA As excel.Application
        Dim XLW As excel.Workbook
        Dim XLS As excel.Worksheet
        Set XLA = New excel.Application
        Set XLW = XLA.Workbooks.Open(ReadINI(App.Path + "\config.ini", "documentos", "plantillas") & "\" & lista.ListItems(lista.selectedItem.Index).SubItems(2))
        Set XLS = XLW.Worksheets(1)
        ' Datos
        Dim letra As String
        Dim NUMERO As String
        Dim fila As Integer
        Dim Col As String
        Dim RESULTADO As Long
        Do
            i = 1
            letra = ""
            NUMERO = ""
            While i <= Len(Trim(rs(3)))
                If IsNumeric(Mid(rs(3), i, 1)) Then
                    NUMERO = NUMERO & Mid(rs(3), i, 1)
                Else
                    If Trim(Mid(rs(3), i, 1)) <> "" Then
                        letra = letra & Mid(rs(3), i, 1)
                    End If
                End If
                i = i + 1
            Wend
            fila = CInt(NUMERO)
            Col = letra
            ' Resultado
            Select Case rs(1) ' Funcion
                Case 1 ' T.M.
                    consulta = "select month(m.fecha_recepcion), count(*) " & _
                               "  from muestras m, tipos_muestra t " & _
                               " where t.nombre = '" & rs(2) & "' " & _
                               "   and m.tipo_muestra_id = t.id_tipo_muestra " & _
                               "   and anno = " & CInt(txtanno2) & _
                               "   and m.anulada = 0 " & _
                               " group by month(m.fecha_recepcion)"
                Case 2 ' Cliente
                    consulta = "select month(m.fecha_recepcion),count(*) " & _
                               "  from muestras m, clientes c " & _
                               " where c.nombre = '" & rs(2) & "' " & _
                               "   and m.cliente_id = c.id_cliente " & _
                               "   and anno = " & CInt(txtanno2) & _
                               "   and m.anulada = 0 " & _
                               " group by month(m.fecha_recepcion)"
                Case 3 ' Familias por número
                    consulta = "select month(m.fecha_recepcion), count(*) " & _
                               "  from muestras m, tipos_muestra t, familias f " & _
                               " where f.nombre = '" & rs(2) & "' " & _
                               "   and m.tipo_muestra_id = t.id_tipo_muestra " & _
                               "   and t.familia_id = f.id_familia " & _
                               "   and anno = " & CInt(txtanno2) & _
                               "   and m.anulada = 0 " & _
                               " group by month(m.fecha_recepcion)"
                Case 4 ' Familias economico
                    consulta = "select month(m.fecha_recepcion), sum(precio) " & _
                               "  from muestras m, tipos_muestra t, familias f " & _
                               " where f.nombre = '" & rs(2) & "' " & _
                               "   and m.tipo_muestra_id = t.id_tipo_muestra " & _
                               "   and t.familia_id = f.id_familia " & _
                               "   and anno = " & CInt(txtanno2) & _
                               "   and m.anulada = 0 " & _
                               " group by month(m.fecha_recepcion)"
                Case 5 ' T.M. Economico
                    consulta = "select month(m.fecha_recepcion), sum(precio) " & _
                               "  from muestras m, tipos_muestra t " & _
                               " where t.nombre = '" & rs(2) & "' " & _
                               "   and m.tipo_muestra_id = t.id_tipo_muestra " & _
                               "   and anno = " & CInt(txtanno2) & _
                               "   and m.anulada = 0 " & _
                               " group by month(m.fecha_recepcion)"
                
                Case 6 ' Clientes Economico
                    consulta = "select month(m.fecha_recepcion),sum(precio) " & _
                               "  from muestras m, clientes c " & _
                               " where c.nombre = '" & rs(2) & "' " & _
                               "   and m.cliente_id = c.id_cliente " & _
                               "   and anno = " & CInt(txtanno2) & _
                               "   and m.anulada = 0 " & _
                               " group by month(m.fecha_recepcion)"
                           
            End Select
            Set rs_resultado = datos_bd(consulta)
            For i = 0 To 11
                XLS.Cells(fila + i, Col) = 0
            Next
            fila = fila - 1
            If rs_resultado.RecordCount <> 0 Then
                Do
                    XLS.Cells(fila + rs_resultado(0), Col) = rs_resultado(1) ' Resultado
                    rs_resultado.MoveNext
                Loop Until rs_resultado.EOF
            End If
            rs.MoveNext
        Loop Until rs.EOF
        XLS.SaveAs ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\" & lista.ListItems(lista.selectedItem.Index).Text & " " & Format(Date, "dd-mm-yyyy")
        XLA.visible = True
    End If
    Set rs = Nothing
    Set XLW = Nothing
    Set XLA = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se han producido errores al generar el indicador: " & Err.Description, vbCritical, "FILA:" & fila & " COL:" & Col
    XLA.visible = True
    Set XLW = Nothing
    Set XLA = Nothing
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 200
    Me.top = 200
    txtanno = Year(Date)
    txtanno2 = Year(Date)
    cabecera
    cargar_lista
End Sub

Public Sub cargar_lista()
    Dim rs As New ADODB.Recordset
    Dim oIndicadores As New clsIndicadores
    Set rs = oIndicadores.lista
    lista.ListItems.Clear
    If rs.RecordCount <> 0 Then
        Do
           With lista.ListItems.Add(, , rs(0))
                .SubItems(1) = rs(1)
                .SubItems(2) = rs(2)
                .SubItems(3) = Format(rs(3), "000")
                .SubItems(4) = rs(4)
           End With
           rs.MoveNext
        Loop Until rs.EOF
    End If
    Set oIndicadores = Nothing
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
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "Nombre", 2500, lvwColumnLeft)
        .Tag = "Nombre"
    End With
    With lista.ColumnHeaders.Add(, , "Frecuencia", 1200, lvwColumnLeft)
        .Tag = "Frecuencia"
    End With
    With lista.ColumnHeaders.Add(, , "Hoja Excel", 2800, lvwColumnLeft)
        .Tag = "Hoja Excel"
    End With
    With lista.ColumnHeaders.Add(, , "ID", 550, lvwColumnCenter)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "ID_FRECUENCIA", 1, lvwColumnCenter)
        .Tag = "ID_FRECUENCIA"
    End With
End Sub
Private Sub lista_Click()
    If lista.ListItems(lista.selectedItem.Index) <> "" Then
        frmMensual.visible = False
        frmAnual.visible = False
        Select Case lista.ListItems(lista.selectedItem.Index).SubItems(4)
        Case 1 ' Mensual
            frmMensual.visible = True
        Case 2 ' Anual
            frmAnual.visible = True
            frmAnual.top = 495
        End Select
    End If
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        gindicadores = lista.ListItems(lista.selectedItem.Index).SubItems(3)
        frmIndicadores.Show 1
        modificar_lista
        gindicadores = 0
    End If
End Sub
Public Sub modificar_lista()
    Dim rs As New ADODB.Recordset
    Dim oIndicadores As New clsIndicadores
    Set rs = oIndicadores.Listado_por_Codigo(gindicadores)
    If rs.RecordCount <> 0 Then
        lista.ListItems(lista.selectedItem.Index).Text = rs(0)
        lista.ListItems(lista.selectedItem.Index).SubItems(1) = rs(1)
        lista.ListItems(lista.selectedItem.Index).SubItems(2) = rs(2)
        lista.ListItems(lista.selectedItem.Index).SubItems(3) = Format(rs(3), "000")
    End If
    Set oIndicadores = Nothing
    If lista.ListItems.Count > 0 Then
        lista_Click
    End If
End Sub

