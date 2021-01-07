VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmEads_Historico 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control histórico de Baños"
   ClientHeight    =   9105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13635
   Icon            =   "frmEads_Historico.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   9105
   ScaleWidth      =   13635
   Begin VB.CommandButton cmdComprobar 
      Caption         =   "Comprobar"
      Height          =   420
      Left            =   8955
      TabIndex        =   32
      Top             =   8640
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Todos Graficos"
      Height          =   285
      Left            =   7650
      TabIndex        =   31
      Top             =   7200
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Grafico"
      Height          =   285
      Left            =   7650
      TabIndex        =   29
      Top             =   7515
      Visible         =   0   'False
      Width           =   1800
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Linea"
      Height          =   375
      Left            =   8775
      TabIndex        =   28
      Top             =   8055
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton cmdTodos 
      Caption         =   "Todos"
      Height          =   375
      Left            =   8100
      TabIndex        =   25
      Top             =   8055
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton cmdOR 
      Caption         =   "OR"
      Height          =   375
      Left            =   7380
      TabIndex        =   24
      Top             =   8055
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.CommandButton cmdVerExcel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Excel"
      Height          =   1005
      Left            =   5535
      Style           =   1  'Graphical
      TabIndex        =   23
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdAnalisis 
      Caption         =   "Analisis"
      Height          =   375
      Left            =   6615
      TabIndex        =   22
      Top             =   8055
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.CommandButton cmdRecarga 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Recargas"
      Height          =   1005
      Left            =   4455
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Enviar informe por E-mail"
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdUSB 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Copiar a USB"
      Height          =   1005
      Left            =   11025
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   8055
      Width           =   1275
   End
   Begin VB.CommandButton cmdGrafico 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Curva"
      Height          =   1005
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdInforme 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe"
      Height          =   1005
      Left            =   2265
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Previsualizar informe de ensayo"
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdCONSOLIDAR 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Consolidar Excel"
      Height          =   1005
      Left            =   9720
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8055
      Width           =   1275
   End
   Begin VB.CommandButton cmdDeter 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Registro"
      Height          =   1005
      Left            =   1185
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   8055
      Width           =   1050
   End
   Begin VB.CommandButton cmdVerMuestra 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ver Muestra"
      Height          =   1005
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   8055
      Width           =   1050
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Criterios de selección"
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
      Height          =   1125
      Left            =   45
      TabIndex        =   14
      Top             =   675
      Width           =   13560
      Begin VB.TextBox txtensayo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   10200
         TabIndex        =   0
         Top             =   270
         Width           =   1215
      End
      Begin VB.CommandButton cmdBuscar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Buscar"
         Height          =   870
         Left            =   11610
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   915
      End
      Begin VB.CommandButton cmdLimpiar 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Limpiar"
         Height          =   870
         Left            =   12555
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   915
      End
      Begin VB.TextBox txtanno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   10200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   660
         Width           =   975
      End
      Begin MSDataListLib.DataCombo cmbBanos 
         Height          =   315
         Left            =   855
         TabIndex        =   2
         Top             =   660
         Width           =   8190
         _ExtentX        =   14446
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
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
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   315
         Left            =   11160
         TabIndex        =   3
         Top             =   660
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   556
         _Version        =   393216
         Value           =   2004
         BuddyControl    =   "txtanno"
         BuddyDispid     =   196628
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
      Begin MSDataListLib.DataCombo cmbLinea 
         Height          =   315
         Left            =   855
         TabIndex        =   1
         Top             =   270
         Width           =   8190
         _ExtentX        =   14446
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Style           =   2
         ListField       =   ""
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
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ensayo"
         Height          =   195
         Index           =   4
         Left            =   9330
         TabIndex        =   30
         Top             =   330
         Width           =   525
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Línea"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Top             =   330
         Width           =   420
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Año Inicial"
         Height          =   225
         Index           =   2
         Left            =   9330
         TabIndex        =   18
         Top             =   690
         Width           =   855
      End
      Begin VB.Label lblCampos 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Baño"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   720
         Width           =   375
      End
   End
   Begin MSComctlLib.ListView lista 
      Height          =   6195
      Left            =   45
      TabIndex        =   16
      Top             =   1815
      Width           =   13560
      _ExtentX        =   23918
      _ExtentY        =   10927
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ColHdrIcons     =   "ImageList1"
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
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Salir"
      Height          =   1005
      Left            =   12315
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   8055
      Width           =   1275
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6615
      Top             =   8505
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEads_Historico.frx":1272
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmEads_Historico.frx":1B4C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label lblbano 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   6885
      TabIndex        =   27
      Top             =   8775
      Width           =   2490
   End
   Begin VB.Label lbllinea 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Height          =   240
      Left            =   6885
      TabIndex        =   26
      Top             =   8505
      Width           =   2490
   End
   Begin VB.Label lblsubtitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Control de resultados y generación en excel."
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   90
      TabIndex        =   20
      Top             =   330
      Width           =   3135
   End
   Begin VB.Label lbltitulo 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Control de histórico de Baños"
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
      TabIndex        =   19
      Top             =   30
      Width           =   3075
   End
   Begin VB.Shape fondo 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      Height          =   675
      Left            =   0
      Top             =   0
      Width           =   13635
   End
End
Attribute VB_Name = "frmEads_Historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const fila_inicial As Integer = 13
Private Enum COLS
        fecha = 2
        D1 = 3
        D2 = 4
        D3 = 5
        D4 = 6
        D5 = 7
        D6 = 8
        D7 = 9
        INFORME_1 = 10
        INFORME_2 = 11
        ACCION = 12
        RESPUESTA = 13
        NIVEL = 14
        TEMPERATURA = 15
        otros = 16
        ERROR = 17
End Enum

Private blnConsolidarEnServidor As Boolean
Private lngIdBano_Consolidacion As Long

Private Sub comprobar_consolidar_en_servidor()


    If Not blnConsolidarEnServidor Then Exit Sub

    Dim oimp As New clsImpresion
    With oimp
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setMUESTRA_ID = lngIdBano_Consolidacion
        .setPUESTO = USUARIO.getUSO
        .setTIPO = 60
        .Insertar
    End With
    
    Set oimp = Nothing
End Sub

Private Sub cmblinea_Change()
    cargar_lista_banos
End Sub

Private Sub cmdAnalisis_Click()
    If cmbLinea.Text = "" Then
        MsgBox "Seleccione una línea.", vbCritical, App.Title
    Else
        If cmbbanos.Text <> "" Then
            genera_excel_analisis cmbbanos.BoundText
        Else
            Dim oBANO As New clsBanos
            Dim rs As ADODB.Recordset
            Set rs = oBANO.Listado_Lineas_Controladas(cmbLinea.BoundText)
            If rs.RecordCount > 0 Then
                Do
                    genera_excel_analisis rs(0)
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            Set oBANO = Nothing
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
    comprobar_consolidar_en_servidor
    
    cargar_bano
    
    blnConsolidarEnServidor = False
End Sub

Private Sub cmdcancel_Click()
    ' por si la última vez, al salir, no se le da a
    comprobar_consolidar_en_servidor
    
    Unload Me
End Sub

Private Sub cmdComprobar_Click()
    Dim hoja As String
    Dim c As String
    c = "SELECT * FROM BANOS_CONTROL"
    Dim rs As ADODB.Recordset
    Dim olinea As New clsLineas
    Dim oBANO As New clsBanos
    Set rs = datos_bd(c)
    If rs.RecordCount > 0 Then
        Do
            hoja = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\"
            oBANO.cargar_bano rs("bano_id")
            olinea.CARGAR oBANO.getID_LINEA
            hoja = hoja & olinea.getNOMBRE & "\"
            hoja = hoja & Replace(oBANO.getNOMBRE, "/", "") & ".xls"
            If Dir(hoja) = "" Then
                MsgBox "La hoja excel no existe : " & olinea.getNOMBRE & "/" & oBANO.getNOMBRE, vbCritical, App.Title
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    MsgBox "FINALIZADO"
End Sub

Private Sub cmdCONSOLIDAR_Click()
    If cmbLinea.Text = "" Then
        MsgBox "Seleccione una línea.", vbCritical, App.Title
    Else
        If cmbbanos.Text <> "" Then
            genera_excel cmbbanos.BoundText, 1
'            genera_excel_bano cmbBanos.BoundText
'            genera_excel_analisis cmbBanos.BoundText
'            genera_excel_OR cmbBanos.BoundText
'            Command2_Click
            cmdVerExcel_Click
        Else
            If MsgBox("Atencion, va a consolidar todos los baños de la línea. ¿Esta seguro?", vbYesNo + vbQuestion, App.Title) = vbYes Then
                Dim oBANO As New clsBanos
                Dim rs As ADODB.Recordset
                Set rs = oBANO.Listado_Lineas_Controladas(cmbLinea.BoundText)
                Dim i As Integer
                i = 1
                If rs.RecordCount > 0 Then
                    Do
                        lbllinea = i & " de " & rs.RecordCount
                        DoEvents
                        genera_excel rs(0), 1
'                        genera_excel_bano rs(0)
'                        genera_excel_analisis rs(0)
'                        genera_excel_OR rs(0)
                        rs.MoveNext
                        i = i + 1
                    Loop Until rs.EOF
                End If
                Set oBANO = Nothing
                MsgBox "Linea consolidada correctamente.", vbInformation, App.Title
            End If
        End If
    End If

End Sub

Private Sub cmdDeter_Click()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).Text
        abrirRegistroMuestra gmuestra
'        frmDeterminaciones.Show 1
        gmuestra = 0
    End If
End Sub


Private Sub cmdGrafico_Click()
    If lista.ListItems.Count > 0 Then
        Dim i As Integer
        Dim j As Integer
        With frmEads_Grafico.lista
            For i = 1 To lista.ColumnHeaders.Count
                .ColumnHeaders.Add , , lista.ColumnHeaders(i).Text, lista.ColumnHeaders(i).Width, lista.ColumnHeaders(i).Alignment
            Next
            For i = 1 To lista.ListItems.Count
                .ListItems.Add , , lista.ListItems(i)
                For j = 1 To .ColumnHeaders.Count - 1
                    .ListItems(i).SubItems(j) = lista.ListItems(i).SubItems(j)
                Next
            Next
        End With
        frmEads_Grafico.lista.Refresh
        frmEads_Grafico.Show 1
    End If

End Sub

Private Sub cmdInforme_Click()
    If lista.ListItems.Count > 0 Then
        Dim oMuestra As New clsMuestra
'C001-I
'        Dim oTD As New clsTipos_documentos
'        If oTD.Nuevo_Formato(omuestra.obtener_tipo_documento(CLng(lista.ListItems(lista.SelectedItem.Index).Text))) Then
'            frmPrevisualizar2.PK_MUESTRA = CLng(lista.ListItems(lista.SelectedItem.Index).Text)
'            frmPrevisualizar2.Show 1
'        Else
'            gmuestra = CLng(lista.ListItems(lista.SelectedItem.Index).Text)
'            frmPrevisualizar.Show 1
            MostrarInforme CLng(lista.ListItems(lista.selectedItem.Index).Text)
'        End If
'C001-F
    End If
End Sub

Private Sub cmdLimpiar_Click()
    cmbLinea.Text = ""
    cmbbanos.Text = ""
    lista.ListItems.Clear
End Sub

Private Sub cmdOR_Click()
    If cmbLinea.Text = "" Then
        MsgBox "Seleccione una línea.", vbCritical, App.Title
    Else
        If cmbbanos.Text <> "" Then
            genera_excel_OR cmbbanos.BoundText
        Else
            Dim oBANO As New clsBanos
            Dim rs As ADODB.Recordset
            Set rs = oBANO.Listado_Lineas_Controladas(cmbLinea.BoundText)
            If rs.RecordCount > 0 Then
                Do
                    genera_excel_OR rs(0)
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            Set oBANO = Nothing
        End If
    End If

End Sub

Private Sub cmdRecarga_Click()
    If lista.ListItems.Count > 0 Then
        gmuestra = lista.ListItems(lista.selectedItem.Index).Text
'        frmEads_Recarga.BANO_ID = getDataComboSel(cmbbanos)
        frmEads_Recarga.Show 1
'        blnConsolidarEnServidor = frmEads_Recarga.MODIFICACIONES_REALIZADAS
'        Unload frmEads_Recarga
        gmuestra = 0
    End If
End Sub

Private Sub cmdtodos_Click()
    Dim olinea As New clsLineas
    Dim rs_lineas As ADODB.Recordset
    Dim oBANO As New clsBanos
    Dim rs As ADODB.Recordset
    Set rs_lineas = olinea.Listado_Historico
    If rs_lineas.RecordCount > 0 Then
        Do
            lbllinea.Caption = rs_lineas("NOMBRE")
            Set rs = oBANO.Listado_Lineas_Controladas(rs_lineas("ID_LINEA"))
            If rs.RecordCount > 0 Then
                Do
                    lblbano.Caption = rs(1)
                    DoEvents
'                    azul rs(0)
'                    genera_excel_bano rs(0)
'                    genera_excel_analisis rs(0)
                    genera_excel_OR rs(0)
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            rs_lineas.MoveNext
        Loop Until rs_lineas.EOF
    End If
    Set oBANO = Nothing
End Sub

Private Sub cmdUSB_Click()
    Dim origen As String
    Dim destino As String
    Dim fso As FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\*.*"
    destino = InputBox("Introduzca la ruta de destino...", App.Title)
    If destino = "" Then
        MsgBox "Debe introducir la ruta de destino.", vbInformation, App.Title
        Exit Sub
    End If
    On Error Resume Next
    Me.MousePointer = 11
    MkDir destino
    DoEvents
'    destino = "c:\copia"
'    fso.CreateFolder (destino)
    On Error GoTo fallo
    fso.CopyFolder origen, destino
    Me.MousePointer = 0
    MsgBox "Copia terminada correctamente.", vbInformation, App.Title
    Exit Sub
fallo:
    Me.MousePointer = 0
    error_grave ("Error al copiar los ficheros. " & Err.Description)

End Sub

Private Sub cmdVerExcel_Click()
   On Error GoTo cmdVerExcel_Click_Error
    If cmbbanos.Text = "" Then
        MsgBox "Seleccione un baño.", vbExclamation, App.Title
        Exit Sub
    End If
    Dim hoja As String
    hoja = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\"
    
    Dim linea As String
'    Select Case cmbLinea.BoundText
'        Case 10
'                        linea = "IFQ-01 INST. FRESADO QUIMICO DE ALUMINIO"
'                    Case 25
'                        linea = "ALUMINIO L1. CBC"
'                    Case 27
'                        linea = "ALUMINIO L2. CBC"
'                    Case 28
'                        linea = "CSP TITANIO L3. CBC"
'                    Case Else
                        linea = cmbLinea.Text
'                End Select
    
    
    
    hoja = hoja & linea & "\"
    hoja = hoja & Replace(cmbbanos.Text, "/", "") & ".xls"
    If Dir(hoja) <> "" Then
        r = Shell("rundll32.exe url.dll,FileProtocolHandler " & hoja, vbMaximizedFocus)
    Else
        MsgBox "La hoja excel no existe.", vbCritical, App.Title
    End If

   On Error GoTo 0
   Exit Sub

cmdVerExcel_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdVerExcel_Click of Formulario frmEads_Historico"
End Sub

Private Sub cmdVerMuestra_Click()
    lista_DblClick
End Sub


Private Sub Command1_Click()
    Dim olinea As New clsLineas
    Dim hoja As String
    Dim rs As ADODB.Recordset
    Set rs = olinea.Listado_Historico
    If rs.RecordCount > 0 Then
        Do
            On Error Resume Next
            MkDir ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\" & rs(1)
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set olinea = Nothing
    
End Sub

Private Sub Command2_Click()
    genera_grafico cmbbanos.BoundText
'    azul cmbBanos.BoundText
End Sub

Private Sub Command3_Click()
    Dim olinea As New clsLineas
    Dim rs_lineas As ADODB.Recordset
    Dim oBANO As New clsBanos
    Dim rs As ADODB.Recordset
    Set rs_lineas = olinea.Listado_Historico
    If rs_lineas.RecordCount > 0 Then
        Do
            lbllinea.Caption = rs_lineas("NOMBRE")
            Set rs = oBANO.Listado_Lineas_Controladas(rs_lineas("ID_LINEA"))
            If rs.RecordCount > 0 Then
                Do
                    lblbano.Caption = rs(1)
                    DoEvents
                    genera_grafico rs(0)
                    rs.MoveNext
                Loop Until rs.EOF
            End If
            rs_lineas.MoveNext
        Loop Until rs_lineas.EOF
    End If
    Set oBANO = Nothing
End Sub

Private Sub Form_Load()
    log (Me.Name)
    cargar_botones Me
    Me.Left = 30
    Me.top = 30
    blnConsolidarEnServidor = False
    
'    txtanno = "2007"
    txtanno = Year(Date)
    cabecera
    cargar_lineas
    If UCase(USUARIO.getNOMBRE) = "JULIO" Then
        cmdAnalisis.visible = True
        cmdOR.visible = True
        cmdtodos.visible = True
        Command1.visible = True
        Command2.visible = True
        Command3.visible = True
        cmdComprobar.visible = True
        cmdtodos.visible = True
    End If
End Sub
Public Sub cabecera()
    With lista.ColumnHeaders.Add(, , "ID", 300, lvwColumnLeft)
        .Tag = "ID"
    End With
    With lista.ColumnHeaders.Add(, , "Código", 1700, lvwColumnCenter)
        .Tag = "Código"
    End With
    With lista.ColumnHeaders.Add(, , "Fecha", 1200, lvwColumnCenter)
        .Tag = "Fecha"
    End With
    With lista.ColumnHeaders.Add(, , "Nivel", 1000, lvwColumnCenter)
        .Tag = "Nivel"
    End With
    With lista.ColumnHeaders.Add(, , "Temp", 800, lvwColumnCenter)
        .Tag = "Temp"
    End With
    With lista.ColumnHeaders.Add(, , "Hora", 800, lvwColumnCenter)
        .Tag = "Hora"
    End With
    With lista.ColumnHeaders.Add(, , "Observaciones", 1500, lvwColumnCenter)
        .Tag = "Observaciones"
    End With
    With lista.ColumnHeaders.Add(, , "Parametros", 5500, lvwColumnLeft)
        .Tag = "Parametros"
    End With
End Sub

Public Sub cargar_bano()
    lista.ListItems.Clear
    If cmbbanos.BoundText = "" Then
        Exit Sub
    End If
    
    lngIdBano_Consolidacion = getDataComboSel(cmbbanos)
    
    ' Cargamos los parametros del baño
'    On Error GoTo fallo
    Me.MousePointer = 11
    Dim oDA As New clsDeterminaciones_analisis
    Dim rs As ADODB.Recordset
    Set rs = oDA.Listado_por_bano_historico(cmbbanos.BoundText)
    Dim i As Integer
    For i = lista.ColumnHeaders.Count To 8 Step -1
        lista.ColumnHeaders.Remove i
    Next
    i = 7
    Dim pos(7 To 30) As Integer
    If rs.RecordCount <> 0 Then
        Do
            With lista.ColumnHeaders.Add(, , rs("NOMBRE"), 900, lvwColumnRight)
                 .Tag = rs("NOMBRE")
            End With
            pos(i) = rs("ID_TIPO_DETERMINACION")
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
        ' Resultados
        Dim oBANO As New clsBanos
        On Error Resume Next
        Set rs = oBANO.Resultados_Banos(cmbbanos.BoundText, txtanno)
        Dim MUESTRA As String
        If rs.RecordCount <> 0 Then
            MUESTRA = ""
            Do
                With lista.ListItems.Add(, , rs(0))
                    .SubItems(1) = "AN-" & Format(rs(5), "0000") & "-" & rs(6) & "-Ed." & rs(7)
                    .SubItems(2) = rs(2)
                    .SubItems(4) = rs(9) ' Temperatura
                    .SubItems(5) = rs(10) ' Hora
                    .SubItems(6) = rs(8) ' Observaciones
                    .SubItems(19) = rs(11) ' Nivel
                    If IsNull(rs(12)) Then
                        lista.ListItems(lista.ListItems.Count).SmallIcon = 1
                    Else
                        If Trim(rs(12)) = "" Then
                            lista.ListItems(lista.ListItems.Count).SmallIcon = 1
                        End If
                    End If
                    
                    MUESTRA = rs(0)
                    Do
                        For j = 7 To 30
                           If rs(4) = pos(j) Then ' ID_DETER
                                If Not rs(3) <> "" And Not IsNull(rs(3)) Then ' resultado
                                   .SubItems(j) = "No informado"
                                Else
                                   .SubItems(j) = Replace(rs(3), ".", ",")
                                End If
                                Exit For
                           End If
                        Next
                       If rs.EOF = False Then
                           rs.MoveNext
                       End If
                    Loop Until MUESTRA <> rs(0)
                End With
                DoEvents
            Loop Until rs.EOF
        End If
    End If
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al recuperar los resultados de los baños.", vbCritical, App.Title
End Sub

'Public Sub cargar_bano_old()
'    lista.ListItems.Clear
'    If cmbbanos.BoundText = "" Then
'        Exit Sub
'    End If
'    ' Cargamos los parametros del baño
''    On Error GoTo fallo
'    Me.MousePointer = 11
'    Dim oDA As New clsDeterminaciones_analisis
'    Dim rs As ADODB.RecordSet
'    Dim rs_valores As ADODB.RecordSet
'    Set rs = oDA.Listado_por_bano(cmbbanos.BoundText)
'    Dim i As Integer
'    For i = lista.ColumnHeaders.Count To 8 Step -1
'        lista.ColumnHeaders.Remove i
'    Next
'    i = 7
'    Dim Pos(7 To 30) As Integer
'    If rs.RecordCount <> 0 Then
'        Do
'            With lista.ColumnHeaders.Add(, , rs(1), 1200, lvwColumnRight)
'                 .Tag = rs(1)
'            End With
'            Pos(i) = rs(7)
'            i = i + 1
'            rs.MoveNext
'        Loop Until rs.EOF
'        ' Resultados
'        Dim oBANO As New clsBanos
'        Dim oDatos_valores As New clsDatos_valores
'        On Error Resume Next
'        Set rs = oBANO.Resultados_Banos(cmbbanos.BoundText, txtanno)
'        Dim MUESTRA As String
'        If rs.RecordCount <> 0 Then
'            MUESTRA = ""
'            Do
'                With lista.ListItems.Add(, , rs(0))
'                    .SubItems(1) = "AN-" & Format(rs(5), "0000") & "-" & rs(6) & "-Ed." & rs(7)
'                    .SubItems(2) = rs(2)
'                    Set rs_valores = oDatos_valores.datos_muestra_historico(rs(0), "1,2,18,19")
'                    If rs_valores.RecordCount > 0 Then
'                        Do
'                            Select Case rs_valores(0)
'                            Case 1 ' Obser
'                                .SubItems(6) = rs_valores(1)
'                            Case 2 ' Temp
'                                .SubItems(4) = rs_valores(1)
'                            Case 18 ' Hora
'                                .SubItems(5) = rs_valores(1)
'                            Case 19 ' Nivel
'                                .SubItems(19) = rs_valores(1)
'                            End Select
'                            rs_valores.MoveNext
'                        Loop Until rs_valores.EOF
'                    End If
'                    MUESTRA = rs(0)
'                    Do
'                        For j = 7 To 30
'                           If rs(4) = Pos(j) Then ' ID_DETER
'                                If Not rs(3) <> "" And Not IsNull(rs(3)) Then ' resultado
'                                   .SubItems(j) = "No informado"
'                                Else
'                                   .SubItems(j) = Replace(rs(3), ".", ",")
'                                End If
'                                Exit For
'                           End If
'                        Next
'                       If rs.EOF = False Then
'                           rs.MoveNext
'                       End If
'                    Loop Until MUESTRA <> rs(0)
'                End With
'                DoEvents
'            Loop Until rs.EOF
'        End If
'    End If
'    Me.MousePointer = 0
'    Exit Sub
'fallo:
'    Me.MousePointer = 0
'    MsgBox "Error al recuperar los resultados de los baños.", vbCritical, App.Title
'End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If UnloadMode = vbFormControlMenu Then
    comprobar_consolidar_en_servidor
End If
    
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        cmdRecarga_Click
'        gmuestra = lista.ListItems(lista.SelectedItem.Index).Text
'        frmVerMuestra.Show 1
    End If
End Sub

Private Sub txtensayo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtensayo <> "" Then
            If IsNumeric(txtensayo) Then
                Dim c As String
                c = "SELECT * FROM MUESTRAS WHERE ID_GENERAL = " & txtensayo & " AND ANNO = " & txtanno
                Dim rs As ADODB.Recordset
                Set rs = datos_bd(c)
                If rs.RecordCount > 0 Then
                    If rs("BANO_ID") = 0 Then
                        txtensayo = ""
                        Exit Sub
                    End If
                    If cmbbanos.BoundText <> rs("BANO_ID") Then
                        Dim oBANO As New clsBanos
                        oBANO.cargar_bano rs("BANO_ID")
                        cmbLinea.BoundText = oBANO.getID_LINEA
                        cmbbanos.BoundText = rs("BANO_ID")
                        cmdBuscar_Click
                    End If
                    Dim i As Integer
                    For i = 1 To lista.ListItems.Count
                        If lista.ListItems(i).Text = rs("ID_MUESTRA") Then
                            lista.ListItems(i).EnsureVisible
                            Set lista.selectedItem = lista.ListItems(i)
                            cmdRecarga_Click
                        End If
                    Next
                End If
            End If
        End If
        txtensayo = ""
    End If
End Sub

Private Sub UpDown1_Change()
'    cargar_bano
End Sub

Private Sub cargar_lineas()
    Dim olinea As New clsLineas
    Dim ooe As New clsEmpleados_Estados
    Set cmbLinea.RowSource = olinea.Listado_Historico
    cmbLinea.ListField = "NOMBRE"
    cmbLinea.BoundColumn = "ID_LINEA"
    Set olinea = Nothing
End Sub

Private Sub cargar_lista_banos()
    cmbbanos.Text = ""
    If cmbLinea.Text <> "" Then
        Dim oBANO As New clsBanos
        Set cmbbanos.RowSource = oBANO.Listado_Lineas_Controladas(cmbLinea.BoundText)
        cmbbanos.ListField = "BANO"
        cmbbanos.BoundColumn = "ID_BANO"
        Set olinea = Nothing
    End If
End Sub

Private Sub genera_excel_bano(BANO As Long)
   On Error GoTo genera_excel_bano_Error

    Me.MousePointer = 11
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    Set XLA = New excel.Application
    Dim hoja As String
    Dim oBANO As New clsBanos
    Dim olinea As New clsLineas
    oBANO.cargar_bano BANO
    olinea.CARGAR oBANO.getID_LINEA
    hoja = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\"
    hoja = hoja & olinea.getNOMBRE & "\"
    hoja = hoja & Replace(oBANO.getNOMBRE, "/", "") & ".xls"
    Set XLW = XLA.Workbooks.Open(hoja)
    Set XLS = XLW.Worksheets(1)
    ' Buscamos la fecha de la última muestra registrada
    Dim fila As Integer
    Dim encontrado As Boolean
    Dim ultima_fecha As String
    fila = fila_inicial
    encontrado = False
    While encontrado = False
        If XLS.Cells(fila, COLS.fecha) <> "" Then
            ultima_fecha = Format(XLS.Cells(fila, COLS.fecha), "yyyy-mm-dd")
            fila = fila + 1
        Else
            encontrado = True
        End If
    Wend
'    MsgBox fila & ": " & ultima_fecha
    ' Recuperamos todas las muestras cerradas superiores a la última
    Dim oDatos_valores As New clsDatos_valores
    Dim rs As ADODB.Recordset
    Dim rs_valores As ADODB.Recordset
    Set rs = oBANO.Resultados_Banos_fecha(BANO, ultima_fecha)
    Dim MUESTRA As String
    On Error Resume Next
    If rs.RecordCount <> 0 Then
        MUESTRA = ""
        Do
            ' Fecha de la muestra
            XLS.Cells(fila, COLS.fecha) = rs(2)
            XLS.Cells(fila, COLS.INFORME_1) = rs(5)
            XLS.Cells(fila, COLS.ACCION) = "Conforme"
            ' Temperatura y Nivel
            Set rs_valores = oDatos_valores.datos_muestra_historico(rs(0), "2,19")
            If rs_valores.RecordCount > 0 Then
                Do
                    Select Case rs_valores(0)
                    Case 2 ' Temp
                        XLS.Cells(fila, COLS.TEMPERATURA) = CDbl(Replace(rs_valores(1), ".", ","))
                    Case 19 ' Nivel
                        XLS.Cells(fila, COLS.NIVEL) = CStr(rs_valores(1))
                    End Select
                    rs_valores.MoveNext
                Loop Until rs_valores.EOF
            End If
            ' Recargas e informes
            Dim oRecarga As New clsRecargas
            If oRecarga.CARGAR(rs(0)) = True Then
                If oRecarga.getAN_RUTA <> "" Then
                    XLS.Hyperlinks.Add XLS.Cells(fila, COLS.INFORME_1), Replace(oRecarga.getAN_RUTA, "\", "/"), , , "AN-" & Mid(rs(1), 4, 13)
                End If
                If oRecarga.getOR_RUTA <> "" Then
                    XLS.Hyperlinks.Add XLS.Cells(fila, COLS.ACCION), Replace(oRecarga.getOR_RUTA, "\", "/"), , , "OR-" & Mid(rs(1), 4, 13)
                Else
                    XLS.Cells(fila, 12) = "Conforme"
                End If
                If oRecarga.getRR_RUTA <> "" Then
                    XLS.Hyperlinks.Add XLS.Cells(fila, COLS.RESPUESTA), Replace(oRecarga.getRR_RUTA, "\", "/"), , , "RR-" & Mid(rs(1), 4, 13)
                End If
            End If
            ' Resultados
            MUESTRA = rs(0)
            Col = COLS.D1
            Do
                    If Not rs(3) <> "" And Not IsNull(rs(3)) Then ' resultado
                        XLS.Cells(fila, Col) = "No informado"
                    Else
                        If IsNumeric(rs(3)) Then
                            XLS.Cells(fila, Col) = CDbl(Replace(rs(3), ".", ","))
                        Else
                            XLS.Cells(fila, Col) = rs(3)
                        End If
                    End If
                   If rs.EOF = False Then
                       rs.MoveNext
                   End If
                   Col = Col + 1
            Loop Until rs.EOF Or MUESTRA <> rs(0)
            DoEvents
            fila = fila + 1
        Loop Until rs.EOF
    End If
    XLW.Save
    XLA.Quit
'    XLA.Visible = True
    Set XLW = Nothing
    Set XLA = Nothing
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

genera_excel_bano_Error:
    Me.MousePointer = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure genera_excel_bano of Formulario frmEads_Historico"
    XLA.visible = True
    Set XLW = Nothing
    Set XLA = Nothing
End Sub

Private Sub genera_excel_analisis(BANO As Long)
    On Error GoTo fallo
    Me.MousePointer = 11
    Dim oGMSO As New Geslab_MSOLink.clsMSOExcel
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    Set XLA = New excel.Application
    Dim oBANO As New clsBanos
    Dim olinea As New clsLineas
    oBANO.cargar_bano BANO
    olinea.CARGAR oBANO.getID_LINEA
    hoja = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\"
    hoja = hoja & olinea.getNOMBRE & "\"
    hoja = hoja & Replace(oBANO.getNOMBRE, "/", "") & ".xls"
    Set XLW = XLA.Workbooks.Open(hoja)
    Set XLS = XLW.Worksheets(1)
    ' Recuperamos todas las muestras cerradas superiores a la última
    Dim rs As ADODB.Recordset
    Dim origen As String
    Dim destino As String
    Dim EDICION As Integer
    fila = fila_inicial
'    fila = 211
    While XLS.Cells(fila, COLS.fecha) <> ""
'      If XLS.Cells(fila, Cols.INFORME_1) <> "" Then
        XLS.Cells(fila, COLS.ERROR) = CStr("")
        Set rs = oBANO.Resultados_Bano_por_fecha(BANO, Format(XLS.Cells(fila, COLS.fecha), "yyyy-mm-dd"))
        If rs.RecordCount > 0 Then
          ' Busco la edicion 1 en el raiz
          EDICION = 1
          origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Análisis\" & rs(6) & "\AN-" & rs(5) & "-" & rs(6) & "-Ed1.pdf"
          If Dir(origen) = "" Then
            ' Busco la edicion 1 en la carpeta del año por si ya se ha copiado
            origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Análisis\" & "AN-" & rs(5) & "-" & rs(6) & "-Ed1.pdf"
            If Dir(origen) = "" Then
                ' Busco la edicion 2 en la carpeta raiz
                origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Análisis\" & "AN-" & rs(5) & "-" & rs(6) & "-Ed2.pdf"
                EDICION = 2
                If Dir(origen) = "" Then
                    ' Busco la edicion 2 en la carpeta del año por si ya se ha copiado
                    origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Análisis\" & rs(6) & "\AN-" & rs(5) & "-" & rs(6) & "-Ed2.pdf"
                    EDICION = 2
                End If
            End If
          End If
          If Dir(origen) <> "" Then
            destino = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Análisis\" & rs(6) & "\AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(EDICION) & ".pdf"
            If origen <> destino Then
               gFSO.CopyFile origen, destino, True
            End If
            origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Análisis\" & "AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(EDICION) & ".pdf"
            On Error Resume Next
            Kill origen
            On Error GoTo fallo
            XLS.Hyperlinks.Add XLS.Cells(fila, COLS.INFORME_1), "../Documentos/Análisis/" & rs(6) & "/AN-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(EDICION) & ".pdf"
            If XLS.Cells(fila, COLS.INFORME_2) <> "" Then
                XLS.Hyperlinks.Add XLS.Cells(fila, COLS.INFORME_2), "../Documentos/Análisis/" & rs(6) & "/AN-" & rs(5) & "-" & rs(6) & "-Ed2.pdf"
            End If
          Else
            XLS.Cells(fila, COLS.ERROR) = "*"
          End If
        Else
          XLS.Cells(fila, COLS.ERROR) = "-"
        End If
'      End If
      fila = fila + 1
    Wend
    
    
    'JONATHAN.2010.07.13. ACTUALIZACION DE LAS GRAFICAS
    
    Set XLW = oGMSO.ActualizarGraficasExcel(XLW, fila_inicial, fila - 1)
    
    'FIN JONATHAN.2010.07.13
    
    XLW.Save
    XLA.Quit
'    XLA.Visible = True
    Set XLW = Nothing
    Set XLA = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se han producido errores al abrir la hoja excel: " & Err.Description, vbCritical, "FILA:" & fila & " COL:" & Col
    XLA.visible = True
    Set XLW = Nothing
    Set XLA = Nothing

End Sub
Private Sub genera_excel_OR(BANO As Long)
    On Error GoTo fallo
    Me.MousePointer = 11
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    Set XLA = New excel.Application
    Dim oBANO As New clsBanos
    Dim olinea As New clsLineas
    oBANO.cargar_bano BANO
    olinea.CARGAR oBANO.getID_LINEA
    hoja = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\"
    hoja = hoja & olinea.getNOMBRE & "\"
    hoja = hoja & Replace(oBANO.getNOMBRE, "/", "") & ".xls"
    Set XLW = XLA.Workbooks.Open(hoja)
    Set XLS = XLW.Worksheets(1)
    ' Recuperamos todas las muestras cerradas superiores a la última
    Dim rs As ADODB.Recordset
    Dim origen As String
    Dim destino As String
    Dim fichero As String
    Dim EDICION As Integer
    Dim p As Integer
    fila = fila_inicial
    While XLS.Cells(fila, COLS.fecha) <> ""
        Set rs = oBANO.Resultados_Bano_por_fecha(BANO, Format(XLS.Cells(fila, COLS.fecha), "yyyy-mm-dd"))
        If rs.RecordCount > 0 Then
          ' Busco la OR
          origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\OR\" & "OR-" & rs(5) & "-" & rs(6) & "*"
          fichero = Dir(origen)
          p = 1
          If fichero = "" Then
           origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Ordenes de Recarga\" & rs(6) & "\" & "OR-" & rs(5) & "-" & rs(6) & "*"
           fichero = Dir(origen)
           p = 2
          End If
          If fichero <> "" Then
            If p = 1 Then
             origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\OR\" & fichero
             destino = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Ordenes de Recarga\" & rs(6) & "\" & fichero
             gFSO.CopyFile origen, destino
             On Error Resume Next
             gFSO.DeleteFile origen, True
            End If
            On Error GoTo fallo
            If InStr(1, UCase(XLS.Cells(fila, COLS.ACCION)), "CONFORME") > 0 Or Trim(XLS.Cells(fila, COLS.ACCION)) = "" Then
                XLS.Cells(fila, COLS.ACCION) = "OR-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(rs(7))
            End If
            XLS.Hyperlinks.Add XLS.Cells(fila, COLS.ACCION), "../Documentos/Ordenes de Recarga/" & rs(6) & "/" & fichero
'            XLS.Cells(fila, Cols.ERROR) = "GGG"&
          End If
          ' Busco la RR
          p = 1
          origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\RR\" & "RR-" & rs(5) & "-" & rs(6) & "*"
          fichero = Dir(origen)
          If fichero = "" Then
           origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Respuesta Recarga\" & rs(6) & "\" & "RR-" & rs(5) & "-" & rs(6) & "*"
           fichero = Dir(origen)
           p = 2
          End If
          If fichero <> "" Then
            If p = 1 Then
             origen = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\RR\" & fichero
             destino = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\Documentos\Respuesta Recarga\" & rs(6) & "\" & fichero
             gFSO.CopyFile origen, destino
             On Error Resume Next
             Kill origen
            End If
            On Error GoTo fallo
            If InStr(1, UCase(XLS.Cells(fila, COLS.RESPUESTA)), "CONFORME") > 0 Or Trim(XLS.Cells(fila, COLS.RESPUESTA)) = "" Then
                XLS.Cells(fila, COLS.RESPUESTA) = "RR-" & rs(5) & "-" & rs(6) & "-Ed" & CStr(rs(7))
            End If
            XLS.Hyperlinks.Add XLS.Cells(fila, COLS.RESPUESTA), "../Documentos/Respuesta Recarga/" & rs(6) & "/" & fichero
'            XLS.Cells(fila, Cols.ERROR) = "GGG"
          End If
        End If
      fila = fila + 1
    Wend
    XLW.Save
    XLA.Quit
'    XLA.Visible = True
    Set XLW = Nothing
    Set XLA = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se han producido errores al abrir la hoja excel: " & Err.Description, vbCritical, "FILA:" & fila & " COL:" & Col
    XLA.visible = True
    Set XLW = Nothing
    Set XLA = Nothing
End Sub

Private Sub genera_grafico(BANO As Long)
    On Error GoTo fallo
    Me.MousePointer = 11
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    Set XLA = New excel.Application
    Dim oBANO As New clsBanos
    Dim olinea As New clsLineas
    oBANO.cargar_bano BANO
    olinea.CARGAR oBANO.getID_LINEA
    hoja = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\"
    
    Dim linea As String
'    Select Case olinea.getID_LINEA
'        Case 10
'                        linea = "IFQ-01 INST. FRESADO QUIMICO DE ALUMINIO"
'                    Case 25
'                        linea = "ALUMINIO L1. CBC"
'                    Case 27
'                        linea = "ALUMINIO L2. CBC"
'                    Case 28
'                        linea = "CSP TITANIO L3. CBC"
'                    Case Else
                        linea = olinea.getNOMBRE
'                End Select
    
    
    hoja = hoja & linea & "\"
    hoja = hoja & Replace(oBANO.getNOMBRE, "/", "") & ".xls"
    Set XLW = XLA.Workbooks.Open(hoja)
    Set XLS = XLW.Worksheets(1)
    
    Dim i As Integer
    i = 13
    Dim posi As Integer
    Dim posf As Integer
    ' Ultima fecha
    While XLS.Cells(i, 2) <> ""
        If IsDate(XLS.Cells(i, 2)) Then
            fecha_ultima = CDate(XLS.Cells(i, 2))
        End If
        i = i + 1
    Wend
    ' Fecha inferior
    Dim fecha_inferior As String
    Dim op As New clsParametros
    Dim DIAS As Integer
    op.Carga parametros.HISTORICO_DIAS_GRAFICAS, ""
    If op.getVALOR = "" Then
        DIAS = 0
    Else
        DIAS = op.getVALOR
    End If
    fecha_inferior = fecha_ultima - DIAS
    i = 13
    posi = 0
    While XLS.Cells(i, 2) <> ""
        If IsDate(XLS.Cells(i, 2)) Then
            If CDate(XLS.Cells(i, 2)) >= CDate(fecha_inferior) And posi = 0 Then
                posi = i
            End If
        End If
        i = i + 1
    Wend
    If posi = 0 Then
        posi = 13
    End If
    posf = i - 1
    ' ACTUALIZA SERIE GRAFICA
    ActualizarGraficasExcel XLW, posi, posf
'    Dim oGMSO As New Geslab_MSOLink.clsMSOExcel
'    Set XLW = oGMSO.ActualizarGraficasExcel(XLW, posi, posf)
'    Set oGMSO = Nothing
    
'    XLA.Visible = True
    XLW.Save
    XLA.Quit
    
    Set XLW = Nothing
    Set XLA = Nothing
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Se han producido errores al abrir la hoja excel: " & Err.Description, vbCritical, "FILA:" & fila & " COL:" & Col
    XLA.visible = True
    Set XLW = Nothing
    Set XLA = Nothing
End Sub
'Private Sub azul(BANO As Long)
'    On Error GoTo fallo
'    Me.MousePointer = 11
'    Dim XLA As Excel.Application
'    Dim XLW As Excel.Workbook
'    Dim XLS As Excel.Worksheet
'    Set XLA = New Excel.Application
'    Dim oBANO As New clsBanos
'    Dim olinea As New clsLineas
'    oBANO.cargar_bano BANO
'    olinea.CARGAR oBANO.getID_LINEA
'    hoja = ReadINI(App.Path + "\config.ini", "documentos", "HistoricoBanos") & "\"
'    hoja = hoja & olinea.getNOMBRE & "\"
'    hoja = hoja & Replace(oBANO.getNOMBRE, "/", "") & ".xls"
'    Set XLW = XLA.Workbooks.Open(hoja)
'    Set XLS = XLW.Worksheets(1)
'    XLS.Range(XLS.Cells(13, 3), XLS.Cells(500, 9)).Font.ColorIndex = 5
'    XLW.Save
'    XLA.Quit
''    XLA.Visible = True
'    Set XLW = Nothing
'    Set XLA = Nothing
'    Me.MousePointer = 0
'    Exit Sub
'fallo:
'    Me.MousePointer = 0
'    MsgBox "Se han producido errores al abrir la hoja excel: " & Err.Description, vbCritical, "FILA:" & fila & " COL:" & Col
'    XLA.Visible = True
'    Set XLW = Nothing
'    Set XLA = Nothing
'End Sub



Public Function ActualizarGraficasExcel(ByRef prmWB As excel.Workbook, ByVal prmFilaIni As Integer, ByVal prmFilaFin As Integer) As excel.Workbook

    ' Los nombres de las hojas
    Dim Hoja1 As String, Hoja2 As String
    Dim oSerie As excel.Series, oChart As excel.ChartObject
    Dim oSheet As excel.Worksheet
    Dim lst_Chart As String, lst_Serie As String
    
    lst_Chart = ""
    lst_Serie = ""
    
    ' Guarda los nombres de las hojas
    Hoja1 = prmWB.Worksheets(1).Name
    Hoja2 = prmWB.Worksheets(2).Name
    Set oSheet = prmWB.Worksheets(2)
    
    
    On Error GoTo ActualizarGraficasExcel_Error

            contChart = 1
            Do While contChart > 0
                On Error Resume Next
                Set oChart = oSheet.ChartObjects(contChart)
                
                If lst_Chart = oChart.Name Then
                    If Err.Number <> 0 Then contSerie = 0
                    On Error GoTo ActualizarGraficasExcel_Error
                
                    contChart = 0 ' para salir del bucle
                Else
                    On Error GoTo ActualizarGraficasExcel_Error
                    ' trabajamos con el Chart
                    ' Guardamos el nombre como último tratado
                    lst_Chart = oChart.Name '(Tip para Excel, por obligacion hay que hacerlo así)
                    
                    contSerie = 1
                    Do While contSerie > 0
                        ' Recogemos la Serie
                        On Error Resume Next
                        Set oSerie = oChart.Chart.SeriesCollection(contSerie)
                        
                        
                        If lst_Serie = oSerie.Name Then
                            If Err.Number <> 0 Then contSerie = 0
                            On Error GoTo ActualizarGraficasExcel_Error
                            contSerie = 0 ' para salir del bucle
                        Else
                            On Error GoTo ActualizarGraficasExcel_Error
                            ' trabajamos con la serie
                            ' guardamos el nombre de la serie como última con la que hemos tratado (Tip para Excel, por obligacion hay que hacerlo así)
                            lst_Serie = oSerie.Name
                            
                            oSerie.Formula = Modificar_FORMULA(oSerie.Formula, prmFilaIni, prmFilaFin, Hoja1, Hoja2)
                            contSerie = contSerie + 1
                        End If
                    Loop
                    contChart = contChart + 1
                End If
            Loop
    
    Set ActualizarGraficasExcel = prmWB

Exit Function

ActualizarGraficasExcel_Error:
    log "Se ha producido un error al actualizar una grafica excel: Estos son los datos: " & vbCrLf & Err.Number & " " & Err.Description & vbCrLf & "Libro: " & oSheet.Parent.Name & "(Reportado por : " & USUARIO.getUSUARIO & ")"
'    Call Enviar_Mail_CDO("informatica@canagrosa", "[ERROR] Actualicion Gráficas Excel", "Se ha producido un error al actualizar una grafica excel: Estos son los datos: " & vbCrLf & Err.Number & " " & Err.Description & vbCrLf & "Libro: " & oSheet.Parent.Name & "(Reportado por : " & USUARIO.getUSUARIO & ")", vbNullString)
    Set ActualizarGraficasExcel = Nothing
    
End Function


Private Function Modificar_FORMULA(ByRef prmFormula As String, ByVal prmFilaIni As Integer, ByVal prmFilaFin As Integer, ByVal Hoja1 As String, ByVal Hoja2 As String) As String

    
    Dim valores As String
    Dim fechas As String
    Dim parte1 As String, parte4 As String
    Dim strCad As String
    Dim pa(1 To 2) As String
    Dim pb(1 To 2) As String
    Dim Col(1 To 2) As String
    
    
    Modificar_FORMULA = prmFormula
    
    
    ' Referencia de la formula
    ' =SERIE("TURCO  NCLT",Hoja1!$B$51:$B$92,Hoja1!$C$51:$C$92,1)
    
'    If InStr(1, prmFormula, Hoja2) > 0 Then
'        Exit Function
'    End If
    
    ' Comienza a Desglosar
    
    parte1 = Split(prmFormula, ",")(0)
    fechas = Split(prmFormula, ",")(1)
    valores = Split(prmFormula, ",")(2)
    parte4 = Split(prmFormula, ",")(3)
    
    ' Reestablece el Rango para FECHAS (Eje X)
    strCad = fechas
    pa(1) = Split(strCad, "!")(0)
    pa(2) = Split(strCad, "!")(1)
    pb(1) = Split(pa(2), ":")(0)
    pb(2) = Split(pa(2), ":")(1)
    Col(1) = Split(pb(1), "$")(1)
    Col(2) = Split(pb(2), "$")(1)
    strCad = pa(1) & "!$" & Col(1) & "$" & CStr(prmFilaIni) & ":$" & Col(2) & "$" & CStr(prmFilaFin)
    fechas = strCad
    
    ' Reestablece el Rango para VALORES
    strCad = valores
    pa(1) = Split(strCad, "!")(0)
    pa(2) = Split(strCad, "!")(1)
    pb(1) = Split(pa(2), ":")(0)
    pb(2) = Split(pa(2), ":")(1)
    Col(1) = Split(pb(1), "$")(1)
    Col(2) = Split(pb(2), "$")(1)
    strCad = pa(1) & "!$" & Col(1) & "$" & CStr(prmFilaIni) & ":$" & Col(2) & "$" & CStr(prmFilaFin)
    valores = strCad
    
    ' Composicion final
    strCad = parte1 & "," & fechas & "," & valores & "," & parte4
    
    Modificar_FORMULA = strCad
End Function



