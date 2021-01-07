VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrevisualizar2 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Previsualizar Datos de Informe"
   ClientHeight    =   9885
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   14805
   Icon            =   "frmPrevisualizar2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9885
   ScaleWidth      =   14805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtplantilla 
      Height          =   285
      Left            =   11880
      TabIndex        =   11
      Top             =   8550
      Visible         =   0   'False
      Width           =   2670
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Opciones"
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
      Height          =   1260
      Left            =   90
      TabIndex        =   1
      Top             =   8550
      Width           =   7035
      Begin VB.CheckBox chkCabecera 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Ocultar cabecera"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4545
         TabIndex        =   10
         Top             =   270
         Width           =   2055
      End
      Begin VB.CheckBox chkimpresa 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Muestra impresa"
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
         Height          =   255
         Left            =   4545
         TabIndex        =   9
         Top             =   540
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CheckBox chkcorreo 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enviada por correo"
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
         Height          =   315
         Left            =   4545
         TabIndex        =   8
         Top             =   765
         Value           =   1  'Checked
         Width           =   2055
      End
      Begin VB.CommandButton cmdmail 
         BackColor       =   &H00E0E0E0&
         Caption         =   "E-mail"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   2835
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Enviar informe de la ultima edición generada por E-mail"
         Top             =   270
         Width           =   1545
      End
      Begin VB.TextBox txtediciongen 
         Alignment       =   2  'Center
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
         Left            =   315
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   585
         Width           =   705
      End
      Begin VB.CommandButton cmdImprimir 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   870
         Left            =   1260
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Imprimir ultima edición generada"
         Top             =   270
         Width           =   1545
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Edición"
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
         Left            =   300
         TabIndex        =   4
         Top             =   315
         Width           =   675
      End
   End
   Begin VB.CommandButton cmdcancel 
      BackColor       =   &H00E0E0E0&
      Caption         =   "ESC-Salir"
      Height          =   870
      Left            =   13635
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   8955
      Width           =   1095
   End
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer crviewer 
      Height          =   8430
      Left            =   90
      TabIndex        =   6
      Top             =   45
      Width           =   14670
      _cx             =   25876
      _cy             =   14870
      DisplayGroupTree=   0   'False
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   0   'False
      EnableNavigationControls=   -1  'True
      EnableStopButton=   0   'False
      EnablePrintButton=   0   'False
      EnableZoomControl=   -1  'True
      EnableCloseButton=   0   'False
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   0   'False
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   0   'False
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   0   'False
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1034
   End
   Begin MSComctlLib.ListView lista 
      Height          =   1335
      Left            =   7200
      TabIndex        =   7
      Top             =   8505
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   2355
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
End
Attribute VB_Name = "frmPrevisualizar2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public PK_MUESTRA As Long

Private Sub chkCabecera_Click()
    If txtplantilla <> "" Then
        cargar_documento txtplantilla, "{muestras.ID_MUESTRA} = " & PK_MUESTRA, "", "", False
    End If
End Sub

Private Sub cmdcancel_Click()
    Unload Me
End Sub

Private Sub cmdImprimir_Click()
    Dim omuestra As New clsMuestra
   On Error GoTo cmdImprimir_Click_Error

    omuestra.informar_impresion PK_MUESTRA, USUARIO.getID_EMPLEADO
    Set omuestra = Nothing
    chkimpresa.value = Checked
    chkimpresa.BackColor = &HC0FFFF
    crviewer.PrintReport

   On Error GoTo 0
   Exit Sub

cmdImprimir_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdImprimir_Click of Formulario frmPrevisualizar2"
End Sub

Private Sub cmdmail_Click()
    On Error GoTo fallo
    Me.MousePointer = 11
    Dim omuestra As New clsMuestra
    omuestra.informar_correo PK_MUESTRA, USUARIO.getID_EMPLEADO
    Set omuestra = Nothing
    chkcorreo.value = Checked
    chkcorreo.BackColor = &HC0FFFF
    On Error Resume Next
    MkDir App.Path & "\temp"
    NOMBRE_DOCUMENTO PK_MUESTRA, True
    ' Creamos el pdf con el report
    Dim report As CRAXDRT.report
    Set report = crviewer.ReportSource
    With report
    .ExportOptions.DestinationType = crEDTDiskFile
    .ExportOptions.FormatType = crEFTPortableDocFormat
    .ExportOptions.DiskFileName = App.Path & "\temp\" & referencia_pdf
    .ExportOptions.PDFExportAllPages = True
    .Export False
    End With
    Set report = Nothing
    ' Lo enviamos por correo
    enviar_informe PK_MUESTRA, App.Path & "\temp\" & referencia_pdf, Me.hWnd
    Me.MousePointer = 0
    Exit Sub
fallo:
    Me.MousePointer = 0
    MsgBox "Error al enviar la muestra por correo. " & Err.Description, vbCritical, App.Title
End Sub

Private Sub Form_Load()
    log Me.Name
    cargar_botones Me
    cabecera
    cargar_muestra
End Sub
Private Sub cargar_documento(PLANTILLA As String, CRITERIO As String, consulta As String, pdf As String, imprimir As Boolean)
    Dim crystal As CRAXDRT.Application
    Dim report As CRAXDRT.report
    Dim total_parametros As Integer, intCont As Integer
   On Error GoTo generar_Error
    Me.MousePointer = 11
    Set crystal = New CRAXDRT.Application
    Set report = crystal.OpenReport(ReadINI(App.Path + "\config.ini", "documentos", "reportes") & "\" & PLANTILLA)
    report.DiscardSavedData
    ' Parametros del informe
    total_parametros = 0
    On Error Resume Next
'    total_parametros = UBound(mvarArrParametrosNombre)
    On Error GoTo generar_Error
    ' Establece los parametros si existen
'    If total_parametros >= 1 Then
'        For intCont = 1 To total_parametros
'            report.ParameterFields.GetItemByName(mvarArrParametrosNombre(intCont)).AddCurrentValue (mvarArrParametrosValores(intCont))
'        Next intCont
'    End If
    report.RecordSelectionFormula = CStr(CRITERIO)
    If consulta <> "" Then
        Dim rs As ADODB.RecordSet
        Set rs = datos_bd(consulta)
        report.database.SetDataSource rs
    End If
'    If chkCabecera.value = Checked Then
'        report.ParameterFields.GetItemByName("LOGO").AddCurrentValue ("0")
'    Else
'        report.ParameterFields.GetItemByName("LOGO").AddCurrentValue ("1")
'    End If
    
    With crviewer
        .ReportSource = report
        .ViewReport
        Do While .IsBusy
            DoEvents
        Loop
        .Zoom 110
        If Trim(pdf) <> "" Then
            report.ExportOptions.DestinationType = crEDTDiskFile
            report.ExportOptions.FormatType = crEFTPortableDocFormat
            report.ExportOptions.DiskFileName = pdf
            report.ExportOptions.PDFExportAllPages = True
            report.Export False
        Else
            If imprimir = True Then
                report.PrintOut False, 1
            End If
        End If
    End With
    Set crystal = Nothing
    Set report = Nothing
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

generar_Error:
    Me.MousePointer = 0
    Set crystal = Nothing
    Set report = Nothing
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar of Formulario frmPrevisualizar"
End Sub

Private Sub cabecera()
    With lista.ColumnHeaders
        .Add , , "Ficheros Adjuntos...", lista.Width, lvwColumnLeft
        .Add , , "Ruta", 1, lvwColumnCenter
    End With
End Sub

Private Sub cargar_muestra()
    Dim omuestra As New clsMuestra
    omuestra.CargaMuestra (PK_MUESTRA)
    cmdImprimir.Enabled = True
    cmdmail.Enabled = True
    ' Edicion de la muestra
    txtediciongen = omuestra.getULT_EDICION_IMP
    If omuestra.getIMPRESA = 0 Then
        chkimpresa.value = Unchecked
        chkimpresa.BackColor = &HFF&
    End If
    If omuestra.getENVIADO_CORREO = 0 Then
        chkcorreo.value = Unchecked
        chkcorreo.BackColor = &HFF&
    End If
    If omuestra.getCERRADA <> 1 Then
        cmdImprimir.Enabled = False
        cmdmail.Enabled = False
    End If
    cargar_adjuntos
    If omuestra.getCERRADA = 1 Then
        txtplantilla = omuestra.obtener_plantilla(PK_MUESTRA)
        If txtplantilla <> "" Then
            cargar_documento txtplantilla, "{muestras.ID_MUESTRA} = " & PK_MUESTRA, "", "", False
        End If
    End If
End Sub
Private Function enviar_informe(ByVal MUESTRA As Long, destino_documento As String, manejador As Long) As Boolean
    Dim matriz(1) As String
    If Dir(destino_documento) = "" Then
        MsgBox "El informe aún no existe. Primero debe generarlo.", vbInformation, App.Title
        Exit Function
    End If
    ' Si es una muestra de ALIMENTOS DIA, la codificación para envio debe ser distinta
    Dim omuestra As New clsMuestra
    omuestra.CargaMuestra (MUESTRA)
    Dim destino_documento_dia As String
    If omuestra.getTIPO_MUESTRA_ID = 56 Then
        destino_documento_dia = omuestra.getREFERENCIA_CLIENTE
        Dim oDA As New clsDatos_valores
        Dim CODIGO_PRODUCTO As String
        CODIGO_PRODUCTO = oDA.Cod_Producto_Dia(MUESTRA)
        If Trim(CODIGO_PRODUCTO) <> "" Then
            destino_documento_dia = destino_documento_dia & " COD " & CODIGO_PRODUCTO
        End If
        destino_documento_dia = destino_documento_dia & " Nº INFORME " & omuestra.getID_GENERAL & "-" & Format(omuestra.getFECHA_RECEPCION, "yyyy") & ".pdf"
        On Error Resume Next
        MkDir App.Path & "\temp"
        FileCopy destino_documento, App.Path & "\temp\" & destino_documento_dia
        destino_documento = App.Path & "\temp\" & destino_documento_dia
    End If
    ' Generación de la Referencia para envio
    Dim ref As String
    matriz(0) = destino_documento
    Dim olb As New clsLineas_Banos
    DOCUMENTO = olb.Buscar_Bano(omuestra.getBANO_ID)
    If DOCUMENTO <> 0 Then ' Agua agrupada
        Dim olinea As New clsLineas
        Dim oBANO As New clsBanos
        oBANO.cargar_bano (omuestra.getBANO_ID)
        olinea.CARGAR (oBANO.getID_LINEA)
        ref = "Informe de la línea " & olinea.getNOMBRE
    Else
        ref = "Informe de la muestra " & omuestra.getREFERENCIA_CLIENTE
    End If
    ' Llamada a la función de envío de correo
    Dim ocliente As New clsCliente
    ocliente.CargaCliente (omuestra.getCLIENTE_ID)
    genera_correo ocliente.getEMAIL, ref, "Adjunto " & ref, destino_documento, manejador
    enviar_informe = True
    Exit Function
fallo:
    Close
    enviar_informe = False
    MsgBox "Error al enviar el informe.", vbCritical, Err.Description
End Function

Private Sub cargar_adjuntos()
    Dim m As Long
    Dim rs As ADODB.RecordSet
    lista.ListItems.Clear
    Dim oMuestra_Adjunto As New clsMuestras_adjuntos
    Set rs = oMuestra_Adjunto.Listado(PK_MUESTRA)
    If rs.RecordCount > 0 Then
        Do
            With lista.ListItems.Add(, , rs(0))
                 .SubItems(1) = ReadINI(App.Path + "\config.ini", "documentos", "ruta") & "\ADJUNTOS\" & rs(3) & "\" & rs(0)
            End With
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    Set oMuestra_Adjunto = Nothing
End Sub

Private Sub lista_DblClick()
    If lista.ListItems.Count > 0 Then
        Dim destino As String
        destino = lista.ListItems(lista.SelectedItem.Index).SubItems(1)
        On Error GoTo fallo
        If Dir(destino) <> "" Then
            r = Shell("rundll32.exe url.dll,FileProtocolHandler " & destino, vbNormalFocus)
        End If
    End If
    Exit Sub
fallo:
    MsgBox "Error al abrir el documento asociado.", vbCritical, App.Title
End Sub


