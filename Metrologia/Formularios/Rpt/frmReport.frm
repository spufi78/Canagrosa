VERSION 5.00
Object = "{FB992564-9055-42B5-B433-FEA84CEA93C4}#11.0#0"; "crviewer.dll"
Begin VB.Form frmReport 
   Caption         =   "Informe"
   ClientHeight    =   6375
   ClientLeft      =   4275
   ClientTop       =   2610
   ClientWidth     =   6555
   Icon            =   "frmReport.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   6555
   WindowState     =   2  'Maximized
   Begin CrystalActiveXReportViewerLib11Ctl.CrystalActiveXReportViewer crviewer 
      Height          =   6045
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   6525
      _cx             =   11509
      _cy             =   10663
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   0   'False
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
      EnableLogonPrompts=   -1  'True
      LocaleID        =   1034
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public informe As String
Public consulta As String
Public CRITERIO As String
Public ordenacion As String
Public imprimir As Boolean
Public copias As Integer
Public pdf As String
Public xml As String
'Private mvarblnMostrarLogo As Boolean
Private mvarblnMostrarTabArbol As Boolean
Private mvarArrParametrosNombre() As String
Private mvarArrParametrosValores() As String
Public Property Let ParametrosNombre(ByRef prmParametros() As String)
    mvarArrParametrosNombre = prmParametros
End Property
Public Property Let ParametrosValores(ByRef prmParametros() As String)
    mvarArrParametrosValores = prmParametros
End Property
Public Property Let MostrarTabArbol(ByVal blnMostrarTabArbol As Boolean)
    mvarblnMostrarTabArbol = blnMostrarTabArbol
End Property

Public Sub iniciar()
    informe = ""
    consulta = ""
    CRITERIO = ""
    ordenacion = ""
    imprimir = False
    mvarblnMostrarTabArbol = False
    pdf = ""
    xml = ""
    Dim vacio() As String
    mvarArrParametrosNombre = vacio
    mvarArrParametrosValores = vacio
End Sub

Private Sub Form_Resize()
    With crviewer
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Public Function generar() As Boolean
    generar = False
    Dim crystal As CRAXDRT.Application
    Dim report As CRAXDRT.report
    Dim total_parametros As Integer, intCont As Integer
   On Error GoTo generar_Error
    Me.MousePointer = 11
    Set crystal = New CRAXDRT.Application
    log "**********************************************************"
    log "COMIENZO GENERACION DE INFORME "
    log "Nombre : " & informe
    log "Criterio : " & CRITERIO
    log "Consulta : " & consulta
    log "Ordenacion : " & ordenacion
    log "**********************************************************"
    Set report = crystal.OpenReport(ReadINI(App.Path + "\config.ini", "documentos", "informes") & "\" & informe & ".rpt")
        
    If xml <> "" Then
        report.database.Tables(1).Location = xml
    End If
    
'    Dim tablas As Integer
'    For tablas = 1 To report.database.Tables.Count
'        report.database.Tables(tablas).Location = "BCA_PRUEBA"
'        report.database.Tables(tablas).SetTableLocation BCA_PRUEBA, "", ""
'    Next
    
        
    If ordenacion <> "" Then
        Dim crxDatabaseField As CRAXDRT.DatabaseFieldDefinition
        For i = 1 To report.database.Tables.Count
            For j = 1 To report.database.Tables.Item(i).Fields.Count
                If report.database.Tables.Item(i).Fields.Item(j).Name = ordenacion Then
                    Set crxDatabaseField = report.database.Tables.Item(i).Fields.Item(j)
                End If
            Next
        Next
        report.RecordSortFields.Add crxDatabaseField, crAscendingOrder
    End If
    
    report.DiscardSavedData
    ' Parametros del informe
    total_parametros = 0
    On Error Resume Next
    total_parametros = UBound(mvarArrParametrosNombre)
    On Error GoTo generar_Error
    ' Establece los parametros si existen
    If total_parametros >= 1 Then
        For intCont = 1 To total_parametros
            report.ParameterFields.GetItemByName(mvarArrParametrosNombre(intCont)).AddCurrentValue (mvarArrParametrosValores(intCont))
            log "PARAMETRO : " & mvarArrParametrosNombre(intCont) & " -> " & mvarArrParametrosValores(intCont)
        Next intCont
    End If
    report.RecordSelectionFormula = ""
    report.RecordSelectionFormula = CRITERIO

    If consulta <> "" Then
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        report.database.SetDataSource rs
    End If
      
    If UCase(informe) = "RPTALBARAN" Or UCase(informe) = "RPTFACTURAPEQUENA" Then
        report.SetUserPaperSize ReadINI(App.Path + "\config.ini", "PARAMETROS", "ALBARAN_ALTO"), ReadINI(App.Path + "\config.ini", "PARAMETROS", "ALBARAN_ANCHO")
        report.PaperSize = crPaperUser
    End If
    If UCase(informe) = "RPTSOBRE" Then
        report.SetUserPaperSize 2340, 1190
        report.PaperSize = crPaperUser
    End If
        
    With crviewer
        .DisplayBorder = False
        .DisplayTabs = False
        .EnableDrilldown = False
        .EnableGroupTree = mvarblnMostrarTabArbol
        .EnableRefreshButton = False
        .ReportSource = report
        .ViewReport
        Do While .IsBusy
            DoEvents
        Loop
        .Zoom 94
        If Trim(pdf) <> "" Then
            report.ExportOptions.DestinationType = crEDTDiskFile
            report.ExportOptions.FormatType = crEFTPortableDocFormat
            report.ExportOptions.DiskFileName = pdf
            report.ExportOptions.PDFExportAllPages = True
            report.Export False
        Else
            If imprimir = True Then
                Dim h As Integer
                For h = 1 To copias
                    report.PrintOut False, 1
                Next
            End If
        End If
    End With
    limpiar_variables
    generar = True
    Set crystal = Nothing
    Set report = Nothing
    Me.MousePointer = 0
    log "**********************************************************"
    log "FIN GENERACION DE INFORME"
    log "**********************************************************"

   On Error GoTo 0
   Exit Function

generar_Error:
    Me.MousePointer = 0
    limpiar_variables
    log "**********************************************************"
    log "---------------ERROR GENERACION DE INFORME----------------"
    log "**********************************************************"
    Set crystal = Nothing
    Set report = Nothing
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar of Formulario frmReport"
End Function

Public Sub Cerrar()
    Unload Me
End Sub

Public Sub limpiar_variables()
    CRITERIO = ""
    ordenacion = ""
    consulta = ""
    pdf = ""
    informe = ""
    xml = ""
End Sub

