VERSION 5.00
Begin VB.Form frmReportSinVisor 
   Caption         =   "Informe"
   ClientHeight    =   6375
   ClientLeft      =   4275
   ClientTop       =   2610
   ClientWidth     =   6555
   Icon            =   "frmReportSinVisor.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6375
   ScaleWidth      =   6555
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmReportSinVisor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public informe As String
Public consulta As String
Public criterio As String
Public imprimir As Boolean
Public pdf As String
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
    criterio = ""
    imprimir = False
    mvarblnMostrarTabArbol = False
    pdf = ""
    Dim vacio() As String
    mvarArrParametrosNombre = vacio
    mvarArrParametrosValores = vacio
End Sub

Private Sub Form_Resize()
    With crviewer
        .top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Public Function generar() As Boolean
    Dim crystal As CRAXDRT.Application
    Dim report As CRAXDRT.report
    Dim total_parametros As Integer, intCont As Integer
   On Error GoTo generar_Error
    Me.MousePointer = 11
    log "frmReportSinVisor : Entrada"
    log "frmReportSinVisor, report : " & ReadINI(App.Path + "\config.ini", "documentos", "reportes") & "\" & informe & ".rpt"
    Set crystal = New CRAXDRT.Application
    Set report = crystal.OpenReport(ReadINI(App.Path + "\config.ini", "documentos", "reportes") & "\" & informe & ".rpt")
    report.DiscardSavedData
    log "frmReportSinVisor : informe"
    ' Parametros del informe
    total_parametros = 0
    On Error Resume Next
    total_parametros = UBound(mvarArrParametrosNombre)
    On Error GoTo generar_Error
    ' Establece los parametros si existen
    If total_parametros >= 1 Then
        For intCont = 1 To total_parametros
            report.ParameterFields.GetItemByName(mvarArrParametrosNombre(intCont)).AddCurrentValue (mvarArrParametrosValores(intCont))
        Next intCont
    End If
    log "frmReportSinVisor : Parametros"
    report.FormulaSyntax = crCrystalSyntaxFormula
    report.RecordSelectionFormula = criterio
    log "frmReportSinVisor - criterio : " & criterio

    If consulta <> "" Then
        log "frmReportSinVisor - consulta : " & consulta
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        report.database.SetDataSource rs
'    Else
'        If UCase(USUARIO.getUSUARIO) = "PRUEBA" Then
'            report.database.LogOnServerEx "Pdsmon.dll", "Geslab", ReadINI(App.Path + "\config.ini", "server", "bd_prueba"), "geslab", "ix1tec", "Active Data (ADO)", ""
'        Else
'            report.database.LogOnServerEx "Pdsmon.dll", "Geslab", ReadINI(App.Path + "\config.ini", "server", "bd"), "geslab", "ix1tec", "Active Data (ADO)", ""
'        End If
    End If
'    With crviewer
'        .DisplayBorder = False
'        .DisplayTabs = False
'        .EnableDrillDown = False
'        .EnableGroupTree = mvarblnMostrarTabArbol
'        .EnableRefreshButton = False
'        .ReportSource = report
'        .ViewReport
'        Do While .IsBusy
'            DoEvents
'        Loop
'        .Zoom 94
        If Trim(pdf) <> "" Then
            log "frmReportSinVisor - pdf : " & pdf
            report.ExportOptions.DestinationType = crEDTDiskFile
            report.ExportOptions.FormatType = crEFTPortableDocFormat
            report.ExportOptions.DiskFileName = pdf
            report.ExportOptions.PDFExportAllPages = True
            report.Export False
'            Do While .IsBusy
'                DoEvents
'            Loop
        Else
            If imprimir = True Then
                report.PrintOut False, 1
            End If
        End If
'    End With
    limpiar_variables
    log "frmReportSinVisor - Fin"
    Set crystal = Nothing
    Set report = Nothing
    Me.MousePointer = 0
    generar = True
   On Error GoTo 0
   Exit Function

generar_Error:
    Me.MousePointer = 0
    Set crystal = Nothing
    Set report = Nothing
    Dim error As String
    error = error & "ERROR Número : " & Err.Number & vbNewLine
    error = error & "ERROR Description : " & Err.Description & vbNewLine
    error = error & "Informe : " & informe & vbNewLine
    error = error & "Criterio : " & criterio & vbNewLine
    error = error & "Pdf : " & pdf & vbNewLine
    error = error & "Consulta : " & consulta & vbNewLine
   
'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar of Formulario frmReportSinVisor"
    log "frmReportSinVisor - ERROR!!!"
    log "frmReportSinVisor : " & error
    enviar_informe_error 0, error
    limpiar_variables
    generar = False
End Function
Public Sub limpiar_variables()
    criterio = ""
    consulta = ""
    pdf = ""
    informe = ""
End Sub

