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
Public criterio As String
Public imprimir As Boolean
Public pdf As String
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
        .Top = 0
        .Left = 0
        .Height = ScaleHeight
        .Width = ScaleWidth
    End With
End Sub

Public Sub generar()
    Dim crystal As CRAXDRT.Application
    Dim report As CRAXDRT.report
    Dim total_parametros As Integer, intCont As Integer
   On Error GoTo generar_Error
    Me.MousePointer = 11
    Set crystal = New CRAXDRT.Application
    Set report = crystal.OpenReport(ReadINI(App.Path + "\config.ini", "documentos", "reportes") & "\" & informe & ".rpt")
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
        Next intCont
    End If
    report.FormulaSyntax = crCrystalSyntaxFormula
    report.RecordSelectionFormula = criterio

    If consulta <> "" Then
        Dim RS As ADODB.Recordset
        Set RS = datos_bd(consulta)
        report.database.SetDataSource RS
'    Else
'        If UCase(USUARIO.getUSUARIO) = "PRUEBA" Then
'            report.database.LogOnServerEx "Pdsmon.dll", "Geslab", ReadINI(App.Path + "\config.ini", "server", "bd_prueba"), "geslab", "ix1tec", "Active Data (ADO)", ""
'        Else
'            report.database.LogOnServerEx "Pdsmon.dll", "Geslab", ReadINI(App.Path + "\config.ini", "server", "bd"), "geslab", "ix1tec", "Active Data (ADO)", ""
'        End If
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
            Do While .IsBusy
                DoEvents
            Loop
        Else
            If imprimir = True Then
                report.PrintOut False, 1
            End If
        End If
    End With
    limpiar_variables
    Set crystal = Nothing
    Set report = Nothing
    Me.MousePointer = 0

   On Error GoTo 0
   Exit Sub

generar_Error:
    Me.MousePointer = 0
    limpiar_variables
    Set crystal = Nothing
    Set report = Nothing
    
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar of Formulario frmReport"
End Sub
Public Sub limpiar_variables()
    criterios = ""
    consulta = ""
    pdf = ""
    informe = ""
End Sub

