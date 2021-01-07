VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCrearPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private objArgs As Collection, fname, tfname, fso, WshShell, oExec, pdfforgePDF, pdfforgePDFEncryptor
Public ReadyState
Private WithEvents PDFCreator As PDFCreator.clsPDFCreator
Attribute PDFCreator.VB_VarHelpID = -1

Public Function convertir_a_pdfCreator(fichero As String, carpeta_destino As String) As Boolean
'    Dim PDFCreator As Object
    Dim objArgs, ifname, DefaultPrinter, I, AppTitle, Scriptname, ScriptBasename
    Dim c As Long
    Dim maxTime As Long
    Dim sleepTime As Long
    maxTime = 180    ' in seconds
    sleepTime = 250 ' in milliseconds
     log "convertir_a_pdfCreator_1"
   On Error GoTo convertir_a_pdfCreator_Error
    convertir_a_pdfCreator = False
    Set fso = CreateObject("Scripting.FileSystemObject")
     log "convertir_a_pdfCreator_2"
    Set PDFCreator = New clsPDFCreator
     log "convertir_a_pdfCreator_3"
    cStart = PDFCreator.cStart("/NoProcessingAtStartup", True)
     log "convertir_a_pdfCreator_4"
    If Not cStart Then
      Set PDFCreator = Nothing
      'Hay más de una instancia. Se puede forzar el arranque con la siguiente línea:
'      PDFCreator.cStart "/NoProcessingAtStartup", True
    End If
     log "convertir_a_pdfCreator_5"
    With PDFCreator
        .cVisible = True
         .cOption("UseAutosave") = 1
         .cOption("UseAutosaveDirectory") = 1
         .cOption("AutosaveFormat") = 0                              ' 0 = PDF
    '     DefaultPrinter = .cDefaultPrinter
         .cDefaultPrinter = "PDFCreator"
         .cClearCache
         .cPrinterStop = False
    End With
     log "convertir_a_pdfCreator_6"
    
'    For i = 0 To objArgs.Count - 1
     With PDFCreator
      ifname = fichero
      If Not fso.FileExists(ifname) Then
       MsgBox "Can't find the file: " & ifname, vbExclamation + vbSystemModal, AppTitle
       Exit Function
      End If
      If Not .cIsPrintable(CStr(ifname)) Then
       MsgBox "Converting: " & ifname & vbCrLf & vbCrLf & _
        "An error is occured: File is not printable!", vbExclamation + vbSystemModal, AppTitle
       Exit Function
      End If
     log "convertir_a_pdfCreator_7"
    
      ReadyState = 0
'      .cOption("AutosaveDirectory") = fso.GetParentFolderName(ifname)
      .cOption("AutosaveDirectory") = carpeta_destino
      .cOption("AutosaveFilename") = fso.GetBaseName(ifname)
      .cPrintFile CStr(ifname)
      log "convertir_a_pdfCreator_8"
  
      c = 0
      Do While (ReadyState = 0) And (c < (maxTime * 1000 / sleepTime))
       c = c + 1
       Espera sleepTime / 1000
      Loop
     log "convertir_a_pdfCreator_9"
      If ReadyState = 0 Then
       log "Converting: " & ifname & vbCrLf & vbCrLf & "An error is occured: Time is up!"
       enviar_informe_error 0, "Generar PDF Excedido de tiempo : " & fichero

       Exit Function
      End If
     End With
'    Next
    
     log "convertir_a_pdfCreator_10"
    With PDFCreator
'     .cDefaultPrinter = DefaultPrinter
     .cClearCache
     Espera 0.5
     .cClose
    End With
     log "convertir_a_pdfCreator_11"
    convertir_a_pdfCreator = True
    Set PDFCreator = Nothing
    
   On Error GoTo 0
   Exit Function

convertir_a_pdfCreator_Error:
    convertir_a_pdfCreator = False
    log "Error " & Err.Number & " (" & Err.Description & ") in procedure convertir_a_pdfCreator of Módulo de clase clsCrearPDF"
    enviar_informe_error 0, "Error al convertir PDFCREATOR : " & "Error " & Err.Number & " (" & Err.Description & ") in procedure convertir_a_pdfCreator of Módulo de clase clsCrearPDF"
End Function
Private Sub PDFCreator_eReady()
 ReadyState = 1
End Sub
Private Sub PDFCreator_eError()
 log "PDFCREATOR: An error is occured!" & vbCrLf & vbCrLf & "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFCreator.cErrorDetail("Description")
 enviar_informe_error 0, "Error al convertir PDFCREATOR : " & "PDFCREATOR: An error is occured!" & vbCrLf & vbCrLf & "Error [" & PDFCreator.cErrorDetail("Number") & "]: " & PDFCreator.cErrorDetail("Description")

End Sub


