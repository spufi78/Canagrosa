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
Const ClaveOwner = "12345678"


Public Function convertir_a_pdfCreator(fichero As String, carpeta_destino As String) As Boolean
'    Dim PDFCreator As Object
    Dim objArgs, ifname, DefaultPrinter, i, AppTitle, Scriptname, ScriptBasename
    Dim c As Long
    Dim maxTime As Long
    Dim sleepTime As Long
    maxTime = 180    ' in seconds
    sleepTime = 250 ' in milliseconds

   On Error GoTo convertir_a_pdfCreator_Error
    convertir_a_pdfCreator = False
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set PDFCreator = New clsPDFCreator
    cStart = PDFCreator.cStart("/NoProcessingAtStartup", True)
    If Not cStart Then
      Set PDFCreator = Nothing
      'Hay más de una instancia. Se puede forzar el arranque con la siguiente línea:
'      PDFCreator.cStart "/NoProcessingAtStartup", True
    End If
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
    
      ReadyState = 0
'      .cOption("AutosaveDirectory") = fso.GetParentFolderName(ifname)
      .cOption("AutosaveDirectory") = carpeta_destino
      .cOption("AutosaveFilename") = fso.GetBaseName(ifname)
      .cPrintFile CStr(ifname)
    
      c = 0
      Do While (ReadyState = 0) And (c < (maxTime * 1000 / sleepTime))
       c = c + 1
       Espera sleepTime / 1000
      Loop
      If ReadyState = 0 Then
       log "Converting: " & ifname & vbCrLf & vbCrLf & "An error is occured: Time is up!"
       enviar_informe_error 0, "Generar PDF Excedido de tiempo : " & fichero

       Exit Function
      End If
     End With
'    Next
    
    With PDFCreator
'     .cDefaultPrinter = DefaultPrinter
     .cClearCache
     Espera 0.5
     .cClose
    End With
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

Public Function protegerPdf(fichero As String) As Boolean
   On Error GoTo protegerPdf_Error

    log "Entrada protegerPdf"
    log "fichero : " & fichero
    Set fso = CreateObject("Scripting.FileSystemObject")
    fname = fichero
    ' Verificamos que sea un pdf
    If UCase(fso.GetExtensionName(fname)) <> "PDF" Then
      log "Error el documento no es un pdf. Doc: " & fname
      protegerPdf = False
      Exit Function
    End If
    On Error Resume Next
    ' Verificamos que este instalado el pdfforge
    Set pdfforgePDF = CreateObject("pdfforge.pdf.PDF")
    If Err.Number = 429 Then
      log "The pdfforge.dll coming with PDFCreator is not installed! A possible reason can be a missing Microsoft .Net 1.1!"
      protegerPdf = False
      Exit Function
    End If
    If Err.Number <> 0 Then
      log "Error PDFFORME : " & Err.Number & Err.Description
      protegerPdf = False
      Exit Function
    End If
   
'    If proteger_documento Then
        log "Inicio Protegiendo documento..."
        ' Añadir opcion de no impresión
        Set pdfforgePDFEncryptor = CreateObject("pdfforge.pdf.PDFEncryptor")
        With pdfforgePDFEncryptor
            .AllowAssembly = False
            .AllowCopy = False
            .AllowPrinting = False 'Impresion
            .AllowPrintingHighResolution = False
'            .AllowDegradedPrinting = False
            .AllowFillIn = False
            .AllowModifyAnnotations = False
            .AllowModifyContents = False
            .AllowScreenReaders = False
            .UserPassword = ""
'            .EncryptionMethod = 2
            .OwnerPassword = ClaveOwner
        End With
        tfname = fso.GetTempName
        log "FNAME : " & fname
        log "TFNAME : " & tfname
        Dim res As Long
        res = pdfforgePDF.EncryptPDFFile(fname, tfname, (pdfforgePDFEncryptor))
        If res = 0 And FileLen(tfname) > 0 Then
            If fso.FileExists(fname) Then
               fso.DeleteFile (fname)
            End If
            fso.MoveFile tfname, fname
        Else
            log "Error al proteger el PDF: " & res
            fso.DeleteFile (tfname)
        End If
        log "Fin de proteger documento..."
'    End If
    Set pdfforgePDFEncryptor = Nothing
    Set objArgs = Nothing
    Set pdfforgePDF = Nothing
    Set fso = Nothing
    protegerPdf = True
    log "Fin protegerPdf"

   On Error GoTo 0
   Exit Function

protegerPdf_Error:
    protegerPdf = False
    log "Error al firmar/proteger fichero: " & fichero & vbNewLine & "PDFCREATOR : Error " & Err.Number & " (" & Err.Description & ") in procedure firmar_documento protegerPdf of Módulo clsCrearPDF"
    enviar_informe_error 0, "Error al firmar/proteger el fichero : " & fichero & " -- PDFCREATOR : Error " & Err.Number & " (" & Err.Description & ") in procedure protegerPdf of Módulo clsCrearPDF"
    
End Function


