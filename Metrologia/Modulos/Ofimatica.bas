Attribute VB_Name = "Ofimatica"
Public Function imprimir_word(doc As String, copias As Integer) As Boolean
    On Error GoTo fallo
    Dim appword As Word.Application
    Dim docword As Word.Document
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(ReadINI(App.path + "\config.ini", "Documentos", "Documentos") & "\" & doc & ".doc")
    appword.Documents(1).PrintOut Background, , , , , , , copias, , , , , , 0
    Do While appword.BackgroundPrintingStatus = 1
    Loop
    appword.Documents.Close (wdDotNotSaveChanges)
    appword.Quit
    Set docword = Nothing
    Set appword = Nothing
    imprimir_word = True
    Exit Function
fallo:
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
    imprimir_word = False
End Function
Public Sub ver_documento_word(ByVal doc As String)
    Dim appword As Word.Application
    Dim docword As Word.Document
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(ReadINI(App.path + "\config.ini", "Documentos", "Documentos") & "\" & doc & ".doc")
    appword.Visible = True
    Set docword = Nothing
    Set appword = Nothing
End Sub
Public Function copiar_plantilla(plantilla As String) As String
    ' Crear copia de la plantilla para su uso
    On Error Resume Next
    Dim ORIGEN As String
    Dim destino As String
    ORIGEN = ReadINI(App.path + "\config.ini", "Documentos", "Plantillas") & "\" & plantilla & ".doc"
    destino = ReadINI(App.path + "\config.ini", "Documentos", "Documentos") & "\" & plantilla & ".doc"
    FileCopy ORIGEN, destino
    copiar_plantilla = destino
End Function
