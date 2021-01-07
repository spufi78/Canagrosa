Attribute VB_Name = "Calidad"
Public Function calidad_nombre_documento_pdf(DOC As Long) As String
    Dim documento As String
    Dim oDeco As New clsDecodificadora
    Dim oPNT As New clsCa_documentos
    Dim EXTENSION As String
    oPNT.Carga DOC
    oDeco.Carga_valor decodificadora.CALIDAD_PLANTILLAS_DOCUMENTOS, oPNT.getPLANTILLA_ID
    
    Dim s() As String
    s = Split(oDeco.getPARAMETROS, ".")
    EXTENSION = "." & s(1)
'    EXTENSION = Right(oDeco.getPARAMETROS, 4)
    ' Nombre del documento
    If UCase(EXTENSION) = ".XLS" Then
        calidad_nombre_documento_pdf = Eliminar_Caracteres_Archivo(Replace(Trim(oPNT.getCODIGO), ".", " ")) & EXTENSION
    Else
        calidad_nombre_documento_pdf = Eliminar_Caracteres_Archivo(Replace(Trim(oPNT.getCODIGO), ".", " ")) & ".pdf"
    End If
    Set oDeco = Nothing
    Set oPNT = Nothing
End Function
Public Function calidad_nombre_documento_trabajo(DOC As Long) As String
    Dim documento As String
    Dim oDeco As New clsDecodificadora
    Dim oPNT As New clsCa_documentos
    Dim EXTENSION As String
    oPNT.Carga DOC
    oDeco.Carga_valor decodificadora.CALIDAD_PLANTILLAS_DOCUMENTOS, oPNT.getPLANTILLA_ID
    Dim s() As String
    s = Split(oDeco.getPARAMETROS, ".")
    EXTENSION = "." & s(1)
'    EXTENSION = Right(oDeco.getPARAMETROS, 4)
    calidad_nombre_documento_trabajo = Eliminar_Caracteres_Archivo(Replace(Trim(oPNT.getCODIGO), ".", " ")) & EXTENSION
    Set oDeco = Nothing
    Set oPNT = Nothing
End Function
Public Function calidad_ruta_documento_trabajo(DOC As Long) As String
    Dim documento As String
    documento = calidad_ruta_trabajo(DOC) & "\" & calidad_nombre_documento_trabajo(DOC)
    calidad_ruta_documento_trabajo = documento
End Function
Public Function calidad_ruta_trabajo(DOC As Long) As String
    Dim oPNT As New clsCa_documentos
    oPNT.Carga DOC
    calidad_ruta_trabajo = calidad_ruta_trabajo_por_familia(oPNT.getFAMILIA_ID)
    On Error Resume Next
    MkDir calidad_ruta_trabajo
    Set oDeco = Nothing
    Set oPNT = Nothing
End Function
Public Function calidad_ruta_pdf(DOC As Long) As String
    Dim oPNT As New clsCa_documentos
    oPNT.Carga DOC
    calidad_ruta_pdf = calidad_ruta_pdf_por_familia(oPNT.getFAMILIA_ID)
    On Error Resume Next
    MkDir calidad_ruta_pdf
    Set oPNT = Nothing
End Function
Private Function calidad_ruta_trabajo_por_familia(familia As Long) As String
    Dim oDeco As New clsDecodificadora
    oDeco.Carga_valor decodificadora.CA_DOCUMENTOS_FAMILIAS, familia
    calidad_ruta_trabajo_por_familia = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Documentos\Trabajo\" & oDeco.getDESCRIPCION
    On Error Resume Next
    MkDir calidad_ruta_trabajo_por_familia
    Set oDeco = Nothing
End Function
Private Function calidad_ruta_pdf_por_familia(familia As Long) As String
    Dim oDeco As New clsDecodificadora
    oDeco.Carga_valor decodificadora.CA_DOCUMENTOS_FAMILIAS, familia
    calidad_ruta_pdf_por_familia = ReadINI(App.Path + "\config.ini", "Documentos", "Ruta") & "\Calidad\Documentos\PDF\" & oDeco.getDESCRIPCION
    On Error Resume Next
    MkDir calidad_ruta_pdf_por_familia
    Set oDeco = Nothing
End Function

