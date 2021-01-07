Attribute VB_Name = "Informes"
Public Function informe_generico(ByVal ID As Long, copias As Integer) As Boolean
    On Error GoTo fallo
    Dim appword As Word.Application
    Dim docword As Word.Document
    ' Crear copia para su uso
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(copiar_plantilla("generico"))
    appword.Visible = False
    appword.WindowState = wdWindowStateMinimize
    ' Ensayo
    Dim oDOCUMENTO As New clsDocumentos
    If oDOCUMENTO.Carga(ID) = True Then
        Dim ocliente As New clsCliente
'        ocliente.CargaCliente (oDOCUMENTO.getCLIENTE_ID)
        Dim oFp As New clsForma_pago
        oFp.Cargar (ocliente.getFORMA_PAGO)
        With docword.Tables(2)
            .Rows(1).Cells(3).Range.Text = ocliente.getNOMBRE
            .Rows(2).Cells(3).Range.Text = ocliente.getDIRECCION
            ' LP005
'            .Rows(3).Cells(3).Range.Text = ocliente.getCP & " " & ocliente.getPROVINCIA
            Dim oProvincia As New clsProvincias
            oProvincia.Carga ocliente.getPROVINCIA_ID
            .Rows(3).Cells(3).Range.Text = ocliente.getCP & " " & oProvincia.getNOMBRE
            .Rows(3).Cells(2).Range.Text = oDOCUMENTO.getFECHA
            .Rows(4).Cells(2).Range.Text = oDOCUMENTO.getNUMERO & "/" & oDOCUMENTO.getANNO
'            .Rows(5).Cells(2).Range.Text = oObra.getDESCRIPCION
            .Rows(6).Cells(2).Range.Text = oFp.getNOMBRE
            .Rows(7).Cells(2).Range.Text = ocliente.getCIF
        End With
        'Detalle
        Dim oDocumento_Detalle As New clsDocumentos_detalle
        Dim rs As ADODB.Recordset
        Set rs = oDocumento_Detalle.Detalle_Documento(ID)
        Dim fila As Integer
        fila = 2
        With docword.Tables(3)
            If rs.RecordCount > 0 Then
                Do
                    If fila = 2 And InStr(1, rs(1), ReadINI(App.Path & "\config.ini", "parametros", "certificacion")) = 0 Then
                        .Rows(fila).Cells(2).Range.Paragraphs.Alignment = wdAlignParagraphLeft
                    End If
                    .Rows(fila).Cells(1).Range.Text = rs(0)
                    .Rows(fila).Cells(2).Range.Text = rs(1)
                    .Rows(fila).Cells(3).Range.Text = rs(2)
                    .Rows(fila).Cells(4).Range.Text = rs(3)
                    fila = fila + 1
                    rs.MoveNext
                Loop Until rs.EOF
            End If
        End With
        ' Total
        With docword.Tables(3)
            .Rows(32).Cells(4).Range.Text = oDOCUMENTO.getTOTAL
        End With
    End If
    docword.Save
    appword.Quit
    If ReadINI(App.Path & "\config.ini", "parametros", "imprimir") = 0 Then
        ver_documento_word "generico"
    Else
        imprimir_word "generico", copias
    End If
    Set docword = Nothing
    Set appword = Nothing
    informe_generico = True
    Exit Function
fallo:
    MsgBox "Error al generar el documento : " & Err.Description, vbCritical, Err.Description
    appword.Quit 0
    informe_generico = False
    Set docword = Nothing
    Set appword = Nothing
End Function

