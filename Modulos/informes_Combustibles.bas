Attribute VB_Name = "informes_Combustibles"
Public Function imprimir_informe_combustible(ByVal muestra As Long, por_impresora As Integer, fecha_impresion As Date) As Boolean
    On Error GoTo fallo
    Dim appword As Word.Application
    Dim docword As Word.Document
    ' Crear copia para su uso
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(copiar_plantilla(docCOMBUSTIBLE, muestra, por_impresora))
    appword.Visible = False
    appword.WindowState = wdWindowStateMinimize
    ' Recuperar Cabecera
    Dim oMuestra As New clsMuestra
    Dim rs As ADODB.Recordset
    Set rs = oMuestra.datos_cabecera_documento(muestra)
    cabecera_bano docword, rs, "CO"
    ' Datos de cabecera
    Dim ovalores_bano As New clsDatos_valores
    Dim rs_dc As ADODB.Recordset
    Dim snivel As String
    Set rs_dc = oMuestra.datos_cabecera_bano(muestra)
    If rs_dc.RecordCount > 0 Then
      With docword.Tables(1)
        .Rows(4).Cells(2).Range.Text = (rs_dc(0))
        .Rows(5).Cells(2).Range.Text = (rs_dc(1))
        .Rows(6).Cells(2).Range.Text = (rs_dc(8)) ' Referencia del cliente
        oMuestra.CargaMuestra muestra
        If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.COMBUSTIBLE Then
            .Rows(6).Cells(1).Range.Text = "SISTEMA (SYSTEM):"
            .Rows(6).Cells(2).Range.Text = (rs_dc(5)) ' Sistema
        End If
        .Rows(7).Cells(2).Range.Text = (rs_dc(2))
      End With
    End If
    ' Datos especificos fluidos (SACA TODOS LOS DATOS ESPECIFICOS)
    Dim rs_de As ADODB.Recordset
    Set rs_de = ovalores_bano.datos_especiales(muestra)
    Dim linea As Integer
    linea = 11
    If rs_de.RecordCount > 0 Then
        Do
            If Trim(rs_de(1)) <> "" And rs_de(2) <> 1 Then
                docword.Tables(1).Rows(linea).Cells(1).Range.InsertAfter rs_de(0) & ": " & rs_de(1)
                linea = linea + 1
                rs_de.MoveNext
                If rs_de.EOF = False Then
                    docword.Tables(1).Rows.Add
                End If
            Else
                rs_de.MoveNext
            End If
        Loop Until rs_de.EOF
    End If
    ' No hay datos especificos
    If linea = 11 Then
        docword.Tables(1).Rows(10).Delete
    End If
    ' Linea, bano, solucion
    With docword.Tables(2)
        .Rows(2).Cells(1).Range.Text = rs_dc(0)
        If Trim(rs(5)) <> "" Then
            .Rows(2).Cells(1).Range.InsertAfter "/" & rs(5)
        End If
        .Rows(2).Cells(2).Range.Text = rs_dc(5)
        log ("5")
        
        If Trim(rs_dc(6)) <> "" Then
            .Rows(2).Cells(2).Range.InsertAfter vbCrLf & vbCrLf & vbCrLf
            .Rows(2).Cells(2).Range.Paragraphs.Add
            .Rows(2).Cells(2).Range.InsertAfter "Volumen = " & rs_dc(6)
        End If
        .Rows(2).Cells(3).Range.Text = rs_dc(3)
        .Rows(2).Cells(2).Range.Paragraphs(1).Range.Bold = True
    End With
    
    ' Determinaciones
    Dim odeterminaciones As New clsDeterminaciones
    Dim rs_deter As ADODB.Recordset
    Dim rango As String
    Dim i As Integer
    Dim nsd As Boolean
    nsd = False
   
    Set rs_deter = odeterminaciones.lista_determinaciones_bano(muestra)
    If rs_deter.RecordCount > 0 Then
        With docword.Tables(2)
            i = 1
            Do
                With .Rows.Last
                    .Cells(4).Range.Bold = False
                    .Cells(4).Range.Text = rs_deter(0)
                    .Cells(4).Range.Paragraphs.Add
                    .Cells(4).Range.Paragraphs(1).Range.Bold = True
                    .Cells(4).Range.InsertAfter rs_deter(1) & vbCrLf
                     rango = ""
                    If Trim(rs_deter(14)) <> "" And Trim(rs_deter(15)) <> "" Then
                        rango = rs_deter(14) & " - " & rs_deter(15)
                    Else
                        rango = Trim(rs_deter(14)) & Trim(rs_deter(15))
                    End If
                    
                    If Trim(rango) <> "" Then
                        rango = rango & " " & rs_deter(7)
                        .Cells(4).Range.InsertAfter rango & vbCrLf
                    End If
                    If IsNumeric(rs_deter(6)) = False Then
                        If rs_deter(6) = "--" Then
                            .Cells(5).Range.Text = rs_deter(6)
                        Else
                            .Cells(5).Range.Text = rs_deter(6) & " " & rs_deter(7)
                        End If
                        .Cells(5).Range.Underline = wdUnderlineNone
                    Else
                        If (CSng(rs_deter(6)) = 0) And (rs_deter(3) = 0) Then
                            .Cells(5).Range.Text = "n.s.d."
                            nsd = True
                        Else
                            .Cells(5).Range.Underline = wdUnderlineNone
                            If Trim(rs_deter(10)) <> "" And IsNumeric(rs_deter(10)) Then
                                If CSng(Replace(rs_deter(6), ".", ",")) < CSng(Replace(rs_deter(10), ".", ",")) Then
                                    .Cells(5).Range.Underline = wdUnderlineSingle
                                End If
                            End If
                            If Trim(rs_deter(11)) <> "" And IsNumeric(rs_deter(11)) Then
                                If CSng(Replace(rs_deter(6), ".", ",")) > CSng(Replace(rs_deter(11), ".", ",")) Then
                                    .Cells(5).Range.Underline = wdUnderlineSingle
                                End If
                            End If
                            .Cells(5).Range.Text = rs_deter(6) & " " & rs_deter(7)
                        End If
                    End If
                    .Cells(6).Range.Text = rs_deter(2)
                    .Cells(7).Range.Text = Format(rs_deter(5), "dd/mm/yy")
                End With
                rs_deter.MoveNext
                If Not rs_deter.EOF Then
                    .Rows.Add
                    i = i + 1
                End If
            Loop Until rs_deter.EOF
            Dim j As Integer
            Dim k As Integer
            ' Merge de las 3 primeras columnas
            For k = 2 To i
                For j = 1 To 3
                    .Cell(2, j).Merge .Cell(k + 1, j)
                Next
            Next
            ' Eliminar bordes del resto de columnas
            For k = 2 To i
                For j = 4 To 7
                    .Cell(k, j).Borders(wdBorderBottom).Visible = False
                Next
            Next
        End With
    End If
    ' NSD
    If nsd = True Then
        docword.Tables(4).Rows(1).Cells(1).Range.Text = "n.s.d. : no se detecta"
    End If
    ' Pie
    pie_banos appword, docword, rs, muestra, fecha_impresion, por_impresora
    Set docword = Nothing
    Set appword = Nothing
    imprimir_informe_combustible = True
    Exit Function
fallo:
    log ("***** Error al generar el documento de COMBUSTIBLE : " & Err.Description)
    enviar_informe_error muestra, "Imprimir_informe_combustible : " & Err.Description
    imprimir_informe_combustible = False
    appword.Documents.Close (wdDotNotSaveChanges)
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
End Function
