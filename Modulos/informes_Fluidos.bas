Attribute VB_Name = "informes_Fluidos"
Public Function imprimir_informe_fluido(ByVal MUESTRA As Long, por_impresora As Integer, fecha_impresion As Date) As Boolean
    On Error GoTo fallo
    Dim appword As Word.Application
    Dim docword As Word.Document
    ' Crear copia para su uso
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(copiar_plantilla(docFLUIDO, MUESTRA, por_impresora))
    appword.Visible = False
    appword.WindowState = wdWindowStateMinimize
    ' Recuperar Cabecera
    Dim oMuestra As New clsMuestra
    Dim rs As ADODB.Recordset
    Set rs = oMuestra.datos_cabecera_documento(MUESTRA)
    cabecera_bano docword, rs, "FH"
    ' Datos de cabecera
    Dim ovalores_bano As New clsDatos_valores
    Dim rs_dc As ADODB.Recordset
    Dim snivel As String
    Set rs_dc = oMuestra.datos_cabecera_bano(MUESTRA)
    If rs_dc.RecordCount > 0 Then
      With docword.Tables(1)
        .Rows(4).Cells(2).Range.Text = (rs_dc(0))
        .Rows(5).Cells(2).Range.Text = (rs_dc(1))
        .Rows(6).Cells(2).Range.Text = (rs_dc(8)) ' Referencia del cliente
' J003-I
        oMuestra.CargaMuestra MUESTRA
        If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.COMBUSTIBLE Then
            .Rows(6).Cells(1).Range.Text = "SISTEMA (SYSTEM):"
            .Rows(6).Cells(2).Range.Text = (rs_dc(5)) ' Sistema
        End If
' J003-F
        .Rows(7).Cells(2).Range.Text = (rs_dc(2))
      End With
    End If
    ' Datos especificos fluidos (SACA TODOS LOS DATOS ESPECIFICOS)
    Dim rs_de As ADODB.Recordset
    Set rs_de = ovalores_bano.datos_especiales(MUESTRA)
    Dim linea As Integer
    linea = 11
    If rs_de.RecordCount > 0 Then
        Do
            If Trim(rs_de(1)) <> "" And rs_de(2) <> 1 Then
                docword.Tables(1).Rows(linea).Cells(1).Range.InsertAfter rs_de(0) & ": " & rs_de(1)
'                docword.Tables(1).Rows(linea).Cells(2).Range.InsertAfter
                linea = linea + 1
                rs_de.MoveNext
                If rs_de.EOF = False Then
    '                docword.Tables(1).Rows(linea).Borders(wdBorderBottom).Visible = False
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
   
    Set rs_deter = odeterminaciones.lista_determinaciones_bano(MUESTRA)
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
'                    If Trim(rs_deter(10)) <> "" And Trim(rs_deter(11)) <> "" Then
'                        rango = rs_deter(10) & " - " & rs_deter(11)
'                    End If
'                    If Trim(rs_deter(10)) <> "" And Trim(rs_deter(11)) = "" Then
'                        rango = " > " & rs_deter(10)
'                    End If
'                    If Trim(rs_deter(10)) = "" And Trim(rs_deter(11)) <> "" Then
'                        rango = " < " & rs_deter(11)
'                    End If
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
    ' Particulas
    Dim oFR As New clsFluidos_resultados
    Dim rs_particulas As ADODB.Recordset
    Set rs_particulas = oFR.Listado(MUESTRA)
    If rs_particulas.RecordCount = 0 Then
        docword.Tables(3).Delete
    Else
        Do
            With docword.Tables(3)
                .Rows(rs_particulas("TAMANO") + 1).Cells(3).Range.Text = rs_particulas("RESULTADO")
                .Rows(rs_particulas("TAMANO") + 1).Cells(4).Range.Text = rs_particulas("CLASIFICACION")
            End With
            rs_particulas.MoveNext
        Loop Until rs_particulas.EOF
        oMuestra.CargaMuestra (MUESTRA)
        Dim oFluido As New clsFluidos_ficha
        oFluido.Carga_por_BANO (oMuestra.getBANO_ID)
        Dim oFN As New clsFluidos_normas
        oFN.Carga (oFluido.getNORMA_ID)
        docword.Tables(3).Rows(2).Cells(5).Range.Text = oFN.getNOMBRE
        ' Merge de parametro y norma
        docword.Tables(3).Cell(2, 1).Merge docword.Tables(3).Cell(6, 1)
        docword.Tables(3).Cell(2, 5).Merge docword.Tables(3).Cell(6, 5)
    End If
    ' Pie
    pie_banos appword, docword, rs, MUESTRA, fecha_impresion, por_impresora
    Set docword = Nothing
    Set appword = Nothing
    imprimir_informe_fluido = True
    Exit Function
fallo:
    log ("***** Error al generar el documento de baño : " & Err.Description)
    enviar_informe_error MUESTRA, "Imprimir_informe_fluido : " & Err.Description
    imprimir_informe_fluido = False
    appword.Documents.Close (wdDotNotSaveChanges)
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
End Function
Public Function imprimir_informe_fluido_columna(ByVal MUESTRA As Long, por_impresora As Integer, fecha_impresion As Date) As Boolean
    On Error GoTo fallo
    Dim appword As Word.Application
    Dim docword As Word.Document
    ' Crear copia para su uso
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(copiar_plantilla("CO-AGRUPADOS", MUESTRA, por_impresora))
    appword.Visible = False
    appword.WindowState = wdWindowStateMinimize
    ' Recuperar Cabecera
    Dim oMuestra As New clsMuestra
    Dim rs As ADODB.Recordset
    Set rs = oMuestra.datos_cabecera_documento(MUESTRA)
    cabecera_bano docword, rs, "AGRUPADA"
    ' Datos de cabecera
    Dim ovalores_bano As New clsDatos_valores
    Dim rs_dc As ADODB.Recordset
    Dim snivel As String
    Set rs_dc = oMuestra.datos_cabecera_bano(MUESTRA)
    If rs_dc.RecordCount > 0 Then
      With docword.Tables(1)
        .Rows(1).Cells(2).Range.Text = "Tomada por " & (rs(15))
        .Rows(2).Cells(2).Range.Text = (rs_dc(0))
        .Rows(3).Cells(2).Range.Text = (rs_dc(1))
        .Rows(4).Cells(2).Range.Text = (rs_dc(2))
      End With
    End If
    appword.Visible = True
    ' Determinaciones agrupadas
    Dim nsd As Boolean
    Dim oFF As New clsFluidos_ficha
    Dim odeterminaciones As New clsDeterminaciones
    Dim rs_deter As ADODB.Recordset
    Dim rs_fluido As ADODB.Recordset
    Set rs_fluido = oFF.Lista_muestras_fluido_columna(MUESTRA)
    If rs_fluido.RecordCount > 0 Then
        ' Sistema
        docword.Tables(2).Rows.Last.Cells(2).Range.Text = rs_dc(5)
        If Trim(rs_dc(6)) <> "" Then
             docword.Tables(2).Rows.Last.Cells(2).Range.InsertAfter vbCrLf & vbCrLf & vbCrLf
             docword.Tables(2).Rows.Last.Cells(2).Range.Paragraphs.Add
             docword.Tables(2).Rows.Last.Cells(2).Range.InsertAfter "Volumen = " & rs_dc(6)
        End If
        ' Matriz
        docword.Tables(2).Rows.Last.Cells(3).Range.Text = rs_dc(3)
        docword.Tables(2).Rows.Last.Cells(2).Range.Paragraphs(1).Range.Bold = True
        Do
            oMuestra.CargaMuestra (rs_fluido("ID_MUESTRA"))
            If (MUESTRA <> rs_fluido("ID_MUESTRA")) Then
                oMuestra.aumentar_edicion_impresa (rs_fluido("ID_MUESTRA"))
            End If
            With docword.Tables(2)
                With .Rows.Last
                    'N.Ensayo
                    If oMuestra.getULT_EDICION_IMP = 0 Then
                        .Cells(1).Range.Text = oMuestra.getID_GENERAL & "/" & Format(oMuestra.getFECHA_RECEPCION, "yyyy") & "/Ed." & oMuestra.getULT_EDICION_IMP + 1
                    Else
                        .Cells(1).Range.Text = oMuestra.getID_GENERAL & "/" & Format(oMuestra.getFECHA_RECEPCION, "yyyy") & "/Ed." & oMuestra.getULT_EDICION_IMP + 1 & " *"
                    End If
                    ' Hora de toma de la muestra
                    .Cells(4).Range.InsertAfter ovalores_bano.HORA(rs_fluido("ID_MUESTRA"))
                    ' Valor teórico
                    .Cells(5).Range.Text = "SIN ESPECIFICAR"
                    ' Valor de la conductividad electrica
                    Set rs_deter = odeterminaciones.lista_determinaciones_bano(rs_fluido("ID_MUESTRA"))
                    If rs_deter.RecordCount <> 0 Then
                        If IsNumeric(rs_deter(6)) = False Then
                            .Cells(6).Range.Text = rs_deter(6)
                            .Cells(6).Range.Underline = wdUnderlineNone
                        Else
                            If (CSng(rs_deter(6)) = 0) And (rs_deter(3) = 0) Then
                                .Cells(6).Range.Text = "n.s.d."
                                nsd = True
                            Else
                                .Cells(6).Range.Underline = wdUnderlineNone
                                If Trim(rs_deter(10)) <> "" And IsNumeric(rs_deter(10)) Then
                                    If CSng(Replace(rs_deter(6), ".", ",")) < CSng(Replace(rs_deter(10), ".", ",")) Then
                                        .Cells(6).Range.Underline = wdUnderlineSingle
                                    End If
                                End If
                                If Trim(rs_deter(11)) <> "" And IsNumeric(rs_deter(11)) Then
                                    If CSng(Replace(rs_deter(6), ".", ",")) > CSng(Replace(rs_deter(11), ".", ",")) Then
                                        .Cells(6).Range.Underline = wdUnderlineSingle
                                    End If
                                End If
                                .Cells(6).Range.Text = rs_deter(6) ' & " " & rs_deter(7)
                            End If
                        End If
                    End If
                End With
            End With
            rs_fluido.MoveNext
            If Not rs_fluido.EOF Then
                docword.Tables(2).Rows.Add
            End If
        Loop Until rs_fluido.EOF
        ' Merge de sistema y matriz
        docword.Tables(2).Cell(3, 2).Merge docword.Tables(2).Cell(docword.Tables(2).Rows.Count, 2)
        docword.Tables(2).Cell(3, 3).Merge docword.Tables(2).Cell(docword.Tables(2).Rows.Count, 3)
        
    End If
    ' Pie
    If nsd = True Then
        docword.Tables(3).Rows(1).Cells(1).Range.Text = "n.s.d. : no se detecta"
    End If
    ' Pie
    pie_banos appword, docword, rs, MUESTRA, fecha_impresion, por_impresora
    Set docword = Nothing
    Set appword = Nothing
    imprimir_informe_fluido_columna = True
    Exit Function
fallo:
    log ("***** Error al generar el documento de aguas : " & Err.Description)
    enviar_informe_error MUESTRA, "Imprimir_informe_fluido_columna : " & Err.Description
    imprimir_informe_fluido_columna = False
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
End Function
Public Function imprimir_informe_fluido_morado(ByVal MUESTRA As Long, por_impresora As Integer, fecha_impresion As Date) As Boolean
    On Error GoTo fallo
    Dim appword As Word.Application
    Dim docword As Word.Document
    ' Crear copia para su uso
    Set appword = CreateObject("word.application")
    Set docword = appword.Documents.Open(copiar_plantilla(docFLUIDO_MORADO, MUESTRA, por_impresora))
    appword.Visible = False
    appword.WindowState = wdWindowStateMinimize
    ' Recuperar Cabecera
    Dim oMuestra As New clsMuestra
    fluidos_cabecera MUESTRA, docword
    fluidos_datos_especificos MUESTRA, docword
    fluidos_determinaciones MUESTRA, docword
    fluidos_particulas MUESTRA, docword
    Dim rs As ADODB.Recordset
    Set rs = oMuestra.datos_cabecera_documento(MUESTRA)
    pie_fluidos appword, docword, rs, MUESTRA, fecha_impresion, por_impresora
    Set docword = Nothing
    Set appword = Nothing
    imprimir_informe_fluido_morado = True
    Exit Function
fallo:
    log ("***** Error al generar el documento imprimir_informe_fluido_morado : " & Err.Description)
'    enviar_informe_error muestra, "imprimir_informe_fluido_morado : " & Err.Description
    imprimir_informe_fluido_morado = False
    appword.Documents.Close (wdDotNotSaveChanges)
    appword.Quit 0
    Set docword = Nothing
    Set appword = Nothing
End Function
Public Sub fluidos_cabecera(MUESTRA As Long, docword As Word.Document)
    Dim oMuestra As New clsMuestra
    Dim rs As ADODB.Recordset
    Set rs = oMuestra.datos_cabecera_documento(MUESTRA)
    ' Datos de Cliente y Fechas de muestreo
    With docword.Sections(1).Headers(1).Range.Tables(1)
        .Rows(3).Cells(1).Range.Text = rs(0)
        .Rows(4).Cells(1).Range.Text = rs(1)
        If rs(3) = "" Then
            .Rows(5).Cells(1).Range.Text = rs(2) & " " & rs(3)
        Else
            .Rows(5).Cells(1).Range.Text = rs(2) & " " & rs(3) & " (" & Trim(rs(4)) & ")"
        End If
        .Rows(6).Cells(1).Range.Text = rs(5)
        If Trim(rs(6)) <> "" Then
            .Rows(7).Cells(1).Range.Text = "A/A de " & rs(6)
        End If
        .Rows(3).Cells(3).Range.Text = rs(8) ' Fecha muestreo
        .Rows(4).Cells(3).Range.Text = rs(7) ' Fecha recepcion
        If Not IsNull(rs(30)) Then
            .Rows(5).Cells(3).Range.Text = rs(30) ' Fecha comiento
        End If
        If Not IsNull(rs(9)) Then
            .Rows(6).Cells(3).Range.Text = rs(9) ' Fecha cierre
        End If
    End With
    ' Número y edición del Fluido
    With docword.Sections(1).Headers(1).Range.Tables(2)
        .Rows(1).Cells(1).Range.InsertAfter (rs(17) & "/" & Format(rs(7), "yyyy") & "/Edición " & rs(10) + 1)
        Dim mensaje_edicion As String
        If rs(10) = 0 Then
            mensaje_edicion = ""
        Else
            mensaje_edicion = ReadINI(App.Path + "\config.ini", "edicion", "mensaje") & "-" & ReadINI(App.Path + "\config.ini", "edicion", "ingles")
        End If
        .Rows(2).Cells(1).Range.Text = mensaje_edicion
        .Rows(4).Cells(2).Range.Text = UCase("TOMADA POR " & (rs(15)))
    End With
    ' Resto de datos del fluido
    Dim ovalores_bano As New clsDatos_valores
    Dim rs_dc As ADODB.Recordset
    Set rs_dc = oMuestra.datos_cabecera_bano(MUESTRA)
    If rs_dc.RecordCount > 0 Then
      With docword.Sections(1).Headers(1).Range.Tables(2)
        ' Ref. cliente, si es combustible se cambia por sistema
        .Rows(3).Cells(2).Range.Text = (rs_dc(8))
        oMuestra.CargaMuestra MUESTRA
        If oMuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.COMBUSTIBLE Then
            .Rows(3).Cells(1).Range.Text = "SISTEMA (SYSTEM):"
            .Rows(3).Cells(2).Range.Text = (rs_dc(5)) ' Sistema
        End If
        .Rows(5).Cells(2).Range.Text = (rs_dc(0)) ' Estación
        .Rows(6).Cells(2).Range.Text = (rs_dc(1)) ' Especificación de control
        Dim oFluido As New clsFluidos_ficha
        If oFluido.Carga_por_BANO(oMuestra.getBANO_ID) Then
            .Rows(7).Cells(2).Range.Text = oFluido.getNORMATIVA_APLICABLE ' Normativa aplicable
            .Rows(8).Cells(2).Range.Text = oFluido.getNORMATIVA_REFERENCIA ' Normativa referencia
        End If
        .Rows(9).Cells(2).Range.Text = rs_dc(2) ' Cadencia de control
      End With
    End If
    ' Estación, Sistema y Matriz
    With docword.Tables(2)
        .Rows(2).Cells(1).Range.Text = rs_dc(0) ' LINEA
        If Trim(rs(5)) <> "" Then
            .Rows(2).Cells(1).Range.InsertAfter "/" & rs(5) ' CLIENTE -> CENTRO
        End If
        .Rows(2).Cells(2).Range.Text = rs_dc(5) ' BAÑO NOMBRE

        If Trim(rs_dc(6)) <> "" Then
            .Rows(2).Cells(2).Range.InsertAfter vbCrLf & vbCrLf & vbCrLf
            .Rows(2).Cells(2).Range.Paragraphs.Add
            .Rows(2).Cells(2).Range.InsertAfter "Volumen = " & rs_dc(6)
        End If
        .Rows(2).Cells(3).Range.Text = rs_dc(3) ' SOLUCION
        .Rows(2).Cells(2).Range.Paragraphs(1).Range.Bold = True
    End With
End Sub

Public Sub fluidos_datos_especificos(MUESTRA As Long, docword As Word.Document)
    ' Datos especificos fluidos (SACA TODOS LOS DATOS ESPECIFICOS)
    Dim ovalores_bano As New clsDatos_valores
    Dim rs As ADODB.Recordset
    Set rs = ovalores_bano.datos_especiales(MUESTRA)
    Dim linea As Integer
    linea = 2
    If rs.RecordCount > 0 Then
        Do
            If Trim(rs(1)) <> "" And rs(2) <> 1 Then
                docword.Tables(1).Rows(linea).Cells(1).Range.InsertAfter rs(0) & ": " & rs(1)
                linea = linea + 1
                rs.MoveNext
                If rs.EOF = False Then
                    docword.Tables(1).Rows.Add
                End If
            Else
                rs.MoveNext
            End If
        Loop Until rs.EOF
    End If
    ' No hay datos especificos
    If linea = 2 Then
        docword.Tables(1).Rows(1).Delete
    End If
End Sub
Public Sub fluidos_determinaciones(MUESTRA As Long, docword As Word.Document)
    Dim odeterminaciones As New clsDeterminaciones
    Dim rs_deter As ADODB.Recordset
    Dim rango As String
    Dim i As Integer
    Set rs_deter = odeterminaciones.lista_determinaciones_bano(MUESTRA)
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
                        .Cells(4).Range.InsertAfter rango & vbCrLf
                    End If
                    If IsNumeric(rs_deter(6)) = False Then
                        .Cells(5).Range.Text = rs_deter(6)
                        .Cells(5).Range.Underline = wdUnderlineNone
                    Else
                        If (CSng(rs_deter(6)) = 0) And (rs_deter(3) = 0) Then
                            .Cells(5).Range.Text = "n.s.d."
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
'    If nsd = True Then
'        docword.Tables(5).Rows(1).Cells(1).Range.Text = "n.s.d. : no se detecta"
'    End If
End Sub
Public Sub fluidos_particulas(MUESTRA As Long, docword As Word.Document)
    Dim oFR As New clsFluidos_resultados
    Dim oFNV As New clsFluidos_normas_valores
    Dim rs_particulas As ADODB.Recordset
    Set rs_particulas = oFR.Listado(MUESTRA)
    If rs_particulas.RecordCount = 0 Then
        docword.Tables(3).Delete
    Else
        Dim oMuestra As New clsMuestra
        oMuestra.CargaMuestra (MUESTRA)
        Dim oFluido As New clsFluidos_ficha
        oFluido.Carga_por_BANO (oMuestra.getBANO_ID)
        Dim oFN As New clsFluidos_normas
        oFN.Carga (oFluido.getNORMA_ID)
        docword.Tables(3).Rows(2).Cells(5).Range.Text = oFN.getNOMBRE
        Do
            With docword.Tables(3)
                .Rows(rs_particulas("TAMANO") + 1).Cells(3).Range.Text = rs_particulas("RESULTADO")
                .Rows(rs_particulas("TAMANO") + 1).Cells(4).Range.Text = rs_particulas("CLASIFICACION")
                fila = oFNV.Calcula_Posicion(oFluido.getNORMA_ID, rs_particulas("TAMANO"), rs_particulas("RESULTADO"))
                Select Case rs_particulas("TAMANO")
                    Case 1
                        docword.Tables(4).Rows(fila).Cells(4).Range.Text = rs_particulas("RESULTADO")
                    Case 2
                        docword.Tables(4).Rows(fila).Cells(6).Range.Text = rs_particulas("RESULTADO")
                    Case 3
                        docword.Tables(4).Rows(fila).Cells(8).Range.Text = rs_particulas("RESULTADO")
                    Case 4
                        docword.Tables(4).Rows(fila).Cells(10).Range.Text = rs_particulas("RESULTADO")
                    Case 5
                        docword.Tables(4).Rows(fila).Cells(12).Range.Text = rs_particulas("RESULTADO")
                End Select
            End With
            rs_particulas.MoveNext
        Loop Until rs_particulas.EOF
        ' Merge de parametro y norma
        docword.Tables(3).Cell(2, 1).Merge docword.Tables(3).Cell(6, 1)
        docword.Tables(3).Cell(2, 5).Merge docword.Tables(3).Cell(6, 5)
    End If
    With docword.Tables(4)
        .Rows(2).Cells(2).Range.Text = "X"
    End With
End Sub
Public Sub pie_fluidos(appword As Word.Application, docword As Word.Document, rs As ADODB.Recordset, MUESTRA As Long, fecha_impresion As Date, por_impresora As Integer)
    Dim ovalores_bano As New clsDatos_valores
    With docword.Sections(1).Footers(1).Range.Tables(1)
        .Rows(1).Cells(1).Range.Text = "Observaciones (Remarks) : " & ovalores_bano.OBSERVACIONES(MUESTRA) & vbCrLf
        .Rows(2).Cells(1).Range.InsertAfter fecha_larga(fecha_impresion)
        Dim oCliente As New clsCliente
        oCliente.CargaCliente rs(18)
        If oCliente.getAIRBUS = 1 Then
            .Rows(3).Cells(3).Range.Text = "Vº.Bº. AIRBUS MILITARY MTQM (AIRBUS MILITARY MTQM approved)"
        Else
            .Rows(3).Cells(3).Range.Text = "Vº.Bº." & rs(0) & " (" & rs(0) & " approved)"
        End If
        ' Firmas
        Dim firma_analista As String
        firma_analista = usuario.firma_analista(MUESTRA)
        If firma_analista <> "" Then
            If Dir(firma_analista) <> "" Then
                .Rows(4).Cells(1).Range.InlineShapes.AddPicture firma_analista
            End If
        End If
    End With
    ' Almacenar e imprimir con logo
    docword.Save
    imprimir_documento MUESTRA, appword, 5
    ' Imprimimos sin el logo
    docword.Sections(1).Headers(1).Range.Tables(1).Rows(1).Cells(1).Range.Delete
    docword.Sections(1).Headers(1).Range.Tables(1).Rows(2).Cells(1).Range.Text = ""
    docword.Sections(1).Headers(1).Range.Tables(1).Rows(1).Cells(3).Range.Delete
    docword.SaveAs NOMBRE_DOCUMENTO(MUESTRA, False) & "--.doc"
    imprimir_documento MUESTRA, appword, por_impresora
End Sub

