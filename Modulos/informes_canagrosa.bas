Attribute VB_Name = "informes_canagrosa"
Public linea1 As String
Public linea2 As String
'Public Function imprimir_informe_alimentos(ByVal MUESTRA As Long, por_impresora As Integer, fecha_impresion As Date) As Boolean
'    On Error GoTo fallo
'    Dim rs As ADODB.Recordset
'    Dim oMuestra As New clsMuestra
'    Dim appword As Word.Application
'    Dim docword As Word.Document
'    ' Crear copia para su uso
'    Set appword = CreateObject("word.application")
'    Set docword = appword.Documents.Open(copiar_plantilla("27", MUESTRA, por_impresora))
'    appword.Visible = False
'    appword.WindowState = wdWindowStateMinimize
'    ' Ensayo
'    oMuestra.CargaMuestra (MUESTRA)
'    Set rs = oMuestra.datos_cabecera_documento(MUESTRA)
'    docword.Sections(1).Headers(1).Range.Tables(2).Rows(1).Cells(1).Range.Text = "Nº INFORME ENSAYO " & rs(17) & "/" & Format(rs(7), "yyyy")
'    ' Cabecera
'    With docword.Sections(1).Headers(1).Range.Tables(3)
'        .Rows(1).Cells(1).Range.Text = "Edición nº " & rs(10) + 1
'        Dim mensaje_edicion As String
'        If rs(10) = 0 Then
'            mensaje_edicion = ""
'        Else
'            mensaje_edicion = ReadINI(App.Path + "\config.ini", "edicion", "mensaje")
'        End If
'        .Rows(2).Cells(1).Range.Text = mensaje_edicion
'        .Rows(1).Cells(2).Range.Text = rs(0)
'        .Rows(2).Cells(2).Range.Text = rs(1)
'        .Rows(3).Cells(1).Range.Text = "Fecha de Recepción: " & Format(rs(7), "dd-mm-yyyy")
'        If rs(4) = "" Then
'            .Rows(3).Cells(2).Range.Text = rs(2) & " " & rs(3)
'        Else
'            .Rows(3).Cells(2).Range.Text = rs(2) & " " & rs(3) & " (" & Trim(rs(4)) & ")"
'        End If
''        .Rows(4).Cells(1).Range.Text = "Fecha de Inicio: " & Format(rs(8), "dd-mm-yyyy") ' F. Muestreo
'        .Rows(4).Cells(1).Range.Text = "Fecha de Inicio: " & Format(rs(30), "dd-mm-yyyy") ' F. Comienzo
'        .Rows(5).Cells(1).Range.Text = "Fecha de Finalización: " & Format(rs(9), "dd-mm-yyyy")
'        .Rows(6).Cells(1).Range.Text = "Registrado en Sevilla."
'        If rs(6) <> "" Then
'            .Rows(6).Cells(2).Range.Text = "A/A de " & rs(6)
'        End If
'    End With
'    ' Descripción de la muestra
'    docword.Tables(1).Rows(1).Cells(2).Range.Text = rs(13)
'    ' Datos específicos, irán en dos columnas
'    Dim rs_valores As ADODB.Recordset
'    Dim ovalores As New clsDatos_valores
'    Dim fila As Integer
'    Dim Col As Integer
'    Col = 1
'    fila = 2
'    Set rs_valores = ovalores.datos_especiales(MUESTRA)
'    If rs_valores.RecordCount <> 0 Then
'        Do
'          If Trim(rs_valores(1)) <> "" Then
'            docword.Tables(1).Rows(fila).Cells(Col).Range.Text = rs_valores(0) & ": " & rs_valores(1)
'            If Col = 1 Then
'                Col = Col + 1
'            Else
'                Col = 1
'                fila = fila + 1
'            End If
'          End If
'          rs_valores.MoveNext
'          If Not rs_valores.EOF Then
'           If Trim(rs_valores(1)) <> "" And Col = 1 Then
'             docword.Tables(1).Rows.Add
'           End If
'          End If
'        Loop Until rs_valores.EOF
''        docword.Tables(1).Rows.Last.Delete
'    End If
'    ' Determinaciones
'    Dim odeterminaciones As New clsDeterminaciones
'    Dim ounidad As New clsUnidades
'    Dim rs_deter As ADODB.Recordset
'    Set rs_deter = odeterminaciones.lista_determinaciones(MUESTRA)
'    Dim des As String
'    Dim pos As Integer
'    Dim oDA As New clsDeterminaciones_analisis
'    With docword.Tables(3)
'        If rs_deter.RecordCount > 0 Then
'        Do
'            With .Rows.Last
'                pos = InStr(1, UCase(rs_deter("nombre")), "S.S.N", vbTextCompare)
'                If pos > 0 Then
'                    des = Left(rs_deter("nombre"), pos - 1)
'                Else
'                    des = rs_deter("nombre")
'                End If
'                .Cells(1).Range.Text = des
'                If Trim(rs_deter("lc")) <> "" Then
'                    .Cells(2).Range.Text = rs_deter("lc")
'                End If
'                .Cells(3).Range.Text = rs_deter("pnt")
'                .Cells(4).Range.Text = rs_deter("resultado")
'                ' Legislación /Rango
'                rango = ""
'                ' TEXTO
'                If oDA.Carga_por_tipo_analisis(oMuestra.getTIPO_ANALISIS_ID, rs_deter("tipo_determinacion_id")) = True Then
'                    With oDA
''                        If Trim(.getMINIMO) <> "" And Trim(.getMAXIMO) <> "" Then
''                            rango = .getMINIMO & " - " & .getMAXIMO
''                        End If
''                        If Trim(.getMINIMO) <> "" And Trim(.getMAXIMO) = "" Then
''                            rango = " > " & .getMINIMO
''                        End If
''                        If Trim(.getMINIMO) = "" And Trim(.getMAXIMO) <> "" Then
''                            rango = " < " & .getMAXIMO
''                        End If
'                        If Trim(.getMINIMO_TEXTO) <> "" And Trim(.getMAXIMO_TEXTO) <> "" Then
'                            rango = .getMINIMO_TEXTO & " - " & .getMAXIMO_TEXTO
'                        Else
'                            rango = Trim(.getMINIMO_TEXTO) & Trim(.getMAXIMO_TEXTO)
'                        End If
'                        If Trim(rango) = "" Then
'                            rango = "--"
'                        End If
'                    End With
'                End If
'                .Cells(6).Range.Text = rango
'                .Cells(7).Range.Text = ounidad.Unidad_Campo_Resultado(rs_deter("id_determinacion"))
'                If InStr(1, UCase(rs_deter("nombre")), "S.S.N", vbTextCompare) > 0 Then
'                    rs_deter.MoveNext
'                    If rs_deter.EOF Then
'                        .Cells(5).Range.Text = "--"
'                    Else
'                        If InStr(1, UCase(rs_deter("nombre")), "S.S.S", vbTextCompare) > 0 Then
'                            .Cells(5).Range.Text = rs_deter("resultado")
'                        Else
'                            .Cells(5).Range.Text = "--"
'                            rs_deter.MovePrevious
'                        End If
'                    End If
'                Else
'                    .Cells(5).Range.Text = "--"
'                End If
'            End With
'            If Not rs_deter.EOF Then
'                rs_deter.MoveNext
'                If Not rs_deter.EOF Then
'                    .Rows.Add
'                End If
'            End If
'        Loop Until rs_deter.EOF
'        End If
'    End With
'    ' Normativa aplicable
'    Dim oTA As New clsTipos_analisis
'    oTA.CARGAR (oMuestra.getTIPO_ANALISIS_ID)
'    If Trim(oTA.getNORMATIVA) <> "" Then
'        With docword.Tables(4)
'            .Rows.Last.Cells(1).Range.Text = "NORMATIVA APLICABLE: " & oTA.getNORMATIVA
'        End With
'    End If
'    ' Pie
'    docword.Sections(1).Footers(1).Range.Tables(1).Rows(1).Cells(1).Range.Text = "Sevilla a, " & Format(fecha_impresion, "d Mmmm yyyy")
'    ' Firma del analista y del responsable del cierre
'    Dim firma As String
'    firma = USUARIO.firma_analista(MUESTRA)
'    If firma <> "" Then
'        If Dir(firma) <> "" Then
'            docword.Sections(1).Footers(1).Range.Tables(2).Rows(2).Cells(1).Range.InlineShapes.AddPicture firma
'        End If
'    End If
'    firma = USUARIO.firma_responsable_cierre(MUESTRA)
'    If firma <> "" Then
'        If Dir(firma) <> "" Then
'            docword.Sections(1).Footers(1).Range.Tables(2).Rows(2).Cells(2).Range.InlineShapes.AddPicture firma
'        End If
'    End If
'    docword.Save
''    imprimir_documento MUESTRA, appword, por_impresora
'    imprimir_documento MUESTRA, appword, 5
'    ' Imprimimos sin el logo
''    docword.Sections(1).Headers(1).Range.Tables(1).Rows(1).Cells(1).Range.Delete
''    docword.Sections(1).Headers(1).Range.Tables(1).Rows(2).Cells(1).Range.Delete
''    docword.SaveAs NOMBRE_DOCUMENTO(MUESTRA, False) & "--.doc"
''    imprimir_documento MUESTRA, appword, por_impresora
'    Set docword = Nothing
'    Set appword = Nothing
'    imprimir_informe_alimentos = True
'    Exit Function
'fallo:
'    log ("Error al generar el documento de alimentos dia : " & Err.Description)
'    enviar_informe_error MUESTRA, "Imprimir_informe_alimentos : " & Err.Description
'    appword.Quit 0
'    imprimir_informe_alimentos = False
'    Set docword = Nothing
'    Set appword = Nothing
''    MsgBox "Se ha producido un error al generar el documento. " & Err.Description, vbCritical, "Error"
'End Function
'M0687-I
'Public Function imprimir_informe_bano(ByVal MUESTRA As Long, por_impresora As Integer, fecha_impresion As Date, TIPO As Integer) As Boolean
'    On Error GoTo fallo
'    Dim appword As Word.Application
'    Dim docword As Word.Document
'    ' Crear copia para su uso
'    Set appword = CreateObject("word.application")
'    Set docword = appword.Documents.Open(copiar_plantilla("6", MUESTRA, por_impresora, TIPO))
'    appword.Visible = True
'    appword.WindowState = wdWindowStateMinimize
'    ' Cabecera
'    Dim oMuestra As New clsMuestra
'    Dim rs As ADODB.Recordset
'    Set rs = oMuestra.datos_cabecera_documento(MUESTRA)
'    cabecera_bano docword, rs, "6"
'    ' Datos Específicos
'    Dim ovalores_bano As New clsDatos_valores
'    Dim rs_dc As ADODB.Recordset
'    Dim snivel As String
'    Set rs_dc = oMuestra.datos_cabecera_bano(MUESTRA)
'    If rs_dc.RecordCount > 0 Then
'        With docword.Tables(1)
'            .Rows(4).Cells(2).Range.Text = (rs_dc(0))
'            .Rows(5).Cells(2).Range.Text = (rs_dc(1))
'            .Rows(6).Cells(2).Range.Text = (rs_dc(2))
'            .Rows(3).Cells(3).Range.InsertAfter ovalores_bano.HORA(MUESTRA)
'            .Rows(5).Cells(3).Range.InsertAfter ovalores_bano.nivel(MUESTRA)
'            .Rows(7).Cells(3).Range.InsertAfter ovalores_bano.TEMPERATURA(MUESTRA)
'            .Rows(8).Cells(2).Range.InsertAfter (ovalores_bano.recarga(MUESTRA))
'        End With
'        ' Linea, bano, solucion
'        With docword.Tables(2)
'            .Rows(2).Cells(1).Range.Text = rs_dc(0)
'            If Trim(rs(5)) <> "" Then
'                .Rows(2).Cells(1).Range.InsertAfter "/" & rs(5)
'            End If
'            .Rows(2).Cells(2).Range.Text = rs_dc(5)
'            If Trim(rs_dc(6)) <> "" Then
'                .Rows(2).Cells(2).Range.InsertAfter vbCrLf & vbCrLf & vbCrLf
'                .Rows(2).Cells(2).Range.Paragraphs.Add
'                .Rows(2).Cells(2).Range.InsertAfter "Volumen = " & rs_dc(6)
'            End If
'            .Rows(2).Cells(3).Range.Text = rs_dc(3)
'            .Rows(2).Cells(2).Range.Paragraphs(1).Range.Bold = True
'        End With
'    End If
'    ' Determinaciones
'    Dim odeterminaciones As New clsDeterminaciones
'    Dim rs_deter As ADODB.Recordset
'    Dim rango As String
'    Dim I As Integer
'    Dim nsd As Boolean
'    nsd = False
'    Set rs_deter = odeterminaciones.lista_determinaciones_bano(MUESTRA)
'    If rs_deter.RecordCount > 0 Then
'        With docword.Tables(2)
'            I = 1
'            Do
'                With .Rows.Last
'                    .Cells(4).Range.Bold = False
'                    .Cells(4).Range.Text = rs_deter(0)
'                    .Cells(4).Range.Paragraphs.Add
'                    .Cells(4).Range.Paragraphs(1).Range.Bold = True
'                    .Cells(4).Range.InsertAfter rs_deter(1) & vbCrLf
'                    ' Rango
'                    rango = ""
'                    If Trim(rs_deter(14)) <> "" And Trim(rs_deter(15)) <> "" Then
'                        rango = rs_deter(14) & " - " & rs_deter(15)
'                    Else
'                        rango = Trim(rs_deter(14)) & Trim(rs_deter(15))
'                    End If
'                    If Trim(rango) <> "" Then
'                        rango = rango & " " & rs_deter(7)
'                        .Cells(4).Range.InsertAfter rango & vbCrLf
'                    End If
'                    ' Resultado
'                    If IsNumeric(rs_deter(6)) = False Then
'                        If rs_deter(6) = "--" Then
'                            .Cells(5).Range.Text = rs_deter(6)
'                        Else
'                            .Cells(5).Range.Text = rs_deter(6) & " " & rs_deter(7)
'                        End If
'                        .Cells(5).Range.Underline = wdUnderlineNone
'                    Else
'                        If (CSng(rs_deter(6)) = 0) And (rs_deter(3) = 0) Then
'                            .Cells(5).Range.Text = "n.s.d."
'                            nsd = True
'                        Else
'                            .Cells(5).Range.Underline = wdUnderlineNone
'                            If Trim(rs_deter(10)) <> "" And IsNumeric(rs_deter(10)) Then
'                                If CSng(Replace(rs_deter(6), ".", ",")) < CSng(Replace(rs_deter(10), ".", ",")) Then
'                                    .Cells(5).Range.Underline = wdUnderlineSingle
'                                End If
'                            End If
'                            If Trim(rs_deter(11)) <> "" And IsNumeric(rs_deter(11)) Then
'                                If CSng(Replace(rs_deter(6), ".", ",")) > CSng(Replace(rs_deter(11), ".", ",")) Then
'                                    .Cells(5).Range.Underline = wdUnderlineSingle
'                                End If
'                            End If
'                            .Cells(5).Range.Text = rs_deter(6) & " " & rs_deter(7)
'                        End If
'                    End If
'                    ' Procedimiento
'                    .Cells(6).Range.Text = rs_deter(2)
'                    ' Fecha
'                    .Cells(7).Range.Text = Format(rs_deter(5), "dd/mm/yy")
'                End With
'                rs_deter.MoveNext
'                If Not rs_deter.EOF Then
'                    .Rows.Add
'                    I = I + 1
'                End If
'            Loop Until rs_deter.EOF
'            Dim J As Integer
'            Dim K As Integer
'            ' Merge de las 3 primeras columnas
'            For K = 2 To I
'                For J = 1 To 3
'                    .Cell(2, J).Merge .Cell(K + 1, J)
'                Next
'            Next
'            ' Eliminar bordes del resto de columnas
'            For K = 2 To I
'                For J = 4 To 7
'                    .Cell(K, J).Borders(wdBorderBottom).Visible = False
'                Next
'            Next
'        End With
'    End If
'    If nsd = True Then
'        docword.Tables(3).Rows(1).Cells(1).Range.Text = "n.s.d. : no se detecta"
'    End If
'    ' Pie
'    pie_banos appword, docword, rs, MUESTRA, fecha_impresion, por_impresora, TIPO
'    Set docword = Nothing
'    Set appword = Nothing
'    imprimir_informe_bano = True
'    Exit Function
'fallo:
'    log ("***** Error al generar el documento de baño : " & Err.Description)
'    enviar_informe_error MUESTRA, "Imprimir_informe_bano : " & Err.Description
'    imprimir_informe_bano = False
'    appword.Documents.Close (wdDotNotSaveChanges)
'    appword.Quit 0
'    Set docword = Nothing
'    Set appword = Nothing
'End Function
'M0687-F

Public Sub cabecera_bano(docword As Word.Document, rs As ADODB.Recordset, PLANTILLA As String)
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
        If Format(rs(8), "yyyy-mm-dd") = "1900-01-01" Then
            .Rows(3).Cells(3).Range.Text = "Sin especificar" ' Fecha muestreo
        Else
            .Rows(3).Cells(3).Range.Text = rs(8) ' Fecha muestreo
        End If
        .Rows(4).Cells(3).Range.Text = rs(7) ' Fecha recepcion
        If Not IsNull(rs(30)) Then
            .Rows(5).Cells(3).Range.Text = rs(30) ' Fecha comiento
        End If
' Sustituir fecha de cierre por fecha de finalizacion
'        If Not IsNull(rs(9)) Then
'            .Rows(6).Cells(3).Range.Text = rs(9) ' Fecha cierre
'        End If
        If Not IsNull(rs(32)) Then
            .Rows(6).Cells(3).Range.Text = rs(32) ' Fecha finalizacion
        End If
    End With
    ' Mensaje de edición baños
    If PLANTILLA <> "AGRUPADA" And PLANTILLA <> "CE" Then
        With docword.Tables(1)
            .Rows(1).Cells(1).Range.InsertAfter (rs(17) & "/" & Format(rs(7), "yyyy") & "/Edición " & rs(10) + 1)
            Dim mensaje_edicion As String
            If rs(10) = 0 Then
                mensaje_edicion = ""
            Else
                mensaje_edicion = ReadINI(App.Path + "\config.ini", "edicion", "mensaje") & "-" & ReadINI(App.Path + "\config.ini", "edicion", "ingles")
            End If
            .Rows(2).Cells(1).Range.Text = mensaje_edicion
            If PLANTILLA <> "CE" And PLANTILLA <> "SELLANTE" And PLANTILLA <> "COMBUSTIBLE" Then
                If PLANTILLA = "HH" Then
                    .Rows(4).Cells(2).Range.Text = "Tomada por " & (rs(15))
                Else
                    .Rows(3).Cells(2).Range.Text = "Tomada por " & (rs(15))
                End If
            End If
        End With
    End If
End Sub
'Public Function imprimir_informe_taladrina(ByVal MUESTRA As Long, por_impresora As Integer, fecha_impresion As Date) As Boolean
'    On Error GoTo fallo
'    Dim appword As Word.Application
'    Dim docword As Word.Document
'    ' Crear copia para su uso
'    Set appword = CreateObject("word.application")
'    Set docword = appword.Documents.Open(copiar_plantilla("23", MUESTRA, por_impresora))
'    appword.Visible = False
'    appword.WindowState = wdWindowStateMinimize
'    ' Cabecera
'    Dim oMuestra As New clsMuestra
'    Dim rs As ADODB.Recordset
'    Set rs = oMuestra.datos_cabecera_documento(MUESTRA)
'    cabecera_bano docword, rs, "23"
'    ' Datos específicos
'    Dim rs_dc As ADODB.Recordset
'    Set rs_dc = oMuestra.datos_cabecera_bano(MUESTRA)
'    With docword.Tables(1)
'        .Rows(4).Cells(2).Range.Text = (rs_dc(2)) ' Periodicidad
'        .Rows(5).Cells(2).Range.Text = (rs_dc(1))
'    End With
'    ' Maquina, Designacion
'    With docword.Tables(2)
''        .Rows(1).Cells(1).Range.Text = rs_dc(5) ' Nombre del baño
'        .Rows(2).Cells(1).Range.Text = rs_dc(5) '
'        .Rows(2).Cells(2).Range.Text = rs_dc(3)
'    End With
'    ' Determinaciones
'    Dim odeterminaciones As New clsDeterminaciones
'    Dim rs_deter As ADODB.Recordset
'    Dim rango As String
'    Dim I As Integer
'    Dim nsd As Boolean
'    nsd = False
'    Set rs_deter = odeterminaciones.lista_determinaciones_bano(MUESTRA)
'    With docword.Tables(2)
'        I = 1
'        If rs_deter.RecordCount > 0 Then
'        Do
'            With .Rows.Last
'                ' Parámetro
'                .Cells(3).Range.Bold = False
'                .Cells(3).Range.Text = rs_deter(0)
'                .Cells(3).Range.Paragraphs.Add
'                .Cells(3).Range.Paragraphs(1).Range.Bold = True
'                .Cells(3).Range.InsertAfter rs_deter(1) & vbCrLf
'                ' Rango
'                ' TEXTO
'                rango = ""
''                If Trim(rs_deter(10)) <> "" And Trim(rs_deter(11)) <> "" Then
''                    rango = rs_deter(10) & " - " & rs_deter(11)
''                End If
''                If Trim(rs_deter(10)) <> "" And Trim(rs_deter(11)) = "" Then
''                    rango = " > " & rs_deter(10)
''                End If
''                If Trim(rs_deter(10)) = "" And Trim(rs_deter(11)) <> "" Then
''                    rango = " < " & rs_deter(11)
''                End If
'
'                If Trim(rs_deter(14)) <> "" And Trim(rs_deter(15)) <> "" Then
'                    rango = rs_deter(14) & " - " & rs_deter(15)
'                Else
'                    rango = Trim(rs_deter(14)) & Trim(rs_deter(15))
'                End If
'
'
'                If Trim(rango) <> "" Then
'                    rango = rango & " " & rs_deter(7)
'                    .Cells(3).Range.InsertAfter rango ' & vbCrLf
'                End If
'                ' Resultado
'                If IsNumeric(rs_deter(6)) = False Then
'                    If rs_deter(6) = "--" Then
'                        .Cells(4).Range.Text = rs_deter(6)
'                    Else
'                        .Cells(4).Range.Text = rs_deter(6) & " " & rs_deter(7)
'                    End If
'                    .Cells(4).Range.Underline = wdUnderlineNone
'                Else
'                    If (CSng(rs_deter(6)) = 0) And (rs_deter(3) = 0) Then
'                        .Cells(4).Range.Text = "n.s.d."
'                        nsd = True
'                    Else
'                        .Cells(4).Range.Underline = wdUnderlineNone
'                        If Trim(rs_deter(10)) <> "" Then
'                            If CSng(Replace(rs_deter(6), ".", ",")) < CSng(Replace(rs_deter(10), ".", ",")) Then
'                                .Cells(4).Range.Underline = wdUnderlineSingle
'                            End If
'                        End If
'                        If Trim(rs_deter(11)) <> "" Then
'                            If CSng(Replace(rs_deter(6), ".", ",")) > CSng(Replace(rs_deter(11), ".", ",")) Then
'                                .Cells(4).Range.Underline = wdUnderlineSingle
'                            End If
'                        End If
'                        .Cells(4).Range.Text = rs_deter(6) & " " & rs_deter(7)
'                    End If
'                End If
'                ' Ref. Ensayo
'                .Cells(5).Range.Text = rs_deter(2)
'                ' Fecha
'                .Cells(6).Range.Text = Format(rs_deter(5), "dd/mm/yy")
'            End With
'            rs_deter.MoveNext
'            If Not rs_deter.EOF Then
'                .Rows.Add
'                I = I + 1
'            End If
'        Loop Until rs_deter.EOF
'        Dim J As Integer
'        Dim K As Integer
'        ' Merge de las 3 primeras columnas
'        For K = 2 To I
'            For J = 1 To 2
'                .Cell(2, J).Merge .Cell(K + 1, J)
'            Next
'        Next
'        End If
'    End With
'    If nsd = True Then
'        docword.Tables(4).Rows(1).Cells(1).Range.Text = "n.s.d. : no se detecta"
'    End If
'    ' Pie
'    firmas_banos docword, rs, MUESTRA, fecha_impresion
'    docword.Save
'    imprimir_documento MUESTRA, appword, por_impresora
'    Set docword = Nothing
'    Set appword = Nothing
'    imprimir_informe_taladrina = True
'    Exit Function
'fallo:
'    log ("Error al generar el documento de taladrinas : " & Err.Description)
'    enviar_informe_error MUESTRA, "Imprimir_informe_taladrina : " & Err.Description
'    imprimir_informe_taladrina = False
'    appword.Quit 0
'    Set docword = Nothing
'    Set appword = Nothing
'End Function
'M00687-I
'Public Function imprimir_informe_aguas(ByVal MUESTRA As Long, por_impresora As Integer, fecha_impresion As Date, TIPO As Integer) As Boolean
'    On Error GoTo fallo
'    Dim appword As Word.Application
'    Dim docword As Word.Document
'    ' Crear copia para su uso
'    Set appword = CreateObject("word.application")
'    Set docword = appword.Documents.Open(copiar_plantilla("2", MUESTRA, por_impresora, TIPO))
'    appword.Visible = False
'    appword.WindowState = wdWindowStateMinimize
'    ' Cabecera
'    Dim oMuestra As New clsMuestra
'    Dim rs As ADODB.Recordset
'    Set rs = oMuestra.datos_cabecera_documento(MUESTRA)
'    cabecera_bano docword, rs, "2"
'    ' Datos específicos
'    Dim rs_dc As ADODB.Recordset
'    Dim ovalores_bano As New clsDatos_valores
'    Set rs_dc = oMuestra.datos_cabecera_bano(MUESTRA)
'    If rs_dc.RecordCount > 0 Then
'        With docword.Tables(1)
'            .Rows(4).Cells(2).Range.Text = (rs_dc(0))
'            .Rows(5).Cells(2).Range.Text = (rs_dc(1))
'            .Rows(6).Cells(2).Range.Text = (rs_dc(2))
'            .Rows(3).Cells(3).Range.InsertAfter ovalores_bano.HORA(MUESTRA)
'            .Rows(5).Cells(3).Range.InsertAfter ovalores_bano.nivel(MUESTRA)
'            .Rows(7).Cells(3).Range.InsertAfter ovalores_bano.TEMPERATURA(MUESTRA)
'            .Rows(8).Cells(2).Range.InsertAfter (ovalores_bano.recarga(MUESTRA))
'        End With
'
'        With docword.Tables(1)
'            .Rows(4).Cells(2).Range.Text = (rs_dc(0))
'            .Rows(5).Cells(2).Range.Text = (rs_dc(1))
'            .Rows(6).Cells(2).Range.Text = (rs_dc(2))
'            .Rows(8).Cells(2).Range.InsertAfter (ovalores_bano.recarga(MUESTRA))
'        End With
'        ' Mirar si es un agua de clase A (59,60 o 100)
'        Dim claseA As Boolean
'        claseA = False
'        If rs_dc(7) = 59 Or rs_dc(7) = 60 Or rs_dc(7) = 100 Then
'            claseA = True
'        End If
'        ' bano , observaciones
'        With docword.Tables(2)
'            ' Nº Baño
'            If claseA = False Then
'                .Rows(2).Cells(1).Range.Text = rs_dc(5)
'            Else
'                .Rows(2).Cells(1).Range.Text = rs_dc(5) & " *"
'            End If
'            If Trim(rs_dc(6)) <> "" Then
'                .Rows(2).Cells(1).Range.InsertAfter vbCrLf
'                .Rows(2).Cells(1).Range.Paragraphs.Add
'                .Rows(2).Cells(1).Range.InsertAfter "Volumen = " & rs_dc(6)
'            End If
'            .Rows(2).Cells(1).Range.Paragraphs(1).Range.Bold = True
'            ' Solución
'            .Rows(2).Cells(2).Range.Text = rs_dc(3)
'            ' Solución de procedencia
'            oMuestra.CargaMuestra (MUESTRA)
'            Dim oBANO As New clsBanos
'            oBANO.cargar_bano (oMuestra.getBANO_ID)
'            Dim osol As New clsSoluciones
'            osol.CARGAR (oBANO.getSOLUCION_PROCEDENCIA_ID)
'            .Rows(2).Cells(3).Range.Text = osol.getNOMBRE
'        End With
'    End If
'    ' Determinaciones
'    Dim odeterminaciones As New clsDeterminaciones
'    Dim rs_deter As ADODB.Recordset
'    Dim rango As String
'    Dim I As Integer
'    Dim nsd As Boolean
'    nsd = False
'    Set rs_deter = odeterminaciones.lista_determinaciones_bano(MUESTRA)
'    If rs_deter.RecordCount > 0 Then
'        With docword.Tables(2)
'            I = 1
'            Do
'                With .Rows.Last
'                    .Cells(4).Range.Bold = False
'                    .Cells(4).Range.Text = rs_deter(0)
'                    .Cells(4).Range.Paragraphs.Add
'                    .Cells(4).Range.Paragraphs(1).Range.Bold = True
'                    .Cells(4).Range.InsertAfter rs_deter(1) & vbCrLf
'                    ' Rango
'                    rango = ""
'                    If Trim(rs_deter(14)) <> "" And Trim(rs_deter(15)) <> "" Then
'                        rango = rs_deter(14) & " - " & rs_deter(15)
'                    Else
'                        rango = Trim(rs_deter(14)) & Trim(rs_deter(15))
'                    End If
'
'                    If Trim(rango) <> "" Then
'                        rango = rango & " " & rs_deter(7)
'                        .Cells(4).Range.InsertAfter rango & vbCrLf
'                    End If
'                    If IsNumeric(rs_deter(6)) = False Then
'                        If rs_deter(6) = "--" Then
'                            .Cells(5).Range.Text = rs_deter(6)
'                        Else
'                            .Cells(5).Range.Text = rs_deter(6) & " " & rs_deter(7)
'                        End If
'                        .Cells(5).Range.Underline = wdUnderlineNone
'                    Else
'                        If (CSng(rs_deter(6)) = 0) And (rs_deter(3) = 0) Then
'                            .Cells(5).Range.Text = "n.s.d."
'                            nsd = True
'                        Else
'                            .Cells(5).Range.Underline = wdUnderlineNone
'                            If Trim(rs_deter(10)) <> "" And IsNumeric(rs_deter(10)) Then
'                                If CSng(Replace(rs_deter(6), ".", ",")) < CSng(Replace(rs_deter(10), ".", ",")) Then
'                                    .Cells(5).Range.Underline = wdUnderlineSingle
'                                End If
'                            End If
'                            If Trim(rs_deter(11)) <> "" And IsNumeric(rs_deter(11)) Then
'                                If CSng(Replace(rs_deter(6), ".", ",")) > CSng(Replace(rs_deter(11), ".", ",")) Then
'                                    .Cells(5).Range.Underline = wdUnderlineSingle
'                                End If
'                            End If
'                            .Cells(5).Range.Text = rs_deter(6) & " " & rs_deter(7)
'                        End If
'                    End If
'                    .Cells(6).Range.Text = rs_deter(2)
'                    .Cells(7).Range.Text = Format(rs_deter(5), "dd/mm/yy")
'                End With
'                rs_deter.MoveNext
'                If Not rs_deter.EOF Then
'                    .Rows.Add
'                    I = I + 1
'                End If
'            Loop Until rs_deter.EOF
'            Dim J As Integer
'            Dim K As Integer
'            ' Merge de las 3 primeras columnas
'            For K = 2 To I
'                For J = 1 To 3
'                    .Cell(2, J).Merge .Cell(K + 1, J)
'                Next
'            Next
'            ' Eliminar bordes del resto de columnas
'            For K = 2 To I
'                For J = 4 To 7
'                    .Cell(K, J).Borders(wdBorderBottom).Visible = False
'                Next
'            Next
'        End With
'    End If
'    If nsd = True Then
'        docword.Tables(3).Rows(1).Cells(1).Range.Text = "n.s.d. : no se detecta"
'    End If
'    If claseA = True Then
'        If nsd = True Then
'            docword.Tables(3).Rows.Add
'            docword.Tables(3).Rows(2).Cells(1).Range.Text = "* Aguas de Clase A"
'        Else
'            docword.Tables(3).Rows(1).Cells(1).Range.Text = "* Aguas de Clase A"
'        End If
'    End If
'    ' Pie
'    pie_banos appword, docword, rs, MUESTRA, fecha_impresion, por_impresora, TIPO
'    Set docword = Nothing
'    Set appword = Nothing
'    imprimir_informe_aguas = True
'    Exit Function
'fallo:
'    log ("Error al generar el documento de aguas : " & Err.Description)
'    enviar_informe_error MUESTRA, "Imprimir_informe_aguas : " & Err.Description
'    imprimir_informe_aguas = False
'    appword.Quit 0
'    Set docword = Nothing
'    Set appword = Nothing
'End Function
'M0687-F
'M0687-I
'Public Function imprimir_informe_aguas_tipo3(ByVal MUESTRA As Long, por_impresora As Integer, documento As Integer, fecha_impresion As Date, TIPO As Integer) As Boolean
'    On Error GoTo fallo
'    Dim rs As ADODB.Recordset
'    Dim oMuestra As New clsMuestra
'    Dim odeterminaciones As New clsDeterminaciones
'    Dim nDeterminaciones As Integer
'    Dim appword As Word.Application
'    Dim docword As Word.Document
'    Set appword = CreateObject("word.application")
'    ' Verificamos si hay que insertar los Fluoruros
'    ' En el caso de que si, seleccionamos la plantilla 2-3F
'    oMuestra.CargaMuestra (MUESTRA)
'    Dim bContieneFluoruro As Boolean
'    bContieneFluoruro = odeterminaciones.Contiene_Determinacion_Fluoruro(oMuestra.getBANO_ID)
'    ' Siempre habilitamos los fluoruros para las aguas de columna
'    bContieneFluoruro = True
'    If Not bContieneFluoruro Then
'        Set docword = appword.Documents.Open(copiar_plantilla("2-3", MUESTRA, por_impresora, TIPO))
'        nDeterminaciones = 4
'    Else
'        Set docword = appword.Documents.Open(copiar_plantilla("2-3F", MUESTRA, por_impresora, TIPO))
'        nDeterminaciones = 5
'    End If
'    appword.Visible = False
'    appword.WindowState = wdWindowStateMinimize
'    ' Ensayo
'    Set rs = oMuestra.datos_cabecera_documento(MUESTRA)
'    cabecera_bano docword, rs, "AGRUPADA"
'    ' Cabecera
'    Dim rs_dc As ADODB.Recordset
'    Set rs_dc = oMuestra.datos_cabecera_bano(MUESTRA)
'    Dim ovalores_bano As New clsDatos_valores
'    With docword.Sections(1).Headers(1).Range.Tables(2)
'        .Rows(1).Cells(2).Range.Text = "Tomada por " & (rs(15))
'        .Rows(4).Cells(2).Range.Text = (rs_dc(2))
'        .Rows(5).Cells(2).Range.InsertAfter (ovalores_bano.recarga(MUESTRA))
'    End With
'    Dim mensaje_edicion As String
'    If rs(10) = 0 Then
'        mensaje_edicion = ""
'    Else
'        mensaje_edicion = "* " & ReadINI(App.Path + "\config.ini", "edicion", "mensaje") & "-" & ReadINI(App.Path + "\config.ini", "edicion", "ingles")
'    End If
'    docword.Sections(1).Footers(1).Range.Tables(1).Rows(1).Cells(1).Range.InsertAfter mensaje_edicion
'    ' Determinaciones
'    Dim olb As New clsLineas_Banos
'    Dim rs_deter As ADODB.Recordset
'    Dim rs_aguas As ADODB.Recordset
'    Dim osol As New clsSoluciones
'    Dim celda As Integer
'    Dim rango As String
'    Dim I As Integer
'    Dim nsd As Boolean
'    Dim muestra_agua As Long
'    nsd = False
'    ' Mirar si es un agua de clase A (59,60 o 100)
'    Dim claseA As Boolean
'    claseA = False
'    Dim fecha As Date
'    Dim oBANO As New clsBanos
'    Dim observ As Integer
'    observ = 2
'    fecha = oMuestra.getFECHA_RECEPCION
'    Set rs_aguas = olb.Buscar_Documento(documento)
'    If rs_aguas.RecordCount > 0 Then
'        With docword.Tables(1)
'            I = 1
'            Do
'                With .Rows.Last
'                    muestra_agua = oMuestra.Cargar_Agua(fecha, rs_aguas("bano_id"))
'                    If muestra_agua <> 0 Then
'                        oMuestra.CargaMuestra (muestra_agua)
'                        If oMuestra.getULT_EDICION_IMP = 0 Then
'                        .Cells(1).Range.Text = oMuestra.getID_GENERAL & "/" & Format(oMuestra.getFECHA_RECEPCION, "yyyy") & "/Ed." & oMuestra.getULT_EDICION_IMP + 1
'                        Else
'                        .Cells(1).Range.Text = oMuestra.getID_GENERAL & "/" & Format(oMuestra.getFECHA_RECEPCION, "yyyy") & "/Ed." & oMuestra.getULT_EDICION_IMP + 1 & " *"
'                        End If
'                        If (MUESTRA <> muestra_agua) Then '  And (documento <> 2) Then
'                            oMuestra.aumentar_edicion_impresa (muestra_agua)
'                        End If
'                        oBANO.cargar_bano (rs_aguas("bano_id"))
'                        If rs_aguas("bano_id") = 59 Or rs_aguas("bano_id") = 60 Or rs_aguas("bano_id") = 100 Then
'                            .Cells(2).Range.Text = oBANO.getNOMBRE & " *"
'                            claseA = True
'                        Else
'                            .Cells(2).Range.Text = oBANO.getNOMBRE
'                        End If
'                        ' Solucion de Procedencia
'                        osol.CARGAR (oBANO.getSOLUCION_PROCEDENCIA_ID)
'                        .Cells(3).Range.Text = osol.getNOMBRE
'                        If muestra_agua <> 0 Then
'                            ' Observaciones
'                            If Trim(ovalores_bano.OBSERVACIONES(muestra_agua)) <> "" Then
'                                docword.Tables(3).Rows(1).Cells(1).Range.Text = "OBSERVACIONES"
'                                docword.Tables(3).Rows(2).Cells(1).Range.Text = "NºENSAYO"
'                                docword.Tables(3).Rows(2).Cells(2).Range.Text = "NºBAÑO"
'                                docword.Tables(3).Rows(2).Cells(3).Range.Text = "OBSERVACION"
'                                docword.Tables(3).Rows(observ).Cells(1).Range.Text = oMuestra.getID_GENERAL & "/" & Format(oMuestra.getFECHA_RECEPCION, "yyyy")
'                                docword.Tables(3).Rows(observ).Cells(2).Range.Text = oBANO.getNOMBRE
'                                docword.Tables(3).Rows(observ).Cells(3).Range.Text = ovalores_bano.OBSERVACIONES(muestra_agua)
'                                observ = observ + 1
'                                docword.Tables(3).Rows.Add
'                            End If
'                            Set rs_deter = odeterminaciones.lista_determinaciones_bano(muestra_agua)
'                            If rs_deter.RecordCount <> 0 Then
'                             Do
'                                Select Case rs_deter(12)
'                                Case 58, 1238 ' SD
'                                    celda = 4
'                                Case 56, 1236 ' Conduc
'                                    celda = 5
'                                Case 59, 1237 ' Cl
'                                    celda = 6
'                                Case 63, 1288 ' pH
'                                    celda = 7
'                                Case 226, 1289 ' Fluoruro
'                                    celda = 8
'                                Case Else
'                                    celda = 4
'                                End Select
'                                .Cells(3 + nDeterminaciones).Range.Text = "N/A"
'                                ' Rango
'                                    rango = ""
'                                    If Trim(rs_deter(14)) <> "" And Trim(rs_deter(15)) <> "" Then
'                                        rango = rs_deter(14) & " - " & rs_deter(15)
'                                    Else
'                                        rango = Trim(rs_deter(14)) & Trim(rs_deter(15))
'                                    End If
'
'                                    .Cells(celda).Range.Text = rango
'                                If Not bContieneFluoruro Then
'                                    .Cells(7 + nDeterminaciones).Range.Text = "N/A"
'                                Else
'                                    .Cells(8 + nDeterminaciones).Range.Text = "N/A"
'                                End If
'                                    If IsNumeric(rs_deter(6)) = False Then
'                                        .Cells(celda + nDeterminaciones).Range.Text = rs_deter(6)
'                                        .Cells(celda + nDeterminaciones).Range.Underline = wdUnderlineNone
'                                    Else
'                                        If (CSng(rs_deter(6)) = 0) And (rs_deter(3) = 0) Then
'                                            .Cells(celda + nDeterminaciones).Range.Text = "n.s.d."
'                                            nsd = True
'                                        Else
'                                            .Cells(celda + nDeterminaciones).Range.Underline = wdUnderlineNone
'                                            If Trim(rs_deter(10)) <> "" And IsNumeric(rs_deter(10)) Then
'                                                If CSng(Replace(rs_deter(6), ".", ",")) < CSng(Replace(rs_deter(10), ".", ",")) Then
'                                                    .Cells(celda + nDeterminaciones).Range.Underline = wdUnderlineSingle
'                                                End If
'                                            End If
'                                            If Trim(rs_deter(11)) <> "" And IsNumeric(rs_deter(11)) Then
'                                                If CSng(Replace(rs_deter(6), ".", ",")) > CSng(Replace(rs_deter(11), ".", ",")) Then
'                                                    .Cells(celda + nDeterminaciones).Range.Underline = wdUnderlineSingle
'                                                End If
'                                            End If
'                                            .Cells(celda + nDeterminaciones).Range.Text = rs_deter(6) ' & " " & rs_deter(7)
'                                        End If
'                                    End If
'    '                            End If
'                             rs_deter.MoveNext
'                             Loop Until rs_deter.EOF
'                            End If
'                        End If
'                    End If ' If muestra_agua <> 0
'                End With
'                rs_aguas.MoveNext
'                If Not rs_aguas.EOF And muestra_agua <> 0 Then
'                    .Rows.Add
'                    I = I + 1
'                End If
'            Loop Until rs_aguas.EOF
'            Dim J As Integer
'            Dim K As Integer
'            Dim columnas As Integer
'            If Not bContieneFluoruro Then
'                columnas = 11
'            Else
'                columnas = 13
'            End If
'            For K = 3 To I + 1
'                For J = 1 To columnas
'                    .Cell(K, J).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
'                Next
'            Next
'        End With
'    End If
'    If nsd = True Then
'        docword.Tables(2).Rows(1).Cells(1).Range.Text = "n.s.d. : no se detecta"
'    End If
'    If claseA = True Then
'        If nsd = True Then
'            docword.Tables(2).Rows.Add
'            docword.Tables(2).Rows(2).Cells(1).Range.Text = "* Aguas de Clase A"
'        Else
'            docword.Tables(2).Rows(1).Cells(1).Range.Text = "* Aguas de Clase A"
'        End If
'    End If
'    ' Pie
'    pie_banos appword, docword, rs, MUESTRA, fecha_impresion, por_impresora, TIPO
'    Set docword = Nothing
'    Set appword = Nothing
'    imprimir_informe_aguas_tipo3 = True
'    Exit Function
'fallo:
'    log ("***** Error al generar el documento de aguas : " & Err.Description)
'    enviar_informe_error MUESTRA, "Imprimir_informe_aguas_tipo3 : " & Err.Description
'    imprimir_informe_aguas_tipo3 = False
'    appword.Quit 0
'    Set docword = Nothing
'    Set appword = Nothing
'End Function
'M0687-F

'Public Sub pie_banos(appword As Word.Application, docword As Word.Document, rs As ADODB.Recordset, MUESTRA As Long, fecha_impresion As Date, por_impresora As Integer, Optional tipo As Integer)
'    Dim ovalores_bano As New clsDatos_valores
'    docword.Sections(1).Footers(1).Range.Tables(1).Rows(1).Cells(1).Range.Text = "Observaciones (Remarks) : " & ovalores_bano.OBSERVACIONES(MUESTRA) & vbCrLf
'    firmas_banos docword, rs, MUESTRA, fecha_impresion
''M00687-I
''    If TIPO = C_TIPOS_IMPRESION.VB_AIRBUS Then ' VºBºAIRBUS
''        visto_bueno_airbus docword, RS, MUESTRA, fecha_impresion
''    Else
'        docword.Tables(docword.Tables.Count).Delete
''    End If
''M00687-F
'    docword.Save
'    imprimir_documento MUESTRA, appword, 5
'End Sub
    
Public Sub firmas_banos(docword As Word.Document, rs As ADODB.Recordset, MUESTRA As Long, fecha_impresion As Date)
    ' Fecha
    docword.Sections(1).Footers(1).Range.Tables(1).Rows(2).Cells(1).Range.InsertAfter fecha_larga(fecha_impresion)
    Dim oCliente As New clsCliente
    oCliente.CargaCliente rs(18)
    If oCliente.getAIRBUS = 1 Then
        docword.Sections(1).Footers(1).Range.Tables(1).Rows(3).Cells(3).Range.Text = "Vº.Bº. AIRBUS MILITARY MTQM (AIRBUS MILITARY MTQM approved)"
    Else
        docword.Sections(1).Footers(1).Range.Tables(1).Rows(3).Cells(3).Range.Text = "Vº.Bº." & rs(0) & " (" & rs(0) & " approved)"
    End If
    ' Firmas
    Dim firma_analista As String
    firma_analista = USUARIO.firma_analista(MUESTRA)
    If firma_analista <> "" Then
        If Dir(firma_analista) <> "" Then
            docword.Sections(1).Footers(1).Range.Tables(1).Rows(4).Cells(1).Range.InlineShapes.AddPicture firma_analista
        End If
    End If
    ' Usuario revision
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra MUESTRA
    If oMuestra.getREVISION_USUARIO <> 0 Then
        If Dir(ReadINI(App.Path + "\config.ini", "documentos", "firmas") & "\" & oMuestra.getREVISION_USUARIO & ".jpg") <> "" Then
            docword.Sections(1).Footers(1).Range.Tables(1).Rows(4).Cells(2).Range.InlineShapes.AddPicture ReadINI(App.Path + "\config.ini", "documentos", "firmas") & "\" & oMuestra.getREVISION_USUARIO & ".jpg"
        End If
    End If
End Sub
'M00687-I
'Public Sub visto_bueno_airbus(docword As Word.Document, RS As ADODB.Recordset, MUESTRA As Long, fecha_impresion As Date)
'    ' Conforme / Texto / Usuario / Fecha y Hora
'    Dim oW As New clsWeb_muestras_revision
'    Dim s As String
'    If oW.Carga(MUESTRA) = True Then
'        If oW.getCONFORME = 0 Then
'            s = "NO CONFORME : "
'        Else
'            If oW.getCONFORME = 1 Then
'                s = "CONFORME : "
'            End If
'        End If
'        If Trim(oW.getCONFORME_TEXTO) <> "" Then
'            s = s & oW.getCONFORME_TEXTO
'        End If
'        s = s & vbNewLine
'        Dim oU As New clsWeb_usuarios
'        oU.Carga oW.getCONFORME_USUARIO_ID
'        s = s & "(Approv: " & oU.getUSUARIO & " date : " & oW.getCONFORME_FS & ")"
'        docword.Sections(1).Footers(1).Range.Tables(1).Rows(4).Cells(3).Range.Text = s
'
'        If oW.getRECARGA = 1 And oW.getRECARGA_TEXTO <> "" Then
'            docword.Tables(docword.Tables.Count).Rows(2).Cells(1).Range.Text = Replace(oW.getRECARGA_TEXTO, "<br />", vbNewLine)
'        Else
'            docword.Tables(docword.Tables.Count).Delete
'        End If
'    Else
'        docword.Tables(docword.Tables.Count).Delete
'    End If
'    Set oW = Nothing
'End Sub
'M00687-F

Public Function imprimir_informe_alodine(ByVal LOTE As Long, por_impresora As Integer) As Boolean
    On Error GoTo fallo
    ' Albaran
    Dim oDocumentacion As New clsDocumentacion
    Dim oAlodine As New clsAlodine
    Dim oAlodine_Lote As New clsAlodine_lotes
    oAlodine_Lote.Carga LOTE
    
    Dim nombre As String
    Dim pdf_albaran As String
    Dim pdf_certificado As String
    nombre = nombre_alodine(LOTE)
    
    Dim ruta As String
    DIRECTORIO_TEMPORAL = App.Path & "\tmp"
    ruta = DIRECTORIO_TEMPORAL & "\"
    
    pdf_certificado = nombre & " CERT.pdf"
    pdf_albaran = nombre & ".pdf"
    
    If oAlodine.Genera_Certificado(LOTE, ruta & pdf_certificado, True, "") = False Then
        GoTo fallo
    Else
        oDocumentacion.SubirAlodine LOTE, oAlodine_Lote.getEDICION, 1, ruta & pdf_certificado, pdf_certificado
    End If
    If oAlodine.Genera_Albaran(LOTE, ruta & pdf_albaran, True, "") = False Then
        GoTo fallo
    Else
        oDocumentacion.SubirAlodine LOTE, oAlodine_Lote.getEDICION, 2, ruta & pdf_albaran, pdf_albaran
    End If
    imprimir_informe_alodine = True
    Set oDocumentacion = Nothing
    Exit Function
fallo:
    imprimir_informe_alodine = False
End Function
'Public Function imprimir_informe_hh(ByVal MUESTRA As Long, por_impresora As Integer, fecha_impresion As Date) As Boolean
'    On Error GoTo fallo
'    Dim appword As Word.Application
'    Dim docword As Word.Document
'    ' Crear copia para su uso
'    Set appword = CreateObject("word.application")
'    Set docword = appword.Documents.Open(copiar_plantilla("HH", MUESTRA, por_impresora))
'    appword.Visible = False
'    appword.WindowState = wdWindowStateMinimize
'    ' Cabecera
'    Dim oMuestra As New clsMuestra
'    Dim rs As ADODB.Recordset
'    Set rs = oMuestra.datos_cabecera_documento(MUESTRA)
'    cabecera_bano docword, rs, "HH"
'    ' Datos especificos del combustible
'    oMuestra.CargaMuestra MUESTRA
'    With docword.Tables(1)
'        Dim oEnvase As New clsformatos
'        oEnvase.CargarFormato oMuestra.getFORMATO_ID
'        Dim oTA As New clsTipos_analisis
'        oTA.CARGAR oMuestra.getTIPO_ANALISIS_ID
'        .Rows(3).Cells(2).Range.Text = oTA.getNOMBRE
'        .Rows(5).Cells(2).Range.Text = oEnvase.getDESCRIPCION
'        .Rows(6).Cells(2).Range.Text = oTA.getNORMATIVA
'    End With
'    ' Determinaciones
'    Dim odeterminaciones As New clsDeterminaciones
'    Dim rs_deter As ADODB.Recordset
'    Dim rango As String
'    Dim I As Integer
'    Dim nsd As Boolean
'    nsd = False
'    Set rs_deter = odeterminaciones.lista_determinaciones_tipo_analisis(MUESTRA)
'    If rs_deter.RecordCount > 0 Then
'        With docword.Tables(2)
'            I = 1
'            Do
'                With .Rows.Last
'                    ' Referencia
'' J002-I
'                    If I = 1 Then
'' J002-F
'                     .Cells(1).Range.Text = oMuestra.getREFERENCIA_CLIENTE
'' J002-I
'                    End If
'' J002-F
'                    ' Parametros
'                    .Cells(2).Range.Bold = False
'                    .Cells(2).Range.Text = rs_deter(0)
'                    .Cells(2).Range.Paragraphs.Add
'                    .Cells(2).Range.Paragraphs(1).Range.Bold = True
'' J002-I
''                    .Cells(2).Range.InsertAfter rs_deter(1) & vbCrLf
'                    .Cells(2).Range.InsertAfter rs_deter(1)
'                    rango = ""
''                    If Trim(rs_deter(10)) <> "" And Trim(rs_deter(11)) <> "" Then
''                        rango = rs_deter(10) & " - " & rs_deter(11)
''                    End If
''                    If Trim(rs_deter(10)) <> "" And Trim(rs_deter(11)) = "" Then
''                        rango = " > " & rs_deter(10)
''                    End If
''                    If Trim(rs_deter(10)) = "" And Trim(rs_deter(11)) <> "" Then
''                        rango = " < " & rs_deter(11)
''                    End If
'                    If Trim(rs_deter(15)) <> "" And Trim(rs_deter(16)) <> "" Then
'                        rango = rs_deter(15) & " - " & rs_deter(16)
'                    Else
'                        rango = Trim(rs_deter(15)) & Trim(rs_deter(16))
'                    End If
'
'                    If Trim(rango) <> "" Then
'                        rango = rango & " " & rs_deter(7)
'                        .Cells(2).Range.InsertAfter vbCrLf & rango
'                    End If
'' J002-F
'                    ' Resultado
'                    Dim RESULTADO As String
'                    RESULTADO = Replace(rs_deter(6), "<", "")
'                    RESULTADO = Replace(RESULTADO, ">", "")
'                    If IsNumeric(RESULTADO) = False Then
'                        .Cells(3).Range.Text = rs_deter(6)
'                    Else
''                        If (CSng(rs_deter(6)) = 0) And (rs_deter(3) = 0) Then
'                        If (CSng(RESULTADO) = 0) And (rs_deter(3) = 0) Then
'                            .Cells(3).Range.Text = "n.s.d."
'                            nsd = True
'                        Else
'                            .Cells(3).Range.Underline = wdUnderlineNone
'                            If Trim(rs_deter(10)) <> "" And IsNumeric(rs_deter(10)) Then
''                                If CSng(Replace(rs_deter(6), ".", ",")) < CSng(Replace(rs_deter(10), ".", ",")) Then
'                                If CSng(Replace(RESULTADO, ".", ",")) < CSng(Replace(rs_deter(10), ".", ",")) Then
'                                    .Cells(3).Range.Underline = wdUnderlineSingle
'                                End If
'                            End If
'                            If Trim(rs_deter(11)) <> "" And IsNumeric(rs_deter(11)) Then
''                                If CSng(Replace(rs_deter(6), ".", ",")) > CSng(Replace(rs_deter(11), ".", ",")) Then
'                                If CSng(Replace(RESULTADO, ".", ",")) > CSng(Replace(rs_deter(11), ".", ",")) Then
'                                    .Cells(3).Range.Underline = wdUnderlineSingle
'                                End If
'                            End If
'                            .Cells(3).Range.Text = rs_deter(6) & " " & rs_deter(7)
''                            .Cells(3).Range.Text = rs_deter(6) & " " & rs_deter(7)
'                        End If
'                    End If
'' J002-I
'                    .Cells(4).Range.Text = rs_deter(14) ' L.C.
'                    .Cells(5).Range.Text = rs_deter(13) ' Ref. Ensayo (PNT)
''                    ' Ref. Ensayo (PNT)
''                    .Cells(4).Range.Text = rs_deter(13)
''                    ' Fecha
''                    .Cells(5).Range.Text = Format(rs_deter(5), "dd/mm/yy")
'' J002-F
'                End With
'                rs_deter.MoveNext
'                If Not rs_deter.EOF Then
'                    .Rows.Add
'                    I = I + 1
'                End If
'            Loop Until rs_deter.EOF
'' J002-I
'            Dim K As Integer
'            ' Merge de la 1 primera columnas
'            For K = 2 To I
'                .Cell(2, 1).Merge .Cell(K + 1, 1)
'            Next
'' J002-F
'        End With
'    End If
'    If nsd = True Then
'        docword.Tables(3).Rows(1).Cells(1).Range.Text = "n.s.d. : no se detecta"
'    End If
'    ' Pie
'    pie_banos appword, docword, rs, MUESTRA, fecha_impresion, por_impresora
'    Set docword = Nothing
'    Set appword = Nothing
'    imprimir_informe_hh = True
'    Exit Function
'fallo:
'    log ("Error al generar el documento HH : " & Err.Description)
'    enviar_informe_error MUESTRA, "Imprimir_informe_hh : " & Err.Description
'    imprimir_informe_hh = False
'    appword.Quit 0
'    Set docword = Nothing
'    Set appword = Nothing
'End Function
'M00687-I
'Public Sub generar_informe_sin_logo(MUESTRA As Long)
'    Dim appword As Word.Application
'    Dim docword As Word.Document
'    ' Crear copia para su uso
'   On Error GoTo generar_informe_sin_logo_Error
'
'    Set appword = CreateObject("word.application")
'    Set docword = appword.Documents.Open(NOMBRE_DOCUMENTO(MUESTRA, False) & ".doc")
'    ' Imprimimos sin el logo
'    docword.Sections(1).Headers(1).Range.Tables(1).Rows(1).Cells(1).Range.Delete
''    docword.Sections(1).Headers(1).Range.Tables(1).Rows(1).Cells(2).Range.Text = "" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
'    docword.Sections(1).Headers(1).Range.Tables(1).Rows(1).Cells(3).Range.Delete
'    docword.SaveAs NOMBRE_DOCUMENTO(MUESTRA, False) & "--.doc"
'    imprimir_documento MUESTRA, appword, 2
'
'   On Error GoTo 0
'   Exit Sub
'
'generar_informe_sin_logo_Error:
'    log ("Error al generar_informe_sin_logo : " & MUESTRA)
'End Sub
'M00687-F

