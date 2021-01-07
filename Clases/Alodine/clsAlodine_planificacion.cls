VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlodine_planificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSALODINE_PLANIFICACION
'***************************************************************
Private LOTE_ID As Long
Private CLIENTE_ID As Long
Private EADS As Long
Private CAPACIDAD_ID As Long
Private ETIQUETA_ID As Long
Private NUMERO_BOTES As Long
Private DOC_ID As Long
Private PRECIO As String
Private pedido As String
Private PEDIDO_ID As Long
Private NORMA As String
Private NORMA_ETIQUETA As Integer
Private fecha As String
Private grupo As Integer
Public Sub Imprimir_Etiquetas(ByVal prm_id_lote As Long, idioma As Long)
    Dim rs As New ADODB.Recordset
    Dim tx As TextStream
    Dim cont As Long, total_etiquetas As Long
    Dim objRpt(1 To 3) As New frmReport
    Dim datos(1 To 3) As String ' donde guarda los xml de los tres tamaños
    Dim intIndice As Integer
    
        
   On Error GoTo Imprimir_Etiquetas_Error

    Set rs = Me.Listado_Etiquetas_Lote(prm_id_lote)
    
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        Set tx = Nothing
        Set objRpt(1) = Nothing
        Set objRpt(2) = Nothing
        Set objRpt(3) = Nothing
        
        Exit Sub
    End If
    
    rs.MoveFirst
    Dim oAE As New clsAlodine_etiqueta
    Dim oD As New clsDecodificadora
    oD.Carga_valor DECODIFICADORA.IDIOMAS, idioma
    If Not rs.EOF Then
        oAE.Carga rs(9), oD.getPARAMETROS
    End If
    ' Normas
    Dim oAlodine As New clsAlodine
    Dim normas As String
    normas = oAlodine.getNormas(rs(9))
    
    ' Pictogramas
    Dim ubicacionPic As String
    ubicacionPic = ReadINI(App.Path & "\config.ini", "Otros", "Pictogramas")
    Dim pos As Integer
    Dim pictogramas(1 To 5) As String
    pos = 1
    If oAE.getPIC1 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC1.jpg"
        pos = pos + 1
    End If
    If oAE.getPIC2 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC2.jpg"
        pos = pos + 1
    End If
    If oAE.getPIC3 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC3.jpg"
        pos = pos + 1
    End If
    If oAE.getPIC4 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC4.jpg"
        pos = pos + 1
    End If
    If oAE.getPIC5 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC5.jpg"
        pos = pos + 1
    End If
    If oAE.getPIC6 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC6.jpg"
        pos = pos + 1
    End If
      
    While Not rs.EOF
        total_etiquetas = rs(7)
        intIndice = rs(10) ' Tamaño de la etiqueta
        
        For cont = 1 To total_etiquetas
            
            datos(intIndice) = datos(intIndice) & "<etiqueta>"
            datos(intIndice) = datos(intIndice) & "<ID_ALODINE>" & rs(9) & "</ID_ALODINE>"
            datos(intIndice) = datos(intIndice) & "<PRODUCTO>" & rs(0) & "</PRODUCTO>"
'            datos(intIndice) = datos(intIndice) & "<NORMA>" & rs(13) & "</NORMA>"
            datos(intIndice) = datos(intIndice) & "<NORMA>" & normas & "</NORMA>"
            datos(intIndice) = datos(intIndice) & "<DESCRIPCION>" & rs(1) & "</DESCRIPCION>"
            datos(intIndice) = datos(intIndice) & "<FECHA_ALTA>" & Format(rs(2), "dd/mm/yyyy") & "</FECHA_ALTA>"
            datos(intIndice) = datos(intIndice) & "<NUMERO_LOTE>" & rs(12) & "-" & rs(3) & "/" & Format(rs(2), "yyyy") & "</NUMERO_LOTE>"
            datos(intIndice) = datos(intIndice) & "<FECHA_CADUCIDAD>" & Format(rs(4), "dd/mm/yyyy") & "</FECHA_CADUCIDAD>"
            datos(intIndice) = datos(intIndice) & "<CAPACIDAD>" & rs(5) & "</CAPACIDAD>"
            datos(intIndice) = datos(intIndice) & "<EADS>" & rs(6) & "</EADS>"
            datos(intIndice) = datos(intIndice) & "<ETIQUETA_ID>" & rs(8) & "</ETIQUETA_ID>"
            datos(intIndice) = datos(intIndice) & "<MOSTRAR>1</MOSTRAR>"
            datos(intIndice) = datos(intIndice) & "<FRASES_PEQ>" & formatearFrase(oAE.getFRASES_PEQ) & "</FRASES_PEQ>"
            datos(intIndice) = datos(intIndice) & "<FRASES_MED>" & formatearFrase(oAE.getFRASES_MED) & "</FRASES_MED>"
            datos(intIndice) = datos(intIndice) & "<FRASES_GRA>" & formatearFrase(oAE.getFRASES_GRA) & "</FRASES_GRA>"
            datos(intIndice) = datos(intIndice) & "<COMPONENTES>" & formatearFrase(oAE.getCOMPONENTES) & "</COMPONENTES>"
            datos(intIndice) = datos(intIndice) & "<ADVERTENCIA>" & formatearFrase(oAE.getADVERTENCIA) & "</ADVERTENCIA>"
            
            Dim h As Integer
            For h = 1 To 5
                datos(intIndice) = datos(intIndice) & "<PIC" & h & ">" & pictogramas(h) & "</PIC" & h & ">"
            Next
            datos(intIndice) = datos(intIndice) & "<IDIOMA>" & oD.getPARAMETROS & "</IDIOMA>"
            datos(intIndice) = datos(intIndice) & "</etiqueta>"
            
        Next cont
        rs.MoveNext
    Wend
    
    
    For cont = 1 To 3
        If Trim(datos(cont)) <> "" Then
            Set tx = gFSO.CreateTextFile((ReadINI(App.Path & "\config.ini", "Documentos", "Reportes") & "\Alodine\datos_t" & CStr(cont) & ".xml"), True)
'            Set tx = gFSO.CreateTextFile("\\servidor\canagrosa\reportes\Alodine\datos_t" & CStr(cont) & ".xml", True, True)
            
'            tx.WriteLine Replace("<?xml version='1.0' encoding='UTF-16'?>", "'", Chr(34))
            tx.WriteLine Replace("<?xml version='1.0' encoding='ISO-8859-1'?>", "'", Chr(34))
            tx.WriteLine "<lista_etiquetas>"
            
            tx.WriteLine datos(cont)
            
            tx.WriteLine "</lista_etiquetas>"
            tx.Close
        
            ' AHORA MUESTRA EL INFORME
    
        End If
    Next
    For cont = 1 To 3
        If Trim(datos(cont)) <> "" Then
    
            With objRpt(cont)
                .Caption = "Etiquetas de alodine, tamaño " & cont
                .iniciar
                .informe = "Alodine\rptEtiquetasAlodine_t" & CStr(cont)
                '.CRITERIO = ""
                .imprimir = False
                .generar
                '.Visible = True
                .Show '1
            End With
            'Unload objRpt
'            Set objRpt(cont) = Nothing
        End If
                
    Next cont
    
    Set objRpt(1) = Nothing
    Set objRpt(2) = Nothing
    Set objRpt(3) = Nothing
    
    Set tx = Nothing
    Set rs = Nothing

   On Error GoTo 0
   Exit Sub

Imprimir_Etiquetas_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Imprimir_Etiquetas of Módulo de clase clsAlodine_planificacion"
    
End Sub
Public Function formatearFrase(FRASE As String) As String
    Dim s As String
    s = FRASE
    s = Replace(s, "<negrita>", "&lt;b&gt;")
    s = Replace(s, "</negrita>", "&lt;/b&gt;")
    s = Replace(s, vbCrLf, "&lt;br&gt;")
    formatearFrase = s
End Function

Public Sub Imprimir_Etiquetas_Lote(ByVal prm_id_lote As Long, ByVal prm_id_cliente As Long, ByVal prm_id_capacidad As Long, idioma As Long)
    Dim rs As New ADODB.Recordset
    Dim tx As TextStream
    Dim blnNEtiquetasOk As Boolean
    Dim cont As Long, total_etiquetas As Long, n_etiquetas As String
    Dim objRpt As New frmReport
    Dim datos As String ' donde guarda los xml de los tres tamaños
    Dim intIndice As Integer
        
    Set rs = Me.Listado_Etiquetas_Lote_Cliente_Capacidad(prm_id_lote, prm_id_cliente, prm_id_capacidad)
    
    If rs.RecordCount = 0 Then
        Set rs = Nothing
        Set tx = Nothing
        Set objRpt = Nothing
        Exit Sub
    End If
    
        
    'Debug.Print pegatina
    
    rs.MoveFirst
    ' NOTA: no itera, sino que se posiciona sobre el primer registro directamente (debe retornar una sola fila conlos parámetros de cliente-lote-capacidad)
    
    total_etiquetas = rs(7)
    n_etiquetas = total_etiquetas
        
    blnNEtiquetasOk = False
    If total_etiquetas = 1 Then blnNEtiquetasOk = True
        
    While Not blnNEtiquetasOk
        n_etiquetas = total_etiquetas
        n_etiquetas = InputBox("Indique el total de etiquetas que desa imprimir para este Cliente de este Lote (max. " & CStr(total_etiquetas) & "): ", "Impresión de Etiquetas Alodine", n_etiquetas)
        If Trim(n_etiquetas) = "" Then ' ha cancelado
            Set rs = Nothing
            Set tx = Nothing
            Set objRpt = Nothing
            Exit Sub
        Else
            If IsNumeric(n_etiquetas) Then ' que sea un numero
                If CLng(n_etiquetas) <= total_etiquetas And CLng(n_etiquetas) > 0 Then ' Que sea igual o menor que el total de etiquetas
                    blnNEtiquetasOk = True
                End If
            End If
        End If
    Wend
    total_etiquetas = CLng(n_etiquetas)
    intIndice = rs(10) ' Tamaño de la etiqueta
    datos = ""
    ' establece la posicion inicial de impresion de pegatinas
    ' SI Y SOLO SI el tipo de pegatinas es 1 (etiquetas pequeñas). IntIndice => alodine_capacidad.tamano_pegatina_id
'    If pegatina > 1 And intIndice = 1 Then
    If intIndice = 1 Then
        ' solicita la posicion inicial de la primera etiqueta
        frmPosicionPegatina.Show 1
        If pegatina = 0 Then Exit Sub
        
        For cont = 1 To pegatina - 1
            datos = datos & "<etiqueta>"
            datos = datos & "<ID_ALODINE></ID_ALODINE>"
            datos = datos & "<PRODUCTO></PRODUCTO>"
            datos = datos & "<NORMA></NORMA>"
            datos = datos & "<DESCRIPCION></DESCRIPCION>"
            datos = datos & "<FECHA_ALTA></FECHA_ALTA>"
            datos = datos & "<NUMERO_LOTE></NUMERO_LOTE>"
            datos = datos & "<FECHA_CADUCIDAD></FECHA_CADUCIDAD>"
            datos = datos & "<CAPACIDAD></CAPACIDAD>"
            datos = datos & "<EADS></EADS>"
            datos = datos & "<ETIQUETA_ID></ETIQUETA_ID>"
            datos = datos & "<MOSTRAR>0</MOSTRAR>"
            datos = datos & "<IDIOMA></IDIOMA>"
            datos = datos & "</etiqueta>"
        Next cont
    End If
    Dim oAE As New clsAlodine_etiqueta
    Dim oD As New clsDecodificadora
    oD.Carga_valor DECODIFICADORA.IDIOMAS, idioma
    oAE.Carga rs(9), oD.getPARAMETROS
    ' Normas
    Dim oAlodine As New clsAlodine
    Dim normas As String
    normas = oAlodine.getNormas(rs(9))
    
    ' Pictogramas
    Dim ubicacionPic As String
    ubicacionPic = ReadINI(App.Path & "\config.ini", "Otros", "Pictogramas")
    Dim pos As Integer
    Dim pictogramas(1 To 5) As String
    pos = 1
    If oAE.getPIC1 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC1.jpg"
        pos = pos + 1
    End If
    If oAE.getPIC2 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC2.jpg"
        pos = pos + 1
    End If
    If oAE.getPIC3 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC3.jpg"
        pos = pos + 1
    End If
    If oAE.getPIC4 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC4.jpg"
        pos = pos + 1
    End If
    If oAE.getPIC5 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC5.jpg"
        pos = pos + 1
    End If
    If oAE.getPIC6 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC6.jpg"
        pos = pos + 1
    End If
    
    
    
    For cont = 1 To total_etiquetas
        datos = datos & "<etiqueta>"
        datos = datos & "<ID_ALODINE>" & rs(9) & "</ID_ALODINE>"
        datos = datos & "<PRODUCTO>" & rs(0) & "</PRODUCTO>"
        datos = datos & "<NORMA>" & normas & "</NORMA>"
'        datos = datos & "<NORMA>" & rs(14) & "</NORMA>"
        datos = datos & "<DESCRIPCION>" & rs(1) & "</DESCRIPCION>"
        datos = datos & "<FECHA_ALTA>" & Format(rs(2), "dd/mm/yyyy") & "</FECHA_ALTA>"
        datos = datos & "<NUMERO_LOTE>" & rs(13) & "-" & rs(3) & "/" & Format(rs(2), "yyyy") & "</NUMERO_LOTE>"
        datos = datos & "<FECHA_CADUCIDAD>" & Format(rs(4), "dd/mm/yyyy") & "</FECHA_CADUCIDAD>"
        datos = datos & "<CAPACIDAD>" & rs(5) & "</CAPACIDAD>"
        datos = datos & "<EADS>" & rs(6) & "</EADS>"
        datos = datos & "<ETIQUETA_ID>" & rs(8) & "</ETIQUETA_ID>"
        datos = datos & "<MOSTRAR>1</MOSTRAR>"
        datos = datos & "<FRASES_PEQ>" & formatearFrase(oAE.getFRASES_PEQ) & "</FRASES_PEQ>"
        datos = datos & "<FRASES_MED>" & formatearFrase(oAE.getFRASES_MED) & "</FRASES_MED>"
        datos = datos & "<FRASES_GRA>" & formatearFrase(oAE.getFRASES_GRA) & "</FRASES_GRA>"
        datos = datos & "<COMPONENTES>" & formatearFrase(oAE.getCOMPONENTES) & "</COMPONENTES>"
        Dim h As Integer
        For h = 1 To 5
            datos = datos & "<PIC" & h & ">" & pictogramas(h) & "</PIC" & h & ">"
        Next
        datos = datos & "<IDIOMA>" & oD.getPARAMETROS & "</IDIOMA>"
        datos = datos & "</etiqueta>"
        
    Next cont
    
    
    
    If Trim(datos) <> "" Then
        Set tx = gFSO.CreateTextFile((ReadINI(App.Path & "\config.ini", "Documentos", "Reportes") & "\Alodine\datos_t" & CStr(intIndice) & ".xml"), True)
        tx.WriteLine Replace("<?xml version='1.0' encoding='ISO-8859-1'?>", "'", Chr(34))
        tx.WriteLine "<lista_etiquetas>"
        tx.WriteLine datos
        tx.WriteLine "</lista_etiquetas>"
        tx.Close
        ' AHORA MUESTRA EL INFORME
        With objRpt
            .iniciar
            .informe = "Alodine\rptEtiquetasAlodine_t" & CStr(intIndice)
            '.CRITERIO = ""
            .imprimir = False
            .generar
            '.Visible = True
            .Show 1
        End With
        
        Unload objRpt
        Set objRpt = Nothing
    
    End If
                    
    
    Set tx = Nothing
    Set rs = Nothing
    


    
End Sub


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSALODINE_PLANIFICACION
'***************************************************************
Public Property Let setLOTE_ID(ByVal dato As Long)
        LOTE_ID = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setEADS(ByVal dato As Long)
        EADS = dato
End Property
Public Property Let setCAPACIDAD_ID(ByVal dato As Long)
        CAPACIDAD_ID = dato
End Property
Public Property Let setETIQUETA_ID(ByVal dato As Long)
        ETIQUETA_ID = dato
End Property
Public Property Let setNUMERO_BOTES(ByVal dato As Long)
        NUMERO_BOTES = dato
End Property
Public Property Let setDOC_ID(ByVal dato As Long)
        DOC_ID = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
Public Property Let setPEDIDO(ByVal dato As String)
        pedido = dato
End Property
Public Property Let setPEDIDO_ID(ByVal dato As Long)
        PEDIDO_ID = dato
End Property
Public Property Let setNORMA(ByVal dato As String)
        NORMA = dato
End Property
Public Property Let setNORMA_ETIQUETA(ByVal dato As Integer)
        NORMA_ETIQUETA = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setGRUPO(ByVal dato As Integer)
        grupo = dato
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSALODINE_PLANIFICACION
'***************************************************************
Public Property Get getLOTE_ID() As Long
        getLOTE_ID = LOTE_ID
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getEADS() As Long
        getEADS = EADS
End Property
Public Property Get getCAPACIDAD_ID() As Long
        getCAPACIDAD_ID = CAPACIDAD_ID
End Property
Public Property Get getETIQUETA_ID() As Long
        getETIQUETA_ID = ETIQUETA_ID
End Property
Public Property Get getNUMERO_BOTES() As Long
        getNUMERO_BOTES = NUMERO_BOTES
End Property
Public Property Get getDOC_ID() As Long
        getDOC_ID = DOC_ID
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
Public Property Get getPEDIDO() As String
        getPEDIDO = pedido
End Property
Public Property Get getPEDIDO_ID() As Long
        getPEDIDO_ID = PEDIDO_ID
End Property
Public Property Get getNORMA() As String
        getNORMA = NORMA
End Property
Public Property Get getNORMA_ETIQUETA() As Integer
        getNORMA_ETIQUETA = NORMA_ETIQUETA
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getGRUPO() As Integer
        getGRUPO = grupo
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSALODINE_PLANIFICACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_PLANIFICACION WHERE LOTE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                LOTE_ID = rs("LOTE_ID")
                CLIENTE_ID = rs("CLIENTE_ID")
                EADS = rs("EADS")
                CAPACIDAD_ID = rs("CAPACIDAD_ID")
                ETIQUETA_ID = rs("ETIQUETA_ID")
                NUMERO_BOTES = rs("NUMERO_BOTES")
                DOC_ID = rs("DOC_ID")
                PRECIO = rs("PRECIO")
                pedido = rs("PEDIDO")
                PEDIDO_ID = rs("PEDIDO_ID")
                NORMA = rs("NORMA")
                NORMA_ETIQUETA = rs("NORMA_ETIQUETA")
                fecha = rs("FECHA")
                grupo = rs("GRUPO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAlodine_planificacion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO ALODINE_PLANIFICACION " & _
                   "  VALUES (" & _
                        LOTE_ID & "," & CLIENTE_ID & "," & EADS & "," & _
                        CAPACIDAD_ID & "," & ETIQUETA_ID & "," & NUMERO_BOTES & "," & DOC_ID & ",'" & PRECIO & _
                "','" & pedido & "'," & PEDIDO_ID & ",'" & NORMA & "'," & NORMA_ETIQUETA & ",'" & Format(fecha, "yyyy-mm-dd") & "'," & grupo & ")"
        execute_bd consulta
        Insertar = LOTE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAlodine_planificacion)", vbCritical, Err.Description
End Function
Public Function Modificar(LOTE As Long, cliente As Long, Capacidad As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
                        
'        consulta = "SELECT * FROM ALODINE_PLANIFICACION " & _
'                   " WHERE LOTE_ID = " & LOTE & _
'                   "   AND CLIENTE_ID = " & cliente & _
'                   "   AND CAPACIDAD_ID = " & Capacidad
'
'        Set rs = datos_bd(consulta)
'        If rs.RecordCount > 0 Then
'            consulta = "UPDATE ALODINE_PLANIFICACION SET " & _
'                        " EADS = " & EADS & "," & _
'                        " ETIQUETA_ID = " & ETIQUETA_ID & "," & _
'                        " NUMERO_BOTES = " & NUMERO_BOTES & "," & _
'                        " PRECIO = '" & PRECIO & "'," & _
'                        " PEDIDO = '" & pedido & "'," & _
'                        " NORMA = '" & norma & "'," & _
'                        " FECHA = '" & Format(fecha, "yyyy-mm-dd") & "'" & _
'                      " WHERE LOTE_ID = " & LOTE & _
'                      " AND CLIENTE_ID = " & cliente & _
'                      " AND CAPACIDAD_ID = " & Capacidad
'            execute_bd consulta
'        Else
            Me.setGRUPO = 0
            Me.setLOTE_ID = LOTE
            Me.setCLIENTE_ID = cliente
            Me.setCAPACIDAD_ID = Capacidad
            Me.Insertar
'        End If
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAlodine_planificacion)", vbCritical, Err.Description
End Function
Public Function InformarGrupo(LOTE As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Dim grupo As Integer
        grupo = 0
        consulta = "SELECT LOTE_ID,CLIENTE_ID,NORMA,FECHA FROM ALODINE_PLANIFICACION " & _
                   " WHERE LOTE_ID = " & LOTE & _
                   " GROUP BY cliente_id, norma, fecha " & _
                   " ORDER BY cliente_id "

        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            Do
                consulta = "UPDATE ALODINE_PLANIFICACION SET GRUPO = " & grupo & _
                           " WHERE LOTE_ID = " & rs(0) & _
                           "   AND CLIENTE_ID = " & rs(1) & _
                           "   AND NORMA = '" & rs(2) & "'" & _
                           "   AND FECHA = '" & Format(rs(3), "yyyy-mm-dd") & "'"
                execute_bd consulta
                grupo = grupo + 1
                rs.MoveNext
            Loop Until rs.EOF
        End If
        InformarGrupo = True
        Exit Function
fallo:
        InformarGrupo = False
        MsgBox "Error al InformarGrupo (clsAlodine_planificacion)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ALODINE_PLANIFICACION " & _
                "    WHERE LOTE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAlodine_planificacion)", vbCritical, Err.Description
End Function
Public Function EliminarNoFacturados(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ALODINE_PLANIFICACION " & _
                "    WHERE LOTE_ID = " & ID & " AND DOC_ID = 0"
        execute_bd consulta
        EliminarNoFacturados = True
        Exit Function
fallo:
        EliminarNoFacturados = False
        MsgBox "Error al Eliminar (clsAlodine_planificacion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_PLANIFICACION ORDER BY LOTE_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Lote(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_PLANIFICACION WHERE LOTE_ID = " & ID & " ORDER BY CLIENTE_ID"
        Set Listado_Lote = datos_bd(consulta)
End Function
Public Function Numero_Clientes_Lote(ID As Long) As Integer
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT * " & _
                   "  FROM ALODINE_PLANIFICACION " & _
                   " WHERE LOTE_ID = " & ID & _
                   " GROUP BY CLIENTE_ID"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            Numero_Clientes_Lote = 0
        Else
            Numero_Clientes_Lote = rs.RecordCount
        End If
        Set rs = Nothing
End Function
Public Function Listado_Clientes_Lote(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT C.NOMBRE,AC.DESCRIPCION,AE.DESCRIPCION,A.NUMERO_BOTES,A.PRECIO,A.CLIENTE_ID,A.CAPACIDAD_ID, " & _
                   "       A.ETIQUETA_ID,A.EADS,C.ID_CLIENTE,A.PEDIDO,A.DOC_ID,A.NORMA,A.FECHA,A.NORMA_ETIQUETA,IFNULL(CONCAT(DP.NUMERO,'/',YEAR(DP.FECHA_FACTURA)),'') AS FACTURA,A.PEDIDO_ID " & _
                   "  FROM ALODINE_PLANIFICACION A " & _
                   " INNER JOIN CLIENTES C ON A.CLIENTE_ID = C.ID_CLIENTE " & _
                   " INNER JOIN ALODINE_CAPACIDAD AC ON A.CAPACIDAD_ID = AC.ID_CAPACIDAD " & _
                   " INNER JOIN ALODINE_ETIQUETAS AE ON A.ETIQUETA_ID = AE.ID_ETIQUETA " & _
                   " LEFT JOIN docs_pago DP ON A.DOC_ID = DP.ID_DOC " & _
                   " WHERE A.LOTE_ID = " & ID & _
                   " ORDER BY C.NOMBRE"
        Set Listado_Clientes_Lote = datos_bd(consulta)
End Function
Public Function Listado_Lotes_Informacion(lotes As String, clientes As String, pedidos As String, EADS As Integer) As ADODB.Recordset
    Dim consulta As String
    consulta = " SELECT AP.CLIENTE_ID, AP.PEDIDO_ID,C.NOMBRE,CP.CODIGO " & _
                " FROM ALODINE_PLANIFICACION AP  " & _
                " INNER JOIN CLIENTES C ON AP.CLIENTE_ID = C.ID_CLIENTE " & _
                " LEFT JOIN CLIENTES_PEDIDOS CP ON AP.PEDIDO_ID = CP.ID_PEDIDO " & _
                " WHERE AP.LOTE_ID IN (" & lotes & ")" & _
                " AND AP.CLIENTE_ID IN (" & clientes & ")" & _
                " AND AP.PEDIDO_ID IN (" & pedidos & ")" & _
                " AND AP.EADS = " & EADS & _
                " AND AP.DOC_ID = 0 " & _
                " GROUP BY AP.CLIENTE_ID,AP.PEDIDO_ID " & _
                " ORDER BY AP.CLIENTE_ID"
    Set Listado_Lotes_Informacion = datos_bd(consulta)
End Function

Public Function Listado_Lotes_Grupo(lotes As String, clientes As String, pedidos As String) As ADODB.Recordset
        Dim consulta As String
'        consulta = "SELECT AP.CLIENTE_ID, A.PRODUCTO,AC.DESCRIPCION,SUM(NUMERO_BOTES), SUM(AP.NUMERO_BOTES * AP.PRECIO), AL.FECHA_ALTA,AP.LOTE_ID "
        consulta = "SELECT AP.CLIENTE_ID, A.PRODUCTO,AC.DESCRIPCION,SUM(NUMERO_BOTES), SUM(AP.NUMERO_BOTES * AP.PRECIO), AL.FECHA_ALTA,AP.LOTE_ID,AP.PRECIO,AP.NORMA,AP.FECHA,AP.PEDIDO_ID " & _
                   "  FROM ALODINE_PLANIFICACION AP,ALODINE A,ALODINE_LOTES AS AL, ALODINE_CAPACIDAD AS AC " & _
                   " WHERE AP.LOTE_ID IN (" & lotes & ")" & _
                   "   AND AP.CLIENTE_ID IN (" & clientes & ")" & _
                   "   AND AP.PEDIDO_ID IN (" & pedidos & ")" & _
                   "   AND AP.DOC_ID = 0 " & _
                   "   AND AP.LOTE_ID = AL.ID_LOTE " & _
                   "   AND AL.ALODINE_ID = A.ID_ALODINE " & _
                   "   AND AP.CAPACIDAD_ID = AC.ID_CAPACIDAD " & _
                   "   AND AP.EADS = 0 " & _
                   " GROUP BY AP.CLIENTE_ID,A.PRODUCTO,AP.CAPACIDAD_ID, AL.FECHA_ALTA " & _
                   " ORDER BY AP.CLIENTE_ID,AP.PEDIDO_ID,AL.FECHA_ALTA"
        Set Listado_Lotes_Grupo = datos_bd(consulta)
End Function
Public Function Listado_Lotes_Grupo_EADS(lotes As String, clientes As String, pedidos As String) As ADODB.Recordset
        Dim consulta As String
'        consulta = "SELECT AP.CLIENTE_ID, A.PRODUCTO,AC.DESCRIPCION,SUM(NUMERO_BOTES), SUM(AP.NUMERO_BOTES * AP.PRECIO), AL.FECHA_ALTA, C.NOMBRE, AP.LOTE_ID "
        consulta = "SELECT AP.CLIENTE_ID, A.PRODUCTO,AC.DESCRIPCION,SUM(NUMERO_BOTES),SUM(AP.NUMERO_BOTES * AP.PRECIO), AL.FECHA_ALTA, C.NOMBRE, AP.LOTE_ID,AP.PRECIO,AP.PEDIDO_ID " & _
                   "  FROM ALODINE_PLANIFICACION AP,ALODINE A,ALODINE_LOTES AS AL, ALODINE_CAPACIDAD AS AC, CLIENTES C " & _
                   " WHERE AP.LOTE_ID IN (" & lotes & ")" & _
                   "   AND AP.CLIENTE_ID IN (" & clientes & ")" & _
                   "   AND AP.PEDIDO_ID IN (" & pedidos & ")" & _
                   "   AND AP.EADS = 1 " & _
                   "   AND AP.DOC_ID = 0 " & _
                   "   AND AP.LOTE_ID = AL.ID_LOTE " & _
                   "   AND AL.ALODINE_ID = A.ID_ALODINE " & _
                   "   AND AP.CAPACIDAD_ID = AC.ID_CAPACIDAD " & _
                   "   AND AP.CLIENTE_ID = C.ID_CLIENTE " & _
                   " GROUP BY AP.CLIENTE_ID,A.PRODUCTO,AP.CAPACIDAD_ID, AL.FECHA_ALTA " & _
                   " ORDER BY AP.CLIENTE_ID,AP.PEDIDO_ID,AL.FECHA_ALTA"
        Set Listado_Lotes_Grupo_EADS = datos_bd(consulta)
End Function
Public Function Numero_Etiquetas_Lote(ID As Long) As Integer
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT SUM(NUMERO_BOTES) " & _
                   "  FROM ALODINE_PLANIFICACION " & _
                   " WHERE LOTE_ID = " & ID
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            Numero_Etiquetas_Lote = 0
        Else
            Numero_Etiquetas_Lote = rs.Fields(0)
        End If
        Set rs = Nothing
End Function

Public Function Listado_Etiquetas_Lote(ID As Long) As ADODB.Recordset
        Dim consulta As String
'        consulta = "SELECT IF(AP.NORMA_ETIQUETA=1,CONCAT(A.PRODUCTO, ' según ',AP.NORMA),A.PRODUCTO),A.DESCRIPCION,AL.FECHA_ALTA,AL.NUMERO_LOTE,AL.FECHA_CADUCIDAD,AC.DESCRIPCION,AP.EADS,AP.NUMERO_BOTES,AP.ETIQUETA_ID, A.ID_ALODINE, AC.TAMANO_ETIQUETA_ID, AP.NORMA_ETIQUETA ,A.MADRE"
        consulta = "SELECT A.PRODUCTO,A.DESCRIPCION,AL.FECHA_ALTA,AL.NUMERO_LOTE,AL.FECHA_CADUCIDAD,AC.DESCRIPCION,AP.EADS,AP.NUMERO_BOTES,AP.ETIQUETA_ID, A.ID_ALODINE, AC.TAMANO_ETIQUETA_ID, AP.NORMA_ETIQUETA ,A.MADRE,AP.NORMA" & _
                   "  FROM ALODINE_PLANIFICACION AP,ALODINE_LOTES AL, ALODINE_CAPACIDAD AC,ALODINE_ETIQUETAS AE,ALODINE A, CLIENTES CLI " & _
                   " WHERE AP.LOTE_ID = " & ID & _
                   "   AND AP.LOTE_ID = AL.ID_LOTE " & _
                   "   AND AP.CAPACIDAD_ID = AC.ID_CAPACIDAD " & _
                   "   AND AP.ETIQUETA_ID = AE.ID_ETIQUETA " & _
                   "   AND AL.ALODINE_ID = A.ID_ALODINE " & _
                   "   AND AP.CLIENTE_ID = CLI.ID_CLIENTE " & _
                   " ORDER BY CLI.NOMBRE "
'                   " ORDER BY AP.CAPACIDAD_ID, AP.CLIENTE_ID"
        Set Listado_Etiquetas_Lote = datos_bd(consulta)
End Function
'Public Function Listado_Etiquetas_Lote_Cliente(ID As Long, cliente As Long) As ADODB.Recordset
'        Dim consulta As String
'        consulta = "SELECT IF(AP.NORMA_ETIQUETA=1,CONCAT(A.PRODUCTO, ' según ',AP.NORMA),A.PRODUCTO),A.DESCRIPCION,AL.FECHA_ALTA,AL.NUMERO_LOTE,AL.FECHA_CADUCIDAD,AC.DESCRIPCION,AP.EADS,AP.NUMERO_BOTES " & _
'                   "  FROM ALODINE_PLANIFICACION AP,ALODINE_LOTES AL, ALODINE_CAPACIDAD AC,ALODINE_ETIQUETAS AE,ALODINE A " & _
'                   " WHERE AP.LOTE_ID = " & ID & _
'                   "   AND AP.LOTE_ID = AL.ID_LOTE " & _
'                   "   AND AP.CAPACIDAD_ID = AC.ID_CAPACIDAD " & _
'                   "   AND AP.ETIQUETA_ID = AE.ID_ETIQUETA " & _
'                   "   AND AL.ALODINE_ID = A.ID_ALODINE " & _
'                   "   AND AP.CLIENTE_ID = " & cliente & _
'                   " ORDER BY AP.CLIENTE_ID"
'        Set Listado_Etiquetas_Lote_Cliente = datos_bd(consulta)
'End Function

Public Function Listado_Etiquetas_Lote_Cliente_Capacidad(ID As Long, cliente As Long, ByVal ID_CAPACIDAD As Long) As ADODB.Recordset
        Dim consulta As String
'        consulta = "SELECT IF(AP.NORMA_ETIQUETA=1,CONCAT(A.PRODUCTO, ' según ',AP.NORMA),A.PRODUCTO),A.DESCRIPCION,AL.FECHA_ALTA,AL.NUMERO_LOTE,AL.FECHA_CADUCIDAD,AC.DESCRIPCION,AP.EADS,AP.NUMERO_BOTES,AP.ETIQUETA_ID, A.ID_ALODINE, AC.TAMANO_ETIQUETA_ID, AP.cliente_id, AP.CAPACIDAD_ID ,A.MADRE "
        consulta = "SELECT A.PRODUCTO,A.DESCRIPCION,AL.FECHA_ALTA,AL.NUMERO_LOTE,AL.FECHA_CADUCIDAD,AC.DESCRIPCION,AP.EADS,AP.NUMERO_BOTES,AP.ETIQUETA_ID, A.ID_ALODINE, AC.TAMANO_ETIQUETA_ID, AP.cliente_id, AP.CAPACIDAD_ID ,A.MADRE,AP.NORMA" & _
                   "  FROM ALODINE_PLANIFICACION AP,ALODINE_LOTES AL, ALODINE_CAPACIDAD AC,ALODINE_ETIQUETAS AE,ALODINE A " & _
                   " WHERE AP.LOTE_ID = " & ID & _
                   "   AND AP.LOTE_ID = AL.ID_LOTE " & _
                   "   AND AP.CAPACIDAD_ID = AC.ID_CAPACIDAD " & _
                   "   AND AP.ETIQUETA_ID = AE.ID_ETIQUETA " & _
                   "   AND AL.ALODINE_ID = A.ID_ALODINE " & _
                   "   AND AP.CLIENTE_ID = " & cliente & _
                   "   AND AP.capacidad_ID = " & ID_CAPACIDAD & _
                   " ORDER BY AP.CLIENTE_ID"
        Set Listado_Etiquetas_Lote_Cliente_Capacidad = datos_bd(consulta)
End Function

Public Function doc_pago(lotes As String, cliente As Long, PEDIDO_ID As Long, DOC As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ALODINE_PLANIFICACION " & _
                   "   SET DOC_ID = " & DOC & _
                   " WHERE LOTE_ID IN (" & lotes & ")" & _
                   "   AND EADS = 0 " & _
                   "   AND CLIENTE_ID = " & cliente & _
                   "   AND PEDIDO_ID = " & PEDIDO_ID & _
                   "   AND DOC_ID = 0"
        execute_bd consulta
        doc_pago = True
        Exit Function
fallo:
        doc_pago = False
        MsgBox "Error al informar el documento de pago (clsAlodine_planificacion)", vbCritical, Err.Description
End Function
Public Function doc_pago_eads(lotes As String, clientes As String, DOC As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ALODINE_PLANIFICACION " & _
                   "   SET DOC_ID = " & DOC & _
                   " WHERE LOTE_ID IN (" & lotes & ")" & _
                   "   AND CLIENTE_ID IN (" & clientes & ")" & _
                   "   AND EADS = 1" & _
                   "   AND DOC_ID = 0"
        execute_bd consulta
        doc_pago_eads = True
        Exit Function
fallo:
        doc_pago_eads = False
        MsgBox "Error al informar el documento de pago eads (clsAlodine_planificacion)", vbCritical, Err.Description
End Function
