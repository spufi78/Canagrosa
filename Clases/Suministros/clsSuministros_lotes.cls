VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSuministros_lotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSUMINISTROS_LOTES
'***************************************************************
Private ID_LOTE As Long
Private TIPO_SUMINISTRO_ID As Long
Private NUMERO_LOTE As String
Private fecha_fabricacion As String
Private fecha_caducidad As String
Private BOTE_PR_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSUMINISTROS_LOTES
'***************************************************************
Public Property Let setID_LOTE(ByVal dato As Long)
        ID_LOTE = dato
End Property
Public Property Let setTIPO_SUMINISTRO_ID(ByVal dato As Long)
        TIPO_SUMINISTRO_ID = dato
End Property
Public Property Let setNUMERO_LOTE(ByVal dato As String)
        NUMERO_LOTE = dato
End Property
Public Property Let setFECHA_FABRICACION(ByVal dato As String)
        fecha_fabricacion = dato
End Property
Public Property Let setFECHA_CADUCIDAD(ByVal dato As String)
        fecha_caducidad = dato
End Property
Public Property Let setBOTE_PR_ID(ByVal dato As Long)
        BOTE_PR_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSUMINISTROS_LOTES
'***************************************************************
Public Property Get getID_LOTE() As Long
        getID_LOTE = ID_LOTE
End Property
Public Property Get getTIPO_SUMINISTRO_ID() As Long
        getTIPO_SUMINISTRO_ID = TIPO_SUMINISTRO_ID
End Property
Public Property Get getNUMERO_LOTE() As String
        getNUMERO_LOTE = NUMERO_LOTE
End Property
Public Property Get getFECHA_FABRICACION() As String
        getFECHA_FABRICACION = fecha_fabricacion
End Property
Public Property Get getFECHA_CADUCIDAD() As String
        getFECHA_CADUCIDAD = fecha_caducidad
End Property
Public Property Get getBOTE_PR_ID() As Long
        getBOTE_PR_ID = BOTE_PR_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSUMINISTROS_LOTES
'***************************************************************
Public Function carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SUMINISTROS_LOTES WHERE ID_LOTE = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                carga = False
        Else
                ID_LOTE = rs("ID_LOTE")
                TIPO_SUMINISTRO_ID = rs("TIPO_SUMINISTRO_ID")
                NUMERO_LOTE = rs("NUMERO_LOTE")
                fecha_fabricacion = rs("FECHA_FABRICACION")
                fecha_caducidad = rs("FECHA_CADUCIDAD")
                BOTE_PR_ID = rs("BOTE_PR_ID")
                carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        carga = False
        MsgBox "Error al cargar los datos(clsSuministros_lotes)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_LOTE) FROM SUMINISTROS_LOTES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_LOTE = 1
        Else
                ID_LOTE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO SUMINISTROS_LOTES " & _
                   "  VALUES (" & _
                        ID_LOTE & "," & TIPO_SUMINISTRO_ID & ",'" & NUMERO_LOTE & "'," & _
                        "'" & fecha_fabricacion & "'" & "," & "'" & fecha_caducidad & "'" & "," & BOTE_PR_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_LOTE
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSuministros_lotes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SUMINISTROS_LOTES SET " & _
                        " TIPO_SUMINISTRO_ID = " & TIPO_SUMINISTRO_ID & "," & _
                        " NUMERO_LOTE = '" & NUMERO_LOTE & "'," & _
                        " FECHA_FABRICACION = '" & fecha_fabricacion & "'," & _
                        " FECHA_CADUCIDAD = '" & fecha_caducidad & "'," & _
                        " BOTE_PR_ID = " & BOTE_PR_ID & "" & _
                " WHERE ID_LOTE = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSuministros_lotes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT * FROM SUMINISTROS_LOTES_CLIENTES WHERE LOTE_ID = " & ID & " AND DOC_ID <> 0 "
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MsgBox "El lote esta insertado en facturación. No se puede eliminar.", vbCritical, App.Title
            Exit Function
        End If
        consulta = "DELETE FROM SUMINISTROS_LOTES_CLIENTES WHERE LOTE_ID = " & ID
        execute_bd consulta
        consulta = "DELETE FROM SUMINISTROS_LOTES WHERE ID_LOTE = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSuministros_lotes)", vbCritical, Err.Description
End Function
Public Function Listado(LOTE As String, producto As String, cliente As String, fdesde As String, fhasta As String, caducidad As Boolean, fcaddesde As String, fcadhasta As String) As ADODB.Recordset
        Dim consulta As String
        Dim s As String
        If LOTE <> "" Then
            s = s & " AND SL.NUMERO_LOTE = '" & LOTE & "'"
        End If
        If producto <> "" Then
            s = s & " AND SL.TIPO_SUMINISTRO_ID = " & CLng(producto)
        End If
        If cliente <> "" Then
            s = s & " AND SLC.CLIENTE_ID = " & CLng(cliente)
        End If
        s = s & " AND SL.FECHA_FABRICACION BETWEEN '" & Format(fdesde, "yyyy-mm-dd") & "' AND '" & Format(fhasta, "yyyy-mm-dd") & "' "
        If caducidad Then
            s = s & " AND SL.FECHA_CADUCIDAD BETWEEN '" & Format(fcaddesde, "yyyy-mm-dd") & "' AND '" & Format(fcadhasta, "yyyy-mm-dd") & "' "
        End If
        consulta = "SELECT ST.PRODUCTO, SL.NUMERO_LOTE,SL.FECHA_FABRICACION, SL.FECHA_CADUCIDAD, SL.ID_LOTE " & _
                   "  FROM SUMINISTROS_LOTES SL, SUMINISTROS_TIPOS ST " & _
                   IIf(cliente <> "", ",SUMINISTROS_LOTES_CLIENTES SLC ", "") & _
                   " WHERE SL.TIPO_SUMINISTRO_ID = ST.ID_TIPO_SUMINISTRO " & _
                   IIf(cliente <> "", " AND SL.ID_LOTE = SLC.LOTE_ID ", "") & _
                   s & _
                   " ORDER BY SL.FECHA_FABRICACION DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_por_ID(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ST.PRODUCTO, SL.NUMERO_LOTE,SL.FECHA_FABRICACION, SL.FECHA_CADUCIDAD, SL.ID_LOTE " & _
                   "  FROM SUMINISTROS_LOTES SL, SUMINISTROS_TIPOS ST " & _
                   " WHERE SL.TIPO_SUMINISTRO_ID = ST.ID_TIPO_SUMINISTRO AND SL.ID_LOTE = " & ID
        Set Listado_por_ID = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_LOTE,TIPO_SUMINISTRO_ID FROM SUMINISTROS_LOTES ORDER BY TIPO_SUMINISTRO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Proximo_NUMERO_LOTE(TIPO_SUMINISTRO_ID As Long, ANNO As Integer) As String
        Dim rs As ADODB.Recordset
        Dim consulta As String
   On Error GoTo Proximo_NUMERO_LOTE_Error

        consulta = "SELECT MAX(CONVERT(NUMERO_LOTE,DECIMAL)) " & _
                   "  FROM SUMINISTROS_LOTES " & _
                   " WHERE YEAR(FECHA_FABRICACION) = " & ANNO & _
                   "   AND TIPO_SUMINISTRO_ID = " & TIPO_SUMINISTRO_ID
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                Proximo_NUMERO_LOTE = 1
        Else
                Proximo_NUMERO_LOTE = rs.Fields(0) + 1
        End If
'        Proximo_NUMERO_LOTE = Proximo_NUMERO_LOTE
        Set rs = Nothing

   On Error GoTo 0
   Exit Function

Proximo_NUMERO_LOTE_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Proximo_NUMERO_LOTE of Módulo de clase clsSuministros_lotes"
End Function
Public Function Listado_Etiquetas_Lote(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = " select st.PRODUCTO, ac.DESCRIPCION, sl.NUMERO_LOTE,sl.FECHA_FABRICACION,sl.FECHA_CADUCIDAD , slc.NUMERO_BOTES, sl.TIPO_SUMINISTRO_ID " & _
                   "  from suministros_lotes sl, suministros_tipos st, suministros_lotes_clientes slc, alodine_capacidad ac " & _
                   " Where sl.ID_LOTE = " & ID & _
                   "   and sl.TIPO_SUMINISTRO_ID=st.ID_TIPO_SUMINISTRO " & _
                   "   and sl.ID_LOTE = slc.LOTE_ID " & _
                   "   and slc.CAPACIDAD_ID = ac.ID_CAPACIDAD "
        Set Listado_Etiquetas_Lote = datos_bd(consulta)
End Function

Public Function ImprimirEtiquetas(LOTE As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim cont As Long
    Dim datos As String ' donde guarda los xml de los tres tamaños

    Set rs = Me.Listado_Etiquetas_Lote(LOTE)

    If rs.RecordCount = 0 Then
        Set rs = Nothing
        Exit Function
    End If
        
    frmPosicionPegatina.Show 1
    
    For cont = 1 To pegatina - 1
        datos = datos & "<etiqueta>"
        datos = datos & "<PRODUCTO></PRODUCTO>"
        datos = datos & "<CAPACIDAD></CAPACIDAD>"
        datos = datos & "<NUMERO_LOTE></NUMERO_LOTE>"
        datos = datos & "<FECHA_ALTA></FECHA_ALTA>"
        datos = datos & "<FECHA_CADUCIDAD></FECHA_CADUCIDAD>"
        datos = datos & "<MOSTRAR>0</MOSTRAR>"
        datos = datos & "<TIPO_SUMINISTRO_ID>0</TIPO_SUMINISTRO_ID>"
        datos = datos & "<FRASES_PEQ></FRASES_PEQ>"
        datos = datos & "<FRASES_MED></FRASES_MED>"
        datos = datos & "<FRASES_GRA></FRASES_GRA>"
        datos = datos & "<COMPONENTES></COMPONENTES>"
        datos = datos & "<ADVERTENCIA></ADVERTENCIA>"
        datos = datos & "<PIC1></PIC1>"
        datos = datos & "<PIC2></PIC2>"
        datos = datos & "<PIC3></PIC3>"
        datos = datos & "<PIC4></PIC4>"
        datos = datos & "<PIC5></PIC5>"
        datos = datos & "</etiqueta>"
    Next cont

    Dim oSe As New clsSuministros_etiqueta
    oSe.carga rs(6)
    ' Pictogramas
    Dim ubicacionPic As String
    ubicacionPic = ReadINI(App.Path & "\config.ini", "Otros", "Pictogramas")
    Dim pos As Integer
    Dim pictogramas(1 To 5) As String
    pos = 1
    If oSe.getPIC1 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC1.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC2 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC2.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC3 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC3.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC4 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC4.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC5 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC5.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC6 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC6.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC7 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC7.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC8 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC8.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC9 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC9.jpg"
        pos = pos + 1
    End If
    
    Do
        For cont = 1 To rs(5)
            datos = datos & "<etiqueta>"
            datos = datos & "<PRODUCTO>" & rs(0) & "</PRODUCTO>"
            datos = datos & "<CAPACIDAD>(Cont. " & rs(1) & ")</CAPACIDAD>"
            datos = datos & "<NUMERO_LOTE>" & rs(2) & "/" & Year(rs(3)) & "</NUMERO_LOTE>"
            datos = datos & "<FECHA_ALTA>" & Format(rs(3), "dd/mm/yyyy") & "</FECHA_ALTA>"
            datos = datos & "<FECHA_CADUCIDAD>" & Format(rs(4), "dd/mm/yyyy") & "</FECHA_CADUCIDAD>"
            datos = datos & "<MOSTRAR>1</MOSTRAR>"
            datos = datos & "<TIPO_SUMINISTRO_ID>" & rs(6) & "</TIPO_SUMINISTRO_ID>"
            datos = datos & "<FRASES_PEQ>" & formatearFrase(oSe.getFRASES_PEQ) & "</FRASES_PEQ>"
            datos = datos & "<FRASES_MED>" & formatearFrase(oSe.getFRASES_MED) & "</FRASES_MED>"
            datos = datos & "<FRASES_GRA>" & formatearFrase(oSe.getFRASES_GRA) & "</FRASES_GRA>"
            datos = datos & "<COMPONENTES>" & formatearFrase(oSe.getCOMPONENTES) & "</COMPONENTES>"
            datos = datos & "<ADVERTENCIA>" & formatearFrase(oSe.getADVERTENCIA) & "</ADVERTENCIA>"
            Dim i As Integer
            For i = 1 To 5
                datos = datos & "<PIC" & i & ">" & pictogramas(i) & "</PIC" & i & ">"
            Next
        
'            datos = datos & "<PIC1>" & oSe.getPIC1 & "</PIC1>"
'            datos = datos & "<PIC2>" & oSe.getPIC2 & "</PIC2>"
'            datos = datos & "<PIC3>" & oSe.getPIC3 & "</PIC3>"
'            datos = datos & "<PIC4>" & oSe.getPIC4 & "</PIC4>"
'            datos = datos & "<PIC5>" & oSe.getPIC5 & "</PIC5>"
'            datos = datos & "<PIC6>" & oSe.getPIC6 & "</PIC6>"
'            datos = datos & "<PIC7>" & oSe.getPIC7 & "</PIC7>"
'            datos = datos & "<PIC8>" & oSe.getPIC8 & "</PIC8>"
'            datos = datos & "<PIC9>" & oSe.getPIC9 & "</PIC9>"
            
            datos = datos & "</etiqueta>"
        Next
        rs.MoveNext
    Loop Until rs.EOF


    If Trim(datos) <> "" Then
       Dim tx As TextStream
'       Set tx = gFSO.CreateTextFile("\\servidor\canagrosa\reportes\suministros\datos.xml", True, True)
        Set tx = gFSO.CreateTextFile((ReadINI(App.Path & "\config.ini", "Documentos", "Reportes") & "\suministros\datos.xml"), True)
'        tx.WriteLine Replace("<?xml version='1.0' encoding='UTF-16'?>", "'", Chr(34))
        tx.WriteLine Replace("<?xml version='1.0' encoding='ISO-8859-1'?>", "'", Chr(34))
        tx.WriteLine "<lista_etiquetas>"
        tx.WriteLine datos
        tx.WriteLine "</lista_etiquetas>"
        tx.Close

        ' AHORA MUESTRA EL INFORME
        Dim objRpt As New frmReport

        With objRpt
            .iniciar
            .informe = "Suministros\rptSuministros_Etiquetas"
            '.CRITERIO = ""
            .imprimir = False
            .generar
            '.Visible = True
            .Show 1
        End With

        Unload objRpt
        Set objRpt = Nothing
        Set tx = Nothing

    End If


    Set rs = Nothing


End Function
Public Function formatearFrase(FRASE As String) As String
    Dim s As String
    s = FRASE
    s = Replace(s, "<negrita>", "&lt;b&gt;")
    s = Replace(s, "</negrita>", "&lt;/b&gt;")
    s = Replace(s, vbCrLf, "&lt;br&gt;")
    formatearFrase = s
End Function

Public Function Listado_Pendientes_Facturar(CLIENTE_ID As Long, fdesde As Date, fhasta As Date) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        If CLIENTE_ID <> 0 Then
            filtro = filtro & " AND SLC.CLIENTE_ID = " & CLIENTE_ID
        End If
        filtro = filtro & " AND SL.FECHA_FABRICACION >= '" & Format(fdesde, "yyyy-mm-dd") & "'"
        filtro = filtro & " AND SL.FECHA_FABRICACION <= '" & Format(fhasta, "yyyy-mm-dd") & "'"
    
        consulta = "SELECT C.NOMBRE,ST.PRODUCTO,SL.NUMERO_LOTE,SL.FECHA_FABRICACION,SL.FECHA_CADUCIDAD,SL.ID_LOTE,C.ID_CLIENTE,SLC.NUMERO_BOTES,SLC.PRECIO,IFNULL(CONCAT(CP.CODIGO,' ',CP.DESCRIPCION),''),SLC.PEDIDO_ID " & _
                   "  FROM SUMINISTROS_LOTES SL " & _
                   " INNER JOIN SUMINISTROS_TIPOS ST ON SL.TIPO_SUMINISTRO_ID = ST.ID_TIPO_SUMINISTRO " & _
                   " INNER JOIN SUMINISTROS_LOTES_CLIENTES SLC ON SL.ID_LOTE = SLC.LOTE_ID " & _
                   " INNER JOIN CLIENTES C ON SLC.CLIENTE_ID = C.ID_CLIENTE " & _
                   "  LEFT JOIN CLIENTES_PEDIDOS CP ON SLC.PEDIDO_ID = CP.ID_PEDIDO " & _
                   " WHERE SLC.DOC_ID = 0 " & _
                   filtro & _
                   " GROUP BY SL.ID_LOTE, SLC.CLIENTE_ID" & _
                   " ORDER BY SL.ID_LOTE ASC,SLC.CLIENTE_ID ASC"
        Set Listado_Pendientes_Facturar = datos_bd(consulta)
End Function

Public Function Listado_Pendientes_Facturar_ID(LOTE_ID As Long, CLIENTE_ID As Long) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        filtro = filtro & " AND SLC.LOTE_ID = " & LOTE_ID
        filtro = filtro & " AND SLC.CLIENTE_ID = " & CLIENTE_ID
    
        consulta = "SELECT C.NOMBRE,ST.PRODUCTO,SL.NUMERO_LOTE,SL.FECHA_FABRICACION,SL.FECHA_CADUCIDAD,SL.ID_LOTE,C.ID_CLIENTE,SLC.NUMERO_BOTES,SLC.PRECIO,CONCAT(CP.CODIGO,' ',CP.DESCRIPCION),SLC.PEDIDO_ID " & _
                   "  FROM SUMINISTROS_LOTES SL " & _
                   " INNER JOIN SUMINISTROS_TIPOS ST ON SL.TIPO_SUMINISTRO_ID = ST.ID_TIPO_SUMINISTRO " & _
                   " INNER JOIN SUMINISTROS_LOTES_CLIENTES SLC ON SL.ID_LOTE = SLC.LOTE_ID " & _
                   " INNER JOIN CLIENTES C ON SLC.CLIENTE_ID = C.ID_CLIENTE " & _
                   "  LEFT JOIN CLIENTES_PEDIDOS CP ON SLC.PEDIDO_ID = CP.ID_PEDIDO " & _
                   " WHERE SLC.DOC_ID = 0 " & _
                   filtro & _
                   " GROUP BY SL.ID_LOTE, SLC.CLIENTE_ID" & _
                   " ORDER BY SL.ID_LOTE ASC,SLC.CLIENTE_ID ASC"
        Set Listado_Pendientes_Facturar_ID = datos_bd(consulta)
End Function
Public Function Listado_Para_Factura(LOTE_ID As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT b.CLIENTE_ID,b.PEDIDO_ID, a.FECHA_FABRICACION, concat(c.PRODUCTO,'.',b.NUMERO_BOTES, ' bote/s de ',d.DESCRIPCION,'. Lote ',a.NUMERO_LOTE,'/',year(a.FECHA_FABRICACION)), " & _
                   "       b.NUMERO_BOTES , b.PRECIO, b.NUMERO_BOTES * b.PRECIO " & _
                   "  FROM SUMINISTROS_LOTES a " & _
                   " inner join suministros_lotes_clientes b on a.ID_LOTE = b.LOTE_ID " & _
                   " inner join suministros_tipos c on a.TIPO_SUMINISTRO_ID = c.ID_TIPO_SUMINISTRO " & _
                   " left join alodine_capacidad d on b.CAPACIDAD_ID = d.ID_CAPACIDAD " & _
                   "  LEFT JOIN CLIENTES_PEDIDOS CP ON b.PEDIDO_ID = CP.ID_PEDIDO " & _
                   " WHERE a.ID_LOTE in (" & LOTE_ID & ") " & _
                   "   AND b.DOC_ID = 0 " & _
                   " order by b.CLIENTE_ID, b.PEDIDO_ID "
        Set Listado_Para_Factura = datos_bd(consulta)
End Function
