VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDeterminaciones_analisis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDETERMINACIONES_ANALISIS
'***************************************************************
Private TIPO_DETERMINACION_ID As Long
Private TIPO_ANALISIS_ID As Long
Private BANO_ID As Long
Private ORDEN As Integer
Private MINIMO As String
Private MAXIMO As String
Private MINIMO_TEXTO As String
Private MAXIMO_TEXTO As String
Private LIM_MINIMO As String
Private LIM_MAXIMO As String
Private REQUERIDA As Integer
Private grafico As Integer
Private TIPO_FRECUENCIA_ID As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDETERMINACIONES_ANALISIS
'***************************************************************
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Long)
        TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setTIPO_ANALISIS_ID(ByVal dato As Long)
        TIPO_ANALISIS_ID = dato
End Property
Public Property Let setBANO_ID(ByVal dato As Long)
        BANO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
Public Property Let setMINIMO(ByVal dato As String)
        MINIMO = dato
End Property
Public Property Let setMAXIMO(ByVal dato As String)
        MAXIMO = dato
End Property
Public Property Let setMINIMO_TEXTO(ByVal dato As String)
        MINIMO_TEXTO = dato
End Property
Public Property Let setMAXIMO_TEXTO(ByVal dato As String)
        MAXIMO_TEXTO = dato
End Property
Public Property Let setLIM_MINIMO(ByVal dato As String)
        LIM_MINIMO = dato
End Property
Public Property Let setLIM_MAXIMO(ByVal dato As String)
        LIM_MAXIMO = dato
End Property
Public Property Let setREQUERIDA(ByVal dato As Integer)
        REQUERIDA = dato
End Property
Public Property Let setGRAFICO(ByVal dato As Integer)
        grafico = dato
End Property
Public Property Let setTIPO_FRECUENCIA_ID(ByVal dato As Integer)
        TIPO_FRECUENCIA_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDETERMINACIONES_ANALISIS
'***************************************************************
Public Property Get getTIPO_DETERMINACION_ID() As Long
        getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getTIPO_ANALISIS_ID() As Long
        getTIPO_ANALISIS_ID = TIPO_ANALISIS_ID
End Property
Public Property Get getBANO_ID() As Long
        getBANO_ID = BANO_ID
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
Public Property Get getMINIMO() As String
        getMINIMO = MINIMO
End Property
Public Property Get getMAXIMO() As String
        getMAXIMO = MAXIMO
End Property
Public Property Get getMINIMO_TEXTO() As String
        getMINIMO_TEXTO = MINIMO_TEXTO
End Property
Public Property Get getMAXIMO_TEXTO() As String
        getMAXIMO_TEXTO = MAXIMO_TEXTO
End Property
Public Property Get getLIM_MINIMO() As String
        getLIM_MINIMO = LIM_MINIMO
End Property
Public Property Get getLIM_MAXIMO() As String
        getLIM_MAXIMO = LIM_MAXIMO
End Property
Public Property Get getREQUERIDA() As Integer
        getREQUERIDA = REQUERIDA
End Property
Public Property Get getGRAFICO() As Integer
        getGRAFICO = grafico
End Property
Public Property Get getTIPO_FRECUENCIA_ID() As Integer
        getTIPO_FRECUENCIA_ID = TIPO_FRECUENCIA_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDETERMINACIONES_ANALISIS
'***************************************************************
Public Function Carga_por_BANO(BANO As Long, TIPO_DETERMINACION As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DETERMINACIONES_ANALISIS WHERE BANO_ID = " & BANO & " AND TIPO_DETERMINACION_ID = " & TIPO_DETERMINACION
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_por_BANO = False
        Else
                TIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
                TIPO_ANALISIS_ID = rs("TIPO_ANALISIS_ID")
                BANO_ID = rs("BANO_ID")
                ORDEN = rs("ORDEN")
                MINIMO = rs("MINIMO")
                MAXIMO = rs("MAXIMO")
                MINIMO_TEXTO = rs("MINIMO_TEXTO")
                MAXIMO_TEXTO = rs("MAXIMO_TEXTO")
                LIM_MINIMO = rs("LIM_MINIMO")
                LIM_MAXIMO = rs("LIM_MAXIMO")
                REQUERIDA = rs("REQUERIDA")
                grafico = rs("GRAFICO")
                If Not IsNull(rs("TIPO_FRECUENCIA_ID")) Then
                    TIPO_FRECUENCIA_ID = rs("TIPO_FRECUENCIA_ID")
                End If
                Carga_por_BANO = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_por_BANO = False
        MsgBox "Error al cargar los datos(clsDeterminaciones_analisis)", vbCritical, Err.Description
End Function
Public Function Carga_por_tipo_analisis(TIPO_ANALISIS As Long, TIPO_DETERMINACION As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DETERMINACIONES_ANALISIS WHERE TIPO_ANALISIS_ID = " & TIPO_ANALISIS & " AND TIPO_DETERMINACION_ID = " & TIPO_DETERMINACION
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_por_tipo_analisis = False
        Else
                TIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
                TIPO_ANALISIS_ID = rs("TIPO_ANALISIS_ID")
                BANO_ID = rs("BANO_ID")
                ORDEN = rs("ORDEN")
                MINIMO = rs("MINIMO")
                MAXIMO = rs("MAXIMO")
                MINIMO_TEXTO = rs("MINIMO_TEXTO")
                MAXIMO_TEXTO = rs("MAXIMO_TEXTO")
                LIM_MINIMO = rs("LIM_MINIMO")
                LIM_MAXIMO = rs("LIM_MAXIMO")
                REQUERIDA = rs("REQUERIDA")
                grafico = rs("GRAFICO")
                If Not IsNull(rs("TIPO_FRECUENCIA_ID")) Then
                    TIPO_FRECUENCIA_ID = rs("TIPO_FRECUENCIA_ID")
                End If
                Carga_por_tipo_analisis = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_por_tipo_analisis = False
        MsgBox "Error al cargar los datos(clsDeterminaciones_analisis)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO DETERMINACIONES_ANALISIS " & _
                   "  VALUES (" & _
                        TIPO_DETERMINACION_ID & "," & TIPO_ANALISIS_ID & "," & BANO_ID & "," & _
                        ORDEN & "," & "'" & MINIMO & "'" & "," & "'" & MAXIMO & "'" & "," & _
                        "'" & MINIMO_TEXTO & "'" & "," & "'" & MAXIMO_TEXTO & "'" & "," & _
                        "'" & LIM_MINIMO & "'" & "," & "'" & LIM_MAXIMO & "'" & "," & REQUERIDA & "," & grafico & "," & _
                        IIf(TIPO_FRECUENCIA_ID = 0, "null", TIPO_FRECUENCIA_ID) & _
                ")"
        execute_bd consulta
        Insertar = TIPO_DETERMINACION_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDeterminaciones_analisis)", vbCritical, Err.Description
End Function
Public Function Eliminar(ANALISIS As Long, BANO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DETERMINACIONES_ANALISIS " & _
                   " WHERE TIPO_ANALISIS_ID = " & ANALISIS & _
                   "   AND BANO_ID = " & BANO
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDeterminaciones_analisis)", vbCritical, Err.Description
End Function
Public Function Duplicar(TA_ORIGEN As Long, BANO_ORIGEN As Long, TA_DESTINO As Long, BANO_DESTINO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO DETERMINACIONES_ANALISIS " & _
                   "SELECT TIPO_DETERMINACION_ID," & TA_DESTINO & "," & BANO_DESTINO & ",ORDEN,MINIMO,MAXIMO,MINIMO_TEXTO,MAXIMO_TEXTO,LIM_MINIMO,LIM_MAXIMO,REQUERIDA,GRAFICO,TIPO_FRECUENCIA_ID " & _
                   " WHERE TIPO_ANALISIS_ID = " & TA_ORIGEN & _
                   "   AND BANO_ID = " & BANO_ORIGEN
        execute_bd consulta
        Duplicar = True
        Exit Function
fallo:
        Duplicar = False
        MsgBox "Error al Duplicar (clsDeterminaciones_analisis)", vbCritical, Err.Description
End Function
Public Function Listado(TIPO_ANALISIS As Long, BANO_ID As Long) As ADODB.Recordset
        Dim consulta As String
        Dim w As String
        If TIPO_ANALISIS <> 0 Then
            w = w & " AND A.TIPO_ANALISIS_ID = " & TIPO_ANALISIS
        End If
        If BANO_ID <> 0 Then
            w = w & " AND A.BANO_ID = " & BANO_ID
        End If
        consulta = "SELECT B.PNT,B.NOMBRE,B.DESCRIPCION,A.MINIMO,A.MAXIMO, " & _
                   "       B.ID_TIPO_DETERMINACION,A.REQUERIDA,A.ORDEN, " & _
                   "       B.FORMULA_ID,B.PRECIO,A.MINIMO_TEXTO,A.MAXIMO_TEXTO,B.METODO, A.LIM_MINIMO, A.LIM_MAXIMO,A.GRAFICO,IFNULL(C.NOMBRE,'') AS TIPO_FRECUENCIA,IFNULL(A.TIPO_FRECUENCIA_ID,'') AS TIPO_FRECUENCIA_ID " & _
                   "  FROM DETERMINACIONES_ANALISIS A " & _
                   "  INNER JOIN TIPOS_DETERMINACION B ON A.TIPO_DETERMINACION_ID = B.ID_TIPO_DETERMINACION " & _
                   "  LEFT JOIN TIPOS_FRECUENCIA C ON A.TIPO_FRECUENCIA_ID = C.ID_TIPO_FRECUENCIA " & _
                   " WHERE 1 = 1 " & _
                   w & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function

Public Function Listado_por_bano_historico(BANO As Long) As ADODB.Recordset
        Dim consulta As String
        Dim limit As String
        consulta = "SELECT B.PNT,B.NOMBRE,B.DESCRIPCION,A.MINIMO,A.MAXIMO, " & _
                   "       B.ID_TIPO_DETERMINACION,A.REQUERIDA,A.ORDEN, " & _
                   "       B.FORMULA_ID,B.PRECIO,A.MINIMO_TEXTO,A.MAXIMO_TEXTO,B.METODO " & _
                   "  FROM DETERMINACIONES_ANALISIS A,TIPOS_DETERMINACION B " & _
                   " WHERE A.BANO_ID = " & BANO & _
                   "   AND A.TIPO_DETERMINACION_ID = B.ID_TIPO_DETERMINACION " & _
                   " ORDER BY A.ORDEN " & limit
        
        Set Listado_por_bano_historico = datos_bd(consulta)
End Function

Public Function Precio_determinaciones_por_tipo_analisis(TIPO_ANALISIS As Long, MUESTRA As Long, TARIFA As Long) As Currency
        Dim consulta As String
        consulta = "SELECT sum(tar.precio) " & _
                   "  from determinaciones d left join determinaciones_analisis df  on d.tipo_determinacion_id = df.tipo_determinacion_id " & _
                   "                         left join tipos_determinacion td       on d.tipo_determinacion_id = td.id_tipo_determinacion " & _
                   "                         left join tarifas_precios tar          on d.tipo_determinacion_id = tar.tipo_determinacion_id " & _
                   " where df.tipo_analisis_id = " & TIPO_ANALISIS & _
                   "   and d.muestra_id = " & MUESTRA & _
                   "   and tar.tarifa_id = " & TARIFA & _
                   "   AND df.requerida = 0"
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            Precio_determinaciones_por_tipo_analisis = Format("0", "currency")
        Else
            Precio_determinaciones_por_tipo_analisis = Format(Replace(rs(0), ".", ","), "currency")
        End If
End Function
Public Function precio_bano(BANO As Long, TARIFA As Integer) As Currency
        Dim consulta As String
        If TARIFA <> 0 Then
            consulta = "SELECT sum(tp.precio) " & _
                       "  from determinaciones_analisis da left join tarifas_precios tp " & _
                       " on da.tipo_determinacion_id = tp.tipo_determinacion_id " & _
                       " where da.bano_id = " & BANO & _
                       "   and tp.tarifa_id = " & TARIFA
        Else
            consulta = "SELECT sum(tp.precio) " & _
                       "  from determinaciones_analisis da left join tipos_determinacion tp " & _
                       " on da.tipo_determinacion_id = tp.id_tipo_determinacion " & _
                       " where da.bano_id = " & BANO
        End If
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            precio_bano = Format("0", "currency")
        Else
            precio_bano = Format(Replace(rs(0), ".", ","), "currency")
        End If
End Function
Public Function Precio_determinaciones_por_bano(BANO As Long, MUESTRA As Long, TARIFA As Long) As Currency
        Dim consulta As String
        consulta = "SELECT sum(tar.precio) " & _
                   "  from determinaciones d left join determinaciones_analisis df  on d.tipo_determinacion_id = df.tipo_determinacion_id " & _
                   "                         left join tipos_determinacion td       on d.tipo_determinacion_id = td.id_tipo_determinacion " & _
                   "                         left join tarifas_precios     tar      on d.tipo_determinacion_id = tar.tipo_determinacion_id " & _
                   " where df.bano_id = " & BANO & _
                   "   and d.muestra_id = " & MUESTRA & _
                   "   and tar.tarifa_id = " & TARIFA & _
                   "   AND df.requerida = 0"
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            Precio_determinaciones_por_bano = Format("0", "currency")
        Else
            Precio_determinaciones_por_bano = Format(Replace(rs(0), ".", ","), "currency")
        End If
End Function
