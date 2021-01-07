VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSellantes_resultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSELLANTES_RESULTADOS
'***************************************************************
Private MUESTRA_ID As Long
Private SELLANTE_ID As Long
Private ORDEN As Long
Private TIPO_DETERMINACION_ID As Long
Private FORMULA_ID As Long
Private RESULTADO As String
Private CONFORME As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSELLANTES_RESULTADOS
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setSELLANTE_ID(ByVal dato As Long)
        SELLANTE_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Long)
        TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setFORMULA_ID(ByVal dato As Long)
        FORMULA_ID = dato
End Property
Public Property Let setRESULTADO(ByVal dato As String)
        RESULTADO = dato
End Property
Public Property Let setCONFORME(ByVal dato As Integer)
        CONFORME = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSELLANTES_RESULTADOS
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getSELLANTE_ID() As Long
        getSELLANTE_ID = SELLANTE_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getTIPO_DETERMINACION_ID() As Long
        getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getFORMULA_ID() As Long
        getFORMULA_ID = FORMULA_ID
End Property
Public Property Get getRESULTADO() As String
        getRESULTADO = RESULTADO
End Property
Public Property Get getCONFORME() As Integer
        getCONFORME = CONFORME
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSELLANTES_RESULTADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SELLANTES_RESULTADOS WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                SELLANTE_ID = rs("SELLANTE_ID")
                ORDEN = rs("ORDEN")
                TIPO_DETERMINACION_ID = rs("TIPO_DETERMINACION_ID")
                FORMULA_ID = rs("FORMULA_ID")
                RESULTADO = rs("RESULTADO")
                CONFORME = rs("CONFORME")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSellantes_resultados)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO SELLANTES_RESULTADOS " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & SELLANTE_ID & "," & ORDEN & "," & TIPO_DETERMINACION_ID & "," & FORMULA_ID & "," & _
                        "'" & RESULTADO & "', curdate(),curtime(),0," & CONFORME & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSellantes_resultados)", vbCritical, Err.Description
End Function
Public Function informar_resultado(MUESTRA As Long, xorden As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SELLANTES_RESULTADOS SET " & _
                        " RESULTADO = '" & RESULTADO & "'," & _
                        " FECHA = curdate(),HORA=curtime(), " & _
                        " EMPLEADO_ID = " & USUARIO.getID_EMPLEADO & "," & _
                        " CONFORME = " & CONFORME & "" & _
                " WHERE MUESTRA_ID = " & MUESTRA & _
                "   AND ORDEN = " & xorden
        execute_bd consulta
        informar_resultado = True
        Exit Function
fallo:
        informar_resultado = False
        MsgBox "Error al informar_resultado (clsSellantes_resultados)", vbCritical, Err.Description
End Function
Public Function Eliminar(MUESTRA As Long, ORDEN As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SELLANTES_RESULTADOS " & _
                "    WHERE MUESTRA_ID = " & MUESTRA & _
                "      AND ORDEN = " & ORDEN
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSellantes_resultados)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SELLANTES_RESULTADOS ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Cerrar(MUESTRA As Long) As Integer
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Cerrar = 1
        consulta = "SELECT * FROM SELLANTES_RESULTADOS " & _
                   " WHERE MUESTRA_ID = " & MUESTRA
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            Do
                If Trim(rs("RESULTADO")) = "" Then
                    Cerrar = 0
                End If
                rs.MoveNext
            Loop Until rs.EOF
        End If
        Set rs = Nothing
End Function
Public Function Listado_Resultados(MUESTRA As Long) As ADODB.Recordset
'M0926-I
'Añadimos:  Al final de la consulta A.SELLANTE_ID,A.FORMULA_ID,A.ORDEN (tras A.FORMULA_ID)
'M0926-F
        Dim consulta As String
        consulta = "SELECT A.ORDEN,B.ENSAYO,B.RANGO_INFERIOR,B.RANGO_SUPERIOR,A.RESULTADO,U.NOMBRE, " & _
                   "       A.CONFORME,B.ENSAYO_INGLES,A.TIPO_DETERMINACION_ID,B.REFERENCIA, " & _
                   "       B.CRITERIO,B.NORMA_CRITERIO,A.FORMULA_ID,A.SELLANTE_ID,A.FORMULA_ID,A.ORDEN " & _
                   "  FROM SELLANTES_RESULTADOS A,SELLANTES_ENSAYOS B,UNIDADES U " & _
                   " WHERE A.MUESTRA_ID = " & MUESTRA & _
                   "   AND A.SELLANTE_ID = B.SELLANTE_ID " & _
                   "   AND A.ORDEN = B.ORDEN " & _
                   "   AND B.UNIDAD_ID = U.ID_UNIDAD " & _
                   " ORDER BY A.ORDEN"
        Set Listado_Resultados = datos_bd(consulta)
End Function

Public Function Listado_Resultados_Vida(MUESTRA As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ORDEN,IFNULL(IF(A.TIPO_DETERMINACION_ID=0,B.ENSAYO,CONCAT(B.ENSAYO,' -> ',FC.NOMBRE)),''),B.RANGO_INFERIOR,B.RANGO_SUPERIOR " & _
                   "       ,IFNULL(IF(A.TIPO_DETERMINACION_ID=0,A.RESULTADO,SD.VALOR_1),''),U.NOMBRE,A.CONFORME,B.ENSAYO_INGLES " & _
                   "       ,B.TIPO_DETERMINACION_ID, A.FECHA, A.HORA, USU.USUARIO, USU.ID_EMPLEADO, A.SELLANTE_ID " & _
                   " FROM SELLANTES_RESULTADOS A " & _
                   " INNER JOIN SELLANTES_ENSAYOS B ON A.SELLANTE_ID = B.SELLANTE_ID AND A.ORDEN = B.ORDEN " & _
                   " INNER JOIN UNIDADES U ON B.UNIDAD_ID = U.ID_UNIDAD " & _
                   " INNER JOIN usuarios USU ON A.EMPLEADO_ID = USU.ID_EMPLEADO " & _
                   " LEFT JOIN tipos_determinacion TD ON A.TIPO_DETERMINACION_ID = TD.ID_TIPO_DETERMINACION AND A.TIPO_DETERMINACION_ID <> 0 " & _
                   " LEFT JOIN formulas_campos FC ON TD.FORMULA_ID = FC.FORMULA_ID " & _
                   " LEFT JOIN sellantes_determinaciones SD ON A.MUESTRA_ID = SD.MUESTRA_ID AND FC.ID_CAMPO = SD.CAMPO_ID " & _
                   " WHERE A.MUESTRA_ID = " & MUESTRA & _
                   " ORDER BY A.ORDEN, SD.CAMPO_ID "
        Set Listado_Resultados_Vida = datos_bd(consulta)
End Function
Public Function ModificarSellante(ID As Long, SELLANTE_ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SELLANTES_RESULTADOS " & _
                   "   SET SELLANTE_ID = " & SELLANTE_ID & _
                   " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        ' Eliminar los resultados que no son validos para el nuevo sellante
        consulta = "delete from sellantes_resultados where muestra_id = " & ID & _
                   " and orden > (SELECT max(orden) from sellantes_ensayos where SELLANTE_ID = " & SELLANTE_ID & ")"
        execute_bd consulta
        ModificarSellante = True
        Exit Function
fallo:
        ModificarSellante = False
        MsgBox "Error al ModificarSellante (clsSellantes_recepcion)", vbCritical, Err.Description
End Function
Public Function ModificarVida(lMUESTRA_ID As Long, SELLANTE_ID As Long, lORDEN As Long, EMPLEADO_ID As Long, fecha As String, hora As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SELLANTES_RESULTADOS SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " HORA = '" & hora & "'," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & _
                " WHERE MUESTRA_ID = " & lMUESTRA_ID & _
                "   AND SELLANTE_ID = " & SELLANTE_ID & _
                "   AND ORDEN = " & lORDEN
        execute_bd consulta
        ModificarVida = True
        Exit Function
fallo:
        ModificarVida = False
        MsgBox "Error al modificarVida (clsSellantes_resultados)", vbCritical, Err.Description
End Function


