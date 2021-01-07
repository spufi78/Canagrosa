VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFluidos_resultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFLUIDOS_RESULTADOS
'***************************************************************
Private MUESTRA_ID As Long
Private TAMANO As Long
Private resultado As Long
Private CLASIFICACION As String
Private FUERA_RANGO As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFLUIDOS_RESULTADOS
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setTAMANO(ByVal dato As Long)
        TAMANO = dato
End Property
Public Property Let setRESULTADO(ByVal dato As Long)
        resultado = dato
End Property
Public Property Let setCLASIFICACION(ByVal dato As String)
        CLASIFICACION = dato
End Property
Public Property Let setFUERA_RANGO(ByVal dato As Long)
        FUERA_RANGO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFLUIDOS_RESULTADOS
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getTAMANO() As Long
        getTAMANO = TAMANO
End Property
Public Property Get getRESULTADO() As Long
        getRESULTADO = resultado
End Property
Public Property Get getCLASIFICACION() As String
        getCLASIFICACION = CLASIFICACION
End Property
Public Property Get getFUERA_RANGO() As Long
        getFUERA_RANGO = FUERA_RANGO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFLUIDOS_RESULTADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM FLUIDOS_RESULTADOS WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                TAMANO = rs("TAMANO")
                resultado = rs("RESULTADO")
                CLASIFICACION = rs("CLASIFICACION")
                FUERA_RANGO = rs("FUERA_RANGO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFluidos_resultados)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO FLUIDOS_RESULTADOS " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & TAMANO & "," & resultado & "," & _
                        "'" & CLASIFICACION & "'" & "," & FUERA_RANGO & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFluidos_resultados)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FLUIDOS_RESULTADOS SET " & _
                        " TAMANO = " & TAMANO & "," & _
                        " RESULTADO = " & resultado & "," & _
                        " CLASIFICACION = '" & CLASIFICACION & "'," & _
                        " FUERA_RANGO = " & FUERA_RANGO & "" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFluidos_resultados)", vbCritical, Err.Description
End Function
Public Function Eliminar(MUESTRA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FLUIDOS_RESULTADOS " & _
                "    WHERE MUESTRA_ID = " & MUESTRA
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFluidos_resultados)", vbCritical, Err.Description
End Function
Public Function Listado(MUESTRA As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM FLUIDOS_RESULTADOS WHERE MUESTRA_ID = " & MUESTRA & " ORDER BY TAMANO"
        Set Listado = datos_bd(consulta)
End Function
