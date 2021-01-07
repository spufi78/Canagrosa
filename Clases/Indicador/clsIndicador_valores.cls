VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndicador_valores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSINDICADOR_VALORES
'***************************************************************
Private INDICADOR_ID As Long
Private ANNO As Long
Private MES As Long
Private VALOR As String
Private OBJETIVO As String
Private COMENTARIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSINDICADOR_VALORES
'***************************************************************
Public Property Let setINDICADOR_ID(ByVal dato As Long)
        INDICADOR_ID = dato
End Property
Public Property Let setANNO(ByVal dato As Long)
        ANNO = dato
End Property
Public Property Let setMES(ByVal dato As Long)
        MES = dato
End Property
Public Property Let setVALOR(ByVal dato As String)
        VALOR = dato
End Property
Public Property Let setOBJETIVO(ByVal dato As String)
        OBJETIVO = dato
End Property
Public Property Let setCOMENTARIO(ByVal dato As String)
        COMENTARIO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSINDICADOR_VALORES
'***************************************************************
Public Property Get getINDICADOR_ID() As Long
        getINDICADOR_ID = INDICADOR_ID
End Property
Public Property Get getANNO() As Long
        getANNO = ANNO
End Property
Public Property Get getMES() As Long
        getMES = MES
End Property
Public Property Get getVALOR() As String
        getVALOR = VALOR
End Property
Public Property Get getOBJETIVO() As String
        getOBJETIVO = OBJETIVO
End Property
Public Property Get getCOMENTARIO() As String
        getCOMENTARIO = COMENTARIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSINDICADOR_VALORES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM INDICADOR_VALORES WHERE INDICADOR_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                INDICADOR_ID = rs("INDICADOR_ID")
                ANNO = rs("ANNO")
                MES = rs("MES")
                VALOR = rs("VALOR")
                OBJETIVO = rs("OBJETIVO")
                COMENTARIO = rs("COMENTARIO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsIndicador_valores)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO INDICADOR_VALORES " & _
                   "  VALUES (" & _
                        INDICADOR_ID & "," & ANNO & "," & MES & "," & _
                        "'" & VALOR & "'" & "," & "'" & OBJETIVO & "'" & "," & "'" & COMENTARIO & "'" & _
                ")" & _
                " ON DUPLICATE KEY UPDATE VALOR='" & VALOR & "'"
        execute_bd consulta
        Insertar = INDICADOR_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsIndicador_valores)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE INDICADOR_VALORES SET " & _
                        " ANNO = " & ANNO & "," & _
                        " MES = " & MES & "," & _
                        " VALOR = '" & VALOR & "'," & _
                        " OBJETIVO = '" & OBJETIVO & "'," & _
                        " COMENTARIO = '" & COMENTARIO & "'" & _
                " WHERE INDICADOR_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsIndicador_valores)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM INDICADOR_VALORES " & _
                "    WHERE INDICADOR_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsIndicador_valores)", vbCritical, Err.Description
End Function
Public Function EliminarAnyo(ID As Long, ANYO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM INDICADOR_VALORES " & _
                "    WHERE INDICADOR_ID = " & ID & " AND ANNO = " & ANYO
        execute_bd consulta
        EliminarAnyo = True
        Exit Function
fallo:
        EliminarAnyo = False
        MsgBox "Error al Eliminar (clsIndicador_valores)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM INDICADOR_VALORES ORDER BY INDICADOR_ID"
        Set Listado = datos_bd(consulta)
End Function
'DGF
Public Function ListadoIndicadorAnno(INDICADOR As Long, ANYO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM INDICADOR_VALORES WHERE INDICADOR_ID=" & INDICADOR & " AND ANNO=" & ANYO
        Set ListadoIndicadorAnno = datos_bd(consulta)
End Function
'DGF

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSINDICADOR_VALORES
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

