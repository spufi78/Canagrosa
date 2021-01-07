VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMuestras_normas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMUESTRAS_NORMAS
'***************************************************************
Private MUESTRA_ID As Long
Private NORMA_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMUESTRAS_NORMAS
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setNORMA_ID(ByVal dato As Long)
        NORMA_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMUESTRAS_NORMAS
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getNORMA_ID() As Long
        getNORMA_ID = NORMA_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMUESTRAS_NORMAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MUESTRAS_NORMAS WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                NORMA_ID = rs("NORMA_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMuestras_normas)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO MUESTRAS_NORMAS " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & NORMA_ID & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMuestras_normas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE MUESTRAS_NORMAS SET " & _
                        " NORMA_ID = " & NORMA_ID & "" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMuestras_normas)", vbCritical, Err.Description
End Function
Public Function Eliminar(muestra As Long, NORMA As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MUESTRAS_NORMAS " & _
                "    WHERE MUESTRA_ID = " & muestra & " AND NORMA_ID = " & NORMA
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMuestras_normas)", vbCritical, Err.Description
End Function
Public Function Listado(muestra As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.NORMA_ID,CONCAT('(',B.CODIGO,') ',B.NOMBRE) FROM MUESTRAS_NORMAS A,CA_NORMAS B " & _
                   " where A.NORMA_ID = B.ID_NORMA AND A.MUESTRA_ID = " & muestra & _
                   " ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSMUESTRAS_NORMAS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

