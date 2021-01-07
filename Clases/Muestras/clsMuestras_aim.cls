VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMuestras_aim"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMUESTRAS_AIM
'***************************************************************
Private MUESTRA_ID As Long
Private AIM_PROGRAMA_ID As Long
Private AIM_CENTRO_ID As Long
Private AIM_TIPO_ENSAYO_ID As Long
Private AIM_SECCION_ID As Long
Private AIM_ESTACION_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMUESTRAS_AIM
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setAIM_PROGRAMA_ID(ByVal dato As Long)
        AIM_PROGRAMA_ID = dato
End Property
Public Property Let setAIM_CENTRO_ID(ByVal dato As Long)
        AIM_CENTRO_ID = dato
End Property
Public Property Let setAIM_TIPO_ENSAYO_ID(ByVal dato As Long)
        AIM_TIPO_ENSAYO_ID = dato
End Property
Public Property Let setAIM_SECCION_ID(ByVal dato As Long)
        AIM_SECCION_ID = dato
End Property
Public Property Let setAIM_ESTACION_ID(ByVal dato As Long)
        AIM_ESTACION_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMUESTRAS_AIM
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getAIM_PROGRAMA_ID() As Long
        getAIM_PROGRAMA_ID = AIM_PROGRAMA_ID
End Property
Public Property Get getAIM_CENTRO_ID() As Long
        getAIM_CENTRO_ID = AIM_CENTRO_ID
End Property
Public Property Get getAIM_TIPO_ENSAYO_ID() As Long
        getAIM_TIPO_ENSAYO_ID = AIM_TIPO_ENSAYO_ID
End Property
Public Property Get getAIM_SECCION_ID() As Long
        getAIM_SECCION_ID = AIM_SECCION_ID
End Property
Public Property Get getAIM_ESTACION_ID() As Long
        getAIM_ESTACION_ID = AIM_ESTACION_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMUESTRAS_AIM
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MUESTRAS_AIM WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                AIM_PROGRAMA_ID = rs("AIM_PROGRAMA_ID")
                AIM_CENTRO_ID = rs("AIM_CENTRO_ID")
                AIM_TIPO_ENSAYO_ID = rs("AIM_TIPO_ENSAYO_ID")
                AIM_SECCION_ID = rs("AIM_SECCION_ID")
                AIM_ESTACION_ID = rs("AIM_ESTACION_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMuestras_aim)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO MUESTRAS_AIM " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & AIM_PROGRAMA_ID & "," & AIM_CENTRO_ID & "," & _
                        AIM_TIPO_ENSAYO_ID & "," & AIM_SECCION_ID & "," & AIM_ESTACION_ID & _
                ") " & _
                "  ON DUPLICATE KEY UPDATE " & _
                " AIM_PROGRAMA_ID = " & AIM_PROGRAMA_ID & "," & _
                " AIM_CENTRO_ID = " & AIM_CENTRO_ID & "," & _
                " AIM_TIPO_ENSAYO_ID = " & AIM_TIPO_ENSAYO_ID & "," & _
                " AIM_SECCION_ID = " & AIM_SECCION_ID & "," & _
                " AIM_ESTACION_ID = " & AIM_ESTACION_ID
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMuestras_aim)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE MUESTRAS_AIM SET " & _
                        " AIM_PROGRAMA_ID = " & AIM_PROGRAMA_ID & "," & _
                        " AIM_CENTRO_ID = " & AIM_CENTRO_ID & "," & _
                        " AIM_TIPO_ENSAYO_ID = " & AIM_TIPO_ENSAYO_ID & "," & _
                        " AIM_SECCION_ID = " & AIM_SECCION_ID & "," & _
                        " AIM_ESTACION_ID = " & AIM_ESTACION_ID & "" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMuestras_aim)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MUESTRAS_AIM " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMuestras_aim)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MUESTRAS_AIM ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSMUESTRAS_AIM
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

