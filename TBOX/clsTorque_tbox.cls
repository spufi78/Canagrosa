VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTorque_tbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTORQUE_TBOX
'***************************************************************
Private EQUIPO_ID As Long
Private N_PUNTOS As Long
Private N_MEDIDAS As Long
Private precargas As String
Private msj As String
Private VARIACION_BRAZO As String
Private VARIACION_TIEMPO As String
Private REPRODUCIBILIDAD As String

'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE CLSTORQUE_TBOX
'***************************************************************
Public Enum LISTA_CLSTORQUE_TBOX_COLS
        ENUM_EQUIPO_ID = 1
        ENUM_N_PUNTOS = 2
        ENUM_N_MEDIDAS = 3
        ENUM_PRECARGAS = 4
        ENUM_MSJ = 5
        ENUM_VARIACION_BRAZO = 6
        ENUM_VARIACION_TIEMPO = 7
        ENUM_REPRODUCIBILIDAD = 8
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTORQUE_TBOX
'***************************************************************
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setN_PUNTOS(ByVal dato As Long)
        N_PUNTOS = dato
End Property
Public Property Let setN_MEDIDAS(ByVal dato As Long)
        N_MEDIDAS = dato
End Property
Public Property Let setPRECARGAS(ByVal dato As String)
        precargas = dato
End Property
Public Property Let setMSJ(ByVal dato As String)
        msj = dato
End Property
Public Property Let setVARIACION_BRAZO(ByVal dato As String)
        VARIACION_BRAZO = dato
End Property
Public Property Let setVARIACION_TIEMPO(ByVal dato As String)
        VARIACION_TIEMPO = dato
End Property
Public Property Let setREPRODUCIBILIDAD(ByVal dato As String)
        REPRODUCIBILIDAD = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTORQUE_TBOX
'***************************************************************
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getN_PUNTOS() As Long
        getN_PUNTOS = N_PUNTOS
End Property
Public Property Get getN_MEDIDAS() As Long
        getN_MEDIDAS = N_MEDIDAS
End Property
Public Property Get getPRECARGAS() As String
        getPRECARGAS = precargas
End Property
Public Property Get getMSJ() As String
        getMSJ = msj
End Property
Public Property Get getVARIACION_BRAZO() As String
        getVARIACION_BRAZO = VARIACION_BRAZO
End Property
Public Property Get getVARIACION_TIEMPO() As String
        getVARIACION_TIEMPO = VARIACION_TIEMPO
End Property
Public Property Get getREPRODUCIBILIDAD() As String
        getREPRODUCIBILIDAD = REPRODUCIBILIDAD
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTORQUE_TBOX
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM geslab_metrologia.TORQUE_TBOX WHERE EQUIPO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                EQUIPO_ID = rs("EQUIPO_ID")
                N_PUNTOS = rs("N_PUNTOS")
                N_MEDIDAS = rs("N_MEDIDAS")
                precargas = rs("PRECARGAS")
                msj = rs("MSJ")
                VARIACION_BRAZO = rs("VARIACION_BRAZO")
                VARIACION_TIEMPO = rs("VARIACION_TIEMPO")
                REPRODUCIBILIDAD = rs("REPRODUCIBILIDAD")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTorque_tbox)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "REPLACE INTO geslab_metrologia.TORQUE_TBOX " & _
                   "  VALUES (" & _
                        EQUIPO_ID & "," & N_PUNTOS & "," & N_MEDIDAS & "," & _
                        "'" & precargas & "'" & "," & "'" & msj & "'" & "," & "'" & VARIACION_BRAZO & "'" & "," & _
                        "'" & VARIACION_TIEMPO & "'" & "," & "'" & REPRODUCIBILIDAD & "'" & _
                ")"
        execute_bd consulta
        Insertar = EQUIPO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTorque_tbox)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE geslab_metrologia.TORQUE_TBOX SET " & _
                        " N_PUNTOS = " & N_PUNTOS & "," & _
                        " N_MEDIDAS = " & N_MEDIDAS & "," & _
                        " PRECARGAS = '" & precargas & "'," & _
                        " MSJ = '" & msj & "'," & _
                        " VARIACION_BRAZO = '" & VARIACION_BRAZO & "'," & _
                        " VARIACION_TIEMPO = '" & VARIACION_TIEMPO & "'," & _
                        " REPRODUCIBILIDAD = '" & REPRODUCIBILIDAD & "'" & _
                " WHERE EQUIPO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTorque_tbox)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM geslab_metrologia.TORQUE_TBOX " & _
                "    WHERE EQUIPO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTorque_tbox)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM geslab_metrologia.TORQUE_TBOX ORDER BY EQUIPO_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSTORQUE_TBOX
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

