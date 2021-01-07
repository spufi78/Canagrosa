VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMuestras_ediciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMUESTRAS_EDICIONES
'***************************************************************
Private ID As Long
Private MUESTRA_ID As Long
Private EDICION As Long
Private OBSERVACIONES As String
Private USUARIO_ID As Long
Private fecha As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMUESTRAS_EDICIONES
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setEDICION(ByVal dato As Long)
        EDICION = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMUESTRAS_EDICIONES
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getEDICION() As Long
        getEDICION = EDICION
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMUESTRAS_EDICIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MUESTRAS_EDICIONES WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                MUESTRA_ID = rs("MUESTRA_ID")
                EDICION = rs("EDICION")
                OBSERVACIONES = rs("OBSERVACIONES")
                USUARIO_ID = rs("USUARIO_ID")
                fecha = rs("FECHA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMuestras_ediciones)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM MUESTRAS_EDICIONES"
        Set rs = datos_bd(consulta, True)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar(no_actualizar As Boolean) As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO MUESTRAS_EDICIONES " & _
                   "  VALUES (" & _
                        ID & "," & MUESTRA_ID & "," & EDICION & "," & _
                        "'" & OBSERVACIONES & "'" & "," & USUARIO_ID & "," & "'" & Format(fecha, "yyyy-mm-dd") & "'" & _
                ")"
        execute_bd consulta, no_actualizar
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMuestras_ediciones)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE MUESTRAS_EDICIONES SET " & _
                        " MUESTRA_ID = " & MUESTRA_ID & "," & _
                        " EDICION = " & EDICION & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " FECHA = '" & Format(fecha, "yyyy-mm-dd") & "'" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMuestras_ediciones)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MUESTRAS_EDICIONES " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMuestras_ediciones)", vbCritical, Err.Description
End Function
Public Function Listado(MUESTRA_ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,MUESTRA_ID,FECHA,EDICION,OBSERVACIONES,CONCAT(B.NOMBRE,' ',B.APELLIDOS) AS USUARIO_ID " & _
                   "  FROM MUESTRAS_EDICIONES A, USUARIOS B" & _
                   " WHERE MUESTRA_ID = " & MUESTRA_ID & _
                   "   AND A.USUARIO_ID = B.ID_EMPLEADO " & _
                   " ORDER BY EDICION,FECHA,ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSMUESTRAS_EDICIONES
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,MUESTRA_ID FROM MUESTRAS_EDICIONES ORDER BY MUESTRA_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
