VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClientes_responsables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCLIENTES_RESPONSABLES
'***************************************************************
Private CLIENTE_ID As Long
Private RESPONSABLE_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCLIENTES_RESPONSABLES
'***************************************************************
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Long)
        RESPONSABLE_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCLIENTES_RESPONSABLES
'***************************************************************
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getRESPONSABLE_ID() As Long
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCLIENTES_RESPONSABLES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CLIENTES_RESPONSABLES WHERE CLIENTE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CLIENTE_ID = rs("CLIENTE_ID")
                RESPONSABLE_ID = rs("RESPONSABLE_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsClientes_responsables)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CLIENTES_RESPONSABLES " & _
                   "  VALUES (" & _
                        CLIENTE_ID & "," & RESPONSABLE_ID & _
                ")"
        execute_bd consulta
        Insertar = CLIENTE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsClientes_responsables)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CLIENTES_RESPONSABLES SET " & _
                        " RESPONSABLE_ID = " & RESPONSABLE_ID & "" & _
                " WHERE CLIENTE_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsClientes_responsables)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CLIENTES_RESPONSABLES " & _
                "    WHERE CLIENTE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsClientes_responsables)", vbCritical, Err.Description
End Function
Public Function Listado(cliente As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.ID_EMPLEADO, CONCAT(B.NOMBRE, ' ', B.APELLIDOS) " & _
                   "  FROM CLIENTES_RESPONSABLES A, usuarios B " & _
                   " WHERE A.CLIENTE_ID =" & cliente & _
                   "   AND A.RESPONSABLE_ID = B.ID_EMPLEADO " & _
                   " ORDER BY B.NOMBRE "
        Set Listado = datos_bd(consulta)
End Function
Public Function duplicar(clienteOrigen As Long, clienteDestino As Long) As ADODB.Recordset
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT * " & _
                   "  FROM CLIENTES_RESPONSABLES " & _
                   " WHERE CLIENTE_ID =" & clienteOrigen
        Set rs = datos_bd(consulta)
        Dim oCR As New clsClientes_responsables
        If rs.RecordCount > 0 Then
            Do
                With oCR
                    .setCLIENTE_ID = clienteDestino
                    .setRESPONSABLE_ID = rs("RESPONSABLE_ID")
                    .Insertar
                End With
                rs.MoveNext
            Loop Until rs.EOF
        End If
End Function
'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSCLIENTES_RESPONSABLES
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

