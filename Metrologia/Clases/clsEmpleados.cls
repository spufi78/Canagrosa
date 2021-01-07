VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEMPLEADOS
'***************************************************************
Private ID_EMPLEADO As Integer
Private NOMBRE As String
Private DIRECCION As String
Private SEG As String
Private IPER As String
Private FECHA As String
Private LUGAR As String
Private CONTRATO As String
Private GRUPO As String
Private CP As String
Private TELEFONO As String
Private OBSERVACIONES As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEMPLEADOS
'***************************************************************
Public Property Let setID_EMPLEADO(ByVal dato As Integer)
        ID_EMPLEADO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setDIRECCION(ByVal dato As String)
        DIRECCION = dato
End Property
Public Property Let setSEG(ByVal dato As String)
        SEG = dato
End Property
Public Property Let setIPER(ByVal dato As String)
        IPER = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        FECHA = dato
End Property
Public Property Let setLUGAR(ByVal dato As String)
        LUGAR = dato
End Property
Public Property Let setCONTRATO(ByVal dato As String)
        CONTRATO = dato
End Property
Public Property Let setGRUPO(ByVal dato As String)
        GRUPO = dato
End Property
Public Property Let setCP(ByVal dato As String)
        CP = dato
End Property
Public Property Let setTELEFONO(ByVal dato As String)
        TELEFONO = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEMPLEADOS
'***************************************************************
Public Property Get getID_EMPLEADO() As Integer
        getID_EMPLEADO = ID_EMPLEADO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getDIRECCION() As String
        getDIRECCION = DIRECCION
End Property
Public Property Get getSEG() As String
        getSEG = SEG
End Property
Public Property Get getIPER() As String
        getIPER = IPER
End Property
Public Property Get getFECHA() As String
        getFECHA = FECHA
End Property
Public Property Get getLUGAR() As String
        getLUGAR = LUGAR
End Property
Public Property Get getCONTRATO() As String
        getCONTRATO = CONTRATO
End Property
Public Property Get getGRUPO() As String
        getGRUPO = GRUPO
End Property
Public Property Get getCP() As String
        getCP = CP
End Property
Public Property Get getTELEFONO() As String
        getTELEFONO = TELEFONO
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEMPLEADOS
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EMPLEADOS WHERE ID_EMPLEADO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_EMPLEADO = rs("ID_EMPLEADO")
                NOMBRE = rs("NOMBRE")
                DIRECCION = rs("DIRECCION")
                SEG = rs("SEG")
                IPER = rs("IPER")
                FECHA = rs("FECHA")
                LUGAR = rs("LUGAR")
                CONTRATO = rs("CONTRATO")
                GRUPO = rs("GRUPO")
                CP = rs("CP")
                TELEFONO = rs("TELEFONO")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEmpleados)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_EMPLEADO) FROM EMPLEADOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_EMPLEADO = 1
        Else
                ID_EMPLEADO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO EMPLEADOS " & _
                   "  VALUES (" & _
                        ID_EMPLEADO & "," & "'" & NOMBRE & "'" & "," & "'" & DIRECCION & "'" & "," & _
                        "'" & SEG & "'" & "," & "'" & IPER & "'" & "," & "'" & FECHA & "'" & "," & _
                        "'" & LUGAR & "'" & "," & "'" & CONTRATO & "'" & "," & "'" & GRUPO & "'" & "," & _
                        "'" & CP & "'" & "," & "'" & TELEFONO & "'" & "," & "'" & OBSERVACIONES & "'" & _
                ")"
        execute_bd consulta
        insertar = ID_EMPLEADO
        Exit Function
fallo:
        insertar = 0
        MsgBox "Error al insertar (clsEmpleados) : " & ID_EMPLEADO & "," & NOMBRE, vbCritical, Err.Description
End Function
Public Function modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EMPLEADOS SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " DIRECCION = '" & DIRECCION & "'," & _
                        " SEG = '" & SEG & "'," & _
                        " IPER = '" & IPER & "'," & _
                        " FECHA = '" & FECHA & "'," & _
                        " LUGAR = '" & LUGAR & "'," & _
                        " CONTRATO = '" & CONTRATO & "'," & _
                        " GRUPO = '" & GRUPO & "'," & _
                        " CP = '" & CP & "'," & _
                        " TELEFONO = '" & TELEFONO & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE ID_EMPLEADO = " & ID
        execute_bd consulta
        modificar = True
        Exit Function
fallo:
        modificar = False
        MsgBox "Error al Modificar (clsEmpleados)", vbCritical, Err.Description
End Function
Public Function eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EMPLEADOS " & _
                "    WHERE ID_EMPLEADO = " & ID
        execute_bd consulta
        eliminar = True
        Exit Function
fallo:
        eliminar = False
        MsgBox "Error al Eliminar (clsEmpleados)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EMPLEADOS ORDER BY ID_EMPLEADO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_EMPLEADO,NOMBRE FROM EMPLEADOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Busqueda(pNOMBRE As String, pNIF As String) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_EMPLEADO,NOMBRE,DIRECCION,SEG,IPER " & _
               "  FROM EMPLEADOS " & _
               " WHERE NOMBRE LIKE '%" & pNOMBRE & "%'" & _
               "   AND IPER LIKE '%" & pNIF & "%'" & _
               " ORDER BY NOMBRE "
    Set Busqueda = datos_bd(consulta)
End Function

