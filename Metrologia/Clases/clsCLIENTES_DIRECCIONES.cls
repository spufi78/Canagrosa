VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCLIENTES_DIRECCIONES"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCLIENTES_DIRECCIONES
'***************************************************************
Private CLIENTE_ID As Long
Private ID_DIRECCION As Long
Private NOMBRE As String
Private DIRECCION As String
Private CP As String
Private MUNICIPIO_ID As Long
Private PROVINCIA_ID As Long
Private TELEFONO As String

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCLIENTES_DIRECCIONES
'***************************************************************
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setID_DIRECCION(ByVal dato As Long)
        ID_DIRECCION = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setDIRECCION(ByVal dato As String)
        DIRECCION = dato
End Property
Public Property Let setCP(ByVal dato As String)
        CP = dato
End Property
Public Property Let setMUNICIPIO_ID(ByVal dato As Long)
        MUNICIPIO_ID = dato
End Property
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
        PROVINCIA_ID = dato
End Property
Public Property Let setTELEFONO(ByVal dato As String)
        TELEFONO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCLIENTES_DIRECCIONES
'***************************************************************
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getID_DIRECCION() As Long
        getID_DIRECCION = ID_DIRECCION
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getDIRECCION() As String
        getDIRECCION = DIRECCION
End Property
Public Property Get getCP() As String
        getCP = CP
End Property
Public Property Get getMUNICIPIO_ID() As Long
        getMUNICIPIO_ID = MUNICIPIO_ID
End Property
Public Property Get getPROVINCIA_ID() As Long
        getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getTELEFONO() As String
        getTELEFONO = TELEFONO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCLIENTES_DIRECCIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CLIENTES_DIRECCIONES WHERE CLIENTE_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CLIENTE_ID = rs("CLIENTE_ID")
                ID_DIRECCION = rs("ID_DIRECCION")
                NOMBRE = rs("NOMBRE")
                CP = rs("CP")
                DIRECCION = rs("DIRECCION")
                MUNICIPIO_ID = rs("MUNICIPIO_ID")
                PROVINCIA_ID = rs("PROVINCIA_ID")
                TELEFONO = rs("TELEFONO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCLIENTES_DIRECCIONES)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CLIENTES_DIRECCIONES " & _
                   "  VALUES (" & _
                        CLIENTE_ID & "," & ID_DIRECCION & ",'" & NOMBRE & "','" & DIRECCION & "'" & "," & "'" & CP & "'" & "," & _
                        MUNICIPIO_ID & "," & PROVINCIA_ID & ",'" & TELEFONO & "')"
        execute_bd consulta
        Insertar = ID_DIRECCION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCLIENTES_DIRECCIONES)", vbCritical, Err.Description
End Function
Public Function Modificar(cliente As Long, ID_DIRECCION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CLIENTES_DIRECCIONES SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " CP = '" & CP & "'," & _
                        " DIRECCION = '" & DIRECCION & "'," & _
                        " MUNICIPIO_ID = " & MUNICIPIO_ID & "," & _
                        " TELEFONO = '" & TELEFONO & "'," & _
                        " PROVINCIA_ID = " & PROVINCIA_ID & _
                " WHERE ID_DIRECCION = " & ID_DIRECCION & _
                "   AND CLIENTE_ID = " & cliente
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCLIENTES_DIRECCIONES)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CLIENTES_DIRECCIONES " & _
                "    WHERE ID_DIRECCION = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCLIENTES_DIRECCIONES)", vbCritical, Err.Description
End Function
Public Function Eliminar_por_Cliente(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CLIENTES_DIRECCIONES " & _
                "    WHERE CLIENTE_ID = " & ID
        execute_bd consulta
        Eliminar_por_Cliente = True
        Exit Function
fallo:
        Eliminar_por_Cliente = False
        MsgBox "Error al Eliminar (clsCLIENTES_DIRECCIONES)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM CLIENTES_DIRECCIONES " & _
                   " WHERE CLIENTE_ID = " & ID & _
                   " ORDER BY ID_DIRECCION"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_DIRECCION AS C1,NOMBRE AS C2 " & _
                   "  FROM CLIENTES_DIRECCIONES " & _
                   " WHERE CLIENTE_ID = " & ID & _
                   " ORDER BY ID_DIRECCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
