VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedores_servicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROVEEDORES_SERVICIOS
'***************************************************************
Private PROVEEDOR_ID As Long
Private ORDEN As Long
Private SERVICIO_ID As Long
Private tipo As String
Private METODO_ID As Long
Private FECHA_INICIAL As String
Private FECHA_APROBACION As String
Private OBSERVACIONES As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROVEEDORES_SERVICIOS
'***************************************************************
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setSERVICIO_ID(ByVal dato As Long)
        SERVICIO_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As String)
        tipo = dato
End Property
Public Property Let setMETODO_ID(ByVal dato As Long)
        METODO_ID = dato
End Property
Public Property Let setFECHA_INICIAL(ByVal dato As String)
        FECHA_INICIAL = dato
End Property
Public Property Let setFECHA_APROBACION(ByVal dato As String)
        FECHA_APROBACION = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROVEEDORES_SERVICIOS
'***************************************************************
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getSERVICIO_ID() As Long
        getSERVICIO_ID = SERVICIO_ID
End Property
Public Property Get getTIPO() As String
        getTIPO = tipo
End Property
Public Property Get getMETODO_ID() As Long
        getMETODO_ID = METODO_ID
End Property
Public Property Get getFECHA_INICIAL() As String
        getFECHA_INICIAL = FECHA_INICIAL
End Property
Public Property Get getFECHA_APROBACION() As String
        getFECHA_APROBACION = FECHA_APROBACION
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPROVEEDORES_SERVICIOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDORES_SERVICIOS WHERE PROVEEDOR_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                ORDEN = rs("ORDEN")
                SERVICIO_ID = rs("SERVICIO_ID")
                tipo = rs("TIPO")
                METODO_ID = rs("METODO_ID")
                FECHA_INICIAL = rs("FECHA_INICIAL")
                FECHA_APROBACION = rs("FECHA_APROBACION")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsProveedores_servicios)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PROVEEDORES_SERVICIOS " & _
                   "  VALUES (" & _
                        PROVEEDOR_ID & "," & ORDEN & "," & SERVICIO_ID & "," & _
                        "'" & tipo & "'" & "," & METODO_ID & "," & "'" & FECHA_INICIAL & "'" & "," & _
                        "'" & FECHA_APROBACION & "'" & "," & "'" & OBSERVACIONES & "'" & _
                ")"
        execute_bd consulta
        Insertar = PROVEEDOR_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsProveedores_servicios)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVEEDORES_SERVICIOS SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " SERVICIO_ID = " & SERVICIO_ID & "," & _
                        " TIPO = '" & tipo & "'," & _
                        " METODO_ID = " & METODO_ID & "," & _
                        " FECHA_INICIAL = '" & FECHA_INICIAL & "'," & _
                        " FECHA_APROBACION = '" & FECHA_APROBACION & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE PROVEEDOR_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsProveedores_servicios)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PROVEEDORES_SERVICIOS " & _
                "    WHERE PROVEEDOR_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsProveedores_servicios)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDORES_SERVICIOS WHERE PROVEEDOR_ID = " & ID & " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function
