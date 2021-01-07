VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProveedores_Adjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPROVEEDORES_ADJUNTOS
'***************************************************************
Private PROVEEDOR_ID As Long
Private ORDEN As Long
Private DESCRIPCION As String
Private RUTA As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPROVEEDORES_ADJUNTOS
'***************************************************************
Public Property Let setPROVEEDOR_ID(ByVal dato As Long)
        PROVEEDOR_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        RUTA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPROVEEDORES_ADJUNTOS
'***************************************************************
Public Property Get getPROVEEDOR_ID() As Long
        getPROVEEDOR_ID = PROVEEDOR_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getRUTA() As String
        getRUTA = RUTA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPROVEEDORES_ADJUNTOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM PROVEEDORES_ADJUNTOS WHERE PROVEEDOR_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                PROVEEDOR_ID = rs("PROVEEDOR_ID")
                ORDEN = rs("ORDEN")
                DESCRIPCION = rs("DESCRIPCION")
                RUTA = rs("RUTA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(CLSPROVEEDORES_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PROVEEDORES_ADJUNTOS " & _
                   "  VALUES (" & _
                        PROVEEDOR_ID & "," & ORDEN & "," & "'" & DESCRIPCION & "'" & "," & _
                        "'" & RUTA & "'" & _
                ")"
        execute_bd consulta
        Insertar = PROVEEDOR_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (CLSPROVEEDORES_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PROVEEDORES_ADJUNTOS SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " RUTA = '" & RUTA & "'" & _
                " WHERE PROVEEDOR_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (CLSPROVEEDORES_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PROVEEDORES_ADJUNTOS " & _
                "    WHERE PROVEEDOR_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (CLSPROVEEDORES_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Listado(ID_CUALIFICACION As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM PROVEEDORES_ADJUNTOS " & _
                   " WHERE PROVEEDOR_ID = " & ID_CUALIFICACION & _
                   " ORDER BY PROVEEDOR_ID"
        Set Listado = datos_bd(consulta)
End Function


