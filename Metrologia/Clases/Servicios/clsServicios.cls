VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServicios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSERVICIOS
'***************************************************************
Private ID_SERVICIO As Long
Private NOMBRE As String
Private DIRECCION As String
Private CP As String
Private PROVINCIA_ID As Long
Private MUNICIPIO_ID As Long
Private OBSERVACIONES As String
Private AVISOS As String
Private PRECIO_SERVICIO As String
Private PRECIO_FUEL As String
Private COSTE_DIETA As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSERVICIOS
'***************************************************************
Public Property Let setID_SERVICIO(ByVal dato As Long)
        ID_SERVICIO = dato
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
Public Property Let setPROVINCIA_ID(ByVal dato As Long)
        PROVINCIA_ID = dato
End Property
Public Property Let setMUNICIPIO_ID(ByVal dato As Long)
        MUNICIPIO_ID = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setAVISOS(ByVal dato As String)
        AVISOS = dato
End Property
Public Property Let setPRECIO_SERVICIO(ByVal dato As String)
        PRECIO_SERVICIO = dato
End Property
Public Property Let setPRECIO_FUEL(ByVal dato As String)
        PRECIO_FUEL = dato
End Property
Public Property Let setCOSTE_DIETA(ByVal dato As String)
        COSTE_DIETA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSERVICIOS
'***************************************************************
Public Property Get getID_SERVICIO() As Long
        getID_SERVICIO = ID_SERVICIO
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
Public Property Get getPROVINCIA_ID() As Long
        getPROVINCIA_ID = PROVINCIA_ID
End Property
Public Property Get getMUNICIPIO_ID() As Long
        getMUNICIPIO_ID = MUNICIPIO_ID
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getAVISOS() As String
        getAVISOS = AVISOS
End Property
Public Property Get getPRECIO_SERVICIO() As String
        getPRECIO_SERVICIO = PRECIO_SERVICIO
End Property
Public Property Get getPRECIO_FUEL() As String
        getPRECIO_FUEL = PRECIO_FUEL
End Property
Public Property Get getCOSTE_DIETA() As String
        getCOSTE_DIETA = COSTE_DIETA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSERVICIOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SERVICIOS WHERE ID_SERVICIO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_SERVICIO = rs("ID_SERVICIO")
                NOMBRE = rs("NOMBRE")
                If Not IsNull(rs("DIRECCION")) Then
                    DIRECCION = rs("DIRECCION")
                End If
                If Not IsNull(rs("CP")) Then
                    CP = rs("CP")
                End If
                If Not IsNull(rs("PROVINCIA_ID")) Then
                    PROVINCIA_ID = rs("PROVINCIA_ID")
                End If
                If Not IsNull(rs("MUNICIPIO_ID")) Then
                    MUNICIPIO_ID = rs("MUNICIPIO_ID")
                End If
                If Not IsNull(rs("OBSERVACIONES")) Then
                    OBSERVACIONES = rs("OBSERVACIONES")
                End If
                If Not IsNull(rs("AVISOS")) Then
                    AVISOS = rs("AVISOS")
                End If
                If Not IsNull(rs("PRECIO_SERVICIO")) Then
                    PRECIO_SERVICIO = rs("PRECIO_SERVICIO")
                End If
                If Not IsNull(rs("PRECIO_FUEL")) Then
                    PRECIO_FUEL = rs("PRECIO_FUEL")
                End If
                If Not IsNull(rs("COSTE_DIETA")) Then
                    COSTE_DIETA = rs("COSTE_DIETA")
                End If
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsServicios)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_SERVICIO) FROM SERVICIOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_SERVICIO = 1
        Else
                ID_SERVICIO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO SERVICIOS " & _
                   "  VALUES (" & _
                        ID_SERVICIO & "," & "'" & NOMBRE & "'" & "," & "'" & DIRECCION & "'" & "," & _
                        "'" & CP & "'" & "," & PROVINCIA_ID & "," & MUNICIPIO_ID & "," & _
                        "'" & OBSERVACIONES & "'" & "," & "'" & AVISOS & "'" & "," & "'" & PRECIO_SERVICIO & "'" & "," & _
                        "'" & PRECIO_FUEL & "','" & COSTE_DIETA & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_SERVICIO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsServicios)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SERVICIOS SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " DIRECCION = '" & DIRECCION & "'," & _
                        " CP = '" & CP & "'," & _
                        " PROVINCIA_ID = " & PROVINCIA_ID & "," & _
                        " MUNICIPIO_ID = " & MUNICIPIO_ID & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " AVISOS = '" & AVISOS & "'," & _
                        " PRECIO_SERVICIO = '" & PRECIO_SERVICIO & "'," & _
                        " PRECIO_FUEL = '" & PRECIO_FUEL & "'," & _
                        " COSTE_DIETA = '" & COSTE_DIETA & "'" & _
                " WHERE ID_SERVICIO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsServicios)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SERVICIOS " & _
                "    WHERE ID_SERVICIO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsServicios)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.ID_SERVICIO,A.NOMBRE,A.DIRECCION,B.NOMBRE,C.NOMBRE " & _
               "  FROM SERVICIOS A, PROVINCIAS B,MUNICIPIOS C " & _
               " WHERE A.PROVINCIA_ID = B.ID_PROVINCIA " & _
               "   AND A.MUNICIPIO_ID = C.ID_MUNICIPIO " & _
               " ORDER BY A.NOMBRE"
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_PK(ID As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.ID_SERVICIO,A.NOMBRE,A.DIRECCION,B.NOMBRE,C.NOMBRE " & _
               "  FROM SERVICIOS A, PROVINCIAS B,MUNICIPIOS C " & _
               " WHERE A.PROVINCIA_ID = B.ID_PROVINCIA " & _
               "   AND A.MUNICIPIO_ID = C.ID_MUNICIPIO " & _
               "   AND A.ID_SERVICIO = " & ID & _
               " ORDER BY A.NOMBRE"
    Set Listado_PK = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_SERVICIO,NOMBRE FROM SERVICIOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Nombre_Servicios(OT As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT A.NOMBRE " & _
               "  FROM SERVICIOS A, OT_SERVICIOS B " & _
               " WHERE A.ID_SERVICIO = B.SERVICIO_ID " & _
               "   AND B.OT_ID = " & OT & _
               "   AND B.ORDEN > 1 " & _
               " ORDER BY B.ORDEN"
    Set Nombre_Servicios = datos_bd(consulta)
End Function


