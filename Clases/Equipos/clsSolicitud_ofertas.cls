VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSolicitud_ofertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSOLICITUD_OFERTAS
'***************************************************************
Private ID_SOLICITUD As Long
Private ID_EQUIPO As Long
Private ID_TIPO_SOLICITUD As Long
Private NOMBRE_COMERCIAL As String
Private OBSERVACIONES As String
Private ID_PROVEEDOR As Long
Private CANTIDAD As Long
Private ID_RESPONSABLE_TECNICO As Long
Private DESCRIPCION As String
Private MODELO As String
Private ID_NORMA As Long
Private FECHA_SOLICITUD As Date
Private RANGO_MIN As Currency
Private RANGO_MAX As Currency
Private INCERTIDUMBRE_MAX As Currency
Private ID_UNIDAD As Long
Private TOLERANCIA_MAX As Currency
Private CODIGO_PRODUCTO As String
Private PORC_VIDA_UTIL As Currency
Private TOLERANCIA_MIN As Currency
Private PRECISION As Currency
Private ID_USO_DESTINO As Long

Private Equipo As clsEquipos
Private Responsable As clsUsuarios

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSOLICITUD_OFERTAS
'***************************************************************
Public Property Let setID_SOLICITUD(ByVal dato As Long)
        ID_SOLICITUD = dato
End Property
Public Property Let setID_EQUIPO(ByVal dato As Long)
        ID_EQUIPO = dato
End Property
Public Property Let setID_TIPO_SOLICITUD(ByVal dato As Long)
        ID_TIPO_SOLICITUD = dato
End Property
Public Property Let setNOMBRE_COMERCIAL(ByVal dato As String)
        NOMBRE_COMERCIAL = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setID_PROVEEDOR(ByVal dato As Long)
        ID_PROVEEDOR = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As Long)
        CANTIDAD = dato
End Property
Public Property Let setID_RESPONSABLE_TECNICO(ByVal dato As Long)
        ID_RESPONSABLE_TECNICO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setMODELO(ByVal dato As String)
        MODELO = dato
End Property
Public Property Let setID_NORMA(ByVal dato As Long)
        ID_NORMA = dato
End Property
Public Property Let setFECHA_SOLICITUD(ByVal dato As Date)
        FECHA_SOLICITUD = dato
End Property
Public Property Let setRANGO_MIN(ByVal dato As Currency)
        RANGO_MIN = dato
End Property
Public Property Let setRANGO_MAX(ByVal dato As Currency)
        RANGO_MAX = dato
End Property
Public Property Let setINCERTIDUMBRE_MAX(ByVal dato As Currency)
        INCERTIDUMBRE_MAX = dato
End Property
Public Property Let setID_UNIDAD(ByVal dato As Long)
        ID_UNIDAD = dato
End Property
Public Property Let setTOLERANCIA_MAX(ByVal dato As Currency)
        TOLERANCIA_MAX = dato
End Property
Public Property Let setCODIGO_PRODUCTO(ByVal dato As String)
        CODIGO_PRODUCTO = dato
End Property
Public Property Let setPORC_VIDA_UTIL(ByVal dato As Currency)
        PORC_VIDA_UTIL = dato
End Property
Public Property Let setTOLERANCIA_MIN(ByVal dato As Currency)
        TOLERANCIA_MIN = dato
End Property
Public Property Let setPRECISION(ByVal dato As Currency)
        PRECISION = dato
End Property
Public Property Let setID_USO_DESTINO(ByVal dato As Long)
        ID_USO_DESTINO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSOLICITUD_OFERTAS
'***************************************************************
Public Property Get getID_SOLICITUD() As Long
        getID_SOLICITUD = ID_SOLICITUD
End Property
Public Property Get getID_EQUIPO() As Long
        getID_EQUIPO = ID_EQUIPO
End Property
Public Property Get getID_TIPO_SOLICITUD() As Long
        getID_TIPO_SOLICITUD = ID_TIPO_SOLICITUD
End Property
Public Property Get getNOMBRE_COMERCIAL() As String
        getNOMBRE_COMERCIAL = NOMBRE_COMERCIAL
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getID_PROVEEDOR() As Long
        getID_PROVEEDOR = ID_PROVEEDOR
End Property
Public Property Get getCANTIDAD() As Long
        getCANTIDAD = CANTIDAD
End Property
Public Property Get getID_RESPONSABLE_TECNICO() As Long
        getID_RESPONSABLE_TECNICO = ID_RESPONSABLE_TECNICO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getMODELO() As String
        getMODELO = MODELO
End Property
Public Property Get getID_NORMA() As Long
        getID_NORMA = ID_NORMA
End Property
Public Property Get getFECHA_SOLICITUD() As Date
        getFECHA_SOLICITUD = FECHA_SOLICITUD
End Property
Public Property Get getRANGO_MIN() As Currency
        getRANGO_MIN = RANGO_MIN
End Property
Public Property Get getRANGO_MAX() As Currency
        getRANGO_MAX = RANGO_MAX
End Property
Public Property Get getINCERTIDUMBRE_MAX() As Currency
        getINCERTIDUMBRE_MAX = INCERTIDUMBRE_MAX
End Property
Public Property Get getID_UNIDAD() As Long
        getID_UNIDAD = ID_UNIDAD
End Property
Public Property Get getTOLERANCIA_MAX() As Currency
        getTOLERANCIA_MAX = TOLERANCIA_MAX
End Property
Public Property Get getCODIGO_PRODUCTO() As String
        getCODIGO_PRODUCTO = CODIGO_PRODUCTO
End Property
Public Property Get getPORC_VIDA_UTIL() As Currency
        getPORC_VIDA_UTIL = PORC_VIDA_UTIL
End Property
Public Property Get getTOLERANCIA_MIN() As Currency
        getTOLERANCIA_MIN = TOLERANCIA_MIN
End Property
Public Property Get getPRECISION() As Currency
        getPRECISION = PRECISION
End Property
Public Property Get getID_USO_DESTINO() As Long
        getID_USO_DESTINO = ID_USO_DESTINO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSOLICITUD_OFERTAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM SOLICITUD_OFERTAS WHERE ID_SOLICITUD = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_SOLICITUD = rs("ID_SOLICITUD")
                ID_EQUIPO = rs("ID_EQUIPO")
                ID_TIPO_SOLICITUD = rs("ID_TIPO_SOLICITUD")
                NOMBRE_COMERCIAL = rs("NOMBRE_COMERCIAL")
                OBSERVACIONES = rs("OBSERVACIONES")
                ID_PROVEEDOR = rs("ID_PROVEEDOR")
                CANTIDAD = rs("CANTIDAD")
                ID_RESPONSABLE_TECNICO = rs("ID_RESPONSABLE_TECNICO")
                DESCRIPCION = rs("DESCRIPCION")
                MODELO = rs("MODELO")
                ID_NORMA = rs("ID_NORMA")
                FECHA_SOLICITUD = rs("FECHA_SOLICITUD")
                RANGO_MIN = rs("RANGO_MIN")
                RANGO_MAX = rs("RANGO_MAX")
                INCERTIDUMBRE_MAX = rs("INCERTIDUMBRE_MAX")
                ID_UNIDAD = rs("ID_UNIDAD")
                TOLERANCIA_MAX = rs("TOLERANCIA_MAX")
                CODIGO_PRODUCTO = rs("CODIGO_PRODUCTO")
                PORC_VIDA_UTIL = rs("PORC_VIDA_UTIL")
                TOLERANCIA_MIN = rs("TOLERANCIA_MIN")
                PRECISION = rs("PRECISION")
                ID_USO_DESTINO = rs("ID_USO_DESTINO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSolicitud_ofertas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_SOLICITUD) FROM SOLICITUD_OFERTAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_SOLICITUD = 1
        Else
                ID_SOLICITUD = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO SOLICITUD_OFERTAS " & _
                   "  VALUES (" & _
                        ID_SOLICITUD & "," & ID_EQUIPO & "," & ID_TIPO_SOLICITUD & "," & _
                        "'" & NOMBRE_COMERCIAL & "'" & "," & "'" & OBSERVACIONES & "'" & "," & ID_PROVEEDOR & "," & _
                        CANTIDAD & "," & ID_RESPONSABLE_TECNICO & "," & "'" & DESCRIPCION & "'" & "," & _
                        "'" & MODELO & "'" & "," & ID_NORMA & "," & FECHA_SOLICITUD & "," & _
                        RANGO_MIN & "," & RANGO_MAX & "," & INCERTIDUMBRE_MAX & "," & _
                        ID_UNIDAD & "," & TOLERANCIA_MAX & "," & "'" & CODIGO_PRODUCTO & "'" & "," & _
                        PORC_VIDA_UTIL & "," & TOLERANCIA_MIN & "," & PRECISION & "," & _
                        ID_USO_DESTINO & _
                ")"
        execute_bd consulta
        Insertar = ID_SOLICITUD
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSolicitud_ofertas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SOLICITUD_OFERTAS SET " & _
                        " ID_EQUIPO = " & ID_EQUIPO & "," & _
                        " ID_TIPO_SOLICITUD = " & ID_TIPO_SOLICITUD & "," & _
                        " NOMBRE_COMERCIAL = '" & NOMBRE_COMERCIAL & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " ID_PROVEEDOR = " & ID_PROVEEDOR & "," & _
                        " CANTIDAD = " & CANTIDAD & "," & _
                        " ID_RESPONSABLE_TECNICO = " & ID_RESPONSABLE_TECNICO & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " MODELO = '" & MODELO & "'," & _
                        " ID_NORMA = " & ID_NORMA & "," & _
                        " FECHA_SOLICITUD = " & FECHA_SOLICITUD & "," & _
                        " RANGO_MIN = " & RANGO_MIN & "," & _
                        " RANGO_MAX = " & RANGO_MAX & "," & _
                        " INCERTIDUMBRE_MAX = " & INCERTIDUMBRE_MAX & "," & _
                        " ID_UNIDAD = " & ID_UNIDAD & "," & _
                        " TOLERANCIA_MAX = " & TOLERANCIA_MAX & "," & _
                        " CODIGO_PRODUCTO = '" & CODIGO_PRODUCTO & "'," & _
                        " PORC_VIDA_UTIL = " & PORC_VIDA_UTIL & "," & _
                        " TOLERANCIA_MIN = " & TOLERANCIA_MIN & "," & _
                        " PRECISION = " & PRECISION & "," & _
                        " ID_USO_DESTINO = " & ID_USO_DESTINO & "" & _
                " WHERE ID_SOLICITUD = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSolicitud_ofertas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SOLICITUD_OFERTAS " & _
                "    WHERE ID_SOLICITUD = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSolicitud_ofertas)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM SOLICITUD_OFERTAS ORDER BY ID_SOLICITUD"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID_SOLICITUD,ID_EQUIPO FROM SOLICITUD_OFERTAS ORDER BY ID_EQUIPO"
        Set Listado_Combo = datos_bd(consulta)
End Function
