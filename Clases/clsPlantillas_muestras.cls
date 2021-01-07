VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlantillas_muestras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*****************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE plantillas_muestras      *
'*****************************************************************
Private ID_PLANTILLA As Integer
Private TIPO_MUESTRA_ID As Integer
Private TIPO_ANALISIS_ID As Integer
Private ANALISIS_MODIFICADO As Integer
Private ENTIDAD_MUESTREO_ID As Integer
Private DETALLE_MUESTREO As String
Private OBSERVACIONES_MUESTREO As String
Private EMPLEADO_ID As Integer
Private FORMATO_ID As Integer
Private ENTIDAD_ENTREGA_ID As Integer
Private DETALLE_ENTREGA As String
Private OBSERVACIONES_ENTREGA As String
Private CLIENTE_ID As Integer
Private REFERENCIA_CLIENTE As String
Private precio As String
Private OBSERVACIONES As String
Private ANALISIS_DUPLICADO As Integer
Private CANTIDAD_MUESTRAS As Integer
Private NOMBRE As String
Private DESCRIPCION As String

Public Property Let setID_PLANTILLA(ByVal dato As Integer)
    ID_PLANTILLA = dato
End Property
Public Property Let setTIPO_MUESTRA_ID(ByVal dato As Integer)
    TIPO_MUESTRA_ID = dato
End Property
Public Property Let setTIPO_ANALISIS_ID(ByVal dato As Integer)
    TIPO_ANALISIS_ID = dato
End Property
Public Property Let setANALISIS_MODIFICADO(ByVal dato As Integer)
    ANALISIS_MODIFICADO = dato
End Property
Public Property Let setENTIDAD_MUESTREO_ID(ByVal dato As Integer)
    ENTIDAD_MUESTREO_ID = dato
End Property
Public Property Let setDETALLE_MUESTREO(ByVal dato As String)
    DETALLE_MUESTREO = Trim(dato)
End Property
Public Property Let setOBSERVACIONES_MUESTREO(ByVal dato As String)
    OBSERVACIONES_MUESTREO = Trim(dato)
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Integer)
    EMPLEADO_ID = dato
End Property
Public Property Let setFORMATO_ID(ByVal dato As Integer)
    FORMATO_ID = dato
End Property
Public Property Let setENTIDAD_ENTREGA_ID(ByVal dato As Integer)
    ENTIDAD_ENTREGA_ID = dato
End Property
Public Property Let setDETALLE_ENTREGA(ByVal dato As String)
    DETALLE_ENTREGA = Trim(dato)
End Property
Public Property Let setOBSERVACIONES_ENTREGA(ByVal dato As String)
    OBSERVACIONES_ENTREGA = Trim(dato)
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Integer)
    CLIENTE_ID = dato
End Property
Public Property Let setREFERENCIA_CLIENTE(ByVal dato As String)
    REFERENCIA_CLIENTE = Trim(dato)
End Property
Public Property Let setPRECIO(ByVal dato As String)
    precio = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
    OBSERVACIONES = Trim(dato)
End Property
Public Property Let setANALISIS_DUPLICADO(ByVal dato As Integer)
    ANALISIS_DUPLICADO = dato
End Property
Public Property Let setCANTIDAD_MUESTRAS(ByVal dato As Integer)
    CANTIDAD_MUESTRAS = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    NOMBRE = Trim(dato)
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = Trim(dato)
End Property

Public Property Get getID_PLANTILLA() As Integer
    getID_PLANTILLA = ID_PLANTILLA
End Property
Public Property Get getTIPO_MUESTRA_ID() As Integer
    getTIPO_MUESTRA_ID = TIPO_MUESTRA_ID
End Property
Public Property Get getTIPO_ANALISIS_ID() As Integer
    getTIPO_ANALISIS_ID = TIPO_ANALISIS_ID
End Property
Public Property Get getANALISIS_MODIFICADO() As Integer
    getANALISIS_MODIFICADO = ANALISIS_MODIFICADO
End Property
Public Property Get getENTIDAD_MUESTREO_ID() As Integer
    getENTIDAD_MUESTREO_ID = ENTIDAD_MUESTREO_ID
End Property
Public Property Get getDETALLE_MUESTREO() As String
    getDETALLE_MUESTREO = DETALLE_MUESTREO
End Property
Public Property Get getOBSERVACIONES_MUESTREO() As String
    getOBSERVACIONES_MUESTREO = OBSERVACIONES_MUESTREO
End Property
Public Property Get getEMPLEADO_ID() As Integer
    getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getFORMATO_ID() As Integer
    getFORMATO_ID = FORMATO_ID
End Property
Public Property Get getENTIDAD_ENTREGA_ID() As Integer
    getENTIDAD_ENTREGA_ID = ENTIDAD_ENTREGA_ID
End Property
Public Property Get getDETALLE_ENTREGA() As String
    getDETALLE_ENTREGA = DETALLE_ENTREGA
End Property
Public Property Get getOBSERVACIONES_ENTREGA() As String
    getOBSERVACIONES_ENTREGA = OBSERVACIONES_ENTREGA
End Property
Public Property Get getCLIENTE_ID() As Integer
    getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getREFERENCIA_CLIENTE() As String
    getREFERENCIA_CLIENTE = REFERENCIA_CLIENTE
End Property
Public Property Get getPRECIO() As String
    getPRECIO = precio
End Property
Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getANALISIS_DUPLICADO() As Integer
    getANALISIS_DUPLICADO = ANALISIS_DUPLICADO
End Property
Public Property Get getCANTIDAD_MUESTRAS() As Integer
    getCANTIDAD_MUESTRAS = CANTIDAD_MUESTRAS
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = NOMBRE
End Property
Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property
Public Function Listado() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT id_plantilla,nombre FROM plantillas_muestras order by nombre"
    Set Listado = datos_bd(consulta)
End Function
Public Function CargaPlantilla(id As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM plantillas_muestras WHERE ID_PLANTILLA=" & id
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            CargaPlantilla = False
        Else
            ID_PLANTILLA = rs.Fields("ID_PLANTILLA")
            If Not IsNull(rs.Fields("TIPO_MUESTRA_ID")) Then
                TIPO_MUESTRA_ID = rs.Fields("TIPO_MUESTRA_ID")
            Else
                TIPO_MUESTRA_ID = 0
            End If
            If Not IsNull(rs.Fields("TIPO_ANALISIS_ID")) Then
                TIPO_ANALISIS_ID = rs.Fields("TIPO_ANALISIS_ID")
            Else
                TIPO_ANALISIS_ID = 0
            End If
            If Not IsNull(rs.Fields("ANALISIS_MODIFICADO")) Then
                ANALISIS_MODIFICADO = rs.Fields("ANALISIS_MODIFICADO")
            Else
                ANALISIS_MODIFICADO = 0
            End If
            If Not IsNull(rs.Fields("ENTIDAD_MUESTREO_ID")) Then
                ENTIDAD_MUESTREO_ID = rs.Fields("ENTIDAD_MUESTREO_ID")
            Else
                ENTIDAD_MUESTREO_ID = 0
            End If
            DETALLE_MUESTREO = rs.Fields("DETALLE_MUESTREO") & ""
            OBSERVACIONES_MUESTREO = rs.Fields("OBSERVACIONES_MUESTREO") & ""
            If Not IsNull(rs.Fields("EMPLEADO_ID")) Then
                EMPLEADO_ID = rs.Fields("EMPLEADO_ID")
            Else
                EMPLEADO_ID = 0
            End If
            If Not IsNull(rs.Fields("FORMATO_ID")) Then
                FORMATO_ID = rs.Fields("FORMATO_ID")
            Else
                FORMATO_ID = 0
            End If
            If Not IsNull(rs.Fields("ENTIDAD_ENTREGA_ID")) Then
                ENTIDAD_ENTREGA_ID = rs.Fields("ENTIDAD_ENTREGA_ID")
            Else
                ENTIDAD_ENTREGA_ID = 0
            End If
            DETALLE_ENTREGA = rs.Fields("DETALLE_ENTREGA") & ""
            OBSERVACIONES_ENTREGA = rs.Fields("OBSERVACIONES_ENTREGA") & ""
            If Not IsNull(rs.Fields("CLIENTE_ID")) Then
                CLIENTE_ID = rs.Fields("CLIENTE_ID")
            Else
                CLIENTE_ID = 0
            End If
            REFERENCIA_CLIENTE = rs.Fields("REFERENCIA_CLIENTE") & ""
            precio = rs.Fields("PRECIO") & ""
            OBSERVACIONES = rs.Fields("OBSERVACIONES") & ""
            If Not IsNull(rs.Fields("ANALISIS_DUPLICADO")) Then
                ANALISIS_DUPLICADO = rs.Fields("ANALISIS_DUPLICADO")
            Else
                ANALISIS_DUPLICADO = 0
            End If
            If IsNull(rs.Fields("CANTIDAD_MUESTRAS")) Then
               CANTIDAD_MUESTRAS = 0
            Else
               CANTIDAD_MUESTRAS = rs.Fields("CANTIDAD_MUESTRAS")
            End If
            NOMBRE = rs.Fields("NOMBRE") & ""
            DESCRIPCION = rs.Fields("DESCRIPCION") & ""
            CargaPlantilla = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaPlantilla = False
        MsgBox "Error al cargar los datos de la Plantilla (CargaPlantilla)", vbCritical, Err.Description
End Function

Public Function Insertar() As Integer
    Dim consulta As String
    On Error GoTo fallo:
    Me.CrearId
    '**************** TABLA MUESTRAS *********************
    consulta = "INSERT INTO PLANTILLAS_MUESTRAS" & _
               " values(" & _
               ID_PLANTILLA & "," & TIPO_MUESTRA_ID & "," & TIPO_ANALISIS_ID & "," & _
               ANALISIS_MODIFICADO & "," & ENTIDAD_MUESTREO_ID & ",'" & DETALLE_MUESTREO & "','" & _
               OBSERVACIONES_MUESTREO & "'," & EMPLEADO_ID & "," & FORMATO_ID & "," & _
               ENTIDAD_ENTREGA_ID & ",'" & DETALLE_ENTREGA & "','" & OBSERVACIONES_ENTREGA & "'," & _
               CLIENTE_ID & ",'" & REFERENCIA_CLIENTE & "','" & precio & "','" & _
               OBSERVACIONES & "'," & ANALISIS_DUPLICADO & "," & CANTIDAD_MUESTRAS & ",'" & _
               NOMBRE & "','" & DESCRIPCION & "')"
    execute_bd consulta
    Insertar = ID_PLANTILLA
    MsgBox "Plantilla almacenada correctamente. Código : " & ID_PLANTILLA, vbInformation, App.Title
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al guardar la plantilla en la bd : " & Err.Description, vbCritical, Err.Number
End Function
Public Sub CrearId()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_PLANTILLA) FROM PLANTILLAS_MUESTRAS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_PLANTILLA = 1
    Else
        ID_PLANTILLA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub

Public Function Modificar(id As Integer) As Boolean
    Dim consulta As String
    On Error GoTo fallo:
    '**************** TABLA MUESTRAS *********************
    consulta = "UPDATE PLANTILLAS_MUESTRAS SET" & _
               " TIPO_MUESTRA_ID = " & TIPO_MUESTRA_ID & "," & _
               " TIPO_ANALISIS_ID = " & TIPO_ANALISIS_ID & "," & _
               " ANALISIS_MODIFICADO = " & ANALISIS_MODIFICADO & "," & _
               " ENTIDAD_MUESTREO_ID = " & ENTIDAD_MUESTREO_ID & "," & _
               " DETALLE_MUESTREO = '" & DETALLE_MUESTREO & "'," & _
               " OBSERVACIONES_MUESTREO = '" & OBSERVACIONES_MUESTREO & "'," & _
               " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
               " FORMATO_ID = " & FORMATO_ID & "," & _
               " ENTIDAD_ENTREGA_ID = " & ENTIDAD_ENTREGA_ID & "," & _
               " DETALLE_ENTREGA = '" & DETALLE_ENTREGA & "'," & _
               " OBSERVACIONES_ENTREGA = '" & OBSERVACIONES_ENTREGA & "'," & _
               " CLIENTE_ID = " & CLIENTE_ID & "," & _
               " REFERENCIA_CLIENTE = '" & REFERENCIA_CLIENTE & "'," & _
               " precio = '" & precio & "'," & _
               " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
               " ANALISIS_DUPLICADO  = " & ANALISIS_DUPLICADO & "," & _
               " CANTIDAD_MUESTRAS = " & CANTIDAD_MUESTRAS & "," & _
               " NOMBRE = '" & NOMBRE & "'," & _
               " DESCRIPCION = '" & DESCRIPCION & "'" & _
               " WHERE ID_PLANTILLA = " & id
    execute_bd consulta
    Modificar = True
    MsgBox "Plantilla modificada correctamente.", vbInformation, App.Title
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al guardar la plantilla en la bd : " & Err.Description, vbCritical, Err.Number
End Function


