VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTareas_incurridos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTAREAS_INCURRIDOS
'***************************************************************
Private ID As Long
Private USUARIO_ID As Long
Private fecha As String
Private TAREA_ID As Long
Private horas As String
Private HORA_INICIO As String
Private HORA_FIN As String
Private REFERENCIA As String
Private OBSERVACIONES As String
Private facturable As Long
Private CLIENTE_ID As Long
Private TIPO_HORA As Long
Private DOC_ID As Long
Private DIETA_MEDIA As Integer
Private DIETA_COMPLETA As Integer
Private ESTANCIA As Integer
Private DESPLAZAMIENTO As Integer
Private URGENCIA As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTAREAS_INCURRIDOS
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setTAREA_ID(ByVal dato As Long)
        TAREA_ID = dato
End Property
Public Property Let setHORAS(ByVal dato As String)
        horas = dato
End Property
Public Property Let setHORA_INICIO(ByVal dato As String)
        HORA_INICIO = dato
End Property
Public Property Let setHORA_FIN(ByVal dato As String)
        HORA_FIN = dato
End Property
Public Property Let setREFERENCIA(ByVal dato As String)
        REFERENCIA = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setFACTURABLE(ByVal dato As Long)
        facturable = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setTIPO_HORA(ByVal dato As Long)
        TIPO_HORA = dato
End Property
Public Property Let setDOC_ID(ByVal dato As Long)
        DOC_ID = dato
End Property
Public Property Let setDIETA_MEDIA(ByVal dato As Integer)
        DIETA_MEDIA = dato
End Property
Public Property Let setDIETA_COMPLETA(ByVal dato As Integer)
        DIETA_COMPLETA = dato
End Property
Public Property Let setESTANCIA(ByVal dato As Integer)
        ESTANCIA = dato
End Property
Public Property Let setDESPLAZAMIENTO(ByVal dato As Integer)
        DESPLAZAMIENTO = dato
End Property
Public Property Let setURGENCIA(ByVal dato As Integer)
        URGENCIA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTAREAS_INCURRIDOS
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getTAREA_ID() As Long
        getTAREA_ID = TAREA_ID
End Property
Public Property Get getHORAS() As String
        getHORAS = horas
End Property
Public Property Get getHORA_FIN() As String
        getHORA_FIN = HORA_FIN
End Property
Public Property Get getHORA_INICIO() As String
        getHORA_INICIO = HORA_INICIO
End Property
Public Property Get getREFERENCIA() As String
        getREFERENCIA = REFERENCIA
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getFACTURABLE() As Long
        getFACTURABLE = facturable
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getTIPO_HORA() As Long
        getTIPO_HORA = TIPO_HORA
End Property
Public Property Get getDOC_ID() As Long
        getDOC_ID = DOC_ID
End Property
Public Property Get getDIETA_MEDIA() As Integer
        getDIETA_MEDIA = DIETA_MEDIA
End Property
Public Property Get getDIETA_COMPLETA() As Integer
        getDIETA_COMPLETA = DIETA_COMPLETA
End Property
Public Property Get getESTANCIA() As Integer
        getESTANCIA = ESTANCIA
End Property
Public Property Get getDESPLAZAMIENTO() As Integer
        getDESPLAZAMIENTO = DESPLAZAMIENTO
End Property
Public Property Get getURGENCIA() As Integer
        getURGENCIA = URGENCIA
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTAREAS_INCURRIDOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM tareas_incurridos WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                USUARIO_ID = rs("USUARIO_ID")
                fecha = rs("FECHA")
                TAREA_ID = rs("TAREA_ID")
                horas = rs("HORAS")
                HORA_INICIO = rs("HORA_INICIO")
                HORA_FIN = rs("HORA_FIN")
                REFERENCIA = rs("REFERENCIA")
                OBSERVACIONES = rs("OBSERVACIONES")
                facturable = rs("FACTURABLE")
                CLIENTE_ID = rs("CLIENTE_ID")
                TIPO_HORA = rs("TIPO_HORA")
                DOC_ID = rs("DOC_ID")
                DIETA_MEDIA = rs("DIETA_MEDIA")
                DIETA_COMPLETA = rs("DIETA_COMPLETA")
                ESTANCIA = rs("ESTANCIA")
                DESPLAZAMIENTO = rs("DESPLAZAMIENTO")
                URGENCIA = rs("URGENCIA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTareas_incurridos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM TAREAS_INCURRIDOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO tareas_incurridos " & _
                   "  VALUES (" & _
                        ID & "," & USUARIO_ID & "," & "'" & fecha & "'" & "," & TAREA_ID & ",'" & _
                        horas & "','" & HORA_INICIO & "','" & HORA_FIN & "','" & REFERENCIA & "'," & "'" & OBSERVACIONES & "'," & facturable & "," & CLIENTE_ID & "," & _
                        TIPO_HORA & "," & DOC_ID & "," & DIETA_MEDIA & "," & DIETA_COMPLETA & "," & _
                        ESTANCIA & "," & DESPLAZAMIENTO & "," & URGENCIA & _
                ")"
        execute_bd consulta
        Insertar = USUARIO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTareas_incurridos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE tareas_incurridos SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " TAREA_ID = " & TAREA_ID & "," & _
                        " HORAS = '" & horas & "'," & _
                        " HORA_INICIO = '" & HORA_INICIO & "'," & _
                        " HORA_FIN = '" & HORA_FIN & "'," & _
                        " REFERENCIA = '" & REFERENCIA & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTareas_incurridos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM tareas_incurridos " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTareas_incurridos)", vbCritical, Err.Description
End Function
Public Function Listado(fecha As String, USUARIO As Integer) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.DESCRIPCION,A.OBSERVACIONES,A.HORAS,A.ID,E.DESCRIPCION,A.FACTURABLE " & _
                   " FROM tareas_incurridos A, tareas B,decodificadora E " & _
                   " WHERE FECHA = '" & Format(fecha, "yyyy-mm-dd") & "' " & _
                   "   AND E.CODIGO = " & DECODIFICADORA.TAREAS_TIPOS_HORAS & _
                   "   AND E.VALOR = A.TIPO_HORA " & _
                   "   AND A.USUARIO_ID = " & USUARIO & _
                   "   AND A.TAREA_ID = B.ID_TAREA " & _
                   " ORDER BY USUARIO_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Filtro(USUARIO As Long, tipo As Long, tarea As Long, desde As Date, hasta As Date, facturable As Boolean) As ADODB.Recordset
        Dim s As String
        If USUARIO <> 0 Then
            s = " AND A.USUARIO_ID = " & USUARIO
        End If
        If tipo <> 0 Then
            s = s & " AND B.MODULO_ID = " & tipo
        End If
        If tarea <> 0 Then
            s = s & " AND B.ID_TAREA = " & tarea
        End If
        If facturable = True Then
            s = s & " AND A.FACTURABLE = 1 "
        End If
        Dim consulta As String
        consulta = "SELECT A.ID,C.USUARIO,D.DESCRIPCION,B.DESCRIPCION,A.FECHA,A.HORAS,A.REFERENCIA,A.OBSERVACIONES,E.DESCRIPCION,A.FACTURABLE " & _
                   " FROM tareas_incurridos A, tareas B,usuarios C,decodificadora D,decodificadora E " & _
                   " WHERE A.TAREA_ID = B.ID_TAREA " & _
                   "   AND A.USUARIO_ID = C.ID_EMPLEADO " & _
                   "   AND D.CODIGO = " & DECODIFICADORA.TAREAS_MODULOS & _
                   "   AND E.CODIGO = " & DECODIFICADORA.TAREAS_TIPOS_HORAS & _
                   "   AND D.VALOR = B.MODULO_ID " & _
                   "   AND E.VALOR = A.TIPO_HORA " & _
                   "   AND A.FECHA >= '" & Format(desde, "yyyy-mm-dd") & "' " & _
                   "   AND A.FECHA <= '" & Format(hasta, "yyyy-mm-dd") & "' " & _
                   s & _
                   " ORDER BY C.USUARIO,D.DESCRIPCION,A.FECHA"
        Set Listado_Filtro = datos_bd(consulta)
End Function
