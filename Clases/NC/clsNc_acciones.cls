VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNc_acciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSNC_ACCIONES
'***************************************************************
Private ID_ACCION As Long
Private NC_ID As Long
Private DESCRIPCION As String
Private FECHA_PREVISTA As String
Private FECHA_RESOLUCION As String
Private RESPONSABLE_ID As Long
Private ESTADO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSNC_ACCIONES
'***************************************************************
Public Property Let setID_ACCION(ByVal dato As Long)
        ID_ACCION = dato
End Property
Public Property Let setNC_ID(ByVal dato As Long)
        NC_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFECHA_PREVISTA(ByVal dato As String)
        FECHA_PREVISTA = dato
End Property
Public Property Let setFECHA_RESOLUCION(ByVal dato As String)
        FECHA_RESOLUCION = dato
End Property
Public Property Let setRESPONSABLE_ID(ByVal dato As Long)
        RESPONSABLE_ID = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Long)
        ESTADO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSNC_ACCIONES
'***************************************************************
Public Property Get getID_ACCION() As Long
        getID_ACCION = ID_ACCION
End Property
Public Property Get getNC_ID() As Long
        getNC_ID = NC_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFECHA_PREVISTA() As String
        getFECHA_PREVISTA = FECHA_PREVISTA
End Property
Public Property Get getFECHA_RESOLUCION() As String
        getFECHA_RESOLUCION = FECHA_RESOLUCION
End Property
Public Property Get getRESPONSABLE_ID() As Long
        getRESPONSABLE_ID = RESPONSABLE_ID
End Property
Public Property Get getESTADO_ID() As Long
        getESTADO_ID = ESTADO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSNC_ACCIONES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM NC_ACCIONES WHERE ID_ACCION = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_ACCION = rs("ID_ACCION")
                NC_ID = rs("NC_ID")
                DESCRIPCION = rs("DESCRIPCION")
                FECHA_PREVISTA = rs("FECHA_PREVISTA")
                FECHA_RESOLUCION = rs("FECHA_RESOLUCION")
                RESPONSABLE_ID = rs("RESPONSABLE_ID")
                ESTADO_ID = rs("ESTADO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsNc_acciones)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_ACCION) FROM NC_ACCIONES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_ACCION = 1
        Else
                ID_ACCION = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO NC_ACCIONES " & _
                   "  VALUES (" & _
                        ID_ACCION & "," & NC_ID & "," & "'" & DESCRIPCION & "'" & "," & _
                        "'" & FECHA_PREVISTA & "'" & "," & "'" & FECHA_RESOLUCION & "'" & "," & RESPONSABLE_ID & "," & _
                        ESTADO_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_ACCION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsNc_acciones)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE NC_ACCIONES SET " & _
                        " NC_ID = " & NC_ID & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FECHA_PREVISTA = '" & FECHA_PREVISTA & "'," & _
                        " FECHA_RESOLUCION = '" & FECHA_RESOLUCION & "'," & _
                        " RESPONSABLE_ID = " & RESPONSABLE_ID & "," & _
                        " ESTADO_ID = " & ESTADO_ID & "" & _
                " WHERE ID_ACCION = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsNc_acciones)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM NC_ACCIONES " & _
                "    WHERE NC_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsNc_acciones)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.FECHA_PREVISTA,A.DESCRIPCION,B.APELLIDOS, B.NOMBRE,C.DESCRIPCION,A.FECHA_RESOLUCION,A.ID_ACCION,A.RESPONSABLE_ID,A.ESTADO_ID " & _
                   " FROM NC_ACCIONES A,usuarios B,decodificadora C " & _
                   " WHERE A.NC_ID = " & ID & _
                   "   AND A.RESPONSABLE_ID = B.ID_EMPLEADO " & _
                   "   AND A.ESTADO_ID = C.VALOR " & _
                   "   AND C.CODIGO = " & DECODIFICADORA.NC_ESTADOS & _
                   " ORDER BY A.ID_ACCION"
        Set Listado = datos_bd(consulta)
End Function
