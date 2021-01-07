VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSellantes_Equipos_Def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSELLANTES_EQUIPOS_DEF
'***************************************************************
Private SELLANTE_ID As Long
Private EQUIPO_ID As Long
Private ORDEN As Integer
Private EN_INFORME As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSELLANTES_EQUIPOS_DEF
'***************************************************************
Public Property Let setSELLANTE_ID(ByVal dato As Long)
        SELLANTE_ID = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
Public Property Let setEN_INFORME(ByVal dato As Integer)
        EN_INFORME = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSELLANTES_EQUIPOS_DEF
'***************************************************************
Public Property Get getSELLANTE_ID() As Long
        getSELLANTE_ID = SELLANTE_ID
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
Public Property Get getEN_INFORME() As Integer
        getEN_INFORME = EN_INFORME
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSELLANTES_EQUIPOS_DEF
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SELLANTES_EQUIPOS_DEF WHERE SELLANTE_ID = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                SELLANTE_ID = RS("SELLANTE_ID")
                EQUIPO_ID = RS("EQUIPO_ID")
                ORDEN = RS("ORDEN")
                EN_INFORME = RS("EN_INFORME")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSELLANTES_EQUIPOS_DEF)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO SELLANTES_EQUIPOS_DEF " & _
                   "  VALUES (" & _
                        SELLANTE_ID & "," & EQUIPO_ID & "," & ORDEN & "," & EN_INFORME & _
                ")"
        execute_bd consulta
        Insertar = SELLANTE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSELLANTES_EQUIPOS_DEF)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SELLANTES_EQUIPOS_DEF SET " & _
                        " EQUIPO_ID = " & EQUIPO_ID & "," & _
                        " ORDEN = " & ORDEN & "," & _
                        " EN_INFORME = " & EN_INFORME & "" & _
                " WHERE SELLANTE_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSELLANTES_EQUIPOS_DEF)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SELLANTES_EQUIPOS_DEF " & _
                "    WHERE SELLANTE_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSELLANTES_EQUIPOS_DEF)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.ID_EQUIPO,B.NOMBRE,B.SERIE,A.EN_INFORME " & _
                   " FROM SELLANTES_EQUIPOS_DEF A,EQUIPOS B " & _
                   " WHERE A.EQUIPO_ID = B.ID_EQUIPO " & _
                   "   AND A.SELLANTE_ID = " & ID & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
'M1137-I
Public Function ListadoActivos(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT B.ID_EQUIPO,B.NOMBRE,B.SERIE,A.EN_INFORME " & _
                   " FROM SELLANTES_EQUIPOS_DEF A,EQUIPOS B " & _
                   " WHERE A.EQUIPO_ID = B.ID_EQUIPO " & _
                   "   AND A.SELLANTE_ID = " & ID & _
                   "   and B.ESTADO_ID = '" & ENUM_EQ_ESTADOS_ACTIVO & "'" & _
                   " ORDER BY A.ORDEN"
        Set ListadoActivos = datos_bd(consulta)
End Function
'M1137-F




