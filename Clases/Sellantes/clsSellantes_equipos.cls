VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSellantes_equipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSELLANTES_EQUIPOS
'***************************************************************
Private MUESTRA_ID As Long
Private ORDEN As Long
Private EQUIPO_ID As Long
Private VERIFICACION_ID As Long
Private EN_INFORME As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSELLANTES_EQUIPOS
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setVERIFICACION_ID(ByVal dato As Long)
        VERIFICACION_ID = dato
End Property
Public Property Let setEN_INFORME(ByVal dato As Integer)
        EN_INFORME = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSELLANTES_EQUIPOS
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getVERIFICACION_ID() As Long
        getVERIFICACION_ID = VERIFICACION_ID
End Property
Public Property Get getEN_INFORME() As Integer
        getEN_INFORME = EN_INFORME
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSELLANTES_EQUIPOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SELLANTES_EQUIPOS WHERE MUESTRA_ID = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = RS("MUESTRA_ID")
                ORDEN = RS("ORDEN")
                EQUIPO_ID = RS("EQUIPO_ID")
                VERIFICACION_ID = RS("VERIFICACION_ID")
                EN_INFORME = RS("EN_INFORME")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(CLSSELLANTES_EQUIPOS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO SELLANTES_EQUIPOS " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & ORDEN & "," & EQUIPO_ID & "," & VERIFICACION_ID & "," & EN_INFORME & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (CLSSELLANTES_EQUIPOS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SELLANTES_EQUIPOS SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " EQUIPO_ID = " & EQUIPO_ID & "," & _
                        " VERIFICACION_ID = " & VERIFICACION_ID & "," & _
                        " EN_INFORME = " & EN_INFORME & "" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (CLSSELLANTES_EQUIPOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SELLANTES_EQUIPOS " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (CLSSELLANTES_EQUIPOS)", vbCritical, Err.Description
End Function
Public Function Listado(MUESTRA As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.EQUIPO_ID,B.NOMBRE,B.SERIE,A.EN_INFORME,A.ORDEN,A.VERIFICACION_ID " & _
                   "  FROM SELLANTES_EQUIPOS A, EQUIPOS B" & _
                   " WHERE A.MUESTRA_ID = " & MUESTRA & _
                   "   AND A.EQUIPO_ID = B.ID_EQUIPO " & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
