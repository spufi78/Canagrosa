VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDecodificadora_adjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDECODIFICADORA_ADJUNTOS
'***************************************************************
Private DECODIFICADORA_ID As Long
Private VALOR As Long
Private ORDEN As Long
Private RUTA As String
Private OBSERVACIONES As String
Private fecha As String
Private EMPLEADO_ID As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDECODIFICADORA_ADJUNTOS
'***************************************************************
Public Property Let setDECODIFICADORA_ID(ByVal dato As Long)
        DECODIFICADORA_ID = dato
End Property
Public Property Let setVALOR(ByVal dato As Long)
        VALOR = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        RUTA = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Integer)
        EMPLEADO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDECODIFICADORA_ADJUNTOS
'***************************************************************
Public Property Get getDECODIFICADORA_ID() As Long
        getDECODIFICADORA_ID = DECODIFICADORA_ID
End Property
Public Property Get getVALOR() As Long
        getVALOR = VALOR
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getRUTA() As String
        getRUTA = RUTA
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getEMPLEADO_ID() As Integer
        getEMPLEADO_ID = EMPLEADO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDECODIFICADORA_ADJUNTOS
'***************************************************************
Public Function Carga(ID As Long, VALOR As Long, ORDEN As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM DECODIFICADORA_ADJUNTOS " & _
                   " where DECODIFICADORA_ID = " & ID & _
                   "   And VALOR = " & VALOR & _
                   "   And ORDEN = " & ORDEN

        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DECODIFICADORA_ID = rs("DECODIFICADORA_ID")
                VALOR = rs("VALOR")
                ORDEN = rs("ORDEN")
                RUTA = rs("RUTA")
                OBSERVACIONES = rs("OBSERVACIONES")
                fecha = rs("FECHA")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDECODIFICADORA_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        On Error GoTo fallo
        consulta = "SELECT MAX(ORDEN) FROM DECODIFICADORA_ADJUNTOS WHERE DECODIFICADORA_ID = " & DECODIFICADORA_ID & " AND VALOR = " & VALOR
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            ORDEN = 1
        Else
            ORDEN = rs.Fields(0) + 1
        End If
        Set rs = Nothing
        consulta = "INSERT INTO DECODIFICADORA_ADJUNTOS " & _
                   "  VALUES (" & _
                        DECODIFICADORA_ID & "," & VALOR & "," & ORDEN & ",'" & RUTA & "'" & "," & _
                        "'" & OBSERVACIONES & "'" & "," & "'" & fecha & "'" & "," & EMPLEADO_ID & _
                ")"
        execute_bd consulta
        Insertar = DECODIFICADORA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDECODIFICADORA_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long, VALOR As Long, ORDEN As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DECODIFICADORA_ADJUNTOS " & _
                "    WHERE DECODIFICADORA_ID = " & ID & _
                "      AND VALOR = " & VALOR & _
                "      AND ORDEN = " & ORDEN
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDECODIFICADORA_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long, VALOR As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT A.RUTA,A.FECHA,E.USUARIO,A.VALOR,A.ORDEN " & _
                   "  FROM DECODIFICADORA_ADJUNTOS A, USUARIOS E " & _
                   " WHERE DECODIFICADORA_ID = " & ID & _
                   "   AND VALOR = " & VALOR & _
                   "   AND A.EMPLEADO_ID = E.ID_EMPLEADO " & _
                   " ORDER BY A.ORDEN"
        Set Listado = datos_bd(consulta)
End Function
