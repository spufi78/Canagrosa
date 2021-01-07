VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_normas_adjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_NORMAS_ADJUNTOS
'***************************************************************
Private NORMA_ID As Long
Private ORDEN As Long
Private RUTA As String
Private OBSERVACIONES As String
Private fecha As String
Private EMPLEADO_ID As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_NORMAS_ADJUNTOS
'***************************************************************
Public Property Let setNORMA_ID(ByVal dato As Long)
        NORMA_ID = dato
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
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_NORMAS_ADJUNTOS
'***************************************************************
Public Property Get getNORMA_ID() As Long
        getNORMA_ID = NORMA_ID
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
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_NORMAS_ADJUNTOS
'***************************************************************
Public Function Carga(ID As Long, o As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CA_NORMAS_ADJUNTOS WHERE NORMA_ID = " & ID & " AND ORDEN = " & o

        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                NORMA_ID = rs("NORMA_ID")
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
        MsgBox "Error al cargar los datos(clsCA_NORMAS_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim rs As New ADODB.RecordSet
        Dim consulta As String
        On Error GoTo fallo
        consulta = "SELECT MAX(ORDEN) FROM CA_NORMAS_ADJUNTOS WHERE NORMA_ID = " & NORMA_ID
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            ORDEN = 1
        Else
            ORDEN = rs.Fields(0) + 1
        End If
        Set rs = Nothing
        consulta = "INSERT INTO CA_NORMAS_ADJUNTOS " & _
                   "  VALUES (" & _
                        NORMA_ID & "," & ORDEN & "," & "'" & RUTA & "'" & "," & _
                        "'" & OBSERVACIONES & "'" & "," & "'" & fecha & "'" & "," & EMPLEADO_ID & _
                ")"
        execute_bd consulta
        ' Se inserta la ruta del documento en las observaciones del documento
        Dim oMuestra As New clsMuestra
        oMuestra.informar_adjunto NORMA_ID, RUTA
        Insertar = NORMA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCA_NORMAS_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_NORMAS_ADJUNTOS SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " RUTA = '" & RUTA & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " FECHA = '" & fecha & "'," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "" & _
                " WHERE NORMA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCA_NORMAS_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long, o As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CA_NORMAS_ADJUNTOS " & _
                "    WHERE NORMA_ID = " & ID & _
                "      AND ORDEN = " & o
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCA_NORMAS_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT A.RUTA,A.FECHA,E.USUARIO,A.NORMA_ID,A.ORDEN " & _
                   "  FROM CA_NORMAS_ADJUNTOS A, USUARIOS E " & _
                   " WHERE NORMA_ID = " & ID & _
                   "   AND A.EMPLEADO_ID = E.ID_EMPLEADO " & _
                   " ORDER BY A.NORMA_ID"
        Set Listado = datos_bd(consulta)
End Function


