VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOfertas_Adjuntos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOFERTAS_ADJUNTOS
'***************************************************************
Private OFERTA_ID As Long
Private ORDEN As Long
Private ruta As String
Private OBSERVACIONES As String
Private fecha As String
Private EMPLEADO_ID As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOFERTAS_ADJUNTOS
'***************************************************************
Public Property Let setOFERTA_ID(ByVal dato As Long)
        OFERTA_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        ruta = dato
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
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOFERTAS_ADJUNTOS
'***************************************************************
Public Property Get getOFERTA_ID() As Long
        getOFERTA_ID = OFERTA_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getRUTA() As String
        getRUTA = ruta
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
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOFERTAS_ADJUNTOS
'***************************************************************
Public Function Carga(ID As Long, o As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OFERTAS_ADJUNTOS WHERE OFERTA_ID = " & ID & " AND ORDEN = " & o

        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                OFERTA_ID = rs("OFERTA_ID")
                ORDEN = rs("ORDEN")
                ruta = rs("RUTA")
                OBSERVACIONES = rs("OBSERVACIONES")
                fecha = rs("FECHA")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsOFERTAS_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim rs As New ADODB.Recordset
        Dim consulta As String
        On Error GoTo fallo
        consulta = "SELECT MAX(ORDEN) FROM OFERTAS_ADJUNTOS WHERE OFERTA_ID = " & OFERTA_ID
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            ORDEN = 1
        Else
            ORDEN = rs.Fields(0) + 1
        End If
        Set rs = Nothing
        consulta = "INSERT INTO OFERTAS_ADJUNTOS " & _
                   "  VALUES (" & _
                        OFERTA_ID & "," & ORDEN & "," & "'" & ruta & "'" & "," & _
                        "'" & OBSERVACIONES & "'" & "," & "'" & fecha & "'" & "," & EMPLEADO_ID & _
                ")"
        execute_bd consulta
        ' Se inserta la ruta del documento en las observaciones del documento
        Dim oMuestra As New clsMuestra
        oMuestra.informar_adjunto OFERTA_ID, ruta
        Insertar = OFERTA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsOFERTAS_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OFERTAS_ADJUNTOS SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " RUTA = '" & ruta & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " FECHA = '" & fecha & "'," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "" & _
                " WHERE OFERTA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsOFERTAS_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long, o As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM OFERTAS_ADJUNTOS " & _
                "    WHERE OFERTA_ID = " & ID & _
                "      AND ORDEN = " & o
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsOFERTAS_ADJUNTOS)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.RUTA,A.FECHA,E.USUARIO,A.OFERTA_ID,A.ORDEN " & _
                   "  FROM OFERTAS_ADJUNTOS A, usuarios E, OFERTAS O " & _
                   " WHERE OFERTA_ID = O.ID_OFERTA " & _
                   "   AND O.NUMERO = " & ID & _
                   "   AND A.EMPLEADO_ID = E.ID_EMPLEADO " & _
                   " ORDER BY A.OFERTA_ID"
        Set Listado = datos_bd(consulta)
End Function


