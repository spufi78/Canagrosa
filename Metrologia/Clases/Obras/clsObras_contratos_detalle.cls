VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObras_contratos_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOBRAS_CONTRATOS_DETALLE
'***************************************************************
Private CONTRATO_ID As Long
Private NOMBRE As String
Private TELEFONO As String
Private OBSERVACIONES As String
Private recepcion_fecha As String
Private recepcion_modo As Long
Private envio_fecha As String
Private envio_modo As Long
Private ADJUDICACION_FECHA As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOBRAS_CONTRATOS_DETALLE
'***************************************************************
Public Property Let setCONTRATO_ID(ByVal dato As Long)
        CONTRATO_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setTELEFONO(ByVal dato As String)
        TELEFONO = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setRECEPCION_FECHA(ByVal dato As String)
        recepcion_fecha = dato
End Property
Public Property Let setRECEPCION_MODO(ByVal dato As Long)
        recepcion_modo = dato
End Property
Public Property Let setENVIO_FECHA(ByVal dato As String)
        envio_fecha = dato
End Property
Public Property Let setENVIO_MODO(ByVal dato As Long)
        envio_modo = dato
End Property
Public Property Let setADJUDICACION_FECHA(ByVal dato As String)
        ADJUDICACION_FECHA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOBRAS_CONTRATOS_DETALLE
'***************************************************************
Public Property Get getCONTRATO_ID() As Long
        getCONTRATO_ID = CONTRATO_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getTELEFONO() As String
        getTELEFONO = TELEFONO
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getRECEPCION_FECHA() As String
        getRECEPCION_FECHA = recepcion_fecha
End Property
Public Property Get getRECEPCION_MODO() As Long
        getRECEPCION_MODO = recepcion_modo
End Property
Public Property Get getENVIO_FECHA() As String
        getENVIO_FECHA = envio_fecha
End Property
Public Property Get getENVIO_MODO() As Long
        getENVIO_MODO = envio_modo
End Property
Public Property Get getADJUDICACION_FECHA() As String
        getADJUDICACION_FECHA = ADJUDICACION_FECHA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOBRAS_CONTRATOS_DETALLE
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OBRAS_CONTRATOS_DETALLE WHERE CONTRATO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CONTRATO_ID = rs("CONTRATO_ID")
                NOMBRE = rs("NOMBRE")
                TELEFONO = rs("TELEFONO")
                OBSERVACIONES = rs("OBSERVACIONES")
                recepcion_fecha = rs("RECEPCION_FECHA")
                recepcion_modo = rs("RECEPCION_MODO")
                If IsNull(rs("ENVIO_FECHA")) Then
                    envio_fecha = "0000-00-00"
                Else
                    envio_fecha = rs("ENVIO_FECHA")
                End If
                envio_modo = rs("ENVIO_MODO")
                If IsNull(rs("ADJUDICACION_FECHA")) Then
                    ADJUDICACION_FECHA = "0000-00-00"
                Else
                    ADJUDICACION_FECHA = rs("ADJUDICACION_FECHA")
                End If
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsObras_contratos_detalle)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim oOCD As New clsObras_contratos_detalle
        If oOCD.Carga(CONTRATO_ID) = False Then
            consulta = "INSERT INTO OBRAS_CONTRATOS_DETALLE " & _
                       "  VALUES (" & _
                            CONTRATO_ID & "," & "'" & NOMBRE & "'" & "," & "'" & TELEFONO & "'" & "," & _
                            "'" & OBSERVACIONES & "'" & "," & "'" & recepcion_fecha & "'" & "," & recepcion_modo & "," & _
                            "'" & envio_fecha & "'" & "," & envio_modo & "," & "'" & ADJUDICACION_FECHA & "'" & _
                    ")"
            execute_bd consulta
        Else
            Me.Modificar CONTRATO_ID
        End If
        Insertar = CONTRATO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsObras_contratos_detalle)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OBRAS_CONTRATOS_DETALLE SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " TELEFONO = '" & TELEFONO & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " RECEPCION_FECHA = '" & recepcion_fecha & "'," & _
                        " RECEPCION_MODO = " & recepcion_modo & "," & _
                        " ENVIO_FECHA = '" & envio_fecha & "'," & _
                        " ENVIO_MODO = " & envio_modo & "," & _
                        " ADJUDICACION_FECHA = '" & ADJUDICACION_FECHA & "'" & _
                " WHERE CONTRATO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsObras_contratos_detalle)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM OBRAS_CONTRATOS_DETALLE " & _
                "    WHERE CONTRATO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsObras_contratos_detalle)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OBRAS_CONTRATOS_DETALLE ORDER BY CONTRATO_ID"
        Set Listado = datos_bd(consulta)
End Function
