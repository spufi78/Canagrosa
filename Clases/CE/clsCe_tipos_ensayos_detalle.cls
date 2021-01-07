VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_tipos_ensayos_detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_TIPOS_ENSAYOS_DETALLE
'***************************************************************
Private TIPO_ENSAYO_ID As Long
Private ENSAYO As String
Private PROBETA As String
Private DIMENSION As String
Private CANTIDAD As String
Private IDENTIFICACION As String
Private UNIDAD_ID As Long
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_TIPOS_ENSAYOS_DETALLE
'***************************************************************
Public Property Let setTIPO_ENSAYO_ID(ByVal dato As Long)
        TIPO_ENSAYO_ID = dato
End Property
Public Property Let setENSAYO(ByVal dato As String)
        ENSAYO = dato
End Property
Public Property Let setPROBETA(ByVal dato As String)
        PROBETA = dato
End Property
Public Property Let setDIMENSION(ByVal dato As String)
        DIMENSION = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As String)
        CANTIDAD = dato
End Property
Public Property Let setIDENTIFICACION(ByVal dato As String)
        IDENTIFICACION = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Long)
        UNIDAD_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_TIPOS_ENSAYOS_DETALLE
'***************************************************************
Public Property Get getTIPO_ENSAYO_ID() As Long
        getTIPO_ENSAYO_ID = TIPO_ENSAYO_ID
End Property
Public Property Get getENSAYO() As String
        getENSAYO = ENSAYO
End Property
Public Property Get getPROBETA() As String
        getPROBETA = PROBETA
End Property
Public Property Get getDIMENSION() As String
        getDIMENSION = DIMENSION
End Property
Public Property Get getCANTIDAD() As String
        getCANTIDAD = CANTIDAD
End Property
Public Property Get getIDENTIFICACION() As String
        getIDENTIFICACION = IDENTIFICACION
End Property
Public Property Get getUNIDAD_ID() As Long
        getUNIDAD_ID = UNIDAD_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_TIPOS_ENSAYOS_DETALLE
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_TIPOS_ENSAYOS_DETALLE WHERE TIPO_ENSAYO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                TIPO_ENSAYO_ID = rs("TIPO_ENSAYO_ID")
                ENSAYO = rs("ENSAYO")
                PROBETA = rs("PROBETA")
                DIMENSION = rs("DIMENSION")
                CANTIDAD = rs("CANTIDAD")
                IDENTIFICACION = rs("IDENTIFICACION")
                UNIDAD_ID = rs("UNIDAD_ID")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_tipos_ensayos_detalle)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CE_TIPOS_ENSAYOS_DETALLE " & _
                   "  VALUES (" & _
                        TIPO_ENSAYO_ID & "," & "'" & ENSAYO & "'" & "," & "'" & PROBETA & "'" & "," & _
                        "'" & DIMENSION & "'" & "," & "'" & CANTIDAD & "'" & "," & "'" & IDENTIFICACION & "'" & "," & _
                        UNIDAD_ID & "," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = TIPO_ENSAYO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_tipos_ensayos_detalle)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_TIPOS_ENSAYOS_DETALLE SET " & _
                        " ENSAYO = '" & ENSAYO & "'," & _
                        " PROBETA = '" & PROBETA & "'," & _
                        " DIMENSION = '" & DIMENSION & "'," & _
                        " CANTIDAD = '" & CANTIDAD & "'," & _
                        " UNIDAD_ID = " & UNIDAD_ID & "," & _
                        " IDENTIFICACION = '" & IDENTIFICACION & "'," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE TIPO_ENSAYO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_tipos_ensayos_detalle)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_TIPOS_ENSAYOS_DETALLE " & _
                "    WHERE TIPO_ENSAYO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_tipos_ensayos_detalle)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ENSAYO,PROBETA,DIMENSION,CANTIDAD,IDENTIFICACION,UNIDAD_ID " & _
                   "  FROM CE_TIPOS_ENSAYOS_DETALLE " & _
                   " WHERE TIPO_ENSAYO_ID = " & ID & _
                   " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function
Public Function lista(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM CE_TIPOS_ENSAYOS_DETALLE " & _
                   " WHERE TIPO_ENSAYO_ID = " & ID & " ORDER BY ORDEN"
        Set lista = datos_bd(consulta)
End Function
