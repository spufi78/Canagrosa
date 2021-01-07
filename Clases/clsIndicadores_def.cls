VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIndicadores_def"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsIndicadores_def
'***************************************************************
Private INDICADOR_ID As Long
Private CAMPO_ID As Long
Private POSICION_EXCEL As String
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsIndicadores_def
'***************************************************************
Public Property Let setINDICADOR_ID(ByVal dato As Long)
        INDICADOR_ID = dato
End Property
Public Property Let setCAMPO_ID(ByVal dato As Long)
        CAMPO_ID = dato
End Property
Public Property Let setPOSICION_EXCEL(ByVal dato As String)
        POSICION_EXCEL = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsIndicadores_def
'***************************************************************
Public Property Get getINDICADOR_ID() As Long
        getINDICADOR_ID = INDICADOR_ID
End Property
Public Property Get getCAMPO_ID() As Long
        getCAMPO_ID = CAMPO_ID
End Property
Public Property Get getPOSICION_EXCEL() As String
        getPOSICION_EXCEL = POSICION_EXCEL
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsIndicadores_def
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Indicadores_def WHERE INDICADOR_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                INDICADOR_ID = rs("INDICADOR_ID")
                CAMPO_ID = rs("CAMPO_ID")
                POSICION_EXCEL = rs("POSICION_EXCEL")
                ORDEN = rs("ORDEN")
                Carga = True
            
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsIndicadores_def)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO Indicadores_def " & _
                   "  VALUES (" & _
                        INDICADOR_ID & "," & CAMPO_ID & "," & "'" & POSICION_EXCEL & "'," & _
                        ORDEN & _
                ")"
        execute_bd consulta
        Insertar = INDICADOR_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsIndicadores_def)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE Indicadores_def SET " & _
                        " INDICADOR_ID = " & INDICADOR_ID & "," & _
                        " CAMPO_ID = " & CAMPO_ID & "," & _
                        " POSICION_EXCEL = '" & POSICION_EXCEL & "'," & _
                        " ORDEN = " & ORDEN & _
                " WHERE INDICADOR_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsIndicadores_def)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM Indicadores_def " & _
                "    WHERE INDICADOR_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsIndicadores_def)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM Indicadores_def ORDER BY INDICADOR_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Campos(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT C.NOMBRE,I.POSICION_EXCEL,I.CAMPO_ID " & _
                   "  FROM Indicadores_def I, indicadores_campos C " & _
                   " WHERE I.CAMPO_ID = C.ID_CAMPO " & _
                   "   AND I.INDICADOR_ID = " & ID & _
                   " ORDER BY I.ORDEN"
        Set Listado_Campos = datos_bd(consulta)
End Function
