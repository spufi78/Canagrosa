VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_pnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_PNT
'***************************************************************
Private CURSO_ID As Long
Private DOCUMENTO_ID As Long
Private FTIMESTP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_PNT
'***************************************************************
Public Property Let setCURSO_ID(ByVal dato As Long)
        CURSO_ID = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_PNT
'***************************************************************
Public Property Get getCURSO_ID() As Long
        getCURSO_ID = CURSO_ID
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_PNT
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PNT WHERE CURSO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CURSO_ID = rs("CURSO_ID")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                FTIMESTP = rs("FTIMESTP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFormacion_pnt)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO FORMACION_PNT " & _
                   "  VALUES (" & _
                        CURSO_ID & "," & DOCUMENTO_ID & "," & "'" & FTIMESTP & "'" & _
                ")"
        execute_bd consulta
        Insertar = CURSO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFormacion_pnt)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FORMACION_PNT SET " & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " FTIMESTP = '" & FTIMESTP & "'" & _
                " WHERE CURSO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFormacion_pnt)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_PNT " & _
                "    WHERE CURSO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFormacion_pnt)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PNT ORDER BY CURSO_ID"
        Set Listado = datos_bd(consulta)
End Function
'M1110-I
Public Function llenar_combo(combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "CA_DOCUMENTOS D,EMPLEADOS_CUALIFICACIONES E"
        .setDESCRIPCION = "CUALIFICACION PNT"
        .setPK = "D.ID_DOCUMENTO"
        .setCAMPO = "D.CODIGO"
        .setMUESTRA_DETALLE = True
        .setFILTRO = " D.ID_DOCUMENTO = E.DOCUMENTO_ID AND E.ES_FORMADOR = 1" & filtro
        Set .FORMULARIO = FORMULARIO
    End With
'M1110-F
End Function
