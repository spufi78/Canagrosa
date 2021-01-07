VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipos_determinacion_contratas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSTIPOS_DETERMINACION_CONTRATAS
'***************************************************************
Private ID_TIPO_DET_CONTRATA As Long
Private TIPO_DETERMINACION_ID As Long
Private CONTRATA_ID As Long
Private VALOR_REFERENCIA As String
Private NORMATIVA_APLICABLE As String
Private PRECIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSTIPOS_DETERMINACION_CONTRATAS
'***************************************************************
Public Property Let setID_TIPO_DET_CONTRATA(ByVal dato As Long)
        ID_TIPO_DET_CONTRATA = dato
End Property
Public Property Let setTIPO_DETERMINACION_ID(ByVal dato As Long)
        TIPO_DETERMINACION_ID = dato
End Property
Public Property Let setCONTRATA_ID(ByVal dato As Long)
        CONTRATA_ID = dato
End Property
Public Property Let setVALOR_REFERENCIA(ByVal dato As String)
        VALOR_REFERENCIA = dato
End Property
Public Property Let setNORMATIVA_APLICABLE(ByVal dato As String)
        NORMATIVA_APLICABLE = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSTIPOS_DETERMINACION_CONTRATAS
'***************************************************************
Public Property Get getID_TIPO_DET_CONTRATA() As Long
        getID_TIPO_DET_CONTRATA = ID_TIPO_DET_CONTRATA
End Property
Public Property Get getTIPO_DETERMINACION_ID() As Long
        getTIPO_DETERMINACION_ID = TIPO_DETERMINACION_ID
End Property
Public Property Get getCONTRATA_ID() As Long
        getCONTRATA_ID = CONTRATA_ID
End Property
Public Property Get getVALOR_REFERENCIA() As String
        getVALOR_REFERENCIA = VALOR_REFERENCIA
End Property
Public Property Get getNORMATIVA_APLICABLE() As String
        getNORMATIVA_APLICABLE = NORMATIVA_APLICABLE
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSTIPOS_DETERMINACION_CONTRATAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_DETERMINACION_CONTRATAS WHERE ID_TIPO_DET_CONTRATA = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                ID_TIPO_DET_CONTRATA = RS("ID_TIPO_DET_CONTRATA")
                TIPO_DETERMINACION_ID = RS("TIPO_DETERMINACION_ID")
                CONTRATA_ID = RS("CONTRATA_ID")
                VALOR_REFERENCIA = RS("VALOR_REFERENCIA")
                NORMATIVA_APLICABLE = RS("NORMATIVA_APLICABLE")
                PRECIO = RS("PRECIO")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsTipos_determinacion_contratas)", vbCritical, Err.Description
End Function


Public Function CargaContrataDeter(ID_CONTRATA As Long, ID_TIPO As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_DETERMINACION_CONTRATAS WHERE CONTRATA_ID = " & ID_CONTRATA & " AND TIPO_DETERMINACION_ID = " & ID_TIPO
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                CargaContrataDeter = False
        Else
                ID_TIPO_DET_CONTRATA = RS("ID_TIPO_DET_CONTRATA")
                TIPO_DETERMINACION_ID = RS("TIPO_DETERMINACION_ID")
                CONTRATA_ID = RS("CONTRATA_ID")
                VALOR_REFERENCIA = RS("VALOR_REFERENCIA")
                NORMATIVA_APLICABLE = RS("NORMATIVA_APLICABLE")
                PRECIO = RS("PRECIO")
                CargaContrataDeter = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        CargaContrataDeter = False
        MsgBox "Error al cargar los datos(clsTipos_determinacion_contratas)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO TIPOS_DETERMINACION_CONTRATAS " & _
                   "  VALUES (" & _
                        ID_TIPO_DET_CONTRATA & "," & TIPO_DETERMINACION_ID & "," & CONTRATA_ID & "," & _
                        "'" & VALOR_REFERENCIA & "'" & "," & "'" & NORMATIVA_APLICABLE & "', '" & PRECIO & "')" & _
                         " ON DUPLICATE KEY UPDATE VALOR_REFERENCIA='" & VALOR_REFERENCIA & "'," & " NORMATIVA_APLICABLE = '" & NORMATIVA_APLICABLE & "', PRECIO = '" & PRECIO & "'"
        execute_bd consulta
        Insertar = ID_TIPO_DET_CONTRATA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsTipos_determinacion_contratas)", vbCritical, Err.Description
End Function
Public Function Modificar() As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE TIPOS_DETERMINACION_CONTRATAS SET " & _
                        " TIPO_DETERMINACION_ID = " & TIPO_DETERMINACION_ID & "," & _
                        " CONTRATA_ID = " & CONTRATA_ID & "," & _
                        " VALOR_REFERENCIA = '" & VALOR_REFERENCIA & "'," & _
                        " NORMATIVA_APLICABLE = '" & NORMATIVA_APLICABLE & "'," & _
                        " PRECIO = '" & PRECIO & "'" & _
                " WHERE ID_TIPO_DET_CONTRATA = " & ID_TIPO_DET_CONTRATA
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsTipos_determinacion_contratas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_DETERMINACION_CONTRATAS " & _
                "    WHERE ID_TIPO_DET_CONTRATA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsTipos_determinacion_contratas)", vbCritical, Err.Description
End Function
Public Function Eliminar_Tipo(DETERMINACION_ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM TIPOS_DETERMINACION_CONTRATAS " & _
                "    WHERE TIPO_DETERMINACION_ID = " & DETERMINACION_ID
        execute_bd consulta
        Eliminar_Tipo = True
        Exit Function
fallo:
        Eliminar_Tipo = False
        MsgBox "Error al Eliminar_Tipo (clsTipos_determinacion_contratas)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_DETERMINACION_CONTRATAS ORDER BY ID_TIPO_DET_CONTRATA"
        Set Listado = datos_bd(consulta)
End Function

Public Function CrearId()
        Dim RS As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_TIPO_DET_CONTRATA) FROM TIPOS_DETERMINACION_CONTRATAS"
        Set RS = datos_bd(consulta)
        If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then
                ID_TIPO_DET_CONTRATA = 1
        Else
                ID_TIPO_DET_CONTRATA = RS.Fields(0) + 1
        End If
        Set RS = Nothing
End Function

Public Function Listado_Tipo(TIPO_DETERMINACION As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM TIPOS_DETERMINACION_CONTRATAS WHERE TIPO_DETERMINACION_ID = " & TIPO_DETERMINACION & " ORDER BY ID_TIPO_DET_CONTRATA"
        Set Listado_Tipo = datos_bd(consulta)
End Function
