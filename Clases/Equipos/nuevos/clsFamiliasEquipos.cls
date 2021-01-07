VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFamiliasEquipos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private ID As Long
Private nombre As String
Private CODIGO As String
Private DESCRIPCION As String
Private OBSERVACIONES As String
'Private ICONO As clsArchivoAdjunto
Private ICONO As String
Private ID_AUX As Integer

Public Property Get getID() As Long
   getID = ID
End Property

Public Property Get getID_AUX() As Integer
    getID_AUX = ID_AUX
End Property

Public Property Let setID(ByVal lngid As Long)
    ID = lngid
End Property

Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property

Public Property Let setDESCRIPCION(ByVal strDescripcion As String)
    DESCRIPCION = strDescripcion
End Property

Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getCODIGO() As String
    getCODIGO = CODIGO
End Property
Public Property Get getICONO() As String
    getICONO = ICONO
End Property
Public Property Let setICONO(ByVal strDescripcion As String)
    ICONO = strDescripcion
End Property

'Public Property Set setICONO(ByRef obj As clsArchivoAdjunto)
'    Set ICONO = obj
'End Property
'Public Property Get getICONO() As clsArchivoAdjunto
'    Set getICONO = ICONO
'End Property
Public Property Let setID_AUX(ByVal VALOR As Integer)
    ID_AUX = VALOR
End Property

Public Property Let setNOMBRE(ByVal strNOMBRE As String)
    nombre = strNOMBRE
End Property
Public Property Let setCODIGO(ByVal strCODIGO As String)
    CODIGO = strCODIGO
End Property


Public Property Get getOBSERVACIONES() As String
    getOBSERVACIONES = OBSERVACIONES
End Property

Public Property Let setOBSERVACIONES(ByVal strOBSERVACIONES As String)
    OBSERVACIONES = strOBSERVACIONES
End Property

Public Function Listado(Optional ByVal filtro As String) As ADODB.Recordset
        
    Set Listado = datos_bd("SELECT * FROM decodificadora WHERE CODIGO = " & DECODIFICADORA.EQ_FAMILIAS & " AND valor > 0 " & filtro)
        
End Function

Public Function Carga(ByVal prmId As Long)

    Dim rs As ADODB.Recordset
    Dim strArr() As String
    
On Error GoTo Carga_Error

    Set rs = Listado("AND valor = " & CStr(prmId))
    
    If rs Is Nothing Then Exit Function
        
    If rs.RecordCount = 0 Then Exit Function

    rs.MoveFirst
        
    ID = prmId
    nombre = rs("DESCRIPCION")
    CODIGO = rs("PARAMETROS")
    ICONO = rs("ICONO")
    ID_AUX = 1
    
'    Set ICONO = New clsArchivoAdjunto
'    strArr = Split(ReadINI(App.Path & "\config.ini", "documentos", "familias_equipos_img"), "\")
'    ICONO.setSUBRUTA_SERVIDOR = strArr(UBound(strArr))
'
'    ICONO.setRUTA_TEMPORAL = ""
'    strRutaImagen = ReadINI(App.Path + "\config.ini", "Documentos", "familias_equipos_img")
'    ICONO.setRUTA = ReadINI(App.Path + "\config.ini", "Documentos", "familias_equipos_img") & "\" & prmId & ".jpg"
    Carga = True
    
On Error GoTo 0
    Exit Function
Carga_Error:
    Carga = False
End Function

Public Function Eliminar(Optional ByVal prmId As Long = -1) As Boolean
        
    Dim rs As ADODB.Recordset
    Dim blnOk As Boolean
    
    blnOk = False
    
    If prmId <> -1 Then ID = prmId
    
    Call Me.Carga(prmId)
    
    Set rs = datos_bd("select * from equipos where familia_id = " & CStr(ID))
    
    If Not rs Is Nothing Then
        blnOk = (rs.RecordCount = 0)
    End If
    
    If blnOk Then
        Call execute_bd("DELETE FROM decodificadora WHERE codigo = " & DECODIFICADORA.EQ_FAMILIAS & " and valor = " & CStr(ID))
    End If


'    Call ICONO.EliminarDeUbicacionFinal
    
    Eliminar = blnOk
    
End Function

Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "decodificadora"
        .setDESCRIPCION = "Familias de Equipo"
        .setPK = "ID_FAmilia"
        .setCAMPO = "NOMBRE"
        .setMUESTRA_DETALLE = True
        .setFILTRO = "ANULADO=0"
        Set .FORMULARIO = FORMULARIO
    End With
End Function

Public Function Insertar() As Boolean
        
    Dim sql As String
    Dim blnOk As Boolean
    
On Error GoTo Insertar_Error

    blnOk = False
    
    sql = "SELECT MAX(VALOR) AS ID FROM decodificadora WHERE CODIGO = " & DECODIFICADORA.EQ_FAMILIAS
    Set rs = datos_bd(sql)
    
    rs.MoveFirst
    
    ID = rs("ID") + 1
        
    sql = "INSERT INTO decodificadora (CODIGO, IDIOMA, VALOR, DESCRIPCION, PARAMETROS)"
    sql = sql & "VALUES(" & DECODIFICADORA.EQ_FAMILIAS & ", 'ES', " & CStr(ID) & ", '" & nombre & "','" & CODIGO & "')"
    
    Call execute_bd(sql)
                
'    Call ICONO.CopiarAUbicacionFinal(CStr(ID) & ".jpg")
                
    Insertar = True

On Error GoTo 0
    Exit Function
Insertar_Error:
    Insertar = False
End Function

Public Function Modificar() As Boolean
On Error GoTo Modificar_Error
        
    Dim sql As String
    Dim blnOk As Boolean
    
    blnOk = False
    
    sql = "UPDATE decodificadora "
    sql = sql & " SET descripcion = '" & nombre & "',"
    sql = sql & "     PARAMETROS = '" & CODIGO & "' "
    sql = sql & "WHERE CODIGO = " & DECODIFICADORA.EQ_FAMILIAS & " AND VALOR = " & CStr(ID)
    
    Call execute_bd(sql)
    
'    Call ICONO.CopiarAUbicacionFinal(CStr(ID) & ".jpg")

    Modificar = True

On Error GoTo 0
    Exit Function
Modificar_Error:
    Modificar = False
End Function

Private Sub Class_Initialize()
'    Set ICONO = New clsArchivoAdjunto
'    Dim strArr() As String
'    strArr = Split(ReadINI(App.Path & "\config.ini", "documentos", "familias_equipos_img"), "\")
'    ICONO.setSUBRUTA_SERVIDOR = strArr(UBound(strArr))
End Sub


