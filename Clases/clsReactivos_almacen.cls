VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReactivos_almacen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSREACTIVOS_ALMACEN
'***************************************************************
Private ID_REGISTRO As Long
Private REACTIVO_ID As Long
Private FECHA_REGISTRO As Date
Private TIPO_REGISTRO As Long
Private USUARIO_ID As Long
Private OBSERVACIONES As String
Private TIPO_REACTIVO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSREACTIVOS_ALMACEN
'***************************************************************
Public Property Let setID_REGISTRO(ByVal dato As Long)
        ID_REGISTRO = dato
End Property
Public Property Let setREACTIVO_ID(ByVal dato As Long)
        REACTIVO_ID = dato
End Property
Public Property Let setFECHA_REGISTRO(ByVal dato As Date)
        FECHA_REGISTRO = dato
End Property
Public Property Let setTIPO_REGISTRO(ByVal dato As Long)
        TIPO_REGISTRO = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setTIPO_REACTIVO(ByVal dato As String)
        TIPO_REACTIVO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSREACTIVOS_ALMACEN
'***************************************************************
Public Property Get getID_REGISTRO() As Long
        getID_REGISTRO = ID_REGISTRO
End Property
Public Property Get getREACTIVO_ID() As Long
        getREACTIVO_ID = REACTIVO_ID
End Property
Public Property Get getFECHA_REGISTRO() As Date
        getFECHA_REGISTRO = FECHA_REGISTRO
End Property
Public Property Get getTIPO_REGISTRO() As Long
        getTIPO_REGISTRO = TIPO_REGISTRO
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getTIPO_REACTIVO() As String
        getTIPO_REACTIVO = TIPO_REACTIVO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSREACTIVOS_ALMACEN
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REACTIVOS_ALMACEN WHERE ID_REGISTRO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_REGISTRO = rs("ID_REGISTRO")
                REACTIVO_ID = rs("REACTIVO_ID")
                FECHA_REGISTRO = rs("FECHA_REGISTRO")
                TIPO_REGISTRO = rs("TIPO_REGISTRO")
                USUARIO_ID = rs("USUARIO_ID")
                OBSERVACIONES = rs("OBSERVACIONES")
                TIPO_REACTIVO = rs("TIPO_REACTIVO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsReactivos_almacen)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_REGISTRO) FROM REACTIVOS_ALMACEN"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_REGISTRO = 1
        Else
                ID_REGISTRO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO REACTIVOS_ALMACEN (ID_REGISTRO)" & _
                   "  VALUES (" & CStr(ID_REGISTRO) & ")"
        execute_bd consulta
        Insertar = ID_REGISTRO
        
        Modificar ID_REGISTRO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsReactivos_almacen)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE REACTIVOS_ALMACEN SET " & _
                        " REACTIVO_ID = " & REACTIVO_ID & "," & _
                        " FECHA_REGISTRO = '" & Format(FECHA_REGISTRO, "yyyy-mm-dd Hh:Nn") & "'," & _
                        " TIPO_REGISTRO = " & TIPO_REGISTRO & "," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " TIPO_REACTIVO = '" & TIPO_REACTIVO & "'" & _
                " WHERE ID_REGISTRO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsReactivos_almacen)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM REACTIVOS_ALMACEN " & _
                "    WHERE ID_REGISTRO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsReactivos_almacen)", vbCritical, Err.Description
End Function
Public Function Listado(Optional ByVal filtro As String = "", Optional ByVal orden_fecha_desc As Boolean = True) As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT T_REACTIVOS.* FROM ("
        consulta = consulta & " SELECT REACTIVOS_ALMACEN.*, concat(USUARIOS.NOMBRE, ' ', USUARIOS.APELLIDOS) as USUARIO"
        consulta = consulta & ", DECODIFICADORA.DESCRIPCION AS TIPO_REGISTRO_DESCRIPCION"
        consulta = consulta & ", 'EXTERNO' AS TIPO_REACTIVO_DESCRIPCION"
        consulta = consulta & " FROM REACTIVOS_ALMACEN"
        consulta = consulta & " INNER JOIN usuarios ON REACTIVOS_ALMACEN.USUARIO_ID = usuarios.ID_EMPLEADO"
        consulta = consulta & " INNER JOIN decodificadora ON REACTIVOS_ALMACEN.TIPO_REGISTRO = DECODIFICADORA.VALOR AND DECODIFICADORA.CODIGO = " & DECODIFICADORA.REX_TIPOS_REGISTRO_ALMACEN
        consulta = consulta & " WHERE TIPO_REACTIVO = 'R'"
        consulta = consulta & " UNION"
        consulta = consulta & " SELECT REACTIVOS_ALMACEN.*, concat(USUARIOS.NOMBRE, ' ', USUARIOS.APELLIDOS) as USUARIO"
        consulta = consulta & ", DECODIFICADORA.DESCRIPCION AS TIPO_REGISTRO_DESCRIPCION"
        consulta = consulta & ", 'PROPIO' AS TIPO_REACTIVO_DESCRIPCION"
        consulta = consulta & " FROM REACTIVOS_ALMACEN"
        consulta = consulta & " INNER JOIN usuarios ON REACTIVOS_ALMACEN.USUARIO_ID = usuarios.ID_EMPLEADO"
        consulta = consulta & " INNER JOIN decodificadora ON REACTIVOS_ALMACEN.TIPO_REGISTRO = DECODIFICADORA.VALOR AND DECODIFICADORA.CODIGO = " & DECODIFICADORA.REX_TIPOS_REGISTRO_ALMACEN
        consulta = consulta & " WHERE TIPO_REACTIVO = 'P'"
        consulta = consulta & " ) AS T_REACTIVOS "
        
        If Trim(filtro) <> "" Then _
            consulta = consulta & " WHERE 1=1 " & filtro
            
        consulta = consulta & " ORDER BY ID_REGISTRO " & IIf(orden_fecha_desc, "DESC", "ASC") '& " " '& IIf(orden_fecha_desc, "DESC", "ASC")

        Set Listado = datos_bd(consulta)
        
End Function
Public Function dar_entrada_reactivo(ByVal prmId_Reactivo As Long, ByVal prmTipo_Reactivo As String, ByVal prmId_Usuario As Long) As Boolean
On Error GoTo dar_entrada_reactivo_Error
    
    REACTIVO_ID = prmId_Reactivo
    FECHA_REGISTRO = Format(Now, "yyyy-mm-dd Hh:Nn")
    TIPO_REGISTRO = C_REX_TIPO_REGISTRO_ALMACEN.REX_ENTRADA_ALMACEN
    USUARIO_ID = prmId_Usuario
    OBSERVACIONES = ""
    TIPO_REACTIVO = UCase(prmTipo_Reactivo)
    
    dar_entrada_reactivo = (Me.Insertar <> 0)
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsReactivos_almacen.dar_entrada_reactivo"
    Exit Function
dar_entrada_reactivo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsReactivos_almacen.dar_entrada_reactivo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure dar_entrada_reactivo of Módulo de clase clsReactivos_almacen" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    dar_entrada_reactivo = False
End Function

Public Function dar_salida_entrada_reactivo(ByVal prmId_Reactivo As Long, ByVal prmTipo_Reactivo As String, ByVal prmId_Usuario As Long) As Boolean
On Error GoTo dar_salida_entrada_reactivo_Error
    
    REACTIVO_ID = prmId_Reactivo
    FECHA_REGISTRO = Format(Now, "yyyy-mm-dd Hh:Nn")
    TIPO_REGISTRO = C_REX_TIPO_REGISTRO_ALMACEN.REX_SALIDA_ENTRADA_ALMACEN
    USUARIO_ID = prmId_Usuario
    OBSERVACIONES = "El reactivo ha salido del almacen, aunque no ha sido registrada su salida"
    TIPO_REACTIVO = UCase(prmTipo_Reactivo)
    
    dar_salida_entrada_reactivo = (Me.Insertar <> 0)
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsReactivos_almacen.dar_salida_entrada_reactivo"
    Exit Function
dar_salida_entrada_reactivo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsReactivos_almacen.dar_salida_entrada_reactivo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure dar_salida_entrada_reactivo of Módulo de clase clsReactivos_almacen" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    dar_salida_entrada_reactivo = False
End Function
Public Function dar_entrada_salida_reactivo(ByVal prmId_Reactivo As Long, ByVal prmTipo_Reactivo As String, ByVal prmId_Usuario As Long, ByVal prmId_Ultimo_Movimiento As Long) As Boolean

    Dim strCad As String, rs As ADODB.Recordset
    
On Error GoTo dar_entrada_salida_reactivo_Error
    
    strCad = " SELECT REACTIVOS_ALMACEN.*, CONCAT(USUARIOS.NOMBRE, ' ', USUARIOS.APELLIDOS) AS USUARIO"
    strCad = strCad & " FROM REACTIVOS_ALMACEN"
    strCad = strCad & " INNER JOIN usuarios ON REACTIVOS_ALMACEN.USUARIO_ID = usuarios.ID_EMPLEADO"
    strCad = strCad & " WHERE REACTIVOS_ALMACEN.ID_REGISTRO = " & CStr(prmId_Ultimo_Movimiento)
    
    Set rs = datos_bd(strCad)
    
    If rs.RecordCount = 0 Then
        strCad = ""
    Else
        rs.MoveFirst
        strCad = "Con Fecha de " & Format(rs("FECHA_REGISTRO"), "dd/mm/yyyy Hh:Nn") & ", el usuario " & rs("USUARIO") & " registró la Salida de este Reactivo, sin registrar la posterior entrada."
    End If
    
    REACTIVO_ID = prmId_Reactivo
    FECHA_REGISTRO = Format(Now, "yyyy-mm-dd Hh:Nn")
    TIPO_REGISTRO = C_REX_TIPO_REGISTRO_ALMACEN.REX_ENTRADA_SALIDA_ALMACEN
    USUARIO_ID = prmId_Usuario
    OBSERVACIONES = strCad
    TIPO_REACTIVO = UCase(prmTipo_Reactivo)
    
    dar_entrada_salida_reactivo = (Me.Insertar <> 0)
    
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsReactivos_almacen.dar_entrada_salida_reactivo"
    Exit Function
dar_entrada_salida_reactivo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsReactivos_almacen.dar_entrada_salida_reactivo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure dar_entrada_salida_reactivo of Módulo de clase clsReactivos_almacen" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    dar_entrada_salida_reactivo = False
    
End Function

Public Function dar_salida_reactivo(ByVal prmId_Reactivo As Long, ByVal prmTipo_Reactivo As String, ByVal prmId_Usuario As Long) As Boolean
    
On Error GoTo dar_salida_reactivo_Error
    
    REACTIVO_ID = prmId_Reactivo
    FECHA_REGISTRO = Format(Now, "yyyy-mm-dd Hh:Nn")
    TIPO_REGISTRO = C_REX_TIPO_REGISTRO_ALMACEN.REX_SALIDA_ALMACEN
    USUARIO_ID = prmId_Usuario
    OBSERVACIONES = ""
    TIPO_REACTIVO = UCase(prmTipo_Reactivo)
    
    dar_salida_reactivo = (Me.Insertar <> 0)
        
On Error GoTo 0
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsReactivos_almacen.dar_salida_reactivo"
    Exit Function
dar_salida_reactivo_Error:
    G_TRAZABILIDAD_ERROR = G_TRAZABILIDAD_ERROR & vbCrLf & " -> clsReactivos_almacen.dar_salida_reactivo"
    error_grave Err.Number & " (" & Err.Description & ") in procedure dar_salida_reactivo of Módulo de clase clsReactivos_almacen" & vbCrLf & "Trazabilidad del Error: " & G_TRAZABILIDAD_ERROR
    G_TRAZABILIDAD_ERROR = ""
    dar_salida_reactivo = False
    
End Function

Public Function dar_apertura_bote_salida_reactivo(ByVal prmId_Reactivo As Long, ByVal prmTipo_Reactivo As String, ByVal prmId_Usuario As Long) As Boolean
    
    REACTIVO_ID = prmId_Reactivo
    FECHA_REGISTRO = Format(Now, "yyyy-mm-dd Hh:Nn")
    TIPO_REGISTRO = C_REX_TIPO_REGISTRO_ALMACEN.REX_ABRIR_BOTE_SALIDA_ALMACEN
    USUARIO_ID = prmId_Usuario
    OBSERVACIONES = ""
    TIPO_REACTIVO = UCase(prmTipo_Reactivo)
    
    If (Me.Insertar <> 0) Then
        'Dim oRec As New clsBotes_ex
        'oRec.abrir prmId_Reactivo, Format(Now, "dd/mm/yyyy")
        'Set oRec = Nothing
        dar_apertura_bote_salida_reactivo = True
    Else
        dar_apertura_bote_salida_reactivo = False
    End If
    'dar_apertura_bote_salida_reactivo = (Me.Insertar <> 0)

End Function

Public Function dar_finalizacion_reactivo(ByVal prmId_Reactivo As Long, ByVal prmTipo_Reactivo As String, ByVal prmId_Usuario As Long) As Boolean
    
    'Dim oReac As Object
    Dim oReac As clsBotes_ex
    
    REACTIVO_ID = prmId_Reactivo
    FECHA_REGISTRO = Format(Now, "yyyy-mm-dd Hh:Nn")
    TIPO_REGISTRO = C_REX_TIPO_REGISTRO_ALMACEN.REX_FINALIZAR_BOTE_BAJA_BOTE
    USUARIO_ID = prmId_Usuario
    OBSERVACIONES = ""
    TIPO_REACTIVO = UCase(prmTipo_Reactivo)
    
    dar_finalizacion_reactivo = (Me.Insertar <> 0)

End Function
Public Function anular_dar_finalizacion_reactivo(ByVal prmId_Reactivo As Long, ByVal prmTipo_Reactivo As String, ByVal prmId_Usuario As Long) As Boolean
    
    'Dim oReac As Object
    Dim oReac As clsBotes_ex
    
    REACTIVO_ID = prmId_Reactivo
    FECHA_REGISTRO = Format(Now, "yyyy-mm-dd Hh:Nn")
    TIPO_REGISTRO = C_REX_TIPO_REGISTRO_ALMACEN.REX_ENTRADA_ALMACEN
    USUARIO_ID = prmId_Usuario
    OBSERVACIONES = "Anular terminación de reactivo"
    TIPO_REACTIVO = UCase(prmTipo_Reactivo)
    
    anular_dar_finalizacion_reactivo = (Me.Insertar <> 0)

End Function
Public Function dar_imprimir_etiqueta_reactivo(ByVal prmId_Reactivo As Long, ByVal prmTipo_Reactivo As String, ByVal prmId_Usuario As Long) As Boolean

End Function

Public Function Listado_por_reactivo(ByVal id_reactivo As Long, Optional ByVal prmTipo_Reactivo As String = "R", Optional ByVal orden_fecha_desc As Boolean = True, Optional ByVal filtro As String = "") As ADODB.Recordset
                
    Set Listado_por_reactivo = Listado(" AND REACTIVO_ID = " & CStr(id_reactivo) & " AND TIPO_REACTIVO = '" & prmTipo_Reactivo & "' " & filtro, orden_fecha_desc)

End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_REGISTRO,REACTIVO_ID FROM REACTIVOS_ALMACEN ORDER BY REACTIVO_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
