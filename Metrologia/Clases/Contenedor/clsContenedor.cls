VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContenedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCONTENEDOR
'***************************************************************
Private ID_CONTENEDOR As Long
Private contenedor As String
Private FECHA_SALIDA As String
Private OT_ID As Long
Private TIPO_CONTENEDOR_ID As Long
Private MODELO_CONTENEDOR_ID As Long
Private ESTADO_CONTENEDOR_ID As Long
Private COMENTARIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCONTENEDOR
'***************************************************************
Public Property Let setID_CONTENEDOR(ByVal dato As Long)
        ID_CONTENEDOR = dato
End Property
Public Property Let setCONTENEDOR(ByVal dato As String)
        contenedor = dato
End Property
Public Property Let setFECHA_SALIDA(ByVal dato As String)
        FECHA_SALIDA = dato
End Property
Public Property Let setCOMENTARIO(ByVal dato As String)
        COMENTARIO = dato
End Property
Public Property Let setOT_ID(ByVal dato As Long)
        OT_ID = dato
End Property
Public Property Let setTIPO_CONTENEDOR_ID(ByVal dato As Long)
        TIPO_CONTENEDOR_ID = dato
End Property
Public Property Let setMODELO_CONTENEDOR_ID(ByVal dato As Long)
        MODELO_CONTENEDOR_ID = dato
End Property
Public Property Let setESTADO_CONTENEDOR_ID(ByVal dato As Long)
        ESTADO_CONTENEDOR_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCONTENEDOR
'***************************************************************
Public Property Get getID_CONTENEDOR() As Long
        getID_CONTENEDOR = ID_CONTENEDOR
End Property
Public Property Get getCONTENEDOR() As String
        getCONTENEDOR = contenedor
End Property
Public Property Get getFECHA_SALIDA() As String
        getFECHA_SALIDA = FECHA_SALIDA
End Property
Public Property Get getCOMENTARIO() As String
        getCOMENTARIO = COMENTARIO
End Property
Public Property Get getOT_ID() As Long
        getOT_ID = OT_ID
End Property
Public Property Get getTIPO_CONTENEDOR_ID() As Long
        getTIPO_CONTENEDOR_ID = TIPO_CONTENEDOR_ID
End Property
Public Property Get getMODELO_CONTENEDOR_ID() As Long
        getMODELO_CONTENEDOR_ID = MODELO_CONTENEDOR_ID
End Property
Public Property Get getESTADO_CONTENEDOR_ID() As Long
        getESTADO_CONTENEDOR_ID = ESTADO_CONTENEDOR_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCONTENEDOR
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CONTENEDOR WHERE ID_CONTENEDOR = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_CONTENEDOR = rs("ID_CONTENEDOR")
                contenedor = rs("CONTENEDOR")
                FECHA_SALIDA = rs("FECHA_SALIDA")
                OT_ID = rs("OT_ID")
                TIPO_CONTENEDOR_ID = rs("TIPO_CONTENEDOR_ID")
                MODELO_CONTENEDOR_ID = rs("MODELO_CONTENEDOR_ID")
                ESTADO_CONTENEDOR_ID = rs("ESTADO_CONTENEDOR_ID")
                COMENTARIO = rs("COMENTARIO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsContenedor)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_CONTENEDOR) FROM CONTENEDOR"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_CONTENEDOR = 1
        Else
                ID_CONTENEDOR = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        If OT_ID <> 0 Then
            consulta = "SELECT * FROM CONTENEDOR WHERE OT_ID = " & OT_ID
            Set rs = datos_bd(consulta)
            If rs.RecordCount <> 0 Then
                Me.modificar rs(0)
            Else
                Me.CrearID
                consulta = "INSERT INTO CONTENEDOR " & _
                       "  VALUES (" & _
                            ID_CONTENEDOR & "," & "'" & contenedor & "','" & FECHA_SALIDA & "'," & OT_ID & "," & TIPO_CONTENEDOR_ID & "," & _
                            MODELO_CONTENEDOR_ID & "," & ESTADO_CONTENEDOR_ID & ",'" & COMENTARIO & "'" & _
                    ")"
                execute_bd consulta
            End If
        Else
            consulta = "SELECT * FROM CONTENEDOR WHERE CONTENEDOR = '" & contenedor & "'"
            Set rs = datos_bd(consulta)
            If rs.RecordCount <> 0 Then
                MsgBox "ATENCION: Ya existe un contenedor con esa descripción.", vbExclamation, App.Title
            End If
            Me.CrearID
            consulta = "INSERT INTO CONTENEDOR " & _
                   "  VALUES (" & _
                        ID_CONTENEDOR & "," & "'" & contenedor & "','" & FECHA_SALIDA & "'," & OT_ID & "," & TIPO_CONTENEDOR_ID & "," & _
                        MODELO_CONTENEDOR_ID & "," & ESTADO_CONTENEDOR_ID & ",'" & COMENTARIO & "'" & _
                ")"
            execute_bd consulta
        End If
        Insertar = ID_CONTENEDOR
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsContenedor)", vbCritical, Err.Description
End Function
Public Function modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CONTENEDOR SET " & _
                        " CONTENEDOR = '" & contenedor & "'," & _
                        " COMENTARIO = '" & COMENTARIO & "'," & _
                        " TIPO_CONTENEDOR_ID = " & TIPO_CONTENEDOR_ID & "," & _
                        " ESTADO_CONTENEDOR_ID = " & ESTADO_CONTENEDOR_ID & "," & _
                        " MODELO_CONTENEDOR_ID = " & MODELO_CONTENEDOR_ID & _
                " WHERE ID_CONTENEDOR = " & ID
        execute_bd consulta
        modificar = True
        Exit Function
fallo:
        modificar = False
        MsgBox "Error al Modificar (clsContenedor)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CONTENEDOR " & _
                "    WHERE ID_CONTENEDOR = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsContenedor)", vbCritical, Err.Description
End Function
Public Function Eliminar_por_contenedor(oContenedor As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CONTENEDOR " & _
                "    WHERE CONTENEDOR = '" & oContenedor & "'"
        execute_bd consulta
        Eliminar_por_contenedor = True
        Exit Function
fallo:
        Eliminar_por_contenedor = False
        MsgBox "Error al Eliminar (clsContenedor)", vbCritical, Err.Description
End Function
Public Function Listado(lTIPO As String, lMODELO As String, lESTADO As String, lContenedor As String, lsalida As Integer) As ADODB.Recordset
        Dim consulta As String
        Dim stipo As String
        Dim sMODELO As String
        Dim sESTADO As String
        Dim sSALIDA As String
        If IsNumeric(lTIPO) Then
            stipo = " AND A.TIPO_CONTENEDOR_ID = " & CInt(lTIPO)
        End If
        If IsNumeric(lMODELO) Then
            sMODELO = " AND A.MODELO_CONTENEDOR_ID = " & CInt(lMODELO)
        End If
        If IsNumeric(lESTADO) Then
            sESTADO = " AND A.ESTADO_CONTENEDOR_ID = " & CInt(lESTADO)
        End If
        If lsalida = 0 Then
            sSALIDA = " AND A.FECHA_SALIDA = ''"
        End If
        consulta = "SELECT A.ID_CONTENEDOR, A.CONTENEDOR,B.NOMBRE,C.NOMBRE,D.NOMBRE,E.WORKORDER,F.NOMBRE,E.FECHA_SERVICIO " & _
                   "  FROM CONTENEDOR A, CONTENEDOR_TIPOS B, CONTENEDOR_MODELOS C, CONTENEDOR_ESTADOS D " & _
                   "  LEFT JOIN OT E  ON A.OT_ID = E.ID_OT " & _
                   "  LEFT JOIN OT_TIPO_ORDEN F  ON E.TIPO_ORDEN_ID = F.ID_TIPO_ORDEN " & _
                   " WHERE A.TIPO_CONTENEDOR_ID = B.ID_TIPO_CONTENEDOR " & _
                   "   AND A.MODELO_CONTENEDOR_ID = C.ID_MODELO_CONTENEDOR " & _
                   "   AND A.ESTADO_CONTENEDOR_ID = D.ID_ESTADO_CONTENEDOR " & _
                   stipo & sMODELO & sESTADO & sSALIDA & _
                   "   AND A.CONTENEDOR LIKE '%" & lContenedor & "%'" & _
                   " ORDER BY A.CONTENEDOR"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_CONTENEDOR,CONTENEDOR FROM CONTENEDOR ORDER BY CONTENEDOR"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Informar_Salida_PK(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CONTENEDOR SET " & _
                   " FECHA_SALIDA = '" & Format(Date, "yyyy-mm-dd") & "'" & _
                   " WHERE ID_CONTENEDOR = " & ID
        execute_bd consulta
        Informar_Salida_PK = True
        Exit Function
fallo:
        Informar_Salida_PK = False
        MsgBox "Error al Informar_Salida (clsContenedor)", vbCritical, Err.Description
End Function
Public Function Informar_Salida_Contenedor(lContenedor As String) As Boolean
        On Error Resume Next
        Dim consulta As String
        consulta = "UPDATE CONTENEDOR SET " & _
                   " ESTADO_CONTENEDOR_ID = " & CONTENEDOR_ESTADOS.UTILIZADO & "," & _
                   " FECHA_SALIDA = '" & Format(Date, "yyyy-mm-dd") & "'" & _
                   " WHERE CONTENEDOR = '" & lContenedor & "'"
        execute_bd consulta
        Informar_Salida_Contenedor = True
End Function
Public Function En_Export(lContenedor As String, lOT As Long) As Boolean
        On Error Resume Next
        Dim consulta As String
        consulta = "UPDATE CONTENEDOR SET " & _
                   " ESTADO_CONTENEDOR_ID = " & CONTENEDOR_ESTADOS.En_Export & "," & _
                   " OT_ID = " & lOT & "," & _
                   " FECHA_SALIDA = '" & Format(Date, "yyyy-mm-dd") & "'" & _
                   " WHERE CONTENEDOR = '" & lContenedor & "'"
        execute_bd consulta
        En_Export = True
End Function
Public Function Marcar_Utilizable(lContenedor As String) As Boolean
        On Error Resume Next
        Dim consulta As String
        consulta = "UPDATE CONTENEDOR SET " & _
                   " ESTADO_CONTENEDOR_ID = " & CONTENEDOR_ESTADOS.UTILIZABLE & "," & _
                   " OT_ID = 0," & _
                   " FECHA_SALIDA = ''" & _
                   " WHERE CONTENEDOR = '" & lContenedor & "'"
        execute_bd consulta
        Marcar_Utilizable = True
End Function


