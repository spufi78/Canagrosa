VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistorial_cambios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSHISTORIAL_CAMBIOS
'***************************************************************
Private tipo As Long
Private IDENTIFICADOR As Long
Private IDENTIFICADOR_TEXTO As String
Private USUARIO_ID As Long
Private MOTIVO As String
Private TS As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSHISTORIAL_CAMBIOS
'***************************************************************
Public Property Let setTIPO(ByVal dato As Long)
        tipo = dato
End Property
Public Property Let setIDENTIFICADOR(ByVal dato As Long)
        IDENTIFICADOR = dato
End Property
Public Property Let setIDENTIFICADOR_TEXTO(ByVal dato As String)
        IDENTIFICADOR_TEXTO = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setMOTIVO(ByVal dato As String)
        MOTIVO = dato
End Property
Public Property Let setTS(ByVal dato As String)
        TS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSHISTORIAL_CAMBIOS
'***************************************************************
Public Property Get getTIPO() As Long
        getTIPO = tipo
End Property
Public Property Get getIDENTIFICADOR() As Long
        getIDENTIFICADOR = IDENTIFICADOR
End Property
Public Property Get getIDENTIFICADOR_TEXTO() As String
        getIDENTIFICADOR_TEXTO = IDENTIFICADOR_TEXTO
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getMOTIVO() As String
        getMOTIVO = MOTIVO
End Property
Public Property Get getTS() As String
        getTS = TS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSHISTORIAL_CAMBIOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM HISTORIAL_CAMBIOS WHERE TIPO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                tipo = rs("TIPO")
                IDENTIFICADOR = rs("IDENTIFICADOR")
                USUARIO_ID = rs("USUARIO_ID")
                MOTIVO = rs("MOTIVO")
                TS = rs("TS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsHistorial_cambios)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        If UCase(USUARIO.getUSUARIO) = "JULIO" And tipo <> HC_INVENTARIO Then
            Exit Function
        End If
        Dim consulta As String
        consulta = "INSERT INTO HISTORIAL_CAMBIOS " & _
                   "  VALUES (" & _
                        tipo & "," & IDENTIFICADOR & "," & USUARIO_ID & "," & _
                        "'" & MOTIVO & "'" & "," & "CURRENT_TIMESTAMP" & _
                ")"
        execute_bd consulta
        Insertar = tipo
        If tipo <> HC_OFERTAS Then
            ' Enviar Correo
            Dim oParametro As New clsParametros
            Dim ASUNTO As String
            Dim mensaje As String
            Dim destinatario As String
            Dim ret As Long
            Dim m As String
            oParametro.Carga parametros.EC_CORREO, ""
            If oParametro.getVALOR <> "" Then
                destinatario = oParametro.getVALOR
                If destinatario <> "" Then
                
                    If MOTIVO = HC_CREACION Then
                        m = "Creación"
                    Else
                        m = "Modificación"
                    End If
                        
                    ASUNTO = m & " de " & decodificar_tipo(CInt(tipo)) & " : " & IDENTIFICADOR_TEXTO & " (Id:" & IDENTIFICADOR & ")"
                    
                    mensaje = "Se ha realizado la " & m & " del siguiente " & decodificar_tipo(CInt(tipo)) & " : " & IDENTIFICADOR_TEXTO & " (Id:" & IDENTIFICADOR & ")" & vbNewLine & vbNewLine
                    
                    mensaje = mensaje & " Usuario : " & USUARIO.getUSUARIO & vbNewLine
                    mensaje = mensaje & " Fecha : " & Date & vbNewLine
                    mensaje = mensaje & " Hora : " & Time & vbNewLine
                    mensaje = mensaje & " Motivo : " & MOTIVO & vbNewLine
                    
                    ret = Enviar_Mail_CDO(destinatario, ASUNTO, mensaje, vbNullString)
                    
                End If
            End If
            Set oParametro = Nothing
        End If
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsHistorial_cambios)", vbCritical, Err.Description
End Function
Public Function decodificar_tipo(tipo As Integer) As String
    Select Case tipo
    Case 1
        decodificar_tipo = "BAÑO"
    Case 2
        decodificar_tipo = "TIPO ENSAYO DE EFICACIA"
    Case 3
        decodificar_tipo = "TIPO DE ANALISIS"
    Case 4
        decodificar_tipo = "FLUIDO"
    Case 5
        decodificar_tipo = "SELLANTE"
    Case 6
        decodificar_tipo = "FICHA CONTROL EFICACIA"
    Case 7
        decodificar_tipo = "FICHA DEL BAÑO DE EFICACIA"
    Case 8
        decodificar_tipo = "TIPO DETERMINACION"
    Case 9
        decodificar_tipo = "FORMULA"
    End Select
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE HISTORIAL_CAMBIOS SET " & _
                        " IDENTIFICADOR = " & IDENTIFICADOR & "," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " MOTIVO = '" & MOTIVO & "'," & _
                        " TS = '" & TS & "'" & _
                " WHERE TIPO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsHistorial_cambios)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM HISTORIAL_CAMBIOS " & _
                "    WHERE TIPO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsHistorial_cambios)", vbCritical, Err.Description
End Function
'M0996-I
Public Function EliminarIdentificador(ID As Long, ID_IDENTIFICADOR As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM HISTORIAL_CAMBIOS " & _
                "    WHERE TIPO = " & ID & " AND IDENTIFICADOR = " & ID_IDENTIFICADOR
        execute_bd consulta
        EliminarIdentificador = True
        Exit Function
fallo:
        EliminarIdentificador = False
        MsgBox "Error al Eliminar (clsHistorial_cambios)", vbCritical, Err.Description
End Function
'M0996-F
'M1108-I
Public Function EliminarIdentificadorTS(tipo As Long, ID_IDENTIFICADOR As Long, TS As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM HISTORIAL_CAMBIOS " & _
                "    WHERE TIPO = " & tipo & " AND IDENTIFICADOR = " & ID_IDENTIFICADOR & " AND TS = '" & TS & "'"
        execute_bd consulta
        EliminarIdentificadorTS = True
        Exit Function
fallo:
        EliminarIdentificadorTS = False
        MsgBox "Error al EliminarIdentificadorTS (clsHistorial_cambios)", vbCritical, Err.Description
End Function
'M1108-F

Public Function Listado(tipo As Integer, ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT H.TS, U.USUARIO,H.MOTIVO " & _
                   " FROM HISTORIAL_CAMBIOS H, usuarios U " & _
                   " WHERE H.USUARIO_ID = U.ID_EMPLEADO " & _
                   "   AND TIPO = " & tipo & _
                   "   AND IDENTIFICADOR = " & ID & _
                   " ORDER BY TS DESC"
        Set Listado = datos_bd(consulta)
End Function
