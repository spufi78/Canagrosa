VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsErrores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSERRORES
'***************************************************************
Private ID As Long
Private PRIORIDAD_ID As Long
Private TIPO_ID As Long
Private ERROR As String
Private texto As String
Private EMPLEADO_ID As Integer
Private FECHA_ALTA As String
Private HORA_ALTA As String
Private fecha_correccion As String
Private CORREGIDO As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSERRORES
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setPRIORIDAD_ID(ByVal dato As Long)
        PRIORIDAD_ID = dato
End Property
Public Property Let setTIPO_ID(ByVal dato As Long)
        TIPO_ID = dato
End Property
Public Property Let setERROR(ByVal dato As String)
        ERROR = dato
End Property
Public Property Let setTEXTO(ByVal dato As String)
        texto = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Integer)
        EMPLEADO_ID = dato
End Property
Public Property Let setFECHA_ALTA(ByVal dato As String)
        FECHA_ALTA = dato
End Property
Public Property Let setHORA_ALTA(ByVal dato As String)
        HORA_ALTA = dato
End Property
Public Property Let setFECHA_CORRECCION(ByVal dato As String)
        fecha_correccion = dato
End Property
Public Property Let setCORREGIDO(ByVal dato As Long)
        CORREGIDO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSERRORES
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getPRIORIDAD_ID() As Long
        getPRIORIDAD_ID = PRIORIDAD_ID
End Property
Public Property Get getTIPO_ID() As Long
        getTIPO_ID = TIPO_ID
End Property
Public Property Get getERROR() As String
        getERROR = ERROR
End Property
Public Property Get getTEXTO() As String
        getTEXTO = texto
End Property
Public Property Get getEMPLEADO_ID() As Integer
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getFECHA_ALTA() As String
        getFECHA_ALTA = FECHA_ALTA
End Property
Public Property Get getHORA_ALTA() As String
        getHORA_ALTA = HORA_ALTA
End Property
Public Property Get getFECHA_CORRECCION() As String
        getFECHA_CORRECCION = fecha_correccion
End Property
Public Property Get getCORREGIDO() As Long
        getCORREGIDO = CORREGIDO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSERRORES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM errores WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                PRIORIDAD_ID = rs("PRIORIDAD_ID")
                TIPO_ID = rs("TIPO_ID")
                ERROR = rs("ERROR")
                texto = rs("TEXTO")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                FECHA_ALTA = rs("FECHA_ALTA")
                HORA_ALTA = rs("HORA_ALTA")
                fecha_correccion = rs("FECHA_CORRECCION")
                CORREGIDO = rs("CORREGIDO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsErrores)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM ERRORES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO errores " & _
                   "  VALUES (" & _
                        ID & "," & PRIORIDAD_ID & "," & TIPO_ID & "," & _
                        "'" & ERROR & "'" & "," & "'" & texto & "'" & "," & EMPLEADO_ID & "," & _
                        "'" & FECHA_ALTA & "'" & "," & "'" & HORA_ALTA & "'" & "," & "'" & fecha_correccion & "'" & "," & _
                        CORREGIDO & _
                ")"
        execute_bd consulta
        ' Crear mensaje para el usuario
        Dim oMensaje As New clsMensajes
        Dim mensaje As Integer
        With oMensaje
            .setASUNTO = "Error geslab : " & ERROR
            .setTEXTO = texto
            .setEMPLEADO_ID = EMPLEADO_ID
            .setFECHA_INICIO = Format(FECHA_ALTA, "yyyy-mm-dd")
            Select Case PRIORIDAD_ID
            Case 1
                .setFECHA_FIN = Format(DateAdd("d", 15, CDate(FECHA_ALTA)), "yyyy-mm-dd")
            Case 2
                .setFECHA_FIN = Format(DateAdd("d", 7, CDate(FECHA_ALTA)), "yyyy-mm-dd")
            Case 3
                .setFECHA_FIN = Format(DateAdd("d", 2, CDate(FECHA_ALTA)), "yyyy-mm-dd")
            Case Else
                .setFECHA_FIN = Format(DateAdd("d", 1, CDate(FECHA_ALTA)), "yyyy-mm-dd")
            End Select
            
            .setHORA_INICIO = Format(Time, "hh:mm:ss")
            .setHORA_FIN = Format(Time, "hh:mm:ss")
            .setCATEGORIA = MENSAJES_CATEGORIAS.MENSAJES_CATEGORIAS_NORMAL
            .setDURACION = 0
            
            mensaje = .Insertar
        End With
        Dim oMensaje_usuarios As New clsMensajes_usuarios
        With oMensaje_usuarios
            .setMENSAJE_ID = mensaje
            .setEMPLEADO_ID = EMPLEADO_ID
            .Insertar
        End With
        
        frmCalendario.cargar_eventos

        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsErrores)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE errores SET " & _
                        " PRIORIDAD_ID = " & PRIORIDAD_ID & "," & _
                        " TIPO_ID = " & TIPO_ID & "," & _
                        " ERROR = '" & ERROR & "'," & _
                        " TEXTO = '" & texto & "'," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
                        " FECHA_ALTA = '" & FECHA_ALTA & "'," & _
                        " HORA_ALTA = '" & HORA_ALTA & "'," & _
                        " FECHA_CORRECCION = '" & fecha_correccion & "'," & _
                        " CORREGIDO = " & CORREGIDO & "" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsErrores)", vbCritical, Err.Description
End Function
Public Function Corregir(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE errores SET " & _
                        " FECHA_CORRECCION = '" & Format(Date, "yyyy-mm-dd") & "'," & _
                        " CORREGIDO = 1" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Corregir = True
        Exit Function
fallo:
        Corregir = False
        MsgBox "Error al Corregir (clsErrores)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM errores " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsErrores)", vbCritical, Err.Description
End Function
Public Function Listado(mis_consultas As Boolean, resueltas As Boolean, tipo As String, texto As String) As ADODB.Recordset
        Dim consulta As String
        Dim s1 As String
        Dim s2 As String
        Dim s3 As String
        If mis_consultas = True Then
            s1 = " AND E.EMPLEADO_ID = " & USUARIO.getID_EMPLEADO
        End If
        If resueltas = False Then
            s2 = " AND E.CORREGIDO = 0"
        End If
        If tipo <> "" Then
            s3 = " AND E.TIPO_ID = " & CInt(tipo)
        End If
        consulta = "SELECT E.ID,E.FECHA_ALTA,D1.DESCRIPCION,D2.DESCRIPCION,E.ERROR,E.TEXTO,U.USUARIO,E.CORREGIDO,E.FECHA_CORRECCION,E.PRIORIDAD_ID,E.TIPO_ID " & _
                   "  FROM errores E, usuarios U,decodificadora D1, decodificadora D2 " & _
                   " WHERE E.EMPLEADO_ID = U.ID_EMPLEADO " & _
                   "   AND D1.CODIGO = " & DECODIFICADORA.ERRORES_PRIORIDADES & _
                   "   AND E.PRIORIDAD_ID = D1.VALOR " & _
                   "   AND D2.CODIGO = " & DECODIFICADORA.ERRORES_TIPOS & _
                   "   AND E.TIPO_ID = D2.VALOR " & _
                   "   AND E.ERROR LIKE '%" & texto & "%'" & _
                   s1 & s2 & s3 & _
                   " ORDER BY E.ID DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID,PRIORIDAD_ID FROM errores ORDER BY PRIORIDAD_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
