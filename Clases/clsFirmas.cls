VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFirmas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFIRMAS
'***************************************************************
Private ID_FIRMA As Long
Private TOBJETO As Long
Private COBJETO As Long
Private USUARIO_ID As Long
Private ROL As String
Private FIRMADA As Long
Private COMENTARIO As String
Private FTIMESTP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFIRMAS
'***************************************************************
Public Property Let setID_FIRMA(ByVal dato As Long)
        ID_FIRMA = dato
End Property
Public Property Let setTOBJETO(ByVal dato As Long)
        TOBJETO = dato
End Property
Public Property Let setCOBJETO(ByVal dato As Long)
        COBJETO = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property

Public Property Let setFIRMADA(ByVal dato As Long)
        FIRMADA = dato
End Property
Public Property Let setCOMENTARIO(ByVal dato As String)
        COMENTARIO = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFIRMAS
'***************************************************************
Public Property Get getID_FIRMA() As Long
        getID_FIRMA = ID_FIRMA
End Property
Public Property Get getTOBJETO() As Long
        getTOBJETO = TOBJETO
End Property
Public Property Get getCOBJETO() As Long
        getCOBJETO = COBJETO
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property

Public Property Get getFIRMADA() As Long
        getFIRMADA = FIRMADA
End Property
Public Property Get getCOMENTARIO() As String
        getCOMENTARIO = COMENTARIO
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFIRMAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM firmas WHERE ID_FIRMA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FIRMA = rs("ID_FIRMA")
                TOBJETO = rs("TOBJETO")
                COBJETO = rs("COBJETO")
                USUARIO_ID = rs("USUARIO_ID")
                FIRMADA = rs("FIRMADA")
                COMENTARIO = rs("COMENTARIO")
                FTIMESTP = rs("FTIMESTP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFirmas)", vbCritical, Err.Description
End Function

Public Function CargaCursoEmpleado(ID_CURSO As Long, ID_USUARIO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM firmas WHERE TOBJETO = " & TOBJETO_CURSO & " AND COBJETO=" & ID_CURSO & " AND USUARIO_ID= " & ID_USUARIO
        
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                CargaCursoEmpleado = False
        Else
                ID_FIRMA = rs("ID_FIRMA")
                TOBJETO = rs("TOBJETO")
                COBJETO = rs("COBJETO")
                USUARIO_ID = rs("USUARIO_ID")
                FIRMADA = rs("FIRMADA")
                COMENTARIO = rs("COMENTARIO")
                FTIMESTP = rs("FTIMESTP")
                CargaCursoEmpleado = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaCursoEmpleado = False
        MsgBox "Error al cargar los datos(clsFirmas)", vbCritical, Err.Description
End Function

Public Function CargaCursoAsistenciaEmpleado(ID_CURSO As Long, ID_USUARIO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM firmas WHERE TOBJETO = " & TOBJETO_INVITACION_CURSO & " AND COBJETO=" & ID_CURSO & " AND USUARIO_ID= " & ID_USUARIO
        
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                CargaCursoAsistenciaEmpleado = False
        Else
                ID_FIRMA = rs("ID_FIRMA")
                TOBJETO = rs("TOBJETO")
                COBJETO = rs("COBJETO")
                USUARIO_ID = rs("USUARIO_ID")
                FIRMADA = rs("FIRMADA")
                COMENTARIO = rs("COMENTARIO")
                FTIMESTP = rs("FTIMESTP")
                CargaCursoAsistenciaEmpleado = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaCursoAsistenciaEmpleado = False
        MsgBox "Error al cargar los datos(clsFirmas)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long

        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO firmas " & _
                   "  VALUES (" & _
                        ID_FIRMA & "," & TOBJETO & "," & COBJETO & "," & _
                        USUARIO_ID & "," & FIRMADA & "," & "'" & COMENTARIO & "'" & ",CURRENT_TIMESTAMP)"
        execute_bd consulta
        Insertar = ID_FIRMA
        Exit Function
fallo:
        Insertar = 0

        MsgBox "Error al insertar (clsFirmas)", vbCritical, Err.Description
         
End Function

Public Function InsertarUsuario() As Long
'Insertar:
'          Devuelve el código del usuario para el que se ha insertado.
'          En caso de inserción duplicada devuelve el código en negativo para gestionar los mensajes.
'          ------------------------------------------------------------------------------------------
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO firmas " & _
                   "  VALUES (" & _
                        ID_FIRMA & "," & TOBJETO & "," & COBJETO & "," & _
                        USUARIO_ID & "," & FIRMADA & "," & "'" & COMENTARIO & "'" & "," & _
                        "CURRENT_TIMESTAMP" & _
                ")"
        execute_bd consulta
        InsertarUsuario = USUARIO_ID
        Exit Function
fallo:
        InsertarUsuario = 0
        
        If Err.Number <> -2147217900 Then
        
            MsgBox "Error al insertar (clsFirmas)", vbCritical, Err.Description
        Else
            InsertarUsuario = -1000 + USUARIO_ID
        End If
End Function


Public Function Modificar() As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE firmas SET " & _
                        " TOBJETO = " & TOBJETO & "," & _
                        " COBJETO = " & COBJETO & "," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " FIRMADA = " & FIRMADA & "," & _
                        " COMENTARIO = '" & COMENTARIO & "'," & _
                        " FTIMESTP = CURRENT_TIMESTAMP" & _
                " WHERE ID_FIRMA = " & ID_FIRMA
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFirmas)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM firmas " & _
                "    WHERE ID_FIRMA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFirmas)", vbCritical, Err.Description
End Function

Public Function Eliminar_Curso(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM firmas " & _
                "    WHERE TOBJETO IN (" & TOBJETO_INVITACION_CURSO & "," & TOBJETO_ASISTENCIA_CURSO_ASISTENTES & "," & TOBJETO_ASISTENCIA_CURSO_FORMADORES & "," & TOBJETO_ASISTENCIA_CURSO_CALIDAD & ") AND COBJETO=" & ID
        execute_bd consulta
        Eliminar_Curso = True
        Exit Function
fallo:
        Eliminar_Curso = False
        MsgBox "Error al Eliminar (clsFirmas)", vbCritical, Err.Description
End Function


Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM firmas ORDER BY ID_FIRMA"
        Set Listado = datos_bd(consulta)
End Function

Public Function ListadoUsuario(ID As Long, todas As Boolean) As ADODB.Recordset
        Dim consulta As String
        Dim strUsuario As String
        Dim oUsuario As New clsUsuarios
        oUsuario.CARGAR USUARIO.getID_EMPLEADO
'        If oUsuario.getPER_RFI = 0 Then
        If todas = False Then
            strUsuario = " AND A.USUARIO_ID=" & ID
        Else
            strUsuario = " AND A.TOBJETO IN (20,21,22)"
        End If
        consulta = "SELECT A.ID_FIRMA,C.NOMBRE,B.DESCRIPCION, IF(A.TOBJETO<>" & TOBJETO_DOCUMENTO_VALORACION & ", D.DESCRIPCION,E.NOMBRE) " & _
                   "  FROM firmas A " & _
                   "  LEFT JOIN decodificadora B ON B.CODIGO = 500 AND A.TOBJETO = B.VALOR " & _
                   "  LEFT JOIN empleados C ON A.USUARIO_ID = C.ID_EMPLEADO " & _
                   "  LEFT JOIN formacion_cursos D ON A.COBJETO = D.ID_CURSO " & _
                   "  LEFT JOIN ca_documentos E ON A.COBJETO = E.ID_DOCUMENTO " & _
                   " WHERE FIRMADA=0  " & _
                   strUsuario & _
                   " ORDER BY A.ID_FIRMA "
        Set ListadoUsuario = datos_bd(consulta)
End Function

Public Function ListadoCurso(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM firmas WHERE TOBJETO IN (" & TOBJETO_INVITACION_CURSO & "," & TOBJETO_ASISTENCIA_CURSO_ASISTENTES & "," & TOBJETO_ASISTENCIA_CURSO_FORMADORES & "," & TOBJETO_ASISTENCIA_CURSO_CALIDAD & ") AND COBJETO=" & ID & " ORDER BY USUARIO_ID ASC"
        Set ListadoCurso = datos_bd(consulta)
End Function

Public Function CrearID() As Boolean

        Dim consulta As String
        Dim rs As New ADODB.Recordset
        
    On Error GoTo fallo:

        consulta = "SELECT MAX(ID_FIRMA) FROM firmas"
        
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
            ID_FIRMA = 1
        Else
            ID_FIRMA = rs.Fields(0) + 1
        End If
        
        Set rs = Nothing
        
        
        CrearID = True
        
        Exit Function
fallo:
        MsgBox "Error al obtener el nuevo id firma (clsFirmas)", vbCritical, Err.Description
        CrearID = False
End Function
