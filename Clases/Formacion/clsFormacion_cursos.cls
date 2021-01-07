VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_cursos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_CURSOS
'***************************************************************
Private ID_CURSO As Long
Private COD_CURSO As Long
Private ANYO As Long
Private ID_RESPONSABLE_CALIDAD As Long
Private DOCUMENTO_ID As Long
Private DESCRIPCION As String
Private OBJETIVOS As String
Private CONTENIDO As String
Private TIPO_MODALIDAD_ID As Long
Private TIPO_NIVEL_ID As Long
Private TIPO_FORMADOR_ID As Long
Private FECHA_PREVISTA_I As String
Private FECHA_PREVISTA_F As String
Private FECHA_REAL_I As String
Private FECHA_REAL_F As String
Private NHORAS As Long
Private Evaluacion As String
Private REALIZADO As Long
Private APROBADO As Long
Private PARADO As Long
Private FTIMESTP As String
Private PLAN_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_CURSOS
'***************************************************************
Public Property Let setID_CURSO(ByVal dato As Long)
        ID_CURSO = dato
End Property
Public Property Let setCOD_CURSO(ByVal dato As Long)
        COD_CURSO = dato
End Property
Public Property Let setANYO(ByVal dato As Long)
        ANYO = dato
End Property
Public Property Let setID_RESPONSABLE_CALIDAD(ByVal dato As Long)
        ID_RESPONSABLE_CALIDAD = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setOBJETIVOS(ByVal dato As String)
        OBJETIVOS = dato
End Property
Public Property Let setCONTENIDO(ByVal dato As String)
        CONTENIDO = dato
End Property
Public Property Let setTIPO_MODALIDAD_ID(ByVal dato As Long)
        TIPO_MODALIDAD_ID = dato
End Property
Public Property Let setTIPO_NIVEL_ID(ByVal dato As Long)
        TIPO_NIVEL_ID = dato
End Property
Public Property Let setTIPO_FORMADOR_ID(ByVal dato As Long)
        TIPO_FORMADOR_ID = dato
End Property
Public Property Let setFECHA_PREVISTA_I(ByVal dato As String)
        FECHA_PREVISTA_I = dato
End Property
Public Property Let setFECHA_PREVISTA_F(ByVal dato As String)
        FECHA_PREVISTA_F = dato
End Property
Public Property Let setFECHA_REAL_I(ByVal dato As String)
        FECHA_REAL_I = dato
End Property
Public Property Let setFECHA_REAL_F(ByVal dato As String)
        FECHA_REAL_F = dato
End Property
Public Property Let setNHORAS(ByVal dato As Long)
        NHORAS = dato
End Property
Public Property Let setEVALUACION(ByVal dato As String)
        Evaluacion = dato
End Property
Public Property Let setREALIZADO(ByVal dato As Long)
        REALIZADO = dato
End Property
Public Property Let setAPROBADO(ByVal dato As Long)
        APROBADO = dato
End Property
Public Property Let setPARADO(ByVal dato As Long)
        PARADO = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
Public Property Let setPLAN_ID(ByVal dato As Long)
        PLAN_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_CURSOS
'***************************************************************
Public Property Get getID_CURSO() As Long
        getID_CURSO = ID_CURSO
End Property
Public Property Get getCOD_CURSO() As Long
        getCOD_CURSO = COD_CURSO
End Property
Public Property Get getANYO() As Long
        getANYO = ANYO
End Property
Public Property Get getID_RESPONSABLE_CALIDAD() As Long
        getID_RESPONSABLE_CALIDAD = ID_RESPONSABLE_CALIDAD
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getOBJETIVOS() As String
        getOBJETIVOS = OBJETIVOS
End Property
Public Property Get getCONTENIDO() As String
        getCONTENIDO = CONTENIDO
End Property
Public Property Get getTIPO_MODALIDAD_ID() As Long
        getTIPO_MODALIDAD_ID = TIPO_MODALIDAD_ID
End Property
Public Property Get getTIPO_NIVEL_ID() As Long
        getTIPO_NIVEL_ID = TIPO_NIVEL_ID
End Property
Public Property Get getTIPO_FORMADOR_ID() As Long
        getTIPO_FORMADOR_ID = TIPO_FORMADOR_ID
End Property
Public Property Get getFECHA_PREVISTA_I() As String
        getFECHA_PREVISTA_I = FECHA_PREVISTA_I
End Property
Public Property Get getFECHA_PREVISTA_F() As String
        getFECHA_PREVISTA_F = FECHA_PREVISTA_F
End Property
Public Property Get getFECHA_REAL_I() As String
        getFECHA_REAL_I = FECHA_REAL_I
End Property
Public Property Get getFECHA_REAL_F() As String
        getFECHA_REAL_F = FECHA_REAL_F
End Property
Public Property Get getNHORAS() As Long
        getNHORAS = NHORAS
End Property
Public Property Get getEVALUACION() As String
        getEVALUACION = Evaluacion
End Property
Public Property Get getREALIZADO() As Long
        getREALIZADO = REALIZADO
End Property
Public Property Get getAPROBADO() As Long
        getAPROBADO = APROBADO
End Property
Public Property Get getPARADO() As Long
        getPARADO = PARADO
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
Public Property Get getPLAN_ID() As Long
        getPLAN_ID = PLAN_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_CURSOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM formacion_cursos WHERE ID_CURSO = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                ID_CURSO = RS("ID_CURSO")
                COD_CURSO = RS("COD_CURSO")
                ANYO = RS("ANYO")
                ID_RESPONSABLE_CALIDAD = RS("ID_RESPONSABLE_CALIDAD")
                DOCUMENTO_ID = RS("DOCUMENTO_ID")
                DESCRIPCION = RS("DESCRIPCION")
                OBJETIVOS = RS("OBJETIVOS")
                CONTENIDO = RS("CONTENIDO")
                TIPO_MODALIDAD_ID = RS("TIPO_MODALIDAD_ID")
                TIPO_NIVEL_ID = RS("TIPO_NIVEL_ID")
                TIPO_FORMADOR_ID = RS("TIPO_FORMADOR_ID")
                FECHA_PREVISTA_I = RS("FECHA_PREVISTA_I")
                FECHA_PREVISTA_F = RS("FECHA_PREVISTA_F")
                
                If IsNull(RS("FECHA_REAL_I")) Then
                    FECHA_REAL_I = RS("FECHA_PREVISTA_I")
                Else
                    FECHA_REAL_I = RS("FECHA_REAL_I")
                End If
                
                If IsNull(RS("FECHA_REAL_F")) Then
                    FECHA_REAL_F = RS("FECHA_PREVISTA_F")
                Else
                    FECHA_REAL_F = RS("FECHA_REAL_F")
                End If
                
                NHORAS = RS("NHORAS")
                Evaluacion = RS("EVALUACION")
                REALIZADO = RS("REALIZADO")
                APROBADO = RS("APROBADO")
                PARADO = RS("PARADO")
                FTIMESTP = RS("FTIMESTP")
                PLAN_ID = RS("PLAN_ID")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFormacion_cursos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearIdCurso
        Me.MaxCodCurso ANYO
        consulta = "INSERT INTO formacion_cursos " & _
                   "  VALUES (" & _
                        ID_CURSO & "," & COD_CURSO & "," & ANYO & "," & ID_RESPONSABLE_CALIDAD & "," & _
                        DOCUMENTO_ID & "," & "'" & DESCRIPCION & "'" & "," & "'" & OBJETIVOS & "'" & "," & _
                        "'" & CONTENIDO & "'" & "," & TIPO_MODALIDAD_ID & "," & TIPO_NIVEL_ID & "," & _
                        TIPO_FORMADOR_ID & "," & "'" & FECHA_PREVISTA_I & "'" & "," & "'" & FECHA_PREVISTA_F & "'" & "," & _
                        "'" & FECHA_REAL_I & "'" & "," & "'" & FECHA_REAL_F & "'" & "," & NHORAS & "," & _
                        "'" & Evaluacion & "'" & "," & REALIZADO & "," & APROBADO & "," & _
                        PARADO & ", CURRENT_TIMESTAMP, " & PLAN_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_CURSO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFormacion_cursos)", vbCritical, Err.Description
End Function

Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE formacion_cursos SET " & _
                        " ID_RESPONSABLE_CALIDAD = " & ID_RESPONSABLE_CALIDAD & "," & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " OBJETIVOS = '" & OBJETIVOS & "'," & _
                        " CONTENIDO = '" & CONTENIDO & "'," & _
                        " TIPO_MODALIDAD_ID = " & TIPO_MODALIDAD_ID & "," & _
                        " TIPO_NIVEL_ID = " & TIPO_NIVEL_ID & "," & _
                        " TIPO_FORMADOR_ID = " & TIPO_FORMADOR_ID & "," & _
                        " FECHA_PREVISTA_I = '" & FECHA_PREVISTA_I & "'," & _
                        " FECHA_PREVISTA_F = '" & FECHA_PREVISTA_F & "'," & _
                        " FECHA_REAL_I = '" & FECHA_REAL_I & "'," & _
                        " FECHA_REAL_F = '" & FECHA_REAL_F & "'," & _
                        " NHORAS = " & NHORAS & "," & _
                        " EVALUACION = '" & Evaluacion & "'," & _
                        " REALIZADO = " & REALIZADO & "," & _
                        " APROBADO = " & APROBADO & "," & _
                        " PARADO = " & PARADO & "," & _
                        " FTIMESTP = CURRENT_TIMESTAMP," & _
                        " PLAN_ID = " & PLAN_ID & _
                " WHERE ID_CURSO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFormacion_cursos)", vbCritical, Err.Description
End Function

Public Function DesasignarPlan(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE formacion_cursos SET " & _
                        " PLAN_ID = 0" & _
                " WHERE ID_CURSO = " & ID
        execute_bd consulta
        DesasignarPlan = True
        Exit Function
fallo:
        DesasignarPlan = False
        MsgBox "Error al Modificar (clsFormacion_cursos)", vbCritical, Err.Description
End Function

Public Function AsignarPlan(ID As Long, PLAN As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE formacion_cursos SET " & _
                        " PLAN_ID = " & PLAN & _
                " WHERE ID_CURSO = " & ID
        execute_bd consulta
        AsignarPlan = True
        Exit Function
fallo:
        AsignarPlan = False
        MsgBox "Error al Modificar (clsFormacion_cursos)", vbCritical, Err.Description
End Function

Public Function Finalizar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE formacion_cursos SET " & _
                        " REALIZADO = 1," & _
                        " APROBADO = 1," & _
                        " PARADO = 0," & _
                        " FTIMESTP = CURRENT_TIMESTAMP" & _
                " WHERE ID_CURSO = " & ID
        execute_bd consulta
        Finalizar = True
        Exit Function
fallo:
        Finalizar = False
        MsgBox "Error al finalizar el curso (clsFormacion_cursos)", vbCritical, Err.Description
End Function

Public Function Parar() As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE formacion_cursos SET " & _
                        " REALIZADO = 0," & _
                        " APROBADO = 1," & _
                        " PARADO = 1," & _
                        " FTIMESTP = CURRENT_TIMESTAMP" & _
                " WHERE ID_CURSO = " & ID_CURSO
        execute_bd consulta
        Parar = True
        Exit Function
fallo:
        Parar = False
        MsgBox "Error al parar el curso (clsFormacion_cursos)", vbCritical, Err.Description
End Function

Public Function Reanudar() As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE formacion_cursos SET " & _
                        " REALIZADO = 0," & _
                        " APROBADO = 1," & _
                        " PARADO = 0," & _
                        " FTIMESTP = CURRENT_TIMESTAMP" & _
                " WHERE ID_CURSO = " & ID_CURSO
        execute_bd consulta
        Reanudar = True
        Exit Function
fallo:
        Reanudar = False
        MsgBox "Error al reanudar el curso (clsFormacion_cursos)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM formacion_cursos " & _
                "    WHERE ID_CURSO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFormacion_cursos)", vbCritical, Err.Description
End Function

Public Function Listado(COD_CURSO As String, DESCRIPCION As String, fecha_i As String, fecha_f As String, MODO As Integer, FORMACION As Integer) As ADODB.Recordset
        
        Dim consulta As String
        Dim strModalidad As String
        Dim strNivel As String
        
        strModalidad = ""
        strNivel = ""
        
        If MODO < 2 Then
           strModalidad = " AND TIPO_MODALIDAD_ID = " & MODO
        Else
           strModalidad = ""
        End If
        
        If FORMACION < 3 Then
           strNivel = " AND TIPO_NIVEL_ID = " & FORMACION
        Else
           strNivel = ""
        End If
        
        consulta = "SELECT ID_CURSO, COD_CURSO, ANYO, DESCRIPCION, TIPO_MODALIDAD_ID, TIPO_NIVEL_ID, FECHA_PREVISTA_I, FECHA_PREVISTA_F, NHORAS, REALIZADO, APROBADO, PARADO, PLAN_ID" & _
                  " FROM formacion_cursos " & _
                  " WHERE FECHA_PREVISTA_I >='" & fecha_i & "'" & " AND FECHA_PREVISTA_F <='" & fecha_f & "'" & _
                  " AND COD_CURSO LIKE '%" & COD_CURSO & "%'" & _
                  " AND DESCRIPCION LIKE '%" & DESCRIPCION & "%'" & _
                   strModalidad & _
                   strNivel & _
                   " ORDER BY COD_CURSO DESC"
        
        Set Listado = datos_bd(consulta)
End Function
Public Sub CrearIdCurso()
    Dim RS As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_CURSO) FROM formacion_cursos"
    Set RS = datos_bd(consulta)
    If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then  'si es nulo No se recupero ninguno
        ID_CURSO = 1
    Else
        ID_CURSO = RS.Fields(0) + 1
    End If
    Set RS = Nothing
End Sub

Public Sub MaxCodCurso(ANYO As Long)

    Dim RS As New ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT MAX(COD_CURSO) FROM formacion_cursos where anyo=" & ANYO
    
    Set RS = datos_bd(consulta)
    
    If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then  'si es nulo No se recupero ninguno
        COD_CURSO = 1
    Else
        COD_CURSO = RS.Fields(0) + 1
    End If
    
    Set RS = Nothing
End Sub

Public Sub MaxCodCursoModalidad(ANYO As Long, tipo As Long)

    Dim RS As New ADODB.Recordset
    Dim consulta As String
    
    consulta = "SELECT MAX(COD_CURSO) FROM formacion_cursos where anyo=" & ANYO
    consulta = consulta + " AND TIPO_MODALIDAD_ID=" & tipo
    
    Set RS = datos_bd(consulta)
    
    If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then  'si es nulo No se recupero ninguno
        COD_CURSO = 1
    Else
        COD_CURSO = RS.Fields(0) + 1
    End If
    
    Set RS = Nothing
End Sub
Public Function generar_firmas_asistencia(Curso As Long) As Boolean
    'Recorrido por los asistentes para generar una firma por cada uno de ellos
    Dim RS As ADODB.Recordset
    Dim oAsistente As New clsFormacion_asistentes
    Dim oFirma As New clsFirmas
    Dim oCurso As New clsFormacion_cursos
   On Error GoTo generar_firmas_Error

    oCurso.Carga Curso
    
    Set RS = oAsistente.Listado(Curso)
    If RS.RecordCount <> 0 Then
        Do
            With oFirma
'                .NuevoIdFirma
                'M1106-I
                '.setTOBJETO = TOBJETO.CURSO
                .setTOBJETO = TOBJETO.TOBJETO_ASISTENCIA_CURSO_ASISTENTES
                'M1106-F
                .setCOBJETO = Curso
                .setUSUARIO_ID = RS(0)
                'M1106-I
                '.setROL = "Alumno"
                'M1106-F
                .setFIRMADA = 0
                .setCOMENTARIO = "Alumno: Confirma realización del curso"
                .InsertarUsuario
            End With
            RS.MoveNext
        Loop Until RS.EOF
    End If
    Set oAsistente = Nothing
    'Recorrido por los formadores internos para generar una firma por cada uno de ellos
'    Dim rsf As ADODB.Recordset
    Dim oFormador As New clsFormacion_Formadores
    Set RS = oFormador.Listado_internos(Curso)
    
    If RS.RecordCount <> 0 Then
        Do
            With oFirma
'            .NuevoIdFirma
                'M1106-I
                '.setTOBJETO = TOBJETO.TOBJETO_CURSO
                .setTOBJETO = TOBJETO.TOBJETO_ASISTENCIA_CURSO_FORMADORES
                'M1106-F
                .setCOBJETO = Curso
                .setUSUARIO_ID = RS(0)
                'M1106-I
                '.setROL = "Formador"
                'M1106-F
                .setFIRMADA = 0
                .setCOMENTARIO = "Formador: Firma de confirmación"
                .InsertarUsuario
            End With
            RS.MoveNext
        Loop Until RS.EOF
    End If
    
    'Control de calidad. Firmado por defecto.
    'Sirve para evitar accesos y tener al responsable de calidad como parte del curso.
    With oFirma
'    oFirma.NuevoIdFirma
        'M1106-I
        '.setTOBJETO = TOBJETO.TOBJETO_CURSO
        .setTOBJETO = TOBJETO.TOBJETO_ASISTENCIA_CURSO_CALIDAD
        'M1106-F
        .setCOBJETO = Curso
        .setUSUARIO_ID = oCurso.getID_RESPONSABLE_CALIDAD
        'M1106-I
        '.setROL = "Control Calidad"
        'M1106-F
        .setFIRMADA = 0
        .setCOMENTARIO = "Control de calidad"
        .InsertarUsuario
    End With
    Set oFirma = Nothing
    Set oFormador = Nothing
    
    oCurso.setREALIZADO = 1
    oCurso.Finalizar (Curso)
    
    generar_firmas_asistencia = True

   On Error GoTo 0
   Exit Function

generar_firmas_Error:
    generar_firmas_asistencia = False

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_firmas of Módulo de clase clsFormacion_cursos"
End Function

Public Function generar_firmas_calidad(Curso As Long) As Boolean
    
    Dim oFirma As New clsFirmas
    Dim oCurso As New clsFormacion_cursos
    On Error GoTo generar_firmas_calidad_Error

    oCurso.Carga Curso
    
    'Control de calidad. Firmado por defecto.
    'Sirve para evitar accesos y tener al responsable de calidad como parte del curso.
    With oFirma
        .setTOBJETO = TOBJETO.TOBJETO_ASISTENCIA_CURSO_CALIDAD
        .setCOBJETO = Curso
        .setUSUARIO_ID = oCurso.getID_RESPONSABLE_CALIDAD
        .setFIRMADA = 0
        .setCOMENTARIO = "Control de calidad"
        .InsertarUsuario
    End With
    Set oFirma = Nothing
    generar_firmas_calidad = True

   On Error GoTo 0
   Exit Function

generar_firmas_calidad_Error:
    generar_firmas_calidad = False

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_firmas_calidad of Módulo de clase clsFormacion_cursos"
End Function
Public Function generar_firmas_invitaciones(Curso As Long) As Boolean
    'Recorrido por los asistentes para generar una firma por cada uno de ellos
    Dim RS As ADODB.Recordset
    Dim oAsistente As New clsFormacion_asistentes
    Dim oFirma As New clsFirmas
    Dim oCurso As New clsFormacion_cursos
    Dim empleadoduplicado As Long
    Dim men As Integer
    
   On Error GoTo generar_firmas_Error

    oCurso.Carga Curso
    men = crea_mensaje(Curso)

    Set RS = oAsistente.Listado(Curso)
    If RS.RecordCount <> 0 Then
        Do
            With oFirma
'                .NuevoIdFirma
                'M1106-I
                '.setTOBJETO = TOBJETO.TOBJETO_ASISTENCIA_CURSO
                .setTOBJETO = TOBJETO.TOBJETO_INVITACION_CURSO
                'M1106-F
                .setCOBJETO = Curso
                .setUSUARIO_ID = RS(0)
                .setFIRMADA = 0
                .setCOMENTARIO = "Alumno: Firma de confirmación de asistencia"
                empleadoduplicado = .InsertarUsuario
                
                If empleadoduplicado > 0 Then
                    'Sólo se envía el mensaje si es la primera vez que se inserta la firma para el usuario
                   envia_mensaje_usuario men, empleadoduplicado
                End If
                
            End With
            RS.MoveNext
        Loop Until RS.EOF
    End If
    
    Set oAsistente = Nothing
    Set oFirma = Nothing
    Set oCurso = Nothing
    Set RS = Nothing
  
    generar_firmas_invitaciones = True

   On Error GoTo 0
   Exit Function

generar_firmas_Error:
    generar_firmas_invitaciones = False

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_firmas_invitaciones of Módulo de clase clsFormacion_cursos"
End Function

Private Function crea_mensaje(Curso As Long) As Integer
'Genera el mensaje destinado a los asistentes al curso

    On Error GoTo crea_mensaje_Error

    Dim oMensaje As New clsMensajes
    Dim oCurso As New clsFormacion_cursos
    
    Dim men As Integer
    Dim texto As String
    
    oCurso.Carga Curso
    
    With oMensaje
        
        texto = "Se requiere respuesta a una invitación de asistencia al curso de formación:" + vbCrLf + vbCrLf
        texto = texto + oCurso.getDESCRIPCION + vbCrLf + vbCrLf
        texto = texto + "Contenido:" + vbCrLf + vbCrLf
        texto = texto + oCurso.getCONTENIDO
        
        .setTEXTO = Trim(texto)
        .setASUNTO = "Invitación a curso de formación"
        .setEMPLEADO_ID = USUARIO.getID_EMPLEADO
        .setFECHA_INICIO = Format(Date, "yyyy-mm-dd")
        .setHORA_INICIO = "00:00:00"
        .setFECHA_FIN = Format(oCurso.getFECHA_PREVISTA_F, "yyyy-mm-dd")
        .setHORA_FIN = "00:00:00"
        men = .Insertar
        
    End With
    
    Set oCurso = Nothing
    Set oMensaje = Nothing
    crea_mensaje = men
    
    Exit Function

crea_mensaje_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure crea_mensaje of Formulario frmFormacion_Curso"

   
End Function

Private Sub envia_mensaje_usuario(men As Integer, ID_EMPLEADO As Long)
'ASocia un determinado mensaje a un usuario concreto

    On Error GoTo envia_mensajes_Error

    Dim omu As New clsMensajes_usuarios
    Dim oempleado As New clsEmpleados
    
    oempleado.CARGAR ID_EMPLEADO
    
    If oempleado.getUSUARIO_ID <> 0 Then
        omu.setEMPLEADO_ID = oempleado.getUSUARIO_ID
        omu.setMENSAJE_ID = men
        omu.Insertar
    End If

    Set oempleado = Nothing
    Set omu = Nothing
    Exit Sub

envia_mensajes_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure envia_mensaje_usuario of Formulario frmFormacion_Curso"
End Sub

'M1268-I
Public Function duplicarCurso(ID As Long) As Long
On Error GoTo ERROR
'Nuevo curso en estado abierto con mismos responsables, formadores y asistentes
    Dim oFormadores As New clsFormacion_Formadores
    Dim oAsistentes As New clsFormacion_asistentes
    Dim NuevoId As Long
'1 - Curso
    Carga ID
    ANYO = Year(Date)
    REALIZADO = 0
    APROBADO = 1
    PARADO = 0
    FECHA_PREVISTA_I = Format(FECHA_PREVISTA_I, "yyyy-mm-dd")
    FECHA_PREVISTA_F = Format(FECHA_PREVISTA_F, "yyyy-mm-dd")
    FECHA_REAL_I = Format(FECHA_REAL_I, "yyyy-mm-dd")
    FECHA_REAL_F = Format(FECHA_REAL_F, "yyyy-mm-dd")
    NuevoId = Insertar
    If NuevoId <> 0 Then
        Me.generar_firmas_calidad NuevoId
'2 - Formadores
        With oFormadores
          oFormadores.duplicarFormadores ID, NuevoId
        End With
'3 - Asistentes
        With oAsistentes
          oAsistentes.duplicarAsistentes ID, NuevoId
        End With
    End If
    duplicarCurso = NuevoId
    Exit Function
ERROR:
    duplicarCurso = 0
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure duplicarCurso of clsfrmFormacion_Cursos"
End Function
'M1268-F



