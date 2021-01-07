VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_asistentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_ASISTENTES
'***************************************************************

Private CURSO_ID As Long
Private EMPLEADO_ID As Long

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_ASISTENTES
'***************************************************************

Public Property Let setCURSO_ID(ByVal dato As Long)
        CURSO_ID = dato
End Property

Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_ASISTENTES
'***************************************************************

Public Property Get getCURSO_ID() As Long
        getCURSO_ID = CURSO_ID
End Property

Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_ASISTENTES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_ASISTENTES WHERE CURSO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CURSO_ID = rs("CURSO_ID")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFormacion_asistentes)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO FORMACION_ASISTENTES " & _
                   "  VALUES (" & _
                        CURSO_ID & "," & EMPLEADO_ID & _
                ")"
        execute_bd consulta
        Insertar = CURSO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFormacion_asistentes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FORMACION_ASISTENTES SET " & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & _
                " WHERE CURSO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFormacion_asistentes)", vbCritical, Err.Description
End Function

Public Function Eliminar(CURSO_ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_ASISTENTES " & _
                "    WHERE CURSO_ID = " & CURSO_ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFormacion_asistentes)", vbCritical, Err.Description
End Function

Public Function Listado(PK_CURSO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT E.ID_EMPLEADO, E.NOMBRE,FASISTE.FIRMADA,DATE(FASISTE.FTIMESTP),FEVAL.FIRMADA,DATE(FEVAL.FTIMESTP) " & _
                   "  FROM FORMACION_ASISTENTES A " & _
                   "  INNER JOIN  empleados E ON A.EMPLEADO_ID = E.ID_EMPLEADO " & _
                   "   LEFT JOIN firmas FASISTE ON FASISTE.TOBJETO = " & TOBJETO.TOBJETO_INVITACION_CURSO & " AND FASISTE.COBJETO=" & PK_CURSO & " AND FASISTE.USUARIO_ID=A.EMPLEADO_ID " & _
                   "   LEFT JOIN firmas FEVAL ON FEVAL.TOBJETO = " & TOBJETO.TOBJETO_ASISTENCIA_CURSO_ASISTENTES & " AND FEVAL.COBJETO=" & PK_CURSO & " AND FEVAL.USUARIO_ID=A.EMPLEADO_ID " & _
                   "  WHERE A.CURSO_ID = " & PK_CURSO
        Set Listado = datos_bd(consulta)
End Function
'M1110-I
Public Function ListadoFirmantes(PK_CURSO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.* " & _
                   "  FROM FORMACION_ASISTENTES A " & _
                   "  INNER JOIN  empleados E ON A.EMPLEADO_ID = E.ID_EMPLEADO " & _
                   "   LEFT JOIN firmas FEVAL ON FEVAL.TOBJETO = " & TOBJETO.TOBJETO_ASISTENCIA_CURSO_ASISTENTES & " AND FEVAL.COBJETO=" & PK_CURSO & " AND FEVAL.USUARIO_ID=A.EMPLEADO_ID AND FEVAL.FIRMADA=1" & _
                   "  WHERE A.CURSO_ID = " & PK_CURSO
        Set ListadoFirmantes = datos_bd(consulta)
End Function
'M1110-F

Public Function EliminarAsistente(CURSO_ID As Long, EMPLEADO_ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_ASISTENTES " & _
                "    WHERE CURSO_ID = " & CURSO_ID & _
                "    AND EMPLEADO_ID = " & EMPLEADO_ID
        execute_bd consulta
        EliminarAsistente = True
        Exit Function
fallo:
        EliminarAsistente = False
        MsgBox "Error al Eliminar (clsFormacion_asistentes)", vbCritical, Err.Description
End Function
'M1268-I
Public Function duplicarAsistentes(ID As Long, NuevoId As Long) As Boolean
On Error GoTo fallo
     Dim rs As ADODB.Recordset
     Set rs = asistentesID(ID)
     If rs.RecordCount > 0 Then
        Do
            CURSO_ID = NuevoId
            EMPLEADO_ID = rs("EMPLEADO_ID")
            Insertar
            rs.MoveNext
        Loop Until rs.EOF
     End If
     Set rs = Nothing
     duplicarAsistentes = True
     Exit Function
fallo:
     duplicarAsistentes = False
     MsgBox "Error al duplicarAsistentes (clsFormacion_Asistentes)", vbCritical, Err.Description
End Function
Private Function asistentesID(PK As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM FORMACION_ASISTENTES" & _
                   "  WHERE CURSO_ID = " & PK
        Set asistentesID = datos_bd(consulta)
End Function
'M1268-F
