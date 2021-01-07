VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_Formadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_Formadores
'***************************************************************
Private CURSO_ID As Long
Private FORMADOR_ID As Long
Private TIPO_FORMADOR_ID As Long


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_Formadores
'***************************************************************

Public Property Let setCURSO_ID(ByVal dato As Long)
        CURSO_ID = dato
End Property
Public Property Let setFORMADOR_ID(ByVal dato As Long)
        FORMADOR_ID = dato
End Property
Public Property Let setTIPO_FORMADOR_ID(ByVal dato As Long)
        TIPO_FORMADOR_ID = dato
End Property


'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_Formadores
'***************************************************************

Public Property Get getCURSO_ID() As Long
        getCURSO_ID = CURSO_ID
End Property
Public Property Get getFORMADOR_ID() As Long
        getFORMADOR_ID = FORMADOR_ID
End Property

Public Property Get getTIPO_FORMADOR_ID() As Long
        getTIPO_FORMADOR_ID = TIPO_FORMADOR_ID
End Property


'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_Formadores
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_Formadores WHERE CURSO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CURSO_ID = rs("CURSO_ID")
                FORMADOR_ID = rs("FORMADOR_ID")
                TIPO_FORMADOR_ID = rs("TIPO_FORMADOR_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFormacion_Formadores)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO FORMACION_Formadores " & _
                   "  VALUES (" & _
                        CURSO_ID & "," & FORMADOR_ID & "," & TIPO_FORMADOR_ID & _
                ")"
        execute_bd consulta
        Insertar = CURSO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFormacion_Formadores)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FORMACION_FORMADORES SET " & _
                        " FORMADOR_ID = " & FORMADOR_ID & "," & _
                        " TIPO_FORMADOR_ID = " & TIPO_FORMADOR_ID & _
                " WHERE CURSO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFormacion_Formadores)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_Formadores " & _
                "    WHERE CURSO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFormacion_Formadores)", vbCritical, Err.Description
End Function
Public Function Listado(PK As Long) As ADODB.Recordset
        Dim consulta As String

        'consulta = "SELECT F.FORMADOR_ID, F.TIPO_FORMADOR_ID, F.NOMBRE FROM FORMACION_Formadores F "
        'consulta = consulta + " WHERE F.CURSO_ID = " & PK
        consulta = "SELECT A.FORMADOR_ID, A.TIPO_FORMADOR_ID, E.NOMBRE,FEVAL.FIRMADA,DATE(FEVAL.FTIMESTP) " & _
                   "  FROM FORMACION_FORMADORES A " & _
                   "  INNER JOIN  empleados E ON A.FORMADOR_ID = E.ID_EMPLEADO " & _
                   "   LEFT JOIN firmas FEVAL ON FEVAL.TOBJETO = " & TOBJETO.TOBJETO_ASISTENCIA_CURSO_FORMADORES & " AND FEVAL.COBJETO=" & PK & " AND FEVAL.USUARIO_ID=A.FORMADOR_ID " & _
                   "  WHERE A.CURSO_ID = " & PK & " AND A.TIPO_FORMADOR_ID = 0"
        consulta = consulta & " UNION "
        consulta = consulta & "SELECT A.FORMADOR_ID, A.TIPO_FORMADOR_ID, P.NOMBRE, 0, '2014-01-01'" & _
                   " FROM FORMACION_FORMADORES A " & _
                   " INNER JOIN  PROVEEDORES P ON A.FORMADOR_ID = P.ID_PROVEEDOR " & _
                   " WHERE A.CURSO_ID = " & PK & " AND A.TIPO_FORMADOR_ID = 1"
          
        Set Listado = datos_bd(consulta)
End Function

Public Function Listado_internos(PK As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT E.ID_EMPLEADO AS ID_EMPLEADO, E.NOMBRE AS NOMBRE, F.TIPO_FORMADOR_ID AS TIPO_FORMADOR_ID FROM FORMACION_Formadores F, empleados E WHERE F.FORMADOR_ID = E.ID_EMPLEADO AND F.CURSO_ID = " & PK
        Set Listado_internos = datos_bd(consulta)
End Function

Public Function Listado_externos(PK As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT P.ID_PROVEEDOR AS ID_EMPLEADO, P.NOMBRE AS NOMBRE, F.TIPO_FORMADOR_ID AS TIPO_FORMADOR_ID FROM FORMACION_Formadores F, PROVEEDORES P WHERE F.FORMADOR_ID = P.ID_PROVEEDOR AND F.CURSO_ID = " & PK
        Set Listado_externos = datos_bd(consulta)
End Function

'M1110-I
Public Function Listado_Internos_Firmantes(PK As Long) As ADODB.Recordset
        Dim consulta As String

        consulta = "SELECT A.* " & _
                   "  FROM FORMACION_FORMADORES A " & _
                   "  INNER JOIN  empleados E ON A.FORMADOR_ID = E.ID_EMPLEADO " & _
                   "   LEFT JOIN firmas FEVAL ON FEVAL.TOBJETO = " & TOBJETO.TOBJETO_ASISTENCIA_CURSO_FORMADORES & " AND FEVAL.COBJETO=" & PK & " AND FEVAL.USUARIO_ID=A.FORMADOR_ID AND FEVAL.FIRMADA = 1" & _
                   "  WHERE A.CURSO_ID = " & PK
        Set Listado_Internos_Firmantes = datos_bd(consulta)
End Function
'M1110-F
Public Function EliminarFormador(ID As Long, ID_FORMADOR As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_Formadores " & _
                "    WHERE CURSO_ID = " & ID & _
                "    AND FORMADOR_ID = " & ID_FORMADOR
        execute_bd consulta
        EliminarFormador = True
        Exit Function
fallo:
        EliminarFormador = False
        MsgBox "Error al Eliminar (clsFormacion_Formadores)", vbCritical, Err.Description
End Function
'M1268-I
Public Function duplicarFormadores(ID As Long, NuevoId As Long) As Boolean
On Error GoTo fallo
     Dim rs As ADODB.Recordset
     Set rs = formadoresID(ID)
     If rs.RecordCount > 0 Then
        Do
            CURSO_ID = NuevoId
            FORMADOR_ID = rs("FORMADOR_ID")
            TIPO_FORMADOR_ID = rs("TIPO_FORMADOR_ID")
            Insertar
            rs.MoveNext
        Loop Until rs.EOF
     End If
     Set rs = Nothing
     duplicarFormadores = True
     Exit Function
fallo:
     duplicarFormadores = False
     MsgBox "Error al duplicarFormadores (clsFormacion_Formadores)", vbCritical, Err.Description
End Function
Private Function formadoresID(PK As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM FORMACION_FORMADORES" & _
                   "  WHERE CURSO_ID = " & PK
        Set formadoresID = datos_bd(consulta)
End Function
'M1268-F
