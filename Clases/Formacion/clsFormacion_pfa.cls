VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_pfa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_PFA
'***************************************************************
Private ID_PFA As Long
Private PLAN_FORMACION_ID As Long
Private CURSO_ID As Long
Private DESCRIPCION As String
Private FECHA_PREVISTA As String
Private OBSERVACIONES As String
Private FTIMESTP As String
Private ANYO As Long
Private MODALIDAD As Long
Private NIVEL As Long
Private FORMACION As Long

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_PFA
'***************************************************************
Public Property Let setID_PFA(ByVal dato As Long)
        ID_PFA = dato
End Property
Public Property Let setPLAN_FORMACION_ID(ByVal dato As Long)
        PLAN_FORMACION_ID = dato
End Property
Public Property Let setCURSO_ID(ByVal dato As Long)
        CURSO_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFECHA_PREVISTA(ByVal dato As String)
        FECHA_PREVISTA = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
Public Property Let setANYO(ByVal dato As Long)
        ANYO = dato
End Property
Public Property Let setMODALIDAD(ByVal dato As Long)
        MODALIDAD = dato
End Property
Public Property Let setNIVEL(ByVal dato As Long)
        NIVEL = dato
End Property
Public Property Let setFORMACION(ByVal dato As Long)
        FORMACION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_PFA
'***************************************************************
Public Property Get getID_PFA() As Long
        getID_PFA = ID_PFA
End Property
Public Property Get getPLAN_FORMACION_ID() As Long
        getPLAN_FORMACION_ID = PLAN_FORMACION_ID
End Property
Public Property Get getCURSO_ID() As Long
        getCURSO_ID = CURSO_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFECHA_PREVISTA() As String
        getFECHA_PREVISTA = FECHA_PREVISTA
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
Public Property Get getANYO() As Long
        getANYO = ANYO
End Property
Public Property Get getMODALIDAD() As Long
        getMODALIDAD = MODALIDAD
End Property
Public Property Get getNIVEL() As Long
        getNIVEL = NIVEL
End Property
Public Property Get getFORMACION() As Long
        getFORMACION = FORMACION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_PFA
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PFA WHERE ID_PFA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PFA = rs("ID_PFA")
                PLAN_FORMACION_ID = rs("PLAN_FORMACION_ID")
                CURSO_ID = rs("CURSO_ID")
                DESCRIPCION = rs("DESCRIPCION")
                FECHA_PREVISTA = rs("FECHA_PREVISTA")
                OBSERVACIONES = rs("OBSERVACIONES")
                FTIMESTP = rs("FTIMESTP")
                ANYO = rs("ANYO")
                MODALIDAD = rs("MODALIDAD")
                NIVEL = rs("NIVEL")
                FORMACION = rs("FORMACION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFormacion_pfa)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO FORMACION_PFA " & _
                   "  VALUES (" & _
                        ID_PFA & "," & PLAN_FORMACION_ID & "," & CURSO_ID & ",'" & DESCRIPCION & "'," & "'" & FECHA_PREVISTA & "'" & "," & _
                        "'" & OBSERVACIONES & "'" & "," & "CURRENT_TIMESTAMP," & ANYO & "," & MODALIDAD & "," & NIVEL & "," & FORMACION & _
                ")"
        execute_bd consulta
        Insertar = ID_PFA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFormacion_pfa)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FORMACION_PFA SET " & _
                        " CURSO_ID = " & CURSO_ID & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FECHA_PREVISTA = '" & FECHA_PREVISTA & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " FTIMESTP = CURRENT_TIMESTAMP," & _
                        " ANYO = " & ANYO & "," & _
                        " MODALIDAD = " & MODALIDAD & "," & _
                        " NIVEL = " & NIVEL & "," & _
                        " FORMACION = " & FORMACION & _
                " WHERE ID_PFA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFormacion_pfa)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_PFA " & _
                "    WHERE ID_PFA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFormacion_pfa)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PFA ORDER BY ID_PFA"
        Set Listado = datos_bd(consulta)
End Function

Public Function ListadoFiltro(ANYO As Long, cod As String, DESCRIPCION As String, MODALIDAD As Integer, NIVEL As Integer, TIPO As Integer) As ADODB.Recordset
        Dim consulta As String
        Dim strModalidad As String
        Dim strNivel As String
        Dim strTipo As String
        Dim strDescripcion As String
        
        consulta = "SELECT A.ID_PFA AS ID_PFA, A.ANYO AS ANYO, A.CURSO_ID AS CURSO_ID, A.FECHA_PREVISTA AS FECHA_PREVISTA, A.DESCRIPCION AS DESCRIPCION, A.MODALIDAD AS MODALIDAD, A.NIVEL AS NIVEL,A.FORMACION AS FORMACION FROM FORMACION_PFA A, FORMACION_PF F" & _
                " WHERE A.PLAN_FORMACION_ID = F.ID_PLAN_FORMACION" & _
                " AND A.ANYO=" & ANYO
                
        If cod <> "" And IsNumeric(cod) Then
             strModalidad = ""
             strNivel = ""
        Else
             strDescripcion = " AND A.DESCRIPCION LIKE '%" & DESCRIPCION & "%'"
             If MODALIDAD = 2 Then
                strModalidad = ""
             Else
                strModalidad = " AND A.MODALIDAD =" & MODALIDAD
             End If
             If NIVEL = 3 Then
                strNivel = ""
             Else
                strNivel = " AND A.FORMACION =" & NIVEL
             End If
                        
             If TIPO = 2 Then
                strTipo = ""
             Else
                strTipo = " AND A.FORMACION=" & TIPO
             End If
        End If
        
        consulta = consulta & strModalidad & strNivel & strTipo & strDescripcion
        Set ListadoFiltro = datos_bd(consulta)
End Function

Public Sub CrearIdPlan()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_PFA) FROM FORMACION_PFA"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_PFA = 1
    Else
        ID_PFA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub

Public Function Actualizar_Curso(ID As Long, Curso As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE  FORMACION_PFA SET " & _
                        " CURSO_ID = " & Curso & _
                " WHERE ID_PFA = " & ID
        execute_bd consulta
        Actualizar_Curso = True
        Exit Function
fallo:
        Actualizar_Curso = False
        MsgBox "Error al Modificar el Código del Curso (cls)", vbCritical, Err.Description
End Function

Public Function ListadoDisponibles(PLAN As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PFA   WHERE CURSO_ID=0 OR ID_PFA = " & PLAN & " ORDER BY ID_PFA"
        Set ListadoDisponibles = datos_bd(consulta)
End Function

