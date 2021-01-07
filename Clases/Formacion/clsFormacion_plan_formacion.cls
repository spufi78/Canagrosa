VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_pf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_PLAN_FORMACION
'***************************************************************
Private ID_PLAN_FORMACION As Long
Private CURSO_ID As Long
Private DESCRIPCION As String
Private fecha_prevista As String
Private MODALIDAD As Long
Private Formacion As Long
Private OBSERVACIONES As String
Private anyo As Long
Private FTIMESTP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_PLAN_FORMACION
'***************************************************************
Public Property Let setID_PLAN_FORMACION(ByVal dato As Long)
        ID_PLAN_FORMACION = dato
End Property
Public Property Let setCURSO_ID(ByVal dato As Long)
        CURSO_ID = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFECHA_PREVISTA(ByVal dato As String)
        fecha_prevista = dato
End Property
Public Property Let setMODALIDAD(ByVal dato As Long)
        MODALIDAD = dato
End Property
Public Property Let setFORMACION(ByVal dato As Long)
        Formacion = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setANYO(ByVal dato As Long)
        anyo = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_PLAN_FORMACION
'***************************************************************
Public Property Get getID_PLAN_FORMACION() As Long
        getID_PLAN_FORMACION = ID_PLAN_FORMACION
End Property
Public Property Get getCURSO_ID() As Long
        getCURSO_ID = CURSO_ID
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFECHA_PREVISTA() As String
        getFECHA_PREVISTA = fecha_prevista
End Property
Public Property Get getMODALIDAD() As Long
        getMODALIDAD = MODALIDAD
End Property
Public Property Get getFORMACION() As Long
        getFORMACION = Formacion
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getANYO() As String
        getANYO = anyo
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_PLAN_FORMACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PLAN_FORMACION WHERE ID_PLAN_FORMACION = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PLAN_FORMACION = rs("ID_PLAN_FORMACION")
                CURSO_ID = rs("CURSO_ID")
                DESCRIPCION = rs("DESCRIPCION")
                fecha_prevista = rs("FECHA_PREVISTA")
                MODALIDAD = rs("MODALIDAD")
                Formacion = rs("FORMACION")
                OBSERVACIONES = rs("OBSERVACIONES")
                anyo = rs("ANYO")
                 
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFormacion_plan_formacion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO FORMACION_PLAN_FORMACION " & _
                   "  VALUES (" & _
                        ID_PLAN_FORMACION & "," & CURSO_ID & ",'" & DESCRIPCION & "'" & ",'" & fecha_prevista & "'," & _
                        MODALIDAD & "," & Formacion & "," & "'" & OBSERVACIONES & "'," & anyo & ",CURRENT_TIMESTAMP" & _
                ")"
        execute_bd consulta
        Insertar = ID_PLAN_FORMACION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFormacion_plan_formacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FORMACION_PLAN_FORMACION SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FECHA_PREVISTA = '" & fecha_prevista & "'," & _
                        " MODALIDAD = " & MODALIDAD & "," & _
                        " FORMACION = " & Formacion & "," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " ANYO = " & anyo & "," & _
                        " FTIMESTP = CURRENT_TIMESTAMP" & _
                " WHERE ID_PLAN_FORMACION = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFormacion_plan_formacion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_PLAN_FORMACION " & _
                "    WHERE ID_PLAN_FORMACION = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFormacion_plan_formacion)", vbCritical, Err.Description
End Function

Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PLAN_FORMACION ORDER BY ID_PLAN_FORMACION"
        Set Listado = datos_bd(consulta)
End Function

Public Function ListadoFiltro(cod As String, DESCRIPCION As String, MODALIDAD As Integer, nivel As Integer) As ADODB.Recordset
        Dim consulta As String
        Dim strModalidad As String
        Dim strNivel As String
        
        If cod <> "" And IsNumeric(cod) Then
         consulta = "SELECT * FROM FORMACION_PLAN_FORMACION WHERE ID_PLAN_FORMACION = " & CLng(cod)
        Else
             If MODALIDAD = 2 Then
                strModalidad = ""
             Else
                strModalidad = " AND MODALIDAD =" & MODALIDAD
             End If
             If nivel = 3 Then
                strNivel = ""
             Else
                strNivel = " AND MODALIDAD =" & nivel
             End If
             consulta = "SELECT * FROM FORMACION_PLAN_FORMACION " & _
                        " WHERE DESCRIPCION LIKE '%" & DESCRIPCION & "%'" & _
                         strModalidad & _
                         strNivel
                        
        End If
        Set ListadoFiltro = datos_bd(consulta)
End Function
Public Function ListadoDisponibles() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PLAN_FORMACION WHERE CURSO_ID=0 ORDER BY ID_PLAN_FORMACION"
        Set ListadoDisponibles = datos_bd(consulta)
End Function
Public Sub CrearIdPlan()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_PLAN_FORMACION) FROM formacion_plan_formacion"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_PLAN_FORMACION = 1
    Else
        ID_PLAN_FORMACION = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub

Public Function Actualizar_Curso(ID As Long, CURSO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FORMACION_PLAN_FORMACION SET " & _
                        " CURSO_ID = " & CURSO & _
                " WHERE ID_PLAN_FORMACION = " & ID
        execute_bd consulta
        Actualizar_Curso = True
        Exit Function
fallo:
        Actualizar_Curso = False
        MsgBox "Error al Modificar el Código del Curso (clsFormacion_plan_formacion)", vbCritical, Err.Description
End Function
