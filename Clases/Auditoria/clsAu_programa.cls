VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAu_programa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSAU_PROGRAMA
'***************************************************************
Private ID_PROGRAMA As Integer
Private VERSION As Integer
Private DESCRIPCION As String
Private ANNO As Integer
Private APROBADO As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSAU_PROGRAMA
'***************************************************************
Public Property Let setID_PROGRAMA(ByVal dato As Integer)
        ID_PROGRAMA = dato
End Property
Public Property Let setVERSION(ByVal dato As Integer)
        VERSION = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setANNO(ByVal dato As Integer)
        ANNO = dato
End Property
Public Property Let setAPROBADO(ByVal dato As Integer)
        APROBADO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSAU_PROGRAMA
'***************************************************************
Public Property Get getID_PROGRAMA() As Integer
        getID_PROGRAMA = ID_PROGRAMA
End Property
Public Property Get getVERSION() As Integer
        getVERSION = VERSION
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getANNO() As Integer
        getANNO = ANNO
End Property
Public Property Get getAPROBADO() As Integer
        getAPROBADO = APROBADO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSAU_PROGRAMA
'***************************************************************
Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM AU_PROGRAMA WHERE ID_PROGRAMA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PROGRAMA = rs("ID_PROGRAMA")
                VERSION = rs("VERSION")
                DESCRIPCION = rs("DESCRIPCION")
                ANNO = rs("ANNO")
                APROBADO = rs("APROBADO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAu_programa)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID_PROGRAMA) FROM AU_PROGRAMA"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PROGRAMA = 1
        Else
                ID_PROGRAMA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Integer
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO AU_PROGRAMA " & _
                   "  VALUES (" & _
                        ID_PROGRAMA & "," & VERSION & "," & "'" & DESCRIPCION & "'" & "," & _
                        ANNO & "," & APROBADO & _
                ")"
        execute_bd consulta
        Insertar = ID_PROGRAMA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAu_programa)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE AU_PROGRAMA SET " & _
                        " VERSION = " & VERSION & "," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " ANNO = " & ANNO & "," & _
                        " APROBADO = " & APROBADO & "" & _
                " WHERE ID_PROGRAMA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAu_programa)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM AU_PROGRAMA " & _
                "    WHERE ID_PROGRAMA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAu_programa)", vbCritical, Err.Description
End Function
Public Function Listado(fDESCRIPCION As String, fAREA As Integer) As ADODB.RecordSet
        Dim consulta As String
        Dim s As String
        If fAREA <> 0 Then
            consulta = "SELECT A.ID_PROGRAMA,A.DESCRIPCION,A.ANNO,A.VERSION,A.APROBADO " & _
                       "  FROM AU_PROGRAMA A, AU_PROGRAMA_AREAS B " & _
                       " WHERE A.DESCRIPCION LIKE '%" & fDESCRIPCION & "%'" & _
                       "   AND A.ID_PROGRAMA = B.PROGRAMA_ID " & _
                       "   AND B.AREA_ID = " & fAREA & _
                       " ORDER BY A.ID_PROGRAMA DESC"
        Else
            consulta = "SELECT ID_PROGRAMA,DESCRIPCION,ANNO,VERSION,APROBADO FROM AU_PROGRAMA " & _
                       " WHERE DESCRIPCION LIKE '%" & fDESCRIPCION & "%'" & _
                       " ORDER BY ID_PROGRAMA DESC"
        End If
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID_PROGRAMA,VERSION FROM AU_PROGRAMA ORDER BY VERSION"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function Enviar_Al_Aprobador(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE AU_PROGRAMA SET " & _
                        " APROBADO = 2 " & _
                " WHERE ID_PROGRAMA = " & ID
        execute_bd consulta
        Enviar_Al_Aprobador = True
        Exit Function
fallo:
        Enviar_Al_Aprobador = False
        MsgBox "Error al Enviar_Al_Aprobador (clsAu_programa)", vbCritical, Err.Description
End Function
Public Function Aprobar(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE AU_PROGRAMA SET " & _
                        " APROBADO = 1 " & _
                " WHERE ID_PROGRAMA = " & ID
        execute_bd consulta
        Aprobar = True
        Exit Function
fallo:
        Aprobar = False
        MsgBox "Error al Aprobar (clsAu_programa)", vbCritical, Err.Description
End Function


