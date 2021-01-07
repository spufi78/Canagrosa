VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_ponentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_PONENTES
'***************************************************************
Private CURSO_ID As Long
Private FORMADOR_ID As Long
Private TIPO_FORMADOR_ID As Long
Private HORARIO_INICIO As String
Private HORARIO_FIN As String
Private FTIMESTP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_PONENTES
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
Public Property Let setHORARIO_INICIO(ByVal dato As String)
        HORARIO_INICIO = dato
End Property
Public Property Let setHORARIO_FIN(ByVal dato As String)
        HORARIO_FIN = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_PONENTES
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
Public Property Get getHORARIO_INICIO() As String
        getHORARIO_INICIO = HORARIO_INICIO
End Property
Public Property Get getHORARIO_FIN() As String
        getHORARIO_FIN = HORARIO_FIN
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_PONENTES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PONENTES WHERE CURSO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CURSO_ID = rs("CURSO_ID")
                FORMADOR_ID = rs("FORMADOR_ID")
                TIPO_FORMADOR_ID = rs("TIPO_FORMADOR_ID")
                HORARIO_INICIO = rs("HORARIO_INICIO")
                HORARIO_FIN = rs("HORARIO_FIN")
                FTIMESTP = rs("FTIMESTP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFormacion_ponentes)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO FORMACION_PONENTES " & _
                   "  VALUES (" & _
                        CURSO_ID & "," & FORMADOR_ID & "," & TIPO_FORMADOR_ID & "," & _
                        "CURRENT_TIMESTAMP" & _
                ")"
        execute_bd consulta
        Insertar = CURSO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFormacion_ponentes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FORMACION_PONENTES SET " & _
                        " FORMADOR_ID = " & FORMADOR_ID & "," & _
                        " TIPO_FORMADOR_ID = " & TIPO_FORMADOR_ID & "," & _
                        " HORARIO_INICIO = '" & HORARIO_INICIO & "'," & _
                        " HORARIO_FIN = '" & HORARIO_FIN & "'," & _
                        " FTIMESTP = '" & FTIMESTP & "'" & _
                " WHERE CURSO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFormacion_ponentes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_PONENTES " & _
                "    WHERE CURSO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFormacion_ponentes)", vbCritical, Err.Description
End Function
Public Function Listado(PK As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_PONENTES WHERE CURSO_ID=" & PK
        Set Listado = datos_bd(consulta)
End Function

Public Function Listado_internos(PK As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT E.ID_EMPLEADO AS ID_EMPLEADO, E.NOMBRE AS NOMBRE, F.TIPO_FORMADOR_ID AS TIPO_FORMADOR_ID FROM FORMACION_PONENTES F, EMPLEADOS E WHERE F.FORMADOR_ID = E.ID_EMPLEADO AND F.CURSO_ID = " & PK
        Set Listado = datos_bd(consulta)
End Function

Public Function Listado_externos() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT P.ID_PROVEEDOR AS ID_EMPLEADO, P.NOMBRE AS NOMBRE, F.TIPO_FORMADOR_ID AS TIPO_FORMADOR_ID FROM FORMACION_PONENTES F, PROVEEDORES P WHERE F.FORMADOR_ID = P.ID_PROVEEDOR AND F.CURSO_ID = " & PK
        Set Listado = datos_bd(consulta)
End Function
