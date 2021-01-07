VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpleados_categorias_historia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEMPLEADOS_CATEGORIAS_HISTORIA
'***************************************************************
Private EMPLEADO_ID As Long
Private CATEGORIA_ID As Long
Private fecha_inicio As String
Private fecha_fin As String
Private ORDEN As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEMPLEADOS_CATEGORIAS_HISTORIA
'***************************************************************
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setCATEGORIA_ID(ByVal dato As Long)
        CATEGORIA_ID = dato
End Property
Public Property Let setFECHA_INICIO(ByVal dato As String)
        fecha_inicio = dato
End Property
Public Property Let setFECHA_FIN(ByVal dato As String)
        fecha_fin = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEMPLEADOS_CATEGORIAS_HISTORIA
'***************************************************************
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getCATEGORIA_ID() As Long
        getCATEGORIA_ID = CATEGORIA_ID
End Property
Public Property Get getFECHA_INICIO() As String
        getFECHA_INICIO = fecha_inicio
End Property
Public Property Get getFECHA_FIN() As String
        getFECHA_FIN = fecha_fin
End Property
Public Property Get getORDEN() As Integer
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEMPLEADOS_CATEGORIAS_HISTORIA
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EMPLEADOS_CATEGORIAS_HISTORIA WHERE EMPLEADO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                EMPLEADO_ID = rs("EMPLEADO_ID")
                CATEGORIA_ID = rs("CATEGORIA_ID")
                fecha_inicio = rs("FECHA_INICIO")
                fecha_fin = rs("FECHA_FIN")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEmpleados_categorias_historia)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO EMPLEADOS_CATEGORIAS_HISTORIA " & _
                   "  VALUES (" & _
                        EMPLEADO_ID & "," & CATEGORIA_ID & ",'" & fecha_inicio & "','" & _
                        fecha_fin & "'," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = EMPLEADO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEmpleados_categorias_historia)", vbCritical, Err.Description
End Function
Public Function Modificar(EMPLEADO As Long, ORDEN As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EMPLEADOS_CATEGORIAS_HISTORIA SET " & _
                        " CATEGORIA_ID = " & CATEGORIA_ID & "," & _
                        " FECHA_INICIO = '" & fecha_inicio & "'," & _
                        " FECHA_FIN = '" & fecha_fin & "'" & _
                " WHERE EMPLEADO_ID = " & EMPLEADO & " AND ORDEN = " & ORDEN
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEmpleados_categorias_historia)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EMPLEADOS_CATEGORIAS_HISTORIA " & _
                "    WHERE EMPLEADO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEmpleados_categorias_historia)", vbCritical, Err.Description
End Function
Public Function EliminarCategoria(ID As Long, ORDEN As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EMPLEADOS_CATEGORIAS_HISTORIA " & _
                "    WHERE EMPLEADO_ID = " & ID & " AND ORDEN = " & ORDEN
        execute_bd consulta
        EliminarCategoria = True
        Exit Function
fallo:
        EliminarCategoria = False
        MsgBox "Error al EliminarCategoria (clsEmpleados_categorias_historia)", vbCritical, Err.Description
End Function
Public Function Listado(EMPLEADO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT CONCAT(EC.CODIGO,'-',EC.DESCRIPCION), D.DESCRIPCION,ECH.FECHA_INICIO,ECH.FECHA_FIN, ECH.CATEGORIA_ID, ECH.ORDEN " & _
                   "  FROM EMPLEADOS_CATEGORIAS EC,EMPLEADOS_CATEGORIAS_HISTORIA ECH, decodificadora D" & _
                   " WHERE ECH.EMPLEADO_ID = " & EMPLEADO & _
                   "   AND ECH.CATEGORIA_ID = EC.ID_CATEGORIA " & _
                   "   AND D.CODIGO = " & DECODIFICADORA.EMPLEADOS_DEPARTAMENTOS & _
                   "   AND EC.DEPARTAMENTO_ID = D.VALOR " & _
                   " ORDER BY ECH.FECHA_INICIO"
        Set Listado = datos_bd(consulta)
End Function

Public Function Listado_Categoria(categoria As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT CONCAT(EC.CODIGO,'-',EC.DESCRIPCION), D.DESCRIPCION,ECH.FECHA_INICIO,ECH.FECHA_FIN, ECH.CATEGORIA_ID, ECH.ORDEN, ECH.EMPLEADO_ID " & _
                   "  FROM EMPLEADOS_CATEGORIAS EC,EMPLEADOS_CATEGORIAS_HISTORIA ECH, empleados E,decodificadora D" & _
                   " WHERE ECH.CATEGORIA_ID = EC.ID_CATEGORIA " & _
                   "   AND ECH.EMPLEADO_ID = E.ID_EMPLEADO " & _
                   "   AND D.CODIGO = " & DECODIFICADORA.EMPLEADOS_DEPARTAMENTOS & _
                   "   AND EC.DEPARTAMENTO_ID = D.VALOR " & _
                   "   AND ECH.CATEGORIA_ID = " & categoria & _
                   "   AND E.ESTADO_ID = 0" & _
                   " ORDER BY ECH.FECHA_INICIO"
        Set Listado_Categoria = datos_bd(consulta)
End Function

