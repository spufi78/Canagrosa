VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_imagenes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_IMAGENES
'***************************************************************
Private MUESTRA_ID As Long
Private ORDEN As Long
Private LEYENDA As String
Private RUTA As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_IMAGENES
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setLEYENDA(ByVal dato As String)
        LEYENDA = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        RUTA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_IMAGENES
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getLEYENDA() As String
        getLEYENDA = LEYENDA
End Property
Public Property Get getRUTA() As String
        getRUTA = RUTA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_IMAGENES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_IMAGENES WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                ORDEN = rs("ORDEN")
                LEYENDA = rs("LEYENDA")
                RUTA = rs("RUTA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_imagenes)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CE_IMAGENES " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & ORDEN & "," & "'" & LEYENDA & "'" & "," & _
                        "'" & RUTA & "'" & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_imagenes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_IMAGENES SET " & _
                        " ORDEN = " & ORDEN & "," & _
                        " LEYENDA = '" & LEYENDA & "'," & _
                        " RUTA = '" & RUTA & "'" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_imagenes)", vbCritical, Err.Description
End Function
Public Function Eliminar(MUESTRA As Long, ORDEN As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_IMAGENES " & _
                "    WHERE MUESTRA_ID = " & MUESTRA & " AND ORDEN = " & ORDEN
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_imagenes)", vbCritical, Err.Description
End Function
Public Function Listado(MUESTRA As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_IMAGENES WHERE MUESTRA_ID = " & MUESTRA & " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function
