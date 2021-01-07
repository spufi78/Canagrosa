VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_dt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_DT
'***************************************************************
Private EQUIPO_ID As Long
Private PRECISIONX As String
Private CAPACIDAD As String
Private TOLERANCIA As String
Private UNIDAD_ID As Long
Private CAMBIO_ESTADO As Long
Private RESOLUCION As String
Private CTE As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_DT
'***************************************************************
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setPRECISIONX(ByVal dato As String)
        PRECISIONX = dato
End Property
Public Property Let setCAPACIDAD(ByVal dato As String)
        CAPACIDAD = dato
End Property
Public Property Let setTOLERANCIA(ByVal dato As String)
        TOLERANCIA = dato
End Property
Public Property Let setUNIDAD_ID(ByVal dato As Long)
        UNIDAD_ID = dato
End Property
Public Property Let setCAMBIO_ESTADO(ByVal dato As Long)
        CAMBIO_ESTADO = dato
End Property
Public Property Let setRESOLUCION(ByVal dato As String)
        RESOLUCION = dato
End Property
Public Property Let setCTE(ByVal dato As String)
        CTE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_DT
'***************************************************************
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getPRECISIONX() As String
        getPRECISIONX = PRECISIONX
End Property
Public Property Get getCAPACIDAD() As String
        getCAPACIDAD = CAPACIDAD
End Property
Public Property Get getTOLERANCIA() As String
        getTOLERANCIA = TOLERANCIA
End Property
Public Property Get getUNIDAD_ID() As Long
        getUNIDAD_ID = UNIDAD_ID
End Property
Public Property Get getCAMBIO_ESTADO() As Long
        getCAMBIO_ESTADO = CAMBIO_ESTADO
End Property
Public Property Get getRESOLUCION() As String
        getRESOLUCION = RESOLUCION
End Property
Public Property Get getCTE() As String
        getCTE = CTE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_DT
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM equipos_dt WHERE EQUIPO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                EQUIPO_ID = rs("EQUIPO_ID")
                PRECISIONX = rs("PRECISIONX")
                CAPACIDAD = rs("CAPACIDAD")
                TOLERANCIA = rs("TOLERANCIA")
                UNIDAD_ID = rs("UNIDAD_ID")
                CAMBIO_ESTADO = rs("CAMBIO_ESTADO")
                RESOLUCION = rs("RESOLUCION")
                CTE = rs("CTE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_dt)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO equipos_dt " & _
                   " (EQUIPO_ID,PRECISIONX,CAPACIDAD,TOLERANCIA,UNIDAD_ID,INSITU,RESOLUCION,CTE) VALUES (" & _
                        EQUIPO_ID & "," & "'" & PRECISIONX & "'" & "," & "'" & CAPACIDAD & "'" & "," & _
                        "'" & TOLERANCIA & "'" & "," & UNIDAD_ID & "," & CAMBIO_ESTADO & "," & _
                        "'" & RESOLUCION & "','" & CTE & "'" & _
                ")"
        consulta = consulta & " ON DUPLICATE KEY UPDATE " & _
                        " PRECISIONX = '" & PRECISIONX & "'," & _
                        " CAPACIDAD = '" & CAPACIDAD & "'," & _
                        " TOLERANCIA = '" & TOLERANCIA & "'," & _
                        " UNIDAD_ID = " & UNIDAD_ID & "," & _
                        " RESOLUCION = '" & RESOLUCION & "'," & _
                        " CTE = '" & CTE & "'"
        execute_bd consulta
        Insertar = EQUIPO_ID
        Exit Function
fallo:
        Insertar = 0
        frmLogin.txtlog = frmLogin.txtlog & consulta & vbNewLine
        MsgBox "Error al insertar (clsEquipos_dt)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE equipos_dt SET " & _
                        " PRECISIONX = '" & PRECISIONX & "'," & _
                        " CAPACIDAD = '" & CAPACIDAD & "'," & _
                        " TOLERANCIA = '" & TOLERANCIA & "'," & _
                        " UNIDAD_ID = " & UNIDAD_ID & "," & _
                        " INSITU = " & CAMBIO_ESTADO & "," & _
                        " RESOLUCION = '" & RESOLUCION & "'" & _
                " WHERE EQUIPO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipos_dt)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM equipos_dt " & _
                "    WHERE EQUIPO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos_dt)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM equipos_dt ORDER BY EQUIPO_ID"
        Set Listado = datos_bd(consulta)
End Function
