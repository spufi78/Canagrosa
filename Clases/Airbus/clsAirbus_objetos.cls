VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAirbus_objetos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsairbus_objetos
'***************************************************************
Private TOBJETO As Long
Private COBJETO As Long
Private ENSAYO_ID As Long
Private PROGRAMA_ID As Long
Private SECTION_ID As Long
Private FLUID_ID As Long
Private FACILITY_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsairbus_objetos
'***************************************************************
Public Property Let setTOBJETO(ByVal dato As Long)
        TOBJETO = dato
End Property
Public Property Let setCOBJETO(ByVal dato As Long)
        COBJETO = dato
End Property
Public Property Let setENSAYO_ID(ByVal dato As Long)
        ENSAYO_ID = dato
End Property
Public Property Let setPROGRAMA_ID(ByVal dato As Long)
        PROGRAMA_ID = dato
End Property
Public Property Let setSECTION_ID(ByVal dato As Long)
        SECTION_ID = dato
End Property
Public Property Let setFLUID_ID(ByVal dato As Long)
        FLUID_ID = dato
End Property
Public Property Let setFACILITY_ID(ByVal dato As Long)
        FACILITY_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsairbus_objetos
'***************************************************************
Public Property Get getTOBJETO() As Long
        getTOBJETO = TOBJETO
End Property
Public Property Get getCOBJETO() As Long
        getCOBJETO = COBJETO
End Property
Public Property Get getENSAYO_ID() As Long
        getENSAYO_ID = ENSAYO_ID
End Property
Public Property Get getPROGRAMA_ID() As Long
        getPROGRAMA_ID = PROGRAMA_ID
End Property
Public Property Get getSECTION_ID() As Long
        getSECTION_ID = SECTION_ID
End Property
Public Property Get getFLUID_ID() As Long
        getFLUID_ID = FLUID_ID
End Property
Public Property Get getFACILITY_ID() As Long
        getFACILITY_ID = FACILITY_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsairbus_objetos
'***************************************************************
Public Function Carga(TOBJETO As Long, COBJETO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM airbus_objetos WHERE TOBJETO = " & TOBJETO & " AND COBJETO = " & COBJETO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                TOBJETO = rs("TOBJETO")
                COBJETO = rs("COBJETO")
                ENSAYO_ID = rs("ENSAYO_ID")
                PROGRAMA_ID = rs("PROGRAMA_ID")
                SECTION_ID = rs("SECTION_ID")
                FLUID_ID = rs("FLUID_ID")
                FACILITY_ID = rs("FACILITY_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsairbus_objetos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO airbus_objetos " & _
                   "  VALUES (" & _
                        TOBJETO & "," & COBJETO & "," & ENSAYO_ID & "," & PROGRAMA_ID & "," & _
                        SECTION_ID & "," & FLUID_ID & "," & FACILITY_ID & _
                ") " & _
                "  ON DUPLICATE KEY UPDATE " & _
                " ENSAYO_ID = " & ENSAYO_ID & "," & _
                " PROGRAMA_ID = " & PROGRAMA_ID & "," & _
                " SECTION_ID = " & SECTION_ID & "," & _
                " FLUID_ID = " & FLUID_ID & "," & _
                " FACILITY_ID = " & FACILITY_ID
        execute_bd consulta
        Insertar = TOBJETO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsairbus_objetos)", vbCritical, Err.Description
End Function
Public Function Eliminar(TOBJETO As Long, COBJETO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM airbus_objetos " & _
                "    WHERE TOBJETO = " & TOBJETO & _
                "      AND COBJETO = " & COBJETO
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsairbus_objetos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM airbus_objetos ORDER BY TOBJETO"
        Set Listado = datos_bd(consulta)
End Function

