VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMuestras_airbus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsMuestras_airbus
'***************************************************************
Private MUESTRA_ID As Long
Private ENSAYO_ID As Long
Private PROGRAMA_ID As Long
Private SECTION_ID As Long
Private FLUID_ID As Long
Private FACILITY_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsMuestras_airbus
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
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
'*   PROPIEDADES DE LECTURA DE LA CLASE clsMuestras_airbus
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
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
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsMuestras_airbus
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MUESTRAS_AIRBUS WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
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
        MsgBox "Error al cargar los datos(clsMuestras_airbus)", vbCritical, Err.Description
End Function
Public Function Insertar(ENSAYO As Boolean, PROGRAMA As Boolean, SECTION As Boolean, FLUID As Boolean, FACILITY As Boolean) As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim oMA As New clsMuestras_airbus
        Dim existe As Boolean
        existe = oMA.Carga(MUESTRA_ID)
        If ENSAYO = False Then
            If existe Then
                ENSAYO_ID = oMA.getENSAYO_ID
            Else
                ENSAYO_ID = 0
            End If
        End If
        If PROGRAMA = False Then
            If existe Then
                PROGRAMA_ID = oMA.getPROGRAMA_ID
            Else
                PROGRAMA_ID = 0
            End If
        End If
        If SECTION = False Then
            If existe Then
                SECTION_ID = oMA.getSECTION_ID
            Else
                SECTION_ID = 0
            End If
        End If
        If FLUID = False Then
            If existe Then
                FLUID_ID = oMA.getFLUID_ID
            Else
                FLUID_ID = 0
            End If
        End If
        If FACILITY = False Then
            If existe Then
                FACILITY_ID = oMA.getFACILITY_ID
            Else
                FACILITY_ID = 0
            End If
        End If
        
        consulta = "INSERT INTO MUESTRAS_AIRBUS " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & ENSAYO_ID & "," & PROGRAMA_ID & "," & _
                        SECTION_ID & "," & FLUID_ID & "," & FACILITY_ID & _
                ") " & _
                "  ON DUPLICATE KEY UPDATE " & _
                " ENSAYO_ID = " & ENSAYO_ID & "," & _
                " PROGRAMA_ID = " & PROGRAMA_ID & "," & _
                " SECTION_ID = " & SECTION_ID & "," & _
                " FLUID_ID = " & FLUID_ID & "," & _
                " FACILITY_ID = " & FACILITY_ID
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMuestras_airbus)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE MUESTRAS_AIRBUS SET " & _
                        " ENSAYO_ID = " & ENSAYO_ID & "," & _
                        " PROGRAMA_ID = " & PROGRAMA_ID & "," & _
                        " SECTION_ID = " & SECTION_ID & "," & _
                        " FLUID_ID = " & FLUID_ID & "," & _
                        " FACILITY_ID = " & FACILITY_ID & "" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMuestras_airbus)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MUESTRAS_AIRBUS " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMuestras_airbus)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MUESTRAS_AIRBUS ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE clsMuestras_airbus
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

