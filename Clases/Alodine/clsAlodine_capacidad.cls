VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlodine_capacidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSALODINE_CAPACIDAD
'***************************************************************
Private ID_CAPACIDAD As Long
Private DESCRIPCION As String
Private TAMANO_ETIQUETA_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSALODINE_CAPACIDAD
'***************************************************************
Public Property Let setID_CAPACIDAD(ByVal dato As Long)
        ID_CAPACIDAD = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setTAMANO_ETIQUETA_ID(ByVal dato As Long)
        TAMANO_ETIQUETA_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSALODINE_CAPACIDAD
'***************************************************************
Public Property Get getID_CAPACIDAD() As Long
        getID_CAPACIDAD = ID_CAPACIDAD
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getTAMANO_ETIQUETA_ID() As Long
        getTAMANO_ETIQUETA_ID = TAMANO_ETIQUETA_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSALODINE_CAPACIDAD
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ALODINE_CAPACIDAD WHERE ID_CAPACIDAD = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_CAPACIDAD = rs("ID_CAPACIDAD")
                DESCRIPCION = rs("DESCRIPCION")
                TAMANO_ETIQUETA_ID = rs("TAMANO_ETIQUETA_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAlodine_capacidad)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_CAPACIDAD) FROM ALODINE_CAPACIDAD"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_CAPACIDAD = 1
        Else
                ID_CAPACIDAD = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO ALODINE_CAPACIDAD " & _
                   "  VALUES (" & _
                        ID_CAPACIDAD & "," & "'" & DESCRIPCION & "'" & "," & TAMANO_ETIQUETA_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_CAPACIDAD
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsAlodine_capacidad)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ALODINE_CAPACIDAD SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " TAMANO_ETIQUETA_ID = " & TAMANO_ETIQUETA_ID & "" & _
                " WHERE ID_CAPACIDAD = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsAlodine_capacidad)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ALODINE_CAPACIDAD " & _
                "    WHERE ID_CAPACIDAD = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAlodine_capacidad)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT AC.DESCRIPCION,E.NOMBRE,AC.ID_CAPACIDAD " & _
                   "  FROM ALODINE_CAPACIDAD AC, TAMANOS_ETIQUETA E " & _
                   " WHERE AC.TAMANO_ETIQUETA_ID = E.ID_TAMANO_ETIQUETA " & _
                   " ORDER BY AC.DESCRIPCION"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_CAPACIDAD,DESCRIPCION FROM ALODINE_CAPACIDAD ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
