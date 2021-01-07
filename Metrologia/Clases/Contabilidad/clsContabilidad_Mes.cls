VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad_Mes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsContabilidad_Mes
'***************************************************************
Private MES As String
Private Sdebe As String
Private Shaber As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsContabilidad_Mes
'***************************************************************
Public Property Let setMES(ByVal dato As String)
        MES = dato
End Property
Public Property Let setSDEBE(ByVal dato As String)
        Sdebe = dato
End Property
Public Property Let setSHABER(ByVal dato As String)
        Shaber = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsContabilidad_Mes
'***************************************************************
Public Property Get getMES() As String
        getMES = MES
End Property
Public Property Get getSDEBE() As String
        getSDEBE = Sdebe
End Property
Public Property Get getSHABER() As String
        getSHABER = Shaber
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsContabilidad_Mes
'***************************************************************
Public Function Carga(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MES WHERE MES = '" & ID & "'"
        Dim oC As New clsContabilidad
        Set rs = oC.datos_bd_contabilidad(consulta)
        Set oC = Nothing
        
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MES = rs("MES")
                Sdebe = rs("Sdebe")
                Shaber = rs("Shaber")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsContabilidad_Mes)", vbCritical, Err.Description
End Function
Public Function Actualizar(lMES As String, IMPORTE As Currency) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        If Me.Carga(lMES) Then
            Sdebe = Sdebe + IMPORTE
            Shaber = Shaber + (-1 * IMPORTE)
            Me.Modificar lMES
        Else
            MES = lMES
            Sdebe = IMPORTE
            Shaber = -1 * IMPORTE
            Me.Insertar
        End If
        Actualizar = True
        Exit Function
fallo:
        Actualizar = False
        MsgBox "Error al actualizar (clsContabilidad_Mes)", vbCritical, Err.Description
End Function

Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO MES " & _
                   "  VALUES (" & _
                        "'" & MES & "'" & "," & "'" & Sdebe & "'" & "," & "'" & Shaber & "'" & _
                ")"
        Dim oC As New clsContabilidad
        oC.execute_bd consulta
        Set oC = Nothing
        Insertar = MES
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsContabilidad_Mes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE MES SET " & _
                        " SDEBE = '" & Sdebe & "'," & _
                        " SHABER = '" & Shaber & "'" & _
                " WHERE MES = '" & ID & "'"
        Dim oC As New clsContabilidad
        oC.execute_bd consulta
        Set oC = Nothing
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsContabilidad_Mes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM MES " & _
                "    WHERE MES = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsContabilidad_Mes)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM MES ORDER BY MES"
        Set Listado = datos_bd(consulta)
End Function
