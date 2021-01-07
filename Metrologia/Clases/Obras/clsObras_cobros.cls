VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObras_cobros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOBRAS_COBROS
'***************************************************************
Private OBRA_ID As Long
Private NOMBRE As String
Private TELEFONO As String
Private OBSERVACIONES As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOBRAS_COBROS
'***************************************************************
Public Property Let setOBRA_ID(ByVal dato As Long)
        OBRA_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setTELEFONO(ByVal dato As String)
        TELEFONO = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOBRAS_COBROS
'***************************************************************
Public Property Get getOBRA_ID() As Long
        getOBRA_ID = OBRA_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getTELEFONO() As String
        getTELEFONO = TELEFONO
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOBRAS_COBROS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OBRAS_COBROS WHERE OBRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                OBRA_ID = rs("OBRA_ID")
                NOMBRE = rs("NOMBRE")
                TELEFONO = rs("TELEFONO")
                OBSERVACIONES = rs("OBSERVACIONES")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsObras_cobros)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim oobra As New clsObras_cobros
        If oobra.Carga(OBRA_ID) = False Then
            consulta = "INSERT INTO OBRAS_COBROS " & _
                       "  VALUES (" & _
                            OBRA_ID & "," & "'" & NOMBRE & "'" & "," & "'" & TELEFONO & "'" & "," & _
                            "'" & OBSERVACIONES & "'" & _
                    ")"
            execute_bd consulta
        Else
            Me.Modificar OBRA_ID
        End If
        Insertar = OBRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsObras_cobros)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OBRAS_COBROS SET " & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " TELEFONO = '" & TELEFONO & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'" & _
                " WHERE OBRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsObras_cobros)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM OBRAS_COBROS " & _
                "    WHERE OBRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsObras_cobros)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OBRAS_COBROS ORDER BY OBRA_ID"
        Set Listado = datos_bd(consulta)
End Function
