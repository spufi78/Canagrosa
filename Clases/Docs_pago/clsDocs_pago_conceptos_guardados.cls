VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocs_pago_conceptos_guardados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDOCS_PAGO_CONCEPTOS_GUARDADOS
'***************************************************************
Private ID As Long
Private texto As String
Private PRECIO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDOCS_PAGO_CONCEPTOS_GUARDADOS
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setTEXTO(ByVal dato As String)
        texto = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDOCS_PAGO_CONCEPTOS_GUARDADOS
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getTEXTO() As String
        getTEXTO = texto
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDOCS_PAGO_CONCEPTOS_GUARDADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM DOCS_PAGO_CONCEPTOS_GUARDADOS WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                texto = rs("TEXTO")
                PRECIO = rs("PRECIO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDocs_pago_conceptos_guardados)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT MAX(ID) FROM DOCS_PAGO_CONCEPTOS_GUARDADOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID = 1
        Else
                ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        consulta = "SELECT * FROM DOCS_PAGO_CONCEPTOS_GUARDADOS WHERE TEXTO = '" & texto & "'"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Me.CrearID
            consulta = "INSERT INTO DOCS_PAGO_CONCEPTOS_GUARDADOS " & _
                       "  VALUES (" & _
                            ID & "," & "'" & texto & "'" & ",'" & PRECIO & _
                    "')"
            execute_bd consulta
            Insertar = ID
        End If
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDocs_pago_conceptos_guardados)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCS_PAGO_CONCEPTOS_GUARDADOS SET " & _
                        " TEXTO = '" & texto & "'," & _
                        " PRECIO = '" & PRECIO & "'" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDocs_pago_conceptos_guardados)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DOCS_PAGO_CONCEPTOS_GUARDADOS " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDocs_pago_conceptos_guardados)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM DOCS_PAGO_CONCEPTOS_GUARDADOS ORDER BY TEXTO"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT ID,TEXTO FROM DOCS_PAGO_CONCEPTOS_GUARDADOS ORDER BY TEXTO"
        Set Listado_Combo = datos_bd(consulta)
End Function
