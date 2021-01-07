VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_probetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_PROBETAS
'***************************************************************
Private FICHA_ID As Long
Private DESIGNACION As String
Private MATERIAL As String
Private TT As String
Private DIMENSION As String
Private NUMERO As String
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_PROBETAS
'***************************************************************
Public Property Let setFICHA_ID(ByVal dato As Long)
        FICHA_ID = dato
End Property
Public Property Let setDESIGNACION(ByVal dato As String)
        DESIGNACION = dato
End Property
Public Property Let setMATERIAL(ByVal dato As String)
        MATERIAL = dato
End Property
Public Property Let setTT(ByVal dato As String)
        TT = dato
End Property
Public Property Let setDIMENSION(ByVal dato As String)
        DIMENSION = dato
End Property
Public Property Let setNUMERO(ByVal dato As String)
        NUMERO = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_PROBETAS
'***************************************************************
Public Property Get getFICHA_ID() As Long
        getFICHA_ID = FICHA_ID
End Property
Public Property Get getDESIGNACION() As String
        getDESIGNACION = DESIGNACION
End Property
Public Property Get getMATERIAL() As String
        getMATERIAL = MATERIAL
End Property
Public Property Get getTT() As String
        getTT = TT
End Property
Public Property Get getDIMENSION() As String
        getDIMENSION = DIMENSION
End Property
Public Property Get getNUMERO() As String
        getNUMERO = NUMERO
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_PROBETAS
'***************************************************************
Public Function Carga(id As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_PROBETAS WHERE FICHA_ID = " & id
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                FICHA_ID = rs("FICHA_ID")
                DESIGNACION = rs("DESIGNACION")
                MATERIAL = rs("MATERIAL")
                TT = rs("TT")
                DIMENSION = rs("DIMENSION")
                NUMERO = rs("NUMERO")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_probetas)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CE_PROBETAS " & _
                   "  VALUES (" & _
                        FICHA_ID & "," & "'" & DESIGNACION & "'" & "," & "'" & MATERIAL & "'" & "," & _
                        "'" & TT & "'" & "," & "'" & DIMENSION & "'" & "," & "'" & NUMERO & "'" & "," & _
                        ORDEN & _
                ")"
        conn.Execute consulta
        Insertar = FICHA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_probetas)", vbCritical, Err.Description
End Function
Public Function Modificar(id As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_PROBETAS SET " & _
                        " DESIGNACION = '" & DESIGNACION & "'," & _
                        " MATERIAL = '" & MATERIAL & "'," & _
                        " TT = '" & TT & "'," & _
                        " DIMENSION = '" & DIMENSION & "'," & _
                        " NUMERO = '" & NUMERO & "'," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE FICHA_ID = " & id
        conn.Execute consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_probetas)", vbCritical, Err.Description
End Function
Public Function Eliminar(id As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_PROBETAS " & _
                "    WHERE FICHA_ID = " & id
        conn.Execute consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_probetas)", vbCritical, Err.Description
End Function
Public Function Listado(id As Integer) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT DESIGNACION,MATERIAL,TT,DIMENSION,NUMERO " & _
                   "  FROM CE_PROBETAS " & _
                   " WHERE FICHA_ID = " & id & " ORDER BY FICHA_ID"
        Set Listado = datos_bd(consulta)
End Function
