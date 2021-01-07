VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_banos_probetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_BANOS_PROBETAS
'***************************************************************
Private BANO_ID As Long
Private DESIGNACION As String
Private MATERIAL As String
Private TT As String
Private DIMENSION As String
Private cantidad As Integer
Private AREAS As Integer
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_BANOS_PROBETAS
'***************************************************************
Public Property Let setBANO_ID(ByVal dato As Long)
        BANO_ID = dato
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
Public Property Let setCANTIDAD(ByVal dato As Integer)
        cantidad = dato
End Property
Public Property Let setAREAS(ByVal dato As Integer)
        AREAS = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_BANOS_PROBETAS
'***************************************************************
Public Property Get getBANO_ID() As Long
        getBANO_ID = BANO_ID
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
Public Property Get getCANTIDAD() As Integer
        getCANTIDAD = cantidad
End Property
Public Property Get getAREAS() As Integer
        getAREAS = AREAS
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_BANOS_PROBETAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_BANOS_PROBETAS WHERE BANO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                BANO_ID = rs("BANO_ID")
                DESIGNACION = rs("DESIGNACION")
                MATERIAL = rs("MATERIAL")
                TT = rs("TT")
                DIMENSION = rs("DIMENSION")
                cantidad = rs("CANTIDAD")
                AREAS = rs("AREAS")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_banos_probetas)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "DELETE FROM CE_BANOS_PROBETAS_MATERIALES WHERE BANO_ID = " & BANO_ID & " AND DESIGNACION = '" & DESIGNACION & "' AND AREA > " & AREAS
        consulta = "INSERT INTO CE_BANOS_PROBETAS " & _
                   "  VALUES (" & _
                        BANO_ID & ",'" & DESIGNACION & "'" & "," & "'" & MATERIAL & "'" & "," & _
                        "'" & TT & "'" & "," & "'" & DIMENSION & "'" & "," & cantidad & "," & AREAS & "," & _
                        ORDEN & _
                ")"
        execute_bd consulta
        Insertar = BANO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_banos_probetas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_BANOS_PROBETAS SET " & _
                        " DESIGNACION = '" & DESIGNACION & "'," & _
                        " MATERIAL = '" & MATERIAL & "'," & _
                        " TT = '" & TT & "'," & _
                        " DIMENSION = '" & DIMENSION & "'," & _
                        " CANTIDAD = " & cantidad & "," & _
                        " AREAS = " & AREAS & "," & _
                        " ORDEN = " & ORDEN & "" & _
                " WHERE BANO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_banos_probetas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_BANOS_PROBETAS " & _
                "    WHERE BANO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_banos_probetas)", vbCritical, Err.Description
End Function
Public Function Listado(BANO As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT DESIGNACION,MATERIAL,TT,DIMENSION,CANTIDAD,AREAS " & _
                   "  FROM CE_BANOS_PROBETAS " & _
                   " WHERE BANO_ID = " & BANO & _
                   " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Completo(BANO As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * " & _
                   "  FROM CE_BANOS_PROBETAS " & _
                   " WHERE BANO_ID = " & BANO & _
                   " ORDER BY ORDEN"
        Set Listado_Completo = datos_bd(consulta)
End Function


