VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_banos_probetas_materiales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_BANOS_PROBETAS_MATERIALES
'***************************************************************
Private BANO_ID As Long
Private DESIGNACION As String
Private AREA As Long
Private MATERIAL_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_BANOS_PROBETAS_MATERIALES
'***************************************************************
Public Property Let setBANO_ID(ByVal dato As Long)
        BANO_ID = dato
End Property
Public Property Let setDESIGNACION(ByVal dato As String)
        DESIGNACION = dato
End Property
Public Property Let setAREA(ByVal dato As Long)
        AREA = dato
End Property
Public Property Let setMATERIAL_ID(ByVal dato As Long)
        MATERIAL_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_BANOS_PROBETAS_MATERIALES
'***************************************************************
Public Property Get getBANO_ID() As Long
        getBANO_ID = BANO_ID
End Property
Public Property Get getDESIGNACION() As String
        getDESIGNACION = DESIGNACION
End Property
Public Property Get getAREA() As Long
        getAREA = AREA
End Property
Public Property Get getMATERIAL_ID() As Long
        getMATERIAL_ID = MATERIAL_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_BANOS_PROBETAS_MATERIALES
'***************************************************************
Public Function Carga(BANO As Long, DESIGN As String, NUM_AREA As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_BANOS_PROBETAS_MATERIALES " & _
                   " WHERE BANO_ID = " & BANO & _
                   "   AND DESIGNACION = '" & DESIGN & "'" & _
                   "   AND AREA = " & NUM_AREA
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                BANO_ID = rs("BANO_ID")
                DESIGNACION = rs("DESIGNACION")
                AREA = rs("AREA")
                MATERIAL_ID = rs("MATERIAL_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_banos_probetas_materiales)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "DELETE FROM CE_BANOS_PROBETAS_MATERIALES WHERE BANO_ID = " & BANO_ID & " AND DESIGNACION = '" & DESIGNACION & "' AND AREA = " & AREA
        consulta = "INSERT INTO CE_BANOS_PROBETAS_MATERIALES " & _
                   "  VALUES (" & _
                        BANO_ID & "," & "'" & DESIGNACION & "'" & "," & AREA & "," & _
                        MATERIAL_ID & _
                ")"
        execute_bd consulta
        Insertar = BANO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_banos_probetas_materiales)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_BANOS_PROBETAS_MATERIALES SET " & _
                        " DESIGNACION = '" & DESIGNACION & "'," & _
                        " AREA = " & AREA & "," & _
                        " MATERIAL_ID = " & MATERIAL_ID & "" & _
                " WHERE BANO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_banos_probetas_materiales)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_BANOS_PROBETAS_MATERIALES " & _
                "    WHERE BANO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_banos_probetas_materiales)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_BANOS_PROBETAS_MATERIALES ORDER BY BANO_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function MATERIAL(lBANO As Long, lDESIGNACION As String, lAREA As Integer) As String
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        consulta = "SELECT B.MATERIAL " & _
                   "  FROM CE_BANOS_PROBETAS_MATERIALES A, CE_BANOS_MATERIALES B " & _
                   " WHERE BANO_ID = " & lBANO & " AND DESIGNACION = '" & lDESIGNACION & "' AND AREA = " & lAREA & _
                   "   AND A.MATERIAL_ID = B.ID_MATERIAL"
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            MATERIAL = rs(0)
        Else
            MATERIAL = ""
        End If
End Function
Public Function CRITERIO(lBANO As Long, lDESIGNACION As String, lAREA As Integer) As String
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        consulta = "SELECT B.CRITERIO " & _
                   "  FROM CE_BANOS_PROBETAS_MATERIALES A, CE_BANOS_MATERIALES B " & _
                   " WHERE BANO_ID = " & lBANO & " AND DESIGNACION = '" & lDESIGNACION & "' AND AREA = " & lAREA & _
                   "   AND A.MATERIAL_ID = B.ID_MATERIAL"
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            CRITERIO = rs(0)
        Else
            CRITERIO = ""
        End If
End Function

