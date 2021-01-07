VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_banos_materiales"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_BANOS_MATERIALES
'***************************************************************
Private ID_MATERIAL As Long
Private MATERIAL As String
Private MINIMO As String
Private MAXIMO As String
Private MINIMO_TEXTO As String
Private MAXIMO_TEXTO As String
Private criterio As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_BANOS_MATERIALES
'***************************************************************
Public Property Let setID_MATERIAL(ByVal dato As Long)
        ID_MATERIAL = dato
End Property
Public Property Let setMATERIAL(ByVal dato As String)
        MATERIAL = dato
End Property
Public Property Let setMINIMO(ByVal dato As String)
        MINIMO = dato
End Property
Public Property Let setMAXIMO(ByVal dato As String)
        MAXIMO = dato
End Property
Public Property Let setMINIMO_TEXTO(ByVal dato As String)
        MINIMO_TEXTO = dato
End Property
Public Property Let setMAXIMO_TEXTO(ByVal dato As String)
        MAXIMO_TEXTO = dato
End Property
Public Property Let setCRITERIO(ByVal dato As String)
        criterio = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_BANOS_MATERIALES
'***************************************************************
Public Property Get getID_MATERIAL() As Long
        getID_MATERIAL = ID_MATERIAL
End Property
Public Property Get getMATERIAL() As String
        getMATERIAL = MATERIAL
End Property
Public Property Get getMINIMO() As String
        getMINIMO = MINIMO
End Property
Public Property Get getMAXIMO() As String
        getMAXIMO = MAXIMO
End Property
Public Property Get getMINIMO_TEXTO() As String
        getMINIMO_TEXTO = MINIMO_TEXTO
End Property
Public Property Get getMAXIMO_TEXTO() As String
        getMAXIMO_TEXTO = MAXIMO_TEXTO
End Property
Public Property Get getCRITERIO() As String
        getCRITERIO = criterio
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_BANOS_MATERIALES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_BANOS_MATERIALES WHERE ID_MATERIAL = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_MATERIAL = rs("ID_MATERIAL")
                MATERIAL = rs("MATERIAL")
                MINIMO = rs("MINIMO")
                MAXIMO = rs("MAXIMO")
                MINIMO_TEXTO = rs("MINIMO_TEXTO")
                MAXIMO_TEXTO = rs("MAXIMO_TEXTO")
                criterio = rs("CRITERIO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_banos_materiales)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_MATERIAL) FROM CE_BANOS_MATERIALES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_MATERIAL = 1
        Else
                ID_MATERIAL = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO CE_BANOS_MATERIALES " & _
                   "  VALUES (" & _
                        ID_MATERIAL & "," & "'" & MATERIAL & "'" & "," & "'" & MINIMO & "'" & "," & _
                        "'" & MAXIMO & "'" & "," & "'" & MINIMO_TEXTO & "'" & "," & "'" & MAXIMO_TEXTO & "'" & "," & _
                        "'" & criterio & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_MATERIAL
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_banos_materiales)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_BANOS_MATERIALES SET " & _
                        " MATERIAL = '" & MATERIAL & "'," & _
                        " MINIMO = '" & MINIMO & "'," & _
                        " MAXIMO = '" & MAXIMO & "'," & _
                        " MINIMO_TEXTO = '" & MINIMO_TEXTO & "'," & _
                        " MAXIMO_TEXTO = '" & MAXIMO_TEXTO & "'," & _
                        " CRITERIO = '" & criterio & "'" & _
                " WHERE ID_MATERIAL = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_banos_materiales)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_BANOS_MATERIALES " & _
                "    WHERE ID_MATERIAL = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_banos_materiales)", vbCritical, Err.Description
End Function
Public Function Listado(MAT As String, CRIT As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_BANOS_MATERIALES " & _
                   " WHERE MATERIAL LIKE '%" & MAT & "%' " & _
                   "   AND CRITERIO LIKE '%" & CRIT & "%' " & _
                   " ORDER BY MATERIAL"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_MATERIAL,MATERIAL FROM CE_BANOS_MATERIALES ORDER BY MATERIAL"
        Set Listado_Combo = datos_bd(consulta)
End Function
