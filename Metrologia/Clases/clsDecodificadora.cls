VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDecodificadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDECODIFICADORA
'***************************************************************
Private CODIGO As Integer
Private IDIOMA As String
Private VALOR As Integer
Private DESCRIPCION As String
Private PARAMETROS As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDECODIFICADORA
'***************************************************************
Public Property Let setCODIGO(ByVal dato As Integer)
        CODIGO = dato
End Property
Public Property Let setIDIOMA(ByVal dato As String)
        IDIOMA = dato
End Property
Public Property Let setVALOR(ByVal dato As Integer)
        VALOR = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setPARAMETROS(ByVal dato As String)
        PARAMETROS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDECODIFICADORA
'***************************************************************
Public Property Get getCODIGO() As Integer
        getCODIGO = CODIGO
End Property
Public Property Get getIDIOMA() As String
        getIDIOMA = IDIOMA
End Property
Public Property Get getVALOR() As Integer
        getVALOR = VALOR
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getPARAMETROS() As String
        getPARAMETROS = PARAMETROS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDECODIFICADORA
'***************************************************************
Public Function Crear_nuevo_valor()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(VALOR) FROM DECODIFICADORA WHERE CODIGO = " & CODIGO
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                VALOR = 1
        Else
                VALOR = CInt(rs.Fields(0)) + 1
        End If
        Set rs = Nothing
End Function

Public Function Carga(ID As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DECODIFICADORA WHERE CODIGO = 0 AND VALOR = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                CODIGO = rs("CODIGO")
                IDIOMA = rs("IDIOMA")
                VALOR = rs("VALOR")
                DESCRIPCION = rs("DESCRIPCION")
                PARAMETROS = rs("PARAMETROS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDecodificadora)", vbCritical, Err.Description
End Function
Public Function Carga_valor(lCODIGO As Long, lVALOR As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DECODIFICADORA WHERE CODIGO = " & lCODIGO & " AND VALOR = " & lVALOR
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_valor = False
        Else
                CODIGO = rs("CODIGO")
                IDIOMA = rs("IDIOMA")
                VALOR = rs("VALOR")
                DESCRIPCION = rs("DESCRIPCION")
                PARAMETROS = rs("PARAMETROS")
                Carga_valor = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_valor = False
        MsgBox "Error al cargar los datos(clsDecodificadora)", vbCritical, Err.Description
End Function

Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT * FROM DECODIFICADORA WHERE CODIGO = " & CODIGO & " AND DESCRIPCION = '" & DESCRIPCION & "'"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            If Trim(VALOR) = 0 Then
                Me.Crear_nuevo_valor
            End If
            consulta = "INSERT INTO DECODIFICADORA " & _
                       "  VALUES (" & _
                            CODIGO & "," & "'" & IDIOMA & "'" & "," & VALOR & "," & _
                            "'" & DESCRIPCION & "','" & PARAMETROS & "'" & _
                    ")"
            execute_bd consulta
        End If
        Insertar = CODIGO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDecodificadora)", vbCritical, Err.Description
End Function
Public Function Modificar(cCODIGO As Integer, cVALOR As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DECODIFICADORA SET " & _
                        " IDIOMA = '" & IDIOMA & "'," & _
                        " PARAMETROS = '" & PARAMETROS & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'" & _
                " WHERE CODIGO = " & cCODIGO & " AND VALOR = " & cVALOR
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDecodificadora)", vbCritical, Err.Description
End Function
Public Function Eliminar(cCODIGO As Integer, cVALOR As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        If Me.Verificar_integridad(cCODIGO, cVALOR) Then
            consulta = "DELETE FROM DECODIFICADORA " & _
                    "    WHERE CODIGO = " & cCODIGO & " AND VALOR = " & cVALOR
            execute_bd consulta
        End If
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDecodificadora)", vbCritical, Err.Description
End Function
Public Function Listado(CODIGO) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DECODIFICADORA " & _
                   " WHERE CODIGO = " & CODIGO & _
                   "   AND VALOR <> 0 " & _
                   " ORDER BY DESCRIPCION"
        Set Listado = datos_bd(consulta)
End Function
Public Function Cargar_Combo(combo As DataCombo, CCODTABL As Integer) As Boolean
    Dim consulta As String
    consulta = "SELECT VALOR,DESCRIPCION FROM DECODIFICADORA WHERE CODIGO = " & CCODTABL & " order by DESCRIPCION"
    Dim rs As ADODB.Recordset
    Set rs = datos_bd(consulta)
    Set combo.RowSource = rs
    combo.ListField = rs(1).Name
    combo.BoundColumn = rs(0).Name
    Set rs = Nothing
End Function
Public Function Cargar_miCombo(combo As miComboBCA, CCODTABL As Integer) As Boolean
    With combo
        .setCONN = conn
        .setQUERY = "SELECT VALOR,DESCRIPCION FROM DECODIFICADORA WHERE CODIGO = " & CCODTABL
        .setFK_CAMPO = ""
        .setFK_VALOR = 0
        .setTABLA = "DECODIFICADORA"
        .setDESCRIPCION = ""
        .setPK = "VALOR"
        .setCAMPO = "DESCRIPCION"
        .setMUESTRA_DETALLE = False
'        .setFILTRO = " AND CODIGO = " & CCODTABL
        Set .FORMULARIO = frmMenu
    End With
    
'    Dim consulta As String
'    consulta = "SELECT VALOR,DESCRIPCION FROM DECODIFICADORA WHERE CODIGO = " & CCODTABL & " order by DESCRIPCION"
'    Dim RS As ADODB.Recordset
'    Set RS = datos_bd(consulta)
'    Set combo.RowSource = RS
'    combo.ListField = RS(1).Name
'    combo.BoundColumn = RS(0).Name
'    Set RS = Nothing
End Function
Public Function Verificar_integridad(cCODIGO As Integer, cVALOR As Integer) As Boolean
        Dim consulta As String
        Select Case cCODIGO
'         Case decodificadora.CA_NORMAS_ESTADOS
'            consulta = "SELECT * FROM CA_NORMAS WHERE ESTADO_ID = " & cVALOR
'         Case decodificadora.CA_NORMAS_SECTORES
'            consulta = "SELECT * FROM CA_NORMAS WHERE SECTOR_ID = " & cVALOR
        End Select
        Dim rs As ADODB.Recordset
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
            Verificar_integridad = True
        Else
            Verificar_integridad = False
            MsgBox "No se puede eliminar el registro. Tiene datos vinculados.", vbExclamation, App.Title
        End If
        Set rs = Nothing
End Function
