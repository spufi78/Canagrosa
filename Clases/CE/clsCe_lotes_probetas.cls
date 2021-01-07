VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_lotes_probetas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_LOTES_PROBETAS
'***************************************************************
Private ID_LOTE As Long
Private Identificacion As String
Private IDENTIFICACION_COMBO As String
Private NUMERO_LOTE As String
Private NUMERO_INFORME As String
Private CARGA_ROTURA As Long
Private MATERIAL As String
Private ESPESOR As Long
Private INDICATIVO_CARGA_75 As Long
Private INDICATIVO_CARGA_80 As Long
Private INDICATIVO_CARGA_85 As Long
Private INDICATIVO_CARGA_90 As Long
Private INDICATIVO_CARGA_73 As Long
Private INDICATIVO_CARGA_74 As Long
Private INDICATIVO_CARGA_76 As Long
Private INDICATIVO_CARGA_77 As Long
Private INDICATIVO_CARGA_81 As Long
Private INDICATIVO_CARGA_79 As Long
Private TE_75 As Long
Private TE_80 As Long
Private TE_85 As Long
Private TE_90 As Long
Private TE_73 As Long
Private TE_74 As Long
Private TE_76 As Long
Private TE_77 As Long
Private TE_81 As Long
Private TE_79 As Long
Private TOTAL_TE As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_LOTES_PROBETAS
'***************************************************************
Public Property Let setID_LOTE(ByVal dato As Long)
        ID_LOTE = dato
End Property
Public Property Let setIDENTIFICACION(ByVal dato As String)
        Identificacion = dato
End Property
Public Property Let setIDENTIFICACION_COMBO(ByVal dato As String)
        IDENTIFICACION_COMBO = dato
End Property
Public Property Let setNUMERO_LOTE(ByVal dato As String)
        NUMERO_LOTE = dato
End Property
Public Property Let setNUMERO_INFORME(ByVal dato As String)
        NUMERO_INFORME = dato
End Property
Public Property Let setCARGA_ROTURA(ByVal dato As Long)
        CARGA_ROTURA = dato
End Property
Public Property Let setMATERIAL(ByVal dato As String)
        MATERIAL = dato
End Property
Public Property Let setESPESOR(ByVal dato As Long)
        ESPESOR = dato
End Property
Public Property Let setINDICATIVO_CARGA_75(ByVal dato As Long)
        INDICATIVO_CARGA_75 = dato
End Property
Public Property Let setINDICATIVO_CARGA_80(ByVal dato As Long)
        INDICATIVO_CARGA_80 = dato
End Property
Public Property Let setINDICATIVO_CARGA_85(ByVal dato As Long)
        INDICATIVO_CARGA_85 = dato
End Property
Public Property Let setINDICATIVO_CARGA_90(ByVal dato As Long)
        INDICATIVO_CARGA_90 = dato
End Property
Public Property Let setINDICATIVO_CARGA_73(ByVal dato As Long)
        INDICATIVO_CARGA_73 = dato
End Property
Public Property Let setINDICATIVO_CARGA_74(ByVal dato As Long)
        INDICATIVO_CARGA_74 = dato
End Property
Public Property Let setINDICATIVO_CARGA_76(ByVal dato As Long)
        INDICATIVO_CARGA_76 = dato
End Property
Public Property Let setINDICATIVO_CARGA_77(ByVal dato As Long)
        INDICATIVO_CARGA_77 = dato
End Property
Public Property Let setINDICATIVO_CARGA_81(ByVal dato As Long)
        INDICATIVO_CARGA_81 = dato
End Property
Public Property Let setINDICATIVO_CARGA_79(ByVal dato As Long)
        INDICATIVO_CARGA_79 = dato
End Property
Public Property Let setTE_75(ByVal dato As Long)
        TE_75 = dato
End Property
Public Property Let setTE_80(ByVal dato As Long)
        TE_80 = dato
End Property
Public Property Let setTE_85(ByVal dato As Long)
        TE_85 = dato
End Property
Public Property Let setTE_90(ByVal dato As Long)
        TE_90 = dato
End Property
Public Property Let setTE_73(ByVal dato As Long)
        TE_73 = dato
End Property
Public Property Let setTE_74(ByVal dato As Long)
        TE_74 = dato
End Property
Public Property Let setTE_76(ByVal dato As Long)
        TE_76 = dato
End Property
Public Property Let setTE_77(ByVal dato As Long)
        TE_77 = dato
End Property
Public Property Let setTE_81(ByVal dato As Long)
        TE_81 = dato
End Property
Public Property Let setTE_79(ByVal dato As Long)
        TE_79 = dato
End Property
Public Property Let setTOTAL_TE(ByVal dato As Long)
        TOTAL_TE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_LOTES_PROBETAS
'***************************************************************
Public Property Get getID_LOTE() As Long
        getID_LOTE = ID_LOTE
End Property
Public Property Get getIDENTIFICACION() As String
        getIDENTIFICACION = Identificacion
End Property
Public Property Get getIDENTIFICACION_COMBO() As String
        getIDENTIFICACION_COMBO = IDENTIFICACION_COMBO
End Property
Public Property Get getNUMERO_LOTE() As String
        getNUMERO_LOTE = NUMERO_LOTE
End Property
Public Property Get getNUMERO_INFORME() As String
        getNUMERO_INFORME = NUMERO_INFORME
End Property
Public Property Get getCARGA_ROTURA() As Long
        getCARGA_ROTURA = CARGA_ROTURA
End Property
Public Property Get getMATERIAL() As String
        getMATERIAL = MATERIAL
End Property
Public Property Get getESPESOR() As Long
        getESPESOR = ESPESOR
End Property
Public Property Get getINDICATIVO_CARGA_75() As Long
        getINDICATIVO_CARGA_75 = INDICATIVO_CARGA_75
End Property
Public Property Get getINDICATIVO_CARGA_80() As Long
        getINDICATIVO_CARGA_80 = INDICATIVO_CARGA_80
End Property
Public Property Get getINDICATIVO_CARGA_85() As Long
        getINDICATIVO_CARGA_85 = INDICATIVO_CARGA_85
End Property
Public Property Get getINDICATIVO_CARGA_90() As Long
        getINDICATIVO_CARGA_90 = INDICATIVO_CARGA_90
End Property
Public Property Get getINDICATIVO_CARGA_73() As Long
        getINDICATIVO_CARGA_73 = INDICATIVO_CARGA_73
End Property
Public Property Get getINDICATIVO_CARGA_74() As Long
        getINDICATIVO_CARGA_74 = INDICATIVO_CARGA_74
End Property
Public Property Get getINDICATIVO_CARGA_76() As Long
        getINDICATIVO_CARGA_76 = INDICATIVO_CARGA_76
End Property
Public Property Get getINDICATIVO_CARGA_77() As Long
        getINDICATIVO_CARGA_77 = INDICATIVO_CARGA_77
End Property
Public Property Get getINDICATIVO_CARGA_81() As Long
        getINDICATIVO_CARGA_81 = INDICATIVO_CARGA_81
End Property
Public Property Get getINDICATIVO_CARGA_79() As Long
        getINDICATIVO_CARGA_79 = INDICATIVO_CARGA_79
End Property
Public Property Get getTE_75() As Long
        getTE_75 = TE_75
End Property
Public Property Get getTE_80() As Long
        getTE_80 = TE_80
End Property
Public Property Get getTE_85() As Long
        getTE_85 = TE_85
End Property
Public Property Get getTE_90() As Long
        getTE_90 = TE_90
End Property
Public Property Get getTE_73() As Long
        getTE_73 = TE_73
End Property
Public Property Get getTE_74() As Long
        getTE_74 = TE_74
End Property
Public Property Get getTE_76() As Long
        getTE_76 = TE_76
End Property
Public Property Get getTE_77() As Long
        getTE_77 = TE_77
End Property
Public Property Get getTE_81() As Long
        getTE_81 = TE_81
End Property
Public Property Get getTE_79() As Long
        getTE_79 = TE_79
End Property
Public Property Get getTOTAL_TE() As Long
        getTOTAL_TE = TOTAL_TE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_LOTES_PROBETAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CE_LOTES_PROBETAS WHERE ID_LOTE = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_LOTE = rs("ID_LOTE")
                Identificacion = rs("IDENTIFICACION")
                IDENTIFICACION_COMBO = rs("IDENTIFICACION_COMBO")
                NUMERO_LOTE = rs("NUMERO_LOTE")
                NUMERO_INFORME = rs("NUMERO_INFORME")
                CARGA_ROTURA = rs("CARGA_ROTURA")
                MATERIAL = rs("MATERIAL")
                ESPESOR = rs("ESPESOR")
                INDICATIVO_CARGA_75 = rs("INDICATIVO_CARGA_75")
                INDICATIVO_CARGA_80 = rs("INDICATIVO_CARGA_80")
                INDICATIVO_CARGA_85 = rs("INDICATIVO_CARGA_85")
                INDICATIVO_CARGA_90 = rs("INDICATIVO_CARGA_90")
                INDICATIVO_CARGA_73 = rs("INDICATIVO_CARGA_73")
                INDICATIVO_CARGA_74 = rs("INDICATIVO_CARGA_74")
                INDICATIVO_CARGA_76 = rs("INDICATIVO_CARGA_76")
                INDICATIVO_CARGA_77 = rs("INDICATIVO_CARGA_77")
                INDICATIVO_CARGA_81 = rs("INDICATIVO_CARGA_81")
                INDICATIVO_CARGA_79 = rs("INDICATIVO_CARGA_79")
                TE_75 = rs("TE_75")
                TE_80 = rs("TE_80")
                TE_85 = rs("TE_85")
                TE_90 = rs("TE_90")
                TE_73 = rs("TE_73")
                TE_74 = rs("TE_74")
                TE_76 = rs("TE_76")
                TE_77 = rs("TE_77")
                TE_77 = rs("TE_81")
                TE_77 = rs("TE_79")
                TOTAL_TE = rs("TOTAL_TE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_lotes_probetas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_LOTE) FROM CE_LOTES_PROBETAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_LOTE = 1
        Else
                ID_LOTE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO CE_LOTES_PROBETAS " & _
                   "  VALUES (" & _
                        ID_LOTE & "," & "'" & Identificacion & "'" & "," & "'" & IDENTIFICACION_COMBO & "'" & "," & _
                        "'" & NUMERO_LOTE & "'" & "," & "'" & NUMERO_INFORME & "'" & "," & CARGA_ROTURA & "," & _
                        "'" & MATERIAL & "'" & "," & ESPESOR & "," & INDICATIVO_CARGA_75 & "," & _
                        INDICATIVO_CARGA_80 & "," & INDICATIVO_CARGA_85 & "," & INDICATIVO_CARGA_90 & "," & _
                        INDICATIVO_CARGA_73 & "," & INDICATIVO_CARGA_74 & "," & INDICATIVO_CARGA_76 & "," & INDICATIVO_CARGA_77 & "," & _
                        INDICATIVO_CARGA_81 & "," & INDICATIVO_CARGA_79 & "," & _
                        TE_75 & "," & _
                        TE_80 & "," & TE_85 & "," & TE_90 & "," & _
                        TE_73 & "," & TE_74 & "," & TE_76 & "," & TE_77 & "," & _
                        TE_81 & "," & TE_79 & "," & _
                        TOTAL_TE & _
                ")"
        execute_bd consulta
        Insertar = ID_LOTE
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_lotes_probetas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CE_LOTES_PROBETAS SET " & _
                        " IDENTIFICACION = '" & Identificacion & "'," & _
                        " IDENTIFICACION_COMBO = '" & IDENTIFICACION_COMBO & "'," & _
                        " NUMERO_LOTE = '" & NUMERO_LOTE & "'," & _
                        " NUMERO_INFORME = '" & NUMERO_INFORME & "'," & _
                        " CARGA_ROTURA = " & CARGA_ROTURA & "," & _
                        " MATERIAL = '" & MATERIAL & "'," & _
                        " ESPESOR = " & ESPESOR & "," & _
                        " INDICATIVO_CARGA_75 = " & INDICATIVO_CARGA_75 & "," & " INDICATIVO_CARGA_80 = " & INDICATIVO_CARGA_80 & "," & _
                        " INDICATIVO_CARGA_85 = " & INDICATIVO_CARGA_85 & "," & " INDICATIVO_CARGA_90 = " & INDICATIVO_CARGA_90 & "," & _
                        " INDICATIVO_CARGA_73 = " & INDICATIVO_CARGA_73 & "," & " INDICATIVO_CARGA_74 = " & INDICATIVO_CARGA_74 & "," & _
                        " INDICATIVO_CARGA_76 = " & INDICATIVO_CARGA_76 & "," & " INDICATIVO_CARGA_77 = " & INDICATIVO_CARGA_77 & "," & _
                        " INDICATIVO_CARGA_81 = " & INDICATIVO_CARGA_81 & "," & " INDICATIVO_CARGA_79 = " & INDICATIVO_CARGA_79 & "," & _
                        " TE_75 = " & TE_75 & "," & " TE_80 = " & TE_80 & "," & _
                        " TE_85 = " & TE_85 & "," & " TE_90 = " & TE_90 & "," & _
                        " TE_73 = " & TE_73 & "," & _
                        " TE_74 = " & TE_74 & "," & _
                        " TE_76 = " & TE_76 & "," & " TE_77 = " & TE_77 & "," & _
                        " TE_81 = " & TE_81 & "," & " TE_79 = " & TE_79 & "," & _
                        " TOTAL_TE = " & TOTAL_TE & "" & _
                " WHERE ID_LOTE = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCe_lotes_probetas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_LOTES_PROBETAS " & _
                "    WHERE ID_LOTE = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_lotes_probetas)", vbCritical, Err.Description
End Function
Public Function Listado(IDEN As String) As ADODB.Recordset
        Dim consulta As String
        'M1126-I
        Dim filtro As String
        filtro = ""
        If IDEN <> "" Then
            filtro = filtro & " AND IDENTIFICACION LIKE '%" & IDEN & "%'"
        End If
        'M1126-F
        consulta = "SELECT ID_LOTE,IDENTIFICACION,NUMERO_LOTE,NUMERO_INFORME,CARGA_ROTURA,MATERIAL,ESPESOR " & _
                   "  FROM CE_LOTES_PROBETAS " & _
                   " WHERE 1 = 1 " & _
                   filtro & _
                   " ORDER BY IDENTIFICACION"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_PK(PK As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_LOTE,IDENTIFICACION,NUMERO_LOTE,NUMERO_INFORME,CARGA_ROTURA,MATERIAL,ESPESOR " & _
                   "  FROM CE_LOTES_PROBETAS " & _
                   " WHERE ID_LOTE = " & PK & _
                   " ORDER BY IDENTIFICACION"
        Set Listado_PK = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_LOTE,IDENTIFICACION FROM CE_LOTES_PROBETAS ORDER BY IDENTIFICACION"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "CE_LOTES_PROBETAS"
        .setDESCRIPCION = "Lotes Probetas"
        .setPK = "ID_LOTE"
        .setCAMPO = "IDENTIFICACION"
        .setMUESTRA_DETALLE = True
        Set .FORMULARIO = FORMULARIO
    End With
End Function

