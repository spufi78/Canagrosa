VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_resultados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPLASMA_RESULTADOS
'***************************************************************
Private MUESTRA_ID As Long
Private tipo As Long
Private BATCH As String
Private MICROESTRUCTURA1 As Long
Private MICROESTRUCTURA2 As Long
Private MICROESTRUCTURA3 As Long
Private MICROESTRUCTURA4 As Long
Private MICROESTRUCTURA5 As Long
Private MICROESTRUCTURA6 As Long
Private MICROESTRUCTURA1_R As String
Private MICROESTRUCTURA2_R As String
Private MICROESTRUCTURA3_R As String
Private MICROESTRUCTURA4_R As String
Private MICROESTRUCTURA5_R As String
Private MICROESTRUCTURA6_R As String
Private MICROESTRUCTURA1_VALOR As String
Private MICROESTRUCTURA2_VALOR As String
Private MICROESTRUCTURA3_VALOR As String
Private MICROESTRUCTURA4_VALOR As String
Private MICROESTRUCTURA5_VALOR As String
Private MICROESTRUCTURA6_VALOR As String
Private TRACCION As String
Private MACRO_DUREZA As String
Private MACRO_DUREZA_DIMENSION As String
Private MACRO_DUREZA_ESPESOR As String
Private MICRO_DUREZA As String
Private ESPESOR As String
Private ESPESOR_MIN As String
Private ESPESOR_MAX As String
Private TRACCION_RES As String
Private MACRO_DUREZA_RES As String
Private MACRO_DUREZA_SD As String
Private MACRO_DUREZA_POR As String
Private MICRO_DUREZA_RES As String
Private MICRO_DUREZA_SD As String
Private MICRO_DUREZA_POR As String
Private ESPESOR_RES As String
Private ESPESOR_SD As String
Private ESPESOR_POR As String
Private TRACCION_PASS As Integer
Private MACRO_DUREZA_PASS As Integer
Private MICRO_DUREZA_PASS As Integer
Private ESPESOR_PASS As Integer


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPLASMA_RESULTADOS
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As Long)
        tipo = dato
End Property
Public Property Let setBATCH(ByVal dato As String)
        BATCH = dato
End Property
Public Property Let setMICROESTRUCTURA1(ByVal dato As Long)
        MICROESTRUCTURA1 = dato
End Property
Public Property Let setMICROESTRUCTURA2(ByVal dato As Long)
        MICROESTRUCTURA2 = dato
End Property
Public Property Let setMICROESTRUCTURA3(ByVal dato As Long)
        MICROESTRUCTURA3 = dato
End Property
Public Property Let setMICROESTRUCTURA4(ByVal dato As Long)
        MICROESTRUCTURA4 = dato
End Property
Public Property Let setMICROESTRUCTURA5(ByVal dato As Long)
        MICROESTRUCTURA5 = dato
End Property
Public Property Let setMICROESTRUCTURA6(ByVal dato As Long)
        MICROESTRUCTURA6 = dato
End Property
Public Property Let setMICROESTRUCTURA1_R(ByVal dato As String)
        MICROESTRUCTURA1_R = dato
End Property
Public Property Let setMICROESTRUCTURA2_R(ByVal dato As String)
        MICROESTRUCTURA2_R = dato
End Property
Public Property Let setMICROESTRUCTURA3_R(ByVal dato As String)
        MICROESTRUCTURA3_R = dato
End Property
Public Property Let setMICROESTRUCTURA4_R(ByVal dato As String)
        MICROESTRUCTURA4_R = dato
End Property
Public Property Let setMICROESTRUCTURA5_R(ByVal dato As String)
        MICROESTRUCTURA5_R = dato
End Property
Public Property Let setMICROESTRUCTURA6_R(ByVal dato As String)
        MICROESTRUCTURA6_R = dato
End Property
Public Property Let setMICROESTRUCTURA1_VALOR(ByVal dato As String)
        MICROESTRUCTURA1_VALOR = dato
End Property
Public Property Let setMICROESTRUCTURA2_VALOR(ByVal dato As String)
        MICROESTRUCTURA2_VALOR = dato
End Property
Public Property Let setMICROESTRUCTURA3_VALOR(ByVal dato As String)
        MICROESTRUCTURA3_VALOR = dato
End Property
Public Property Let setMICROESTRUCTURA4_VALOR(ByVal dato As String)
        MICROESTRUCTURA4_VALOR = dato
End Property
Public Property Let setMICROESTRUCTURA5_VALOR(ByVal dato As String)
        MICROESTRUCTURA5_VALOR = dato
End Property
Public Property Let setMICROESTRUCTURA6_VALOR(ByVal dato As String)
        MICROESTRUCTURA6_VALOR = dato
End Property
Public Property Let setTRACCION(ByVal dato As String)
        TRACCION = dato
End Property
Public Property Let setMACRO_DUREZA(ByVal dato As String)
        MACRO_DUREZA = dato
End Property
Public Property Let setMACRO_DUREZA_DIMENSION(ByVal dato As String)
        MACRO_DUREZA_DIMENSION = dato
End Property
Public Property Let setMACRO_DUREZA_ESPESOR(ByVal dato As String)
        MACRO_DUREZA_ESPESOR = dato
End Property
Public Property Let setMICRO_DUREZA(ByVal dato As String)
        MICRO_DUREZA = dato
End Property
Public Property Let setESPESOR(ByVal dato As String)
        ESPESOR = dato
End Property
Public Property Let setESPESOR_MIN(ByVal dato As String)
        ESPESOR_MIN = dato
End Property
Public Property Let setESPESOR_MAX(ByVal dato As String)
        ESPESOR_MAX = dato
End Property
Public Property Let setTRACCION_RES(ByVal dato As String)
        TRACCION_RES = dato
End Property
Public Property Let setMACRO_DUREZA_RES(ByVal dato As String)
        MACRO_DUREZA_RES = dato
End Property
Public Property Let setMACRO_DUREZA_SD(ByVal dato As String)
        MACRO_DUREZA_SD = dato
End Property
Public Property Let setMACRO_DUREZA_POR(ByVal dato As String)
        MACRO_DUREZA_POR = dato
End Property
Public Property Let setMICRO_DUREZA_RES(ByVal dato As String)
        MICRO_DUREZA_RES = dato
End Property
Public Property Let setMICRO_DUREZA_SD(ByVal dato As String)
        MICRO_DUREZA_SD = dato
End Property
Public Property Let setMICRO_DUREZA_POR(ByVal dato As String)
        MICRO_DUREZA_POR = dato
End Property
Public Property Let setESPESOR_RES(ByVal dato As String)
        ESPESOR_RES = dato
End Property
Public Property Let setESPESOR_SD(ByVal dato As String)
        ESPESOR_SD = dato
End Property
Public Property Let setESPESOR_POR(ByVal dato As String)
        ESPESOR_POR = dato
End Property
Public Property Let setTRACCION_PASS(ByVal dato As Integer)
        TRACCION_PASS = dato
End Property
Public Property Let setMACRO_DUREZA_PASS(ByVal dato As Integer)
        MACRO_DUREZA_PASS = dato
End Property
Public Property Let setMICRO_DUREZA_PASS(ByVal dato As Integer)
        MICRO_DUREZA_PASS = dato
End Property
Public Property Let setESPEROR_PASS(ByVal dato As Integer)
        ESPESOR_PASS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPLASMA_RESULTADOS
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getTIPO() As Long
        getTIPO = tipo
End Property
Public Property Get getBATCH() As String
        getBATCH = BATCH
End Property
Public Property Get getMICROESTRUCTURA1() As Long
        getMICROESTRUCTURA1 = MICROESTRUCTURA1
End Property
Public Property Get getMICROESTRUCTURA2() As Long
        getMICROESTRUCTURA2 = MICROESTRUCTURA2
End Property
Public Property Get getMICROESTRUCTURA3() As Long
        getMICROESTRUCTURA3 = MICROESTRUCTURA3
End Property
Public Property Get getMICROESTRUCTURA4() As Long
        getMICROESTRUCTURA4 = MICROESTRUCTURA4
End Property
Public Property Get getMICROESTRUCTURA5() As Long
        getMICROESTRUCTURA5 = MICROESTRUCTURA5
End Property
Public Property Get getMICROESTRUCTURA6() As Long
        getMICROESTRUCTURA6 = MICROESTRUCTURA6
End Property
Public Property Get getMICROESTRUCTURA1_R() As String
        getMICROESTRUCTURA1_R = MICROESTRUCTURA1_R
End Property
Public Property Get getMICROESTRUCTURA2_R() As String
        getMICROESTRUCTURA2_R = MICROESTRUCTURA2_R
End Property
Public Property Get getMICROESTRUCTURA3_R() As String
        getMICROESTRUCTURA3_R = MICROESTRUCTURA3_R
End Property
Public Property Get getMICROESTRUCTURA4_R() As String
        getMICROESTRUCTURA4_R = MICROESTRUCTURA4_R
End Property
Public Property Get getMICROESTRUCTURA5_R() As String
        getMICROESTRUCTURA5_R = MICROESTRUCTURA5_R
End Property
Public Property Get getMICROESTRUCTURA6_R() As String
        getMICROESTRUCTURA6_R = MICROESTRUCTURA6_R
End Property
Public Property Get getMICROESTRUCTURA1_VALOR() As String
        getMICROESTRUCTURA1_VALOR = MICROESTRUCTURA1_VALOR
End Property
Public Property Get getMICROESTRUCTURA2_VALOR() As String
        getMICROESTRUCTURA2_VALOR = MICROESTRUCTURA2_VALOR
End Property
Public Property Get getMICROESTRUCTURA3_VALOR() As String
        getMICROESTRUCTURA3_VALOR = MICROESTRUCTURA3_VALOR
End Property
Public Property Get getMICROESTRUCTURA4_VALOR() As String
        getMICROESTRUCTURA4_VALOR = MICROESTRUCTURA4_VALOR
End Property
Public Property Get getMICROESTRUCTURA5_VALOR() As String
        getMICROESTRUCTURA5_VALOR = MICROESTRUCTURA5_VALOR
End Property
Public Property Get getMICROESTRUCTURA6_VALOR() As String
        getMICROESTRUCTURA6_VALOR = MICROESTRUCTURA6_VALOR
End Property
Public Property Get getTRACCION() As String
        getTRACCION = TRACCION
End Property
Public Property Get getMACRO_DUREZA() As String
        getMACRO_DUREZA = MACRO_DUREZA
End Property
Public Property Get getMACRO_DUREZA_DIMENSION() As String
        getMACRO_DUREZA_DIMENSION = MACRO_DUREZA_DIMENSION
End Property
Public Property Get getMACRO_DUREZA_ESPESOR() As String
        getMACRO_DUREZA_ESPESOR = MACRO_DUREZA_ESPESOR
End Property
Public Property Get getMICRO_DUREZA() As String
        getMICRO_DUREZA = MICRO_DUREZA
End Property
Public Property Get getESPESOR() As String
        getESPESOR = ESPESOR
End Property
Public Property Get getESPESOR_MIN() As String
        getESPESOR_MIN = ESPESOR_MIN
End Property
Public Property Get getESPESOR_MAX() As String
        getESPESOR_MAX = ESPESOR_MAX
End Property
Public Property Get getTRACCION_RES() As String
        getTRACCION_RES = TRACCION_RES
End Property
Public Property Get getMACRO_DUREZA_RES() As String
        getMACRO_DUREZA_RES = MACRO_DUREZA_RES
End Property
Public Property Get getMACRO_DUREZA_SD() As String
        getMACRO_DUREZA_SD = MACRO_DUREZA_SD
End Property
Public Property Get getMACRO_DUREZA_POR() As String
        getMACRO_DUREZA_POR = MACRO_DUREZA_POR
End Property
Public Property Get getMICRO_DUREZA_RES() As String
        getMICRO_DUREZA_RES = MICRO_DUREZA_RES
End Property
Public Property Get getMICRO_DUREZA_SD() As String
        getMICRO_DUREZA_SD = MICRO_DUREZA_SD
End Property
Public Property Get getMICRO_DUREZA_POR() As String
        getMICRO_DUREZA_POR = MICRO_DUREZA_POR
End Property
Public Property Get getESPESOR_RES() As String
        getESPESOR_RES = ESPESOR_RES
End Property
Public Property Get getESPESOR_SD() As String
        getESPESOR_SD = ESPESOR_SD
End Property
Public Property Get getESPESOR_POR() As String
        getESPESOR_POR = ESPESOR_POR
End Property

Public Property Get getTRACCION_PASS() As Integer
        getTRACCION_PASS = TRACCION_PASS
End Property
Public Property Get getMACRO_DUREZA_PASS() As Integer
        getMACRO_DUREZA_PASS = MACRO_DUREZA_PASS
End Property
Public Property Get getMICRO_DUREZA_PASS() As Integer
        getMICRO_DUREZA_PASS = MICRO_DUREZA_PASS
End Property
Public Property Get getESPESOR_PASS() As Integer
        getESPESOR_PASS = ESPESOR_PASS
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPLASMA_RESULTADOS
'***************************************************************
Public Function Carga(ID As Long, tipo As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_RESULTADOS " & _
                   " WHERE MUESTRA_ID = " & ID & _
                   "   AND TIPO = " & tipo
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                tipo = rs("TIPO")
                BATCH = rs("BATCH")
                MICROESTRUCTURA1 = rs("MICROESTRUCTURA1")
                MICROESTRUCTURA2 = rs("MICROESTRUCTURA2")
                MICROESTRUCTURA3 = rs("MICROESTRUCTURA3")
                MICROESTRUCTURA4 = rs("MICROESTRUCTURA4")
                MICROESTRUCTURA5 = rs("MICROESTRUCTURA5")
                MICROESTRUCTURA6 = rs("MICROESTRUCTURA6")
                MICROESTRUCTURA1_R = rs("MICROESTRUCTURA1_R")
                MICROESTRUCTURA2_R = rs("MICROESTRUCTURA2_R")
                MICROESTRUCTURA3_R = rs("MICROESTRUCTURA3_R")
                MICROESTRUCTURA4_R = rs("MICROESTRUCTURA4_R")
                MICROESTRUCTURA5_R = rs("MICROESTRUCTURA5_R")
                MICROESTRUCTURA6_R = rs("MICROESTRUCTURA6_R")
                MICROESTRUCTURA1_VALOR = rs("MICROESTRUCTURA1_VALOR")
                MICROESTRUCTURA2_VALOR = rs("MICROESTRUCTURA2_VALOR")
                MICROESTRUCTURA3_VALOR = rs("MICROESTRUCTURA3_VALOR")
                MICROESTRUCTURA4_VALOR = rs("MICROESTRUCTURA4_VALOR")
                MICROESTRUCTURA5_VALOR = rs("MICROESTRUCTURA5_VALOR")
                MICROESTRUCTURA6_VALOR = rs("MICROESTRUCTURA6_VALOR")
                TRACCION = rs("TRACCION")
                MACRO_DUREZA = rs("MACRO_DUREZA")
                MACRO_DUREZA_DIMENSION = rs("MACRO_DUREZA_DIMENSION")
                MACRO_DUREZA_ESPESOR = rs("MACRO_DUREZA_ESPESOR")
                MICRO_DUREZA = rs("MICRO_DUREZA")
                ESPESOR = rs("ESPESOR")
                ESPESOR_MIN = rs("ESPESOR_MIN")
                ESPESOR_MAX = rs("ESPESOR_MAX")
                TRACCION_RES = rs("TRACCION_RES")
                MACRO_DUREZA_RES = rs("MACRO_DUREZA_RES")
                MACRO_DUREZA_SD = rs("MACRO_DUREZA_SD")
                MACRO_DUREZA_POR = rs("MACRO_DUREZA_POR")
                MICRO_DUREZA_RES = rs("MICRO_DUREZA_RES")
                MICRO_DUREZA_SD = rs("MICRO_DUREZA_SD")
                MICRO_DUREZA_POR = rs("MICRO_DUREZA_POR")
                ESPESOR_RES = rs("ESPESOR_RES")
                ESPESOR_SD = rs("ESPESOR_SD")
                ESPESOR_POR = rs("ESPESOR_POR")
                TRACCION_PASS = rs("TRACCION_PASS")
                MACRO_DUREZA_PASS = rs("MACRO_DUREZA_PASS")
                MICRO_DUREZA_PASS = rs("MICRO_DUREZA_PASS")
                ESPESOR_PASS = rs("ESPESOR_PASS")
                
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPlasma_resultados)", vbCritical, Err.Description
End Function
Public Function actualizarTraccion(MUESTRA_ID As Long, tipo As Integer, TRACCION As String, TRACCION_RES As String) As Boolean
    Dim consulta As String
    If Me.Carga(MUESTRA_ID, tipo) = True Then
        consulta = "UPDATE PLASMA_RESULTADOS " & _
                   "   SET TRACCION = '" & TRACCION & "'" & _
                   "      ,TRACCION_RES = '" & TRACCION_RES & "'" & _
                   " WHERE MUESTRA_ID = " & MUESTRA_ID & _
                   "   AND TIPO = " & tipo
        execute_bd consulta
    Else
        With Me
         .setMUESTRA_ID = MUESTRA_ID
         .setTIPO = tipo
         .setTRACCION = TRACCION
         .setTRACCION_RES = TRACCION_RES
         .setMICROESTRUCTURA1 = 0
         .setMICROESTRUCTURA2 = 0
         .setMICROESTRUCTURA3 = 0
         .setMICROESTRUCTURA4 = 0
         .setMICROESTRUCTURA5 = 0
         .setMICROESTRUCTURA6 = 0
         .setTRACCION_PASS = 0
         .setMACRO_DUREZA_PASS = 0
         .setMICRO_DUREZA_PASS = 0
         .setESPEROR_PASS = 0
         .Insertar
        End With
    End If
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "REPLACE INTO PLASMA_RESULTADOS " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & tipo & "," & "'" & BATCH & "'" & "," & _
                        MICROESTRUCTURA1 & "," & MICROESTRUCTURA2 & "," & MICROESTRUCTURA3 & "," & _
                        MICROESTRUCTURA4 & "," & MICROESTRUCTURA5 & "," & MICROESTRUCTURA6 & "," & _
                        "'" & MICROESTRUCTURA1_R & "','" & MICROESTRUCTURA2_R & "','" & MICROESTRUCTURA3_R & "'," & _
                        "'" & MICROESTRUCTURA4_R & "','" & MICROESTRUCTURA5_R & "','" & MICROESTRUCTURA6_R & "'," & _
                        "'" & MICROESTRUCTURA1_VALOR & "','" & MICROESTRUCTURA2_VALOR & "','" & MICROESTRUCTURA3_VALOR & "'," & _
                        "'" & MICROESTRUCTURA4_VALOR & "','" & MICROESTRUCTURA5_VALOR & "','" & MICROESTRUCTURA6_VALOR & "'," & _
                        "'" & TRACCION & "'" & "," & "'" & MACRO_DUREZA & "','" & MACRO_DUREZA_DIMENSION & "','" & MACRO_DUREZA_ESPESOR & "','" & MICRO_DUREZA & "','" & ESPESOR & "','" & ESPESOR_MIN & "','" & ESPESOR_MAX & "'," & _
                        "'" & TRACCION_RES & "'," & _
                        "'" & MACRO_DUREZA_RES & "','" & MACRO_DUREZA_SD & "','" & MACRO_DUREZA_POR & "'," & _
                        "'" & MICRO_DUREZA_RES & "','" & MICRO_DUREZA_SD & "','" & MICRO_DUREZA_POR & "'," & _
                        "'" & ESPESOR_RES & "','" & ESPESOR_SD & "','" & ESPESOR_POR & "'," & _
                        TRACCION_PASS & "," & MACRO_DUREZA_PASS & "," & MICRO_DUREZA_PASS & "," & ESPESOR_PASS & _
                ")"
'                " ON DUPLICATE KEY UPDATE "
'                        " BATCH = '" & BATCH & "',"
'                        " MICROESTRUCTURA1 = " & MICROESTRUCTURA1 & "," & " MICROESTRUCTURA2 = " & MICROESTRUCTURA2 & "," & " MICROESTRUCTURA3 = " & MICROESTRUCTURA3 & ","
'                        " MICROESTRUCTURA4 = " & MICROESTRUCTURA4 & "," & " MICROESTRUCTURA5 = " & MICROESTRUCTURA5 & "," & " MICROESTRUCTURA6 = " & MICROESTRUCTURA6 & ","
'                        " MICROESTRUCTURA1_R = '" & MICROESTRUCTURA1_R & "'," & " MICROESTRUCTURA2_R = '" & MICROESTRUCTURA2_R & "'," & " MICROESTRUCTURA3_R = '" & MICROESTRUCTURA3_R & "',"
'                        " MICROESTRUCTURA4_R = '" & MICROESTRUCTURA4_R & "'," & " MICROESTRUCTURA5_R = '" & MICROESTRUCTURA5_R & "'," & " MICROESTRUCTURA6_R = '" & MICROESTRUCTURA6_R & "',"
'                        " TRACCION = '" & TRACCION & "'," & " MACRO_DUREZA = '" & MACRO_DUREZA & "'," & " MICRO_DUREZA = '" & MICRO_DUREZA & "'," & " ESPESOR_MIN = '" & ESPESOR_MIN & "',"
'                        " ESPESOR_MAX = '" & ESPESOR_MAX & "'," & " TRACCION_RES = '" & TRACCION_RES & "'," & " MACRO_DUREZA_RES = '" & MACRO_DUREZA_RES & "'," & " MICRO_DUREZA_RES = '" & MICRO_DUREZA_RES & "',"
'                        " TRACCION_PASS = " & TRACCION_PASS & "," & "MACRO_DUREZA_PASS = " & MACRO_DUREZA_PASS & "," & "MICRO_DUREZA_PASS = " & MICRO_DUREZA_PASS & "," & "ESPESOR_PASS = " & ESPESOR_PASS
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlasma_resultados)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PLASMA_RESULTADOS " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlasma_resultados)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_RESULTADOS ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPLASMA_RESULTADOS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
         consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function
