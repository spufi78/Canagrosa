VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEq_usos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQ_USOS
'***************************************************************
Private EQUIPO_ID As Long
Private MUESTRA_ID As Long
Private DETERMINACION_ID As Long
Private USOS As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQ_USOS
'***************************************************************
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setDETERMINACION_ID(ByVal dato As Long)
        DETERMINACION_ID = dato
End Property
Public Property Let setUSOS(ByVal dato As Long)
        USOS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQ_USOS
'***************************************************************
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getDETERMINACION_ID() As Long
        getDETERMINACION_ID = DETERMINACION_ID
End Property

Public Property Get getUSOS() As Long
        getUSOS = USOS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQ_USOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EQ_USOS WHERE EQUIPO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                EQUIPO_ID = rs("EQUIPO_ID")
                MUESTRA_ID = rs("MUESTRA_ID")
                DETERMINACION_ID = rs("DETERMINACION_ID")
                USOS = rs("USOS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEq_usos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "UPDATE EQUIPOS SET NUMERO_USOS_CONTADOR = NUMERO_USOS_CONTADOR + " & USOS & " WHERE ID_EQUIPO = " & EQUIPO_ID

        consulta = "INSERT INTO EQ_USOS " & _
                   "  VALUES (" & _
                        EQUIPO_ID & "," & MUESTRA_ID & "," & DETERMINACION_ID & "," & USOS & _
                ")"
        execute_bd consulta
        Insertar = EQUIPO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEq_usos)", vbCritical, Err.Description
End Function
Public Function Modificar(EQUIPO As Long, MUESTRA As Long, USOS As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EQ_USOS SET " & _
                        " USOS = " & USOS & "" & _
                  " WHERE EQUIPO_ID = " & EQUIPO & _
                  "   AND MUESTRA_ID = " & MUESTRA
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEq_usos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long, DETERMINACION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        'JGM-I
'        consulta = "SELECT * FROM EQ_USOS WHERE MUESTRA_ID = " & ID & " AND DETERMINACION_ID = " & DETERMINACION
'        Set rs = datos_bd(consulta)
'        If rs.RecordCount > 0 Then
'            Do
'                execute_bd "UPDATE EQUIPOS SET NUMERO_USOS_CONTADOR = NUMERO_USOS_CONTADOR - " & rs("USOS") & " WHERE ID_EQUIPO = " & rs("EQUIPO_ID")
'                rs.MoveNext
'            Loop Until rs.EOF
'        End If
        consulta = "UPDATE EQUIPOS AA, " & _
                   " (SELECT EQUIPO_ID AS C1, USOS AS C2 FROM EQ_USOS WHERE MUESTRA_ID = " & ID & " AND DETERMINACION_ID = " & DETERMINACION & ") BB  " & _
                   " Set AA.NUMERO_USOS_CONTADOR = IF(AA.NUMERO_USOS_CONTADOR - BB.c2 > 0, AA.NUMERO_USOS_CONTADOR - BB.c2,0) " & _
                   " where AA.ID_EQUIPO = BB.c1 "
        execute_bd consulta
        'JGM-F
        
        consulta = "DELETE FROM EQ_USOS WHERE MUESTRA_ID = " & ID & " AND DETERMINACION_ID = " & DETERMINACION
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEq_usos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EQ_USOS ORDER BY EQUIPO_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Numero_Total_Usos(EQUIPO As Long) As Long
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT SUM(USOS) FROM EQ_USOS WHERE EQUIPO_ID = " & EQUIPO
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            If IsNull(rs(0)) Then
                Numero_Total_Usos = 0
            Else
                Numero_Total_Usos = rs(0)
            End If
        Else
            Numero_Total_Usos = 0
        End If
        Set rs = Nothing
End Function
