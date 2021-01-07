VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_Normas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_NORMAS
'***************************************************************
Private NORMA_ID As Long
Private EQUIPO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_NORMAS
'***************************************************************
Public Property Let setNORMA_ID(ByVal dato As Long)
        NORMA_ID = dato
End Property
Public Property Let setEQUIPO_ID(ByVal dato As Long)
        EQUIPO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_NORMAS
'***************************************************************
Public Property Get getNORMA_ID() As Long
        getNORMA_ID = NORMA_ID
End Property
Public Property Get getEQUIPO_ID() As Long
        getEQUIPO_ID = EQUIPO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_NORMAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM EQUIPOS_NORMAS WHERE EQUIPO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Carga = False
        Else
            NORMA_ID = rs("NORMA_ID")
            EQUIPO_ID = rs("EQUIPO_ID")
            Carga = True
        End If
        Set rs = Nothing
        
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_Normas)", vbCritical, Err.Description
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        If Not duplicado Then
            consulta = "INSERT INTO EQUIPOS_NORMAS VALUES (" & _
                                    EQUIPO_ID & ", " & NORMA_ID & ", " & _
                                    "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP " & _
                       ")"
            execute_bd consulta
            Insertar = NORMA_ID ' Puede que haya que añadir el código del equipo (o crear un id nuevo para equipos_norma)
        End If
        
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos_Normas)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID_N As Long, ID_E As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "DELETE FROM EQUIPOS_NORMAS " & _
                   " WHERE NORMA_ID = " & ID_N & _
                   "   AND EQUIPO_ID = " & ID_E
        execute_bd consulta
        Eliminar = True
        
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos_Normas)", vbCritical, Err.Description
End Function

Public Function lista_por_equipos(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT CA_N.ID_NORMA, CA_N.CODIGO, CA_N.NOMBRE " & _
                   "  FROM EQUIPOS_NORMAS EN, CA_NORMAS CA_N" & _
                   " WHERE CA_N.ID_NORMA = EN.NORMA_ID " & _
                   "   AND EN.EQUIPO_ID = " & ID
        Set lista_por_equipos = datos_bd(consulta)
End Function

Public Function duplicado() As Boolean
    Dim rs As New ADODB.RecordSet
    Dim consulta As String
    
    consulta = "SELECT COUNT(*) AS TOTAL " & _
               "  FROM EQUIPOS_NORMAS " & _
               " WHERE NORMA_ID = " & NORMA_ID & _
               "   AND EQUIPO_ID = " & EQUIPO_ID & ""
    Set rs = datos_bd(consulta)
    If rs.Fields("TOTAL") = 0 Then
        duplicado = False
    Else
        MsgBox "Esta norma ya está asociada al equipo.", vbExclamation, App.Title
        duplicado = True
    End If
    Set rs = Nothing
End Function
