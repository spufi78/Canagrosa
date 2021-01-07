VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_normas_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_NORMAS_HISTORICO
'***************************************************************
Private NORMA_ID As Long
Private fecha As String
Private MOTIVO As String
Private RUTA As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_NORMAS_HISTORICO
'***************************************************************
Public Property Let setNORMA_ID(ByVal dato As Long)
        NORMA_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setMOTIVO(ByVal dato As String)
        MOTIVO = dato
End Property
Public Property Let setRUTA(ByVal dato As String)
        RUTA = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_NORMAS_HISTORICO
'***************************************************************
Public Property Get getNORMA_ID() As Long
        getNORMA_ID = NORMA_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getMOTIVO() As String
        getMOTIVO = MOTIVO
End Property
Public Property Get getRUTA() As String
        getRUTA = RUTA
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_NORMAS_HISTORICO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CA_NORMAS_HISTORICO WHERE NORMA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                NORMA_ID = rs("NORMA_ID")
                fecha = rs("FECHA")
                MOTIVO = rs("MOTIVO")
                RUTA = rs("RUTA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_normas_historico)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CA_NORMAS_HISTORICO " & _
                   "  VALUES (" & _
                        NORMA_ID & "," & "'" & fecha & "'" & "," & "'" & MOTIVO & "'" & "," & _
                        "'" & RUTA & "'" & _
                ")"
        execute_bd consulta
        Insertar = NORMA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_normas_historico)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_NORMAS_HISTORICO SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " MOTIVO = '" & MOTIVO & "'," & _
                        " RUTA = '" & RUTA & "'" & _
                " WHERE NORMA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_normas_historico)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CA_NORMAS_HISTORICO " & _
                "    WHERE NORMA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_normas_historico)", vbCritical, Err.Description
End Function
Public Function Listado(NORMA As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT FECHA,MOTIVO,RUTA FROM CA_NORMAS_HISTORICO WHERE NORMA_ID = " & NORMA & " ORDER BY FECHA DESC"
        Set Listado = datos_bd(consulta)
End Function
