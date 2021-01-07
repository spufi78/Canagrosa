VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocumentos_anulados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDOCUMENTOS_ANULADOS
'***************************************************************
Private DOCUMENTO_ID As Long
Private USUARIO_ID As Long
Private fecha As String
Private hora As String
Private MOTIVO As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDOCUMENTOS_ANULADOS
'***************************************************************
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setHORA(ByVal dato As String)
        hora = dato
End Property
Public Property Let setMOTIVO(ByVal dato As String)
        MOTIVO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDOCUMENTOS_ANULADOS
'***************************************************************
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getHORA() As String
        getHORA = hora
End Property
Public Property Get getMOTIVO() As String
        getMOTIVO = MOTIVO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDOCUMENTOS_ANULADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS_ANULADOS WHERE DOCUMENTO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                USUARIO_ID = rs("USUARIO_ID")
                fecha = rs("FECHA")
                hora = rs("HORA")
                MOTIVO = rs("MOTIVO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDocumentos_anulados)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT DOCUMENTO_ID FROM DOCUMENTOS_ANULADOS WHERE DOCUMENTO_ID = " & DOCUMENTO_ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount > 0 Then
            Me.Modificar (DOCUMENTO_ID)
        Else
            consulta = "INSERT INTO DOCUMENTOS_ANULADOS " & _
                       "  VALUES (" & _
                            DOCUMENTO_ID & "," & USUARIO_ID & "," & "'" & fecha & "'" & "," & _
                            "'" & hora & "'" & "," & "'" & MOTIVO & "'" & _
                    ")"
            execute_bd consulta
        End If
        Dim oDOC As New clsDocumentos
        oDOC.Anular DOCUMENTO_ID
        Set oDOC = Nothing
        Insertar = DOCUMENTO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDocumentos_anulados)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS_ANULADOS SET " & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " HORA = '" & hora & "'," & _
                        " MOTIVO = '" & MOTIVO & "'" & _
                " WHERE DOCUMENTO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDocumentos_anulados)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DOCUMENTOS_ANULADOS " & _
                "    WHERE DOCUMENTO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDocumentos_anulados)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS_ANULADOS ORDER BY DOCUMENTO_ID"
        Set Listado = datos_bd(consulta)
End Function
