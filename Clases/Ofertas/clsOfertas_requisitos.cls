VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOfertas_requisitos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOFERTAS_REQUISITOS
'***************************************************************
Private OFERTA_ID As Long
Private texto As String
Private sn As String
Private acreditaciones As String
Private aprobaciones As String
Private USUARIO_ID As Long
Private TS As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOFERTAS_REQUISITOS
'***************************************************************
Public Property Let setOFERTA_ID(ByVal dato As Long)
        OFERTA_ID = dato
End Property
Public Property Let setTEXTO(ByVal dato As String)
        texto = dato
End Property
Public Property Let setSN(ByVal dato As String)
        sn = dato
End Property
Public Property Let setACREDITACIONES(ByVal dato As String)
        acreditaciones = dato
End Property
Public Property Let setAPROBACIONES(ByVal dato As String)
        aprobaciones = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOFERTAS_REQUISITOS
'***************************************************************
Public Property Get getOFERTA_ID() As Long
        getOFERTA_ID = OFERTA_ID
End Property
Public Property Get getTEXTO() As String
        getTEXTO = texto
End Property
Public Property Get getSN() As String
        getSN = sn
End Property
Public Property Get getACREDITACIONES() As String
        getACREDITACIONES = acreditaciones
End Property
Public Property Get getAPROBACIONES() As String
        getAPROBACIONES = aprobaciones
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTS() As String
        getTS = TS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOFERTAS_REQUISITOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OFERTAS_REQUISITOS WHERE OFERTA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                OFERTA_ID = rs("OFERTA_ID")
                texto = rs("TEXTO")
                sn = rs("SN")
                acreditaciones = rs("ACREDITACIONES")
                aprobaciones = rs("APROBACIONES")
                USUARIO_ID = rs("USUARIO_ID")
                TS = rs("TS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsOfertas_requisitos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO OFERTAS_REQUISITOS " & _
                   "  VALUES (" & _
                        OFERTA_ID & "," & "'" & texto & "'" & "," & "'" & sn & "'" & "," & _
                        "'" & acreditaciones & "'" & "," & "'" & aprobaciones & "'," & USUARIO.getID_EMPLEADO & ",CURRENT_TIMESTAMP" & _
                ") ON DUPLICATE KEY UPDATE " & _
                " TEXTO = '" & texto & "'," & _
                " SN = '" & sn & "'," & _
                " ACREDITACIONES = '" & acreditaciones & "'," & _
                " APROBACIONES = '" & aprobaciones & "'"
        execute_bd consulta
        Insertar = OFERTA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsOfertas_requisitos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OFERTAS_REQUISITOS SET " & _
                        " TEXTO = '" & texto & "'," & _
                        " SN = '" & sn & "'," & _
                        " ACREDITACIONES = '" & acreditaciones & "'," & _
                        " APROBACIONES = '" & aprobaciones & "'" & _
                " WHERE OFERTA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsOfertas_requisitos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM OFERTAS_REQUISITOS " & _
                "    WHERE OFERTA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsOfertas_requisitos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OFERTAS_REQUISITOS ORDER BY OFERTA_ID"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSOFERTAS_REQUISITOS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

