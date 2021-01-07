VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_documentos_val"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_DOCUMENTOS_VAL
'***************************************************************
Private DOCUMENTO_ID As Long
Private USUARIO_ID As Long
Private TS As String
Private P1 As Long
Private P2 As Long
Private P3 As Long
Private COMENTARIOS As String
Private CUMPLIMENTADO As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_DOCUMENTOS_VAL
'***************************************************************
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTS(ByVal dato As String)
        TS = dato
End Property
Public Property Let setP1(ByVal dato As Long)
        P1 = dato
End Property
Public Property Let setP2(ByVal dato As Long)
        P2 = dato
End Property
Public Property Let setP3(ByVal dato As Long)
        P3 = dato
End Property
Public Property Let setCOMENTARIOS(ByVal dato As String)
        COMENTARIOS = dato
End Property
Public Property Let setCUMPLIMENTADO(ByVal dato As Integer)
        CUMPLIMENTADO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_DOCUMENTOS_VAL
'***************************************************************
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTS() As String
        getTS = TS
End Property
Public Property Get getP1() As Long
        getP1 = P1
End Property
Public Property Get getP2() As Long
        getP2 = P2
End Property
Public Property Get getP3() As Long
        getP3 = P3
End Property
Public Property Get getCOMENTARIOS() As String
        getCOMENTARIOS = COMENTARIOS
End Property
Public Property Get getCUMPLIMENTADO() As Integer
        getCUMPLIMENTADO = CUMPLIMENTADO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_DOCUMENTOS_VAL
'***************************************************************
Public Function Carga(ID As Long, USUARIO_ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CA_DOCUMENTOS_VAL WHERE DOCUMENTO_ID = " & ID
        If USUARIO_ID = 0 Then
           consulta = consulta & " AND USUARIO_ID = " & USUARIO.getID_EMPLEADO
        Else
           consulta = consulta & " AND USUARIO_ID = " & USUARIO_ID
        End If
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                USUARIO_ID = rs("USUARIO_ID")
                TS = rs("TS")
                P1 = rs("P1")
                P2 = rs("P2")
                P3 = rs("P3")
                CUMPLIMENTADO = rs("CUMPLIMENTADO")
                COMENTARIOS = rs("COMENTARIOS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_documentos_val)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CA_DOCUMENTOS_VAL " & _
                   "  VALUES (" & _
                        DOCUMENTO_ID & "," & USUARIO_ID & ",CURRENT_TIMESTAMP," & _
                        P1 & "," & P2 & "," & P3 & "," & _
                        "'" & COMENTARIOS & "',0" & _
                ")"
        execute_bd consulta
        Insertar = DOCUMENTO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_documentos_val)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_DOCUMENTOS_VAL SET " & _
                        " TS = CURRENT_TIMESTAMP," & _
                        " P1 = " & P1 & "," & _
                        " P2 = " & P2 & "," & _
                        " P3 = " & P3 & "," & _
                        " COMENTARIOS = '" & COMENTARIOS & "'," & _
                        " CUMPLIMENTADO = 1 " & _
                " WHERE DOCUMENTO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_documentos_val)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CA_DOCUMENTOS_VAL " & _
                "    WHERE DOCUMENTO_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_documentos_val)", vbCritical, Err.Description
End Function
Public Function Listado(ID_DOCUMENTO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.USUARIO_ID,A.TS,CONCAT(B.NOMBRE,' ',B.APELLIDOS),A.P1,A.P2,A.P3,A.COMENTARIOS " & _
                   "  FROM CA_DOCUMENTOS_VAL A " & _
                   "  LEFT JOIN USUARIOS B ON A.USUARIO_ID = B.ID_EMPLEADO " & _
                   " WHERE A.DOCUMENTO_ID = " & ID_DOCUMENTO & _
                   "   AND A.CUMPLIMENTADO = 1 " & _
                   " ORDER BY TS"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSCA_DOCUMENTOS_VAL
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function
Public Function verificarExisteEvaluacion(ID_DOCUMENTO As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CA_DOCUMENTOS_VAL WHERE DOCUMENTO_ID = " & ID_DOCUMENTO & " AND USUARIO_ID = " & USUARIO.getID_EMPLEADO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            ' No existe, generamos el registro y la firma para el usuario
            With Me
                .setDOCUMENTO_ID = ID_DOCUMENTO
                .setUSUARIO_ID = USUARIO.getID_EMPLEADO
                .setP1 = 0
                .setP2 = 0
                .setP3 = 0
                .Insertar
            End With
            ' Firma
            Dim oFirma As New clsFirmas
            With oFirma
                .setTOBJETO = TOBJETO.TOBJETO_DOCUMENTO_VALORACION
                .setCOBJETO = ID_DOCUMENTO
                Dim oempleado As New clsEmpleados
                oempleado.CARGAR_POR_USUARIO USUARIO.getID_EMPLEADO
                .setUSUARIO_ID = oempleado.getID_EMPLEADO
                .setFIRMADA = 0
                Dim oCD As New clsCa_documentos
                oCD.Carga ID_DOCUMENTO
                .setCOMENTARIO = "Cuestionario sobre el procedimiento : " & oCD.getNOMBRE
                .InsertarUsuario
                Set oCD = Nothing
            End With
            Set oFirma = Nothing
        End If
        verificarExisteEvaluacion = True
        Set rs = Nothing
        Exit Function
fallo:
        verificarExisteEvaluacion = False
        MsgBox "Error al cargar los datos(clsCa_documentos_val)", vbCritical, Err.Description

End Function
