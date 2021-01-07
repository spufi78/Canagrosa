VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormacion_certificados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSFORMACION_CERTIFICADOS
'***************************************************************
Private ID_FORMACION_CERTIFICADO As Long
Private CUALIFICACION_ID As Long
Private COD_CERTIFICADO As Long
Private ANYO As Long
Private DOCUMENTO_ID As Long
Private CERTIFICADOR_ID As Long
Private BASE As String
Private CURSO_ID As Long
Private FECHA_CERTIFICACION As String
Private CUSERID As Long
Private FTIMESTP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSFORMACION_CERTIFICADOS
'***************************************************************
Public Property Let setID_FORMACION_CERTIFICADO(ByVal dato As Long)
        ID_FORMACION_CERTIFICADO = dato
End Property
Public Property Let setCUALIFICACION_ID(ByVal dato As Long)
        CUALIFICACION_ID = dato
End Property
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setANYO(ByVal dato As Long)
        ANYO = dato
End Property
Public Property Let setCOD_CERTIFICADO(ByVal dato As Long)
        COD_CERTIFICADO = dato
End Property
Public Property Let setCERTIFICADOR_ID(ByVal dato As Long)
        CERTIFICADOR_ID = dato
End Property
Public Property Let setBASE(ByVal dato As String)
        BASE = dato
End Property
Public Property Let setCURSO_ID(ByVal dato As Long)
        CURSO_ID = dato
End Property
Public Property Let setFECHA_CERTIFICACION(ByVal dato As String)
        FECHA_CERTIFICACION = dato
End Property
Public Property Let setCUSERID(ByVal dato As Long)
        CUSERID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSFORMACION_CERTIFICADOS
'***************************************************************
Public Property Get getID_FORMACION_CERTIFICADO() As Long
        getID_FORMACION_CERTIFICADO = ID_FORMACION_CERTIFICADO
End Property
Public Property Get getCUALIFICACION_ID() As Long
        getCUALIFICACION_ID = CUALIFICACION_ID
End Property
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getANYO() As Long
        getANYO = ANYO
End Property
Public Property Get getCOD_CERTIFICADO() As Long
        getCOD_CERTIFICADO = COD_CERTIFICADO
End Property
Public Property Get getCERTIFICADOR_ID() As Long
        getCERTIFICADOR_ID = CERTIFICADOR_ID
End Property
Public Property Get getBASE() As String
        getBASE = BASE
End Property
Public Property Get getCURSO_ID() As Long
        getCURSO_ID = CURSO_ID
End Property
Public Property Get getFECHA_CERTIFICACION() As String
        getFECHA_CERTIFICACION = FECHA_CERTIFICACION
End Property
Public Property Get getCUSERID() As Long
        getCUSERID = CUSERID
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSFORMACION_CERTIFICADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_CERTIFICADOS WHERE ID_FORMACION_CERTIFICADO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_FORMACION_CERTIFICADO = rs("ID_FORMACION_CERTIFICADO")
                CUALIFICACION_ID = rs("CUALIFICACION_ID")
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                ANYO = rs("ANYO")
                COD_CERTIFICADO = rs("COD_CERTIFICADO")
                CERTIFICADOR_ID = rs("CERTIFICADOR_ID")
                BASE = rs("BASE")
                CURSO_ID = rs("CURSO_ID")
                FECHA_CERTIFICACION = rs("FECHA_CERTIFICACION")
                CUSERID = rs("CUSERID")
                FTIMESTP = rs("FTIMESTP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsFormacion_certificados)", vbCritical, Err.Description
End Function
Public Function CargaCualificacion(ID As Long) As Long
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_CERTIFICADOS WHERE CUALIFICACION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                CargaCualificacion = 0
        Else
                CargaCualificacion = rs("ID_FORMACION_CERTIFICADO")
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaCualificacion = 0
        MsgBox "Error al cargar los datos(clsFormacion_certificados)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_FORMACION_CERTIFICADO) FROM FORMACION_CERTIFICADOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_FORMACION_CERTIFICADO = 1
        Else
                ID_FORMACION_CERTIFICADO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function CrearCodigo()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(COD_CERTIFICADO) FROM FORMACION_CERTIFICADOS WHERE ANYO=" & CLng(Format(Date, "yyyy"))
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                COD_CERTIFICADO = 1
        Else
                COD_CERTIFICADO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        Me.CrearCodigo
        consulta = "INSERT INTO FORMACION_CERTIFICADOS " & _
                   "  VALUES (" & _
                        ID_FORMACION_CERTIFICADO & "," & CUALIFICACION_ID & "," & COD_CERTIFICADO & "," & _
                        ANYO & "," & DOCUMENTO_ID & "," & CERTIFICADOR_ID & "," & _
                        "'" & BASE & "'" & "," & CURSO_ID & "," & "'" & FECHA_CERTIFICACION & "'," & _
                        CUSERID & "," & "CURRENT_TIMESTAMP" & _
                ")"
        execute_bd consulta
        Insertar = ID_FORMACION_CERTIFICADO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsFormacion_certificados)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE FORMACION_CERTIFICADOS SET " & _
                        " CUALIFICACION_ID = " & CUALIFICACION_ID & "," & _
                        " DOCUMENTO_ID = " & DOCUMENTO_ID & "," & _
                        " ANYO = " & ANYO & "," & _
                        " COD_CERTIFICADO = " & COD_CERTIFICADO & "," & _
                        " CERTIFICADOR_ID = " & CERTIFICADOR_ID & "," & _
                        " BASE = '" & BASE & "'," & _
                        " CURSO_ID = " & CURSO_ID & "," & _
                        " FECHA_CERTIFICACION = '" & FECHA_CERTIFICACION & "'," & _
                        " CUSERID = " & CUSERID & "," & _
                        " FTIMESTP = '" & FTIMESTP & "'" & _
                " WHERE ID_FORMACION_CERTIFICADO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsFormacion_certificados)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_CERTIFICADOS " & _
                "    WHERE ID_FORMACION_CERTIFICADO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsFormacion_certificados)", vbCritical, Err.Description
End Function
Public Function EliminarCualificacion(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM FORMACION_CERTIFICADOS " & _
                "    WHERE CUALIFICACION_ID = " & ID
        execute_bd consulta
        EliminarCualificacion = True
        Exit Function
fallo:
        EliminarCualificacion = False
        MsgBox "Error al Eliminar por Cualificación(clsFormacion_certificados)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM FORMACION_CERTIFICADOS ORDER BY ID_FORMACION_CERTIFICADO"
        Set Listado = datos_bd(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_FORMACION_CERTIFICADO,CUALIFICACION_ID FROM FORMACION_CERTIFICADOS ORDER BY CUALIFICACION_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function

Public Function ListadoFiltro(COD_CERTIFICADO As Long, fecha_i As String, fecha_f As String, CERTIFICADOR_ID As Long, RESPONSABLE_ID As Long, CUALIFICACION_ID As Long) As ADODB.Recordset
        
        Dim consulta As String
        Dim strCod As String
        Dim strCert As String
        Dim strResp As String
        Dim strEmpl As String
        
        If COD_CERTIFICADO = 0 Then
            strCod = ""
        Else
            strCod = " AND COD_CERTIFICADO = " & COD_CERTIFICADO
        End If
        
        If CERTIFICADOR_ID = 0 Then
            strCert = ""
        Else
            strCert = " AND CERTIFICADOR_ID = " & CERTIFICADOR_ID
        End If
        
        If RESPONSABLE_ID = 0 Then
            strResp = ""
        Else
            strResp = " AND RESPONSABLE_ID = " & RESPONSABLE_ID
        End If
        
        If CUALIFICACION_ID = 0 Then
            strEmpl = ""
        Else
            strEmpl = " AND CUALIFICACION_ID = " & CUALIFICACION_ID
        End If
        
        consulta = "SELECT FC.ID_FORMACION_CERTIFICADO, FC.COD_CERTIFICADO AS CODIGO, FC.ANYO, FC.BASE, FC.FECHA_CERTIFICACION, E.DEPARTAMENTOS, E.NOMBRE, D.CODIGO AS COD_DOC" & _
                  " FROM empleados_cualificaciones EC, FORMACION_CERTIFICADOS FC, empleados E, ca_documentos AS D" & _
                  " WHERE FECHA_CERTIFICACION >='" & fecha_i & "'" & " AND FECHA_CERTIFICACION <='" & fecha_f & "'" & _
                  " AND FC.CUALIFICACION_ID = EC.ID_CUALIFICACION" & _
                  " AND FC.ANYO >= " & CLng(Format(fecha_i, "yyyy")) & " AND FC.ANYO <= " & CLng(Format(fecha_f, "yyyy")) & _
                  " AND FC.DOCUMENTO_ID = D.ID_DOCUMENTO" & _
                  " AND EC.EMPLEADO_ID = E.ID_EMPLEADO" & _
                  strCod & strCert & strResp & strEmpl
        
        Set ListadoFiltro = datos_bd(consulta)
End Function

Public Function generar_firmas(PK As Long) As Boolean
    'Recorrido por los asistentes para generar una firma por cada uno de ellos
    Dim rs As ADODB.Recordset
    Dim oCertificado As New clsFormacion_certificados
    Dim oCualificacion As New clsEmpleados_cualificaciones
    Dim oFirma As New clsFirmas
 
    On Error GoTo generar_firmas_Error

    oCertificado.Carga PK
    oCualificacion.Carga oCertificado.getCUALIFICACION_ID
    With oFirma
        .setTOBJETO = TOBJETO.TOBJETO_CERTIFICACION
        .setCOBJETO = PK
        .setUSUARIO_ID = oCualificacion.getEMPLEADO_ID
        .setFIRMADA = 1
        .setCOMENTARIO = "Certificado(empleado): Firma del documento"
        .InsertarUsuario
    End With
    With oFirma
        .setTOBJETO = TOBJETO.TOBJETO_CERTIFICACION_CERTIFICADOR
        .setCOBJETO = PK
        .setUSUARIO_ID = oCertificado.getCERTIFICADOR_ID
        .setFIRMADA = 1
        .setCOMENTARIO = "Empleado que certifica: Firma del documento"
        .InsertarUsuario
    End With
    With oFirma
        .setTOBJETO = TOBJETO.TOBJETO_CERTIFICACION_RESPONSABLE
        .setCOBJETO = PK
        .setUSUARIO_ID = oCualificacion.getEMPLEADO_ID_FORMADOR
        .setFIRMADA = 1
        .setCOMENTARIO = "Responsable de DPTO: Firma del documento"
        .InsertarUsuario
    End With
    Set oCertificado = Nothing
    Set oFirma = Nothing

    generar_firmas = True

   On Error GoTo 0
   Exit Function

generar_firmas_Error:
    generar_firmas = False

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure generar_firmas of Módulo de clase clsFormacion_certificados"
End Function

Public Function Listado_Combo_Todos(PK As Long)
        Dim consulta As String
   
        consulta = "SELECT CONCAT(DOCS.ID_DOCUMENTO, '-', FA.CURSO_ID)  AS ID, DOCS.NOMBRE, FC.ID_CURSO " & _
                    " FROM formacion_cursos FC " & _
                    " INNER JOIN FORMACION_ASISTENTES FA ON FC.ID_CURSO = FA.CURSO_ID " & _
                    " INNER JOIN FORMACION_PFA PFA ON FC.PLAN_ID = PFA.ID_PFA " & _
                    " INNER JOIN FORMACION_PF_DOCS D ON PFA.PLAN_FORMACION_ID = D.PLAN_FORMACION_ID " & _
                    " INNER JOIN ca_documentos DOCS ON D.DOCUMENTO_ID = DOCS.ID_DOCUMENTO " & _
                    " WHERE FA.EMPLEADO_ID = " & PK
        Set Listado_Combo_Todos = datos_bd(consulta)
End Function
