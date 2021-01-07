VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_pnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_PNT
'***************************************************************
Private DOCUMENTO_ID As Long
Private EDICION As Integer
Private CODIGO As String
Private DESCRIPCION As String
Private MODIFICACION As String
Private fecha_creacion As String
Private FECHA_REVISION As String
Private FECHA_APROBACION As String
Private USUARIO_CREACION As Integer
Private USUARIO_REVISION As Integer
Private USUARIO_APROBACION As Integer
Private MODIFICACION_FECHA As String
Private MODIFICACION_FECHA_REVISION As String
Private MODIFICACION_FECHA_APROBACION As String
Private MODIFICACION_USUARIO_ As Integer
Private MODIFICACION_USUARIO_REVISION As Integer
Private MODIFICACION_USUARIO_APROBACION As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_PNT
'***************************************************************
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setEDICION(ByVal dato As Integer)
        EDICION = dato
End Property
Public Property Let setCODIGO(ByVal dato As String)
        CODIGO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setMODIFICACION(ByVal dato As String)
        MODIFICACION = dato
End Property
Public Property Let setFECHA_CREACION(ByVal dato As String)
        fecha_creacion = dato
End Property
Public Property Let setFECHA_REVISION(ByVal dato As String)
        FECHA_REVISION = dato
End Property
Public Property Let setFECHA_APROBACION(ByVal dato As String)
        FECHA_APROBACION = dato
End Property
Public Property Let setUSUARIO_CREACION(ByVal dato As Integer)
        USUARIO_CREACION = dato
End Property
Public Property Let setUSUARIO_REVISION(ByVal dato As Integer)
        USUARIO_REVISION = dato
End Property
Public Property Let setUSUARIO_APROBACION(ByVal dato As Integer)
        USUARIO_APROBACION = dato
End Property
Public Property Let setMODIFICACION_FECHA(ByVal dato As String)
        MODIFICACION_FECHA = dato
End Property
Public Property Let setMODIFICACION_FECHA_REVISION(ByVal dato As String)
        MODIFICACION_FECHA_REVISION = dato
End Property
Public Property Let setMODIFICACION_FECHA_APROBACION(ByVal dato As String)
        MODIFICACION_FECHA_APROBACION = dato
End Property
Public Property Let setMODIFICACION_USUARIO_(ByVal dato As Integer)
        MODIFICACION_USUARIO_ = dato
End Property
Public Property Let setMODIFICACION_USUARIO_REVISION(ByVal dato As Integer)
        MODIFICACION_USUARIO_REVISION = dato
End Property
Public Property Let setMODIFICACION_USUARIO_APROBACION(ByVal dato As Integer)
        MODIFICACION_USUARIO_APROBACION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_PNT
'***************************************************************
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getEDICION() As Integer
        getEDICION = EDICION
End Property
Public Property Get getCODIGO() As String
        getCODIGO = CODIGO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getMODIFICACION() As String
        getMODIFICACION = MODIFICACION
End Property
Public Property Get getFECHA_CREACION() As String
        getFECHA_CREACION = fecha_creacion
End Property
Public Property Get getFECHA_REVISION() As String
        getFECHA_REVISION = FECHA_REVISION
End Property
Public Property Get getFECHA_APROBACION() As String
        getFECHA_APROBACION = FECHA_APROBACION
End Property
Public Property Get getUSUARIO_CREACION() As Integer
        getUSUARIO_CREACION = USUARIO_CREACION
End Property
Public Property Get getUSUARIO_REVISION() As Integer
        getUSUARIO_REVISION = USUARIO_REVISION
End Property
Public Property Get getUSUARIO_APROBACION() As Integer
        getUSUARIO_APROBACION = USUARIO_APROBACION
End Property
Public Property Get getMODIFICACION_FECHA() As String
        getMODIFICACION_FECHA = MODIFICACION_FECHA
End Property
Public Property Get getMODIFICACION_FECHA_REVISION() As String
        getMODIFICACION_FECHA_REVISION = MODIFICACION_FECHA_REVISION
End Property
Public Property Get getMODIFICACION_FECHA_APROBACION() As String
        getMODIFICACION_FECHA_APROBACION = MODIFICACION_FECHA_APROBACION
End Property
Public Property Get getMODIFICACION_USUARIO_() As Integer
        getMODIFICACION_USUARIO_ = MODIFICACION_USUARIO_
End Property
Public Property Get getMODIFICACION_USUARIO_REVISION() As Integer
        getMODIFICACION_USUARIO_REVISION = MODIFICACION_USUARIO_REVISION
End Property
Public Property Get getMODIFICACION_USUARIO_APROBACION() As Integer
        getMODIFICACION_USUARIO_APROBACION = MODIFICACION_USUARIO_APROBACION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_PNT
'***************************************************************
Public Function Carga(documento As Long, lEDICION As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CA_PNT WHERE DOCUMENTO_ID = " & documento & " AND EDICION = " & lEDICION
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                EDICION = rs("EDICION")
                CODIGO = rs("CODIGO")
                DESCRIPCION = rs("DESCRIPCION")
                MODIFICACION = rs("MODIFICACION")
                If Not IsNull(rs("FECHA_CREACION")) Then
                    fecha_creacion = rs("FECHA_CREACION")
                End If
                If Not IsNull(rs("FECHA_REVISION")) Then
                    FECHA_REVISION = rs("FECHA_REVISION")
                End If
                If Not IsNull(rs("FECHA_APROBACION")) Then
                    FECHA_APROBACION = rs("FECHA_APROBACION")
                End If
                USUARIO_CREACION = rs("USUARIO_CREACION")
                USUARIO_REVISION = rs("USUARIO_REVISION")
                USUARIO_APROBACION = rs("USUARIO_APROBACION")
                If Not IsNull(rs("MODIFICACION_FECHA")) Then
                    MODIFICACION_FECHA = rs("MODIFICACION_FECHA")
                End If
                If Not IsNull(rs("MODIFICACION_FECHA_REVISION")) Then
                    MODIFICACION_FECHA_REVISION = rs("MODIFICACION_FECHA_REVISION")
                End If
                If Not IsNull(rs("MODIFICACION_FECHA_APROBACION")) Then
                    MODIFICACION_FECHA_APROBACION = rs("MODIFICACION_FECHA_APROBACION")
                End If
                MODIFICACION_USUARIO_ = rs("MODIFICACION_USUARIO_")
                MODIFICACION_USUARIO_REVISION = rs("MODIFICACION_USUARIO_REVISION")
                MODIFICACION_USUARIO_APROBACION = rs("MODIFICACION_USUARIO_APROBACION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCa_pnt)", vbCritical, Err.Description
End Function
Public Function Carga_Ultima_edicion(documento As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM CA_PNT WHERE DOCUMENTO_ID = " & documento & " ORDER BY EDICION DESC"
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga_Ultima_edicion = False
        Else
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                EDICION = rs("EDICION")
                CODIGO = rs("CODIGO")
                DESCRIPCION = rs("DESCRIPCION")
                MODIFICACION = rs("MODIFICACION")
                If Not IsNull(rs("FECHA_CREACION")) Then
                    fecha_creacion = rs("FECHA_CREACION")
                End If
                If Not IsNull(rs("FECHA_REVISION")) Then
                    FECHA_REVISION = rs("FECHA_REVISION")
                End If
                If Not IsNull(rs("FECHA_APROBACION")) Then
                    FECHA_APROBACION = rs("FECHA_APROBACION")
                End If
                USUARIO_CREACION = rs("USUARIO_CREACION")
                USUARIO_REVISION = rs("USUARIO_REVISION")
                USUARIO_APROBACION = rs("USUARIO_APROBACION")
                If Not IsNull(rs("MODIFICACION_FECHA")) Then
                    MODIFICACION_FECHA = rs("MODIFICACION_FECHA")
                End If
                If Not IsNull(rs("MODIFICACION_FECHA_REVISION")) Then
                    MODIFICACION_FECHA_REVISION = rs("MODIFICACION_FECHA_REVISION")
                End If
                If Not IsNull(rs("MODIFICACION_FECHA_APROBACION")) Then
                    MODIFICACION_FECHA_APROBACION = rs("MODIFICACION_FECHA_APROBACION")
                End If
                MODIFICACION_USUARIO_ = rs("MODIFICACION_USUARIO_")
                MODIFICACION_USUARIO_REVISION = rs("MODIFICACION_USUARIO_REVISION")
                MODIFICACION_USUARIO_APROBACION = rs("MODIFICACION_USUARIO_APROBACION")
                Carga_Ultima_edicion = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga_Ultima_edicion = False
        MsgBox "Error al cargar los Carga_Ultima_edicion(clsCa_pnt)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO CA_PNT " & _
                   "  VALUES (" & _
                        DOCUMENTO_ID & "," & EDICION & "," & "'" & CODIGO & "'" & "," & _
                        "'" & DESCRIPCION & "'" & "," & "'" & MODIFICACION & "'" & "," & "'" & fecha_creacion & "'" & "," & _
                        "'" & FECHA_REVISION & "'" & "," & "'" & FECHA_APROBACION & "'" & "," & USUARIO_CREACION & "," & _
                        USUARIO_REVISION & "," & USUARIO_APROBACION & ",'" & MODIFICACION_FECHA & "','" & MODIFICACION_FECHA_REVISION & "','" & MODIFICACION_FECHA_APROBACION & "'," & _
                        MODIFICACION_USUARIO_ & "," & MODIFICACION_USUARIO_REVISION & "," & MODIFICACION_USUARIO_APROBACION & _
                ")"
        execute_bd consulta
        Insertar = DOCUMENTO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCa_pnt)", vbCritical, Err.Description
End Function
Public Function Modificar(documento As Long, lEDICION As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE CA_PNT SET " & _
                        " CODIGO = '" & CODIGO & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " MODIFICACION = '" & MODIFICACION & "'," & _
                        " FECHA_CREACION = '" & fecha_creacion & "'," & _
                        " FECHA_REVISION = '" & FECHA_REVISION & "'," & _
                        " FECHA_APROBACION = '" & FECHA_APROBACION & "'," & _
                        " USUARIO_CREACION = " & USUARIO_CREACION & "," & _
                        " USUARIO_REVISION = " & USUARIO_REVISION & "," & _
                        " USUARIO_APROBACION = " & USUARIO_APROBACION & "," & _
                        " MODIFICACION_FECHA = '" & MODIFICACION_FECHA & "'," & _
                        " MODIFICACION_FECHA_REVISION = '" & MODIFICACION_FECHA_REVISION & "'," & _
                        " MODIFICACION_FECHA_APROBACION = '" & MODIFICACION_FECHA_APROBACION & "'," & _
                        " MODIFICACION_USUARIO_ = " & MODIFICACION_USUARIO_ & "," & _
                        " MODIFICACION_USUARIO_REVISION = " & MODIFICACION_USUARIO_REVISION & "," & _
                        " MODIFICACION_USUARIO_APROBACION = " & MODIFICACION_USUARIO_APROBACION & "" & _
                " WHERE DOCUMENTO_ID = " & documento & _
                "   AND EDICION = " & lEDICION
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCa_pnt)", vbCritical, Err.Description
End Function
Public Function Eliminar(DOCUMENTO_ID As Long, EDICION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CA_PNT " & _
                "    WHERE DOCUMENTO_ID = " & DOCUMENTO_ID & _
                "      AND EDICION = " & EDICION
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCa_pnt)", vbCritical, Err.Description
End Function
Public Function Listado(documento As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.EDICION,A.MODIFICACION_FECHA,A.MODIFICACION,B.USUARIO" & _
                   "      ,IF(A.EDICION=1,A.FECHA_REVISION,A.MODIFICACION_FECHA_REVISION),C.USUARIO " & _
                   "      ,IF(A.EDICION=1,A.FECHA_APROBACION,A.MODIFICACION_FECHA_APROBACION),D.USUARIO " & _
                   "      ,B.ID_EMPLEADO,C.ID_EMPLEADO,D.ID_EMPLEADO " & _
                   "  FROM CA_PNT A,usuarios B,usuarios C,usuarios D " & _
                   " WHERE A.DOCUMENTO_ID = " & documento & _
                   "   AND A.MODIFICACION_USUARIO_ = B.ID_EMPLEADO " & _
                   "   AND A.MODIFICACION_USUARIO_REVISION = C.ID_EMPLEADO " & _
                   "   AND A.MODIFICACION_USUARIO_APROBACION = D.ID_EMPLEADO " & _
                   " ORDER BY A.EDICION DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Informar_Aprobacion(documento As Long, lEDICION As Integer, lFECHA As Date) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim campo As String
        If lEDICION = 1 Then
            campo = " FECHA_APROBACION "
        Else
            campo = " MODIFICACION_FECHA_APROBACION "
        End If
        consulta = "UPDATE CA_PNT SET " & _
                        campo & " = '" & Format(lFECHA, "yyyy-mm-dd") & "'" & _
                 " WHERE DOCUMENTO_ID = " & documento & _
                 "   AND EDICION = " & lEDICION
        execute_bd consulta
        Informar_Aprobacion = True
        Exit Function
fallo:
        Informar_Aprobacion = False
        MsgBox "Error al Informar_Aprobacion (clsCa_pnt)", vbCritical, Err.Description
End Function
Public Function Informar_Revision(documento As Long, lEDICION As Integer, lFECHA As Date) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim campo As String
        If lEDICION = 1 Then
            campo = " FECHA_REVISION "
        Else
            campo = " MODIFICACION_FECHA_REVISION "
        End If
        consulta = "UPDATE CA_PNT SET " & _
                        campo & " = '" & Format(lFECHA, "yyyy-mm-dd") & "'" & _
                 " WHERE DOCUMENTO_ID = " & documento & _
                 "   AND EDICION = " & lEDICION
        execute_bd consulta
        Informar_Revision = True
        Exit Function
fallo:
        Informar_Revision = False
        MsgBox "Error al Informar_Revision (clsCa_pnt)", vbCritical, Err.Description
End Function
Public Function Informar_Modificacion(documento As Long, lEDICION As Integer, lFECHA As Date) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim campo As String
        If lEDICION = 1 Then
            campo = " FECHA_CREACION "
        Else
            campo = " MODIFICACION_FECHA "
        End If
        consulta = "UPDATE CA_PNT SET " & _
                        campo & " = '" & Format(lFECHA, "yyyy-mm-dd") & "'" & _
                 " WHERE DOCUMENTO_ID = " & documento & _
                 "   AND EDICION = " & lEDICION
        execute_bd consulta
        Informar_Modificacion = True
        Exit Function
fallo:
        Informar_Modificacion = False
        MsgBox "Error al Informar_Modificacion (clsCa_pnt)", vbCritical, Err.Description
End Function


