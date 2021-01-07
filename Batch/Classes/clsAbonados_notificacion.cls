VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAbonados_notificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSABONADOS_NOTIFICACION
'***************************************************************
Private ABONADO_ID As Long
Private CALLE As String
Private POBLACION As String
Private NUMERO As String
Private ESCALERA As String
Private PISO As String
Private PUERTA As String
Private CP As String
Private PROVINCIA As String
Private OFICINA_ID As Integer
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSABONADOS_NOTIFICACION
'***************************************************************
Public Property Let setABONADO_ID(ByVal dato As Long)
        ABONADO_ID = dato
End Property
Public Property Let setCALLE(ByVal dato As String)
        CALLE = dato
End Property
Public Property Let setPOBLACION(ByVal dato As String)
        POBLACION = dato
End Property
Public Property Let setNUMERO(ByVal dato As String)
        NUMERO = dato
End Property
Public Property Let setESCALERA(ByVal dato As String)
        ESCALERA = dato
End Property
Public Property Let setPISO(ByVal dato As String)
        PISO = dato
End Property
Public Property Let setPUERTA(ByVal dato As String)
        PUERTA = dato
End Property
Public Property Let setCP(ByVal dato As String)
        CP = dato
End Property
Public Property Let setPROVINCIA(ByVal dato As String)
        PROVINCIA = dato
End Property
Public Property Let setOFICINA_ID(ByVal dato As Integer)
        OFICINA_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSABONADOS_NOTIFICACION
'***************************************************************
Public Property Get getABONADO_ID() As Long
        getABONADO_ID = ABONADO_ID
End Property
Public Property Get getCALLE() As String
        getCALLE = CALLE
End Property
Public Property Get getPOBLACION() As String
        getPOBLACION = POBLACION
End Property
Public Property Get getNUMERO() As String
        getNUMERO = NUMERO
End Property
Public Property Get getESCALERA() As String
        getESCALERA = ESCALERA
End Property
Public Property Get getPISO() As String
        getPISO = PISO
End Property
Public Property Get getPUERTA() As String
        getPUERTA = PUERTA
End Property
Public Property Get getCP() As String
        getCP = CP
End Property
Public Property Get getPROVINCIA() As String
        getPROVINCIA = PROVINCIA
End Property
Public Property Get getOFICINA_ID() As Integer
        getOFICINA_ID = OFICINA_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSABONADOS_NOTIFICACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ABONADOS_NOTIFICACION WHERE ABONADO_ID = " & ID
        Set rs = datos_bd(consulta, conn2)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ABONADO_ID = rs("ABONADO_ID")
                CALLE = rs("CALLE")
                POBLACION = rs("POBLACION")
                NUMERO = rs("NUMERO")
                ESCALERA = rs("ESCALERA")
                PISO = rs("PISO")
                PUERTA = rs("PUERTA")
                CP = rs("CP")
                PROVINCIA = rs("PROVINCIA")
                OFICINA_ID = rs("OFICINA_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAbonados_notificacion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT ABONADO_ID FROM ABONADOS_NOTIFICACION WHERE ABONADO_ID = " & ABONADO_ID
        Set rs = datos_bd(consulta, conn2)
        If rs.RecordCount = 0 Then
            consulta = "INSERT INTO ABONADOS_NOTIFICACION " & _
                       "  VALUES (" & _
                            ABONADO_ID & "," & "'" & CALLE & "'" & "," & "'" & POBLACION & "'" & "," & _
                            "'" & NUMERO & "'" & "," & "'" & ESCALERA & "'" & "," & "'" & PISO & "'" & "," & _
                            "'" & PUERTA & "'" & "," & "'" & CP & "'" & "," & "'" & PROVINCIA & "'," & OFICINA_ID & _
                    ")"
            execute_bd consulta, conn2
        Else
            ABONADO_ID = rs(0)
            Me.Modificar ABONADO_ID
        End If
        Insertar = ABONADO_ID
        Exit Function
fallo:
        Insertar = 0
        Dim s As String
        s = "Error al insertar el abonado_notificacion. " & vbNewLine
        s = s & vbNewLine & " ABONADO_ID : " & ABONADO_ID
        s = s & vbNewLine & " CALLE : " & CALLE
        s = s & vbNewLine & " NUMERO : " & NUMERO
        s = s & vbNewLine & " Poblacion : " & POBLACION
        s = s & vbNewLine & " Provincia : " & PROVINCIA
        log s
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ABONADOS_NOTIFICACION SET " & _
                        " CALLE = '" & CALLE & "'," & _
                        " POBLACION = '" & POBLACION & "'," & _
                        " NUMERO = '" & NUMERO & "'," & _
                        " ESCALERA = '" & ESCALERA & "'," & _
                        " PISO = '" & PISO & "'," & _
                        " PUERTA = '" & PUERTA & "'," & _
                        " CP = '" & CP & "'," & _
                        " PROVINCIA = '" & PROVINCIA & "'," & _
                        " OFICINA_ID = " & OFICINA_ID & _
                " WHERE ABONADO_ID = " & ID
        execute_bd consulta, conn2
        Modificar = True
        Exit Function
fallo:
        Modificar = False

End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ABONADOS_NOTIFICACION " & _
                "    WHERE ABONADO_ID = " & ID
        execute_bd consulta, conn2
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAbonados_notificacion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ABONADOS_NOTIFICACION ORDER BY ABONADO_ID"
        Set Listado = datos_bd(consulta, conn2)
End Function
