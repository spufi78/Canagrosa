VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsReclamaciones_notificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSRECLAMACIONES_NOTIFICACION
'***************************************************************
Private RECLAMACION_ID As Long
Private nombre As String
Private TELEFONO As String
Private DNI As String
Private EMAIL As String
Private CALLE As String
Private POBLACION As String
Private NUMERO As String
Private ESCALERA As String
Private PISO As String
Private PUERTA As String
Private CP As String
Private PROVINCIA As String
Private CALLE2 As String
Private POBLACION2 As String
Private NUMERO2 As String
Private ESCALERA2 As String
Private PISO2 As String
Private PUERTA2 As String
Private CP2 As String
Private PROVINCIA2 As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSRECLAMACIONES_NOTIFICACION
'***************************************************************
Public Property Let setRECLAMACION_ID(ByVal dato As Long)
        RECLAMACION_ID = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setTELEFONO(ByVal dato As String)
        TELEFONO = dato
End Property
Public Property Let setDNI(ByVal dato As String)
        DNI = dato
End Property
Public Property Let setEMAIL(ByVal dato As String)
        EMAIL = dato
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
Public Property Let setCALLE2(ByVal dato As String)
        CALLE2 = dato
End Property
Public Property Let setPOBLACION2(ByVal dato As String)
        POBLACION2 = dato
End Property
Public Property Let setNUMERO2(ByVal dato As String)
        NUMERO2 = dato
End Property
Public Property Let setESCALERA2(ByVal dato As String)
        ESCALERA2 = dato
End Property
Public Property Let setPISO2(ByVal dato As String)
        PISO2 = dato
End Property
Public Property Let setPUERTA2(ByVal dato As String)
        PUERTA2 = dato
End Property
Public Property Let setCP2(ByVal dato As String)
        CP2 = dato
End Property
Public Property Let setPROVINCIA2(ByVal dato As String)
        PROVINCIA2 = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSRECLAMACIONES_NOTIFICACION
'***************************************************************
Public Property Get getRECLAMACION_ID() As Long
        getRECLAMACION_ID = RECLAMACION_ID
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getTELEFONO() As String
        getTELEFONO = TELEFONO
End Property
Public Property Get getDNI() As String
        getDNI = DNI
End Property
Public Property Get getEMAIL() As String
        getEMAIL = EMAIL
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
Public Property Get getCALLE2() As String
        getCALLE2 = CALLE2
End Property
Public Property Get getPOBLACION2() As String
        getPOBLACION2 = POBLACION2
End Property
Public Property Get getNUMERO2() As String
        getNUMERO2 = NUMERO2
End Property
Public Property Get getESCALERA2() As String
        getESCALERA2 = ESCALERA2
End Property
Public Property Get getPISO2() As String
        getPISO2 = PISO2
End Property
Public Property Get getPUERTA2() As String
        getPUERTA2 = PUERTA2
End Property
Public Property Get getCP2() As String
        getCP2 = CP2
End Property
Public Property Get getPROVINCIA2() As String
        getPROVINCIA2 = PROVINCIA2
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSRECLAMACIONES_NOTIFICACION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM RECLAMACIONES_NOTIFICACION WHERE RECLAMACION_ID = " & ID
        Set rs = datos_bd(consulta, conn2)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                RECLAMACION_ID = rs("RECLAMACION_ID")
                nombre = rs("NOMBRE")
                TELEFONO = rs("TELEFONO")
                DNI = rs("DNI")
                EMAIL = rs("EMAIL")
                CALLE = rs("CALLE")
                POBLACION = rs("POBLACION")
                NUMERO = rs("NUMERO")
                ESCALERA = rs("ESCALERA")
                PISO = rs("PISO")
                PUERTA = rs("PUERTA")
                CP = rs("CP")
                PROVINCIA = rs("PROVINCIA")
                CALLE2 = rs("CALLE2")
                POBLACION2 = rs("POBLACION2")
                NUMERO2 = rs("NUMERO2")
                ESCALERA2 = rs("ESCALERA2")
                PISO2 = rs("PISO2")
                PUERTA2 = rs("PUERTA2")
                CP2 = rs("CP2")
                PROVINCIA2 = rs("PROVINCIA2")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsReclamaciones_notificacion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO RECLAMACIONES_NOTIFICACION " & _
                   "  VALUES (" & _
                        RECLAMACION_ID & "," & "'" & nombre & "'" & "," & "'" & TELEFONO & "'" & "," & _
                        "'" & DNI & "'" & "," & "'" & EMAIL & "'" & "," & "'" & CALLE & "'" & "," & _
                        "'" & POBLACION & "'" & "," & "'" & NUMERO & "'" & "," & "'" & ESCALERA & "'" & "," & _
                        "'" & PISO & "'" & "," & "'" & PUERTA & "'" & "," & "'" & CP & "'" & "," & _
                        "'" & PROVINCIA & "'" & "," & "'" & CALLE2 & "'" & "," & "'" & POBLACION2 & "'" & "," & _
                        "'" & NUMERO2 & "'" & "," & "'" & ESCALERA2 & "'" & "," & "'" & PISO2 & "'" & "," & _
                        "'" & PUERTA2 & "'" & "," & "'" & CP2 & "'" & "," & "'" & PROVINCIA2 & "'" & _
                ")"
        execute_bd consulta, conn2
        Insertar = RECLAMACION_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsReclamaciones_notificacion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE RECLAMACIONES_NOTIFICACION SET " & _
                        " NOMBRE = '" & nombre & "'," & _
                        " TELEFONO = '" & TELEFONO & "'," & _
                        " DNI = '" & DNI & "'," & _
                        " EMAIL = '" & EMAIL & "'," & _
                        " CALLE = '" & CALLE & "'," & _
                        " POBLACION = '" & POBLACION & "'," & _
                        " NUMERO = '" & NUMERO & "'," & _
                        " ESCALERA = '" & ESCALERA & "'," & _
                        " PISO = '" & PISO & "'," & _
                        " PUERTA = '" & PUERTA & "'," & _
                        " CP = '" & CP & "'," & _
                        " PROVINCIA = '" & PROVINCIA & "'," & _
                        " CALLE2 = '" & CALLE2 & "'," & _
                        " POBLACION2 = '" & POBLACION2 & "'," & _
                        " NUMERO2 = '" & NUMERO2 & "'," & _
                        " ESCALERA2 = '" & ESCALERA2 & "'," & _
                        " PISO2 = '" & PISO2 & "'," & _
                        " PUERTA2 = '" & PUERTA2 & "'," & _
                        " CP2 = '" & CP2 & "'," & _
                        " PROVINCIA2 = '" & PROVINCIA2 & "'" & _
                " WHERE RECLAMACION_ID = " & ID
        execute_bd consulta, conn2
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsReclamaciones_notificacion)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM RECLAMACIONES_NOTIFICACION " & _
                "    WHERE RECLAMACION_ID = " & ID
        execute_bd consulta, conn2
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsReclamaciones_notificacion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM RECLAMACIONES_NOTIFICACION ORDER BY RECLAMACION_ID"
        Set Listado = datos_bd(consulta, conn2)
End Function
