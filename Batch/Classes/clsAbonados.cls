VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAbonados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSABONADOS
'***************************************************************
Private ID_ABONADO As Long
Private nombre As String
Private DNI As String
Private CALLE As String
Private NUMERO As String
Private ESCALERA As String
Private PISO As String
Private PUERTA As String
Private CP As String
Private POBLACION As String
Private PROVINCIA As String
Private TELEFONO As String
Private MOVIL As String
Private DERECHO_USO As String
Private BAJA As Integer
Private NUMERO_ABONADO As String
Private OFICINA_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSABONADOS
'***************************************************************
Public Property Let setID_ABONADO(ByVal dato As Long)
        ID_ABONADO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        nombre = dato
End Property
Public Property Let setDNI(ByVal dato As String)
        DNI = dato
End Property
Public Property Let setCALLE(ByVal dato As String)
        CALLE = dato
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
Public Property Let setPOBLACION(ByVal dato As String)
        POBLACION = dato
End Property
Public Property Let setPROVINCIA(ByVal dato As String)
        PROVINCIA = dato
End Property
Public Property Let setTELEFONO(ByVal dato As String)
        TELEFONO = dato
End Property
Public Property Let setMOVIL(ByVal dato As String)
        MOVIL = dato
End Property
Public Property Let setDERECHO_USO(ByVal dato As String)
        DERECHO_USO = dato
End Property
Public Property Let setBAJA(ByVal dato As Integer)
        BAJA = dato
End Property
Public Property Let setNUMERO_ABONADO(ByVal dato As String)
        NUMERO_ABONADO = dato
End Property
Public Property Let setOFICINA_ID(ByVal dato As Long)
        OFICINA_ID = dato
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSABONADOS
'***************************************************************
Public Property Get getID_ABONADO() As Long
        getID_ABONADO = ID_ABONADO
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = nombre
End Property
Public Property Get getDNI() As String
        getDNI = DNI
End Property
Public Property Get getCALLE() As String
        getCALLE = CALLE
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
Public Property Get getPOBLACION() As String
        getPOBLACION = POBLACION
End Property
Public Property Get getPROVINCIA() As String
        getPROVINCIA = PROVINCIA
End Property
Public Property Get getTELEFONO() As String
        getTELEFONO = TELEFONO
End Property
Public Property Get getMOVIL() As String
        getMOVIL = MOVIL
End Property
Public Property Get getDERECHO_USO() As String
        getDERECHO_USO = DERECHO_USO
End Property
Public Property Get getBAJA() As Integer
        getBAJA = BAJA
End Property
Public Property Get getNUMERO_ABONADO() As String
        getNUMERO_ABONADO = NUMERO_ABONADO
End Property
Public Property Get getOFICINA_ID() As Long
        getOFICINA_ID = OFICINA_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSABONADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ABONADOS WHERE ID_ABONADO = " & ID
        Set rs = datos_bd(consulta, conn2)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_ABONADO = rs("ID_ABONADO")
                nombre = rs("NOMBRE")
                DNI = rs("DNI")
                CALLE = rs("CALLE")
                NUMERO = rs("NUMERO")
                ESCALERA = rs("ESCALERA")
                PISO = rs("PISO")
                PUERTA = rs("PUERTA")
                CP = rs("CP")
                POBLACION = rs("POBLACION")
                PROVINCIA = rs("PROVINCIA")
                TELEFONO = rs("TELEFONO")
                MOVIL = rs("MOVIL")
                DERECHO_USO = rs("DERECHO_USO")
                BAJA = rs("BAJA")
                NUMERO_ABONADO = rs("NUMERO_ABONADO")
                OFICINA_ID = rs("OFICINA_ID")
                
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsAbonados)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_ABONADO) FROM ABONADOS WHERE ID_ABONADO < 1000000"
        Set rs = datos_bd(consulta, conn2)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_ABONADO = 1
        Else
                ID_ABONADO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT ID_ABONADO FROM ABONADOS WHERE NUMERO_ABONADO ='" & NUMERO_ABONADO & "' AND OFICINA_ID = " & OFICINA_ID
        Set rs = datos_bd(consulta, conn2)
        If rs.RecordCount = 0 Then
            Me.CrearID
            consulta = "INSERT INTO ABONADOS " & _
                       "  VALUES (" & _
                            ID_ABONADO & "," & "'" & nombre & "'" & "," & "'" & DNI & "'" & "," & _
                            "'" & CALLE & "'" & "," & "'" & NUMERO & "'" & "," & "'" & ESCALERA & "'" & "," & _
                            "'" & PISO & "'" & "," & "'" & PUERTA & "'" & "," & "'" & CP & "'" & "," & _
                            "'" & POBLACION & "'" & "," & "'" & PROVINCIA & "'" & "," & "'" & TELEFONO & "'" & "," & _
                            "'" & MOVIL & "','" & DERECHO_USO & "'," & BAJA & "," & _
                        "'" & NUMERO_ABONADO & "'" & "," & OFICINA_ID & _
                    ")"
            execute_bd consulta, conn2
        Else
            ID_ABONADO = rs(0)
            Me.Modificar ID_ABONADO
        End If
        Insertar = ID_ABONADO
        Exit Function
fallo:
        Insertar = 0
        Dim s As String
        s = "Error al insertar el abonado. " & vbNewLine
        s = s & vbNewLine & " Nombre : " & nombre
        s = s & vbNewLine & " DNI : " & DNI
        s = s & vbNewLine & " Calle : " & CALLE
        s = s & vbNewLine & " Poblacion : " & POBLACION
        s = s & vbNewLine & " Provincia : " & PROVINCIA
        log s
'        MsgBox s, vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ABONADOS SET " & _
                        " NOMBRE = '" & nombre & "'," & _
                        " DNI = '" & DNI & "'," & _
                        " CALLE = '" & CALLE & "'," & _
                        " NUMERO = '" & NUMERO & "'," & _
                        " ESCALERA = '" & ESCALERA & "'," & _
                        " PISO = '" & PISO & "'," & _
                        " PUERTA = '" & PUERTA & "'," & _
                        " CP = '" & CP & "'," & _
                        " POBLACION = '" & POBLACION & "'," & _
                        " PROVINCIA = '" & PROVINCIA & "'," & _
                        " TELEFONO = '" & TELEFONO & "'," & _
                        " MOVIL = '" & MOVIL & "'," & _
                        " DERECHO_USO = '" & DERECHO_USO & "'," & _
                        " BAJA = " & BAJA & "," & _
                        " NUMERO_ABONADO = '" & NUMERO_ABONADO & "'," & _
                        " OFICINA_ID = " & OFICINA_ID & _
                " WHERE ID_ABONADO = " & ID
        execute_bd consulta, conn2
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        Dim s As String
        s = "Error al modificar el abonado. " & vbNewLine
        s = s & vbNewLine & " Nombre : " & nombre
        s = s & vbNewLine & " DNI : " & DNI
        s = s & vbNewLine & " Calle : " & CALLE
        s = s & vbNewLine & " Poblacion : " & POBLACION
        s = s & vbNewLine & " Provincia : " & PROVINCIA
        log s
'        MsgBox s, vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ABONADOS " & _
                "    WHERE ID_ABONADO = " & ID
        execute_bd consulta, conn2
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsAbonados)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ABONADOS ORDER BY ID_ABONADO"
        Set Listado = datos_bd(consulta, conn2)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_ABONADO,NOMBRE FROM ABONADOS ORDER BY NOMBRE"
        Set Listado_Combo = datos_bd(consulta, conn2)
End Function
