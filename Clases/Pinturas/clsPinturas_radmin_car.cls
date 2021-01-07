VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPinturas_radmin_car"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsPINTURAS_RADMIN_CAR
'***************************************************************
Private RADMIN_ID As Long
Private tipo As Long
Private NOMBRE_BASE As String
Private fabricante As String
Private ESPECIFICACIONES As String
Private N_LATAS As String
Private CANTIDAD_LATA As String
Private LOTE As String
Private FECHA_ENVASADO As String
Private fecha_caducidad As String
Private CONDICIONES_ALMACENAMIENTO As String
Private ESTADO_EMBALAJE As String
'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE clsPINTURAS_RADMIN_CAR
'***************************************************************
Public Enum LISTA_clsPINTURAS_RADMIN_CAR_COLS
        ENUM_RADMIN_ID = 1
        ENUM_TIPO = 2
        ENUM_NOMBRE_BASE = 3
        ENUM_FABRICANTE = 4
        ENUM_ESPECIFICACIONES = 5
        ENUM_N_LATAS = 6
        ENUM_CANTIDAD_LATA = 7
        ENUM_LOTE = 8
        ENUM_FECHA_ENVASADO = 9
        ENUM_FECHA_CADUCIDAD = 10
        ENUM_CONDICIONES_ALMACENAMIENTO = 11
        ENUM_ESTADO_EMBALAJE = 12
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsPINTURAS_RADMIN_CAR
'***************************************************************
Public Property Let setRADMIN_ID(ByVal dato As Long)
        RADMIN_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As Long)
        tipo = dato
End Property
Public Property Let setNOMBRE_BASE(ByVal dato As String)
        NOMBRE_BASE = dato
End Property
Public Property Let setFABRICANTE(ByVal dato As String)
        fabricante = dato
End Property
Public Property Let setESPECIFICACIONES(ByVal dato As String)
        ESPECIFICACIONES = dato
End Property
Public Property Let setN_LATAS(ByVal dato As String)
        N_LATAS = dato
End Property
Public Property Let setCANTIDAD_LATA(ByVal dato As String)
        CANTIDAD_LATA = dato
End Property
Public Property Let setLOTE(ByVal dato As String)
        LOTE = dato
End Property
Public Property Let setFECHA_ENVASADO(ByVal dato As String)
        FECHA_ENVASADO = dato
End Property
Public Property Let setFECHA_CADUCIDAD(ByVal dato As String)
        fecha_caducidad = dato
End Property
Public Property Let setCONDICIONES_ALMACENAMIENTO(ByVal dato As String)
        CONDICIONES_ALMACENAMIENTO = dato
End Property
Public Property Let setESTADO_EMBALAJE(ByVal dato As String)
        ESTADO_EMBALAJE = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsPINTURAS_RADMIN_CAR
'***************************************************************
Public Property Get getRADMIN_ID() As Long
        getRADMIN_ID = RADMIN_ID
End Property
Public Property Get getTIPO() As Long
        getTIPO = tipo
End Property
Public Property Get getNOMBRE_BASE() As String
        getNOMBRE_BASE = NOMBRE_BASE
End Property
Public Property Get getFABRICANTE() As String
        getFABRICANTE = fabricante
End Property
Public Property Get getESPECIFICACIONES() As String
        getESPECIFICACIONES = ESPECIFICACIONES
End Property
Public Property Get getN_LATAS() As String
        getN_LATAS = N_LATAS
End Property
Public Property Get getCANTIDAD_LATA() As String
        getCANTIDAD_LATA = CANTIDAD_LATA
End Property
Public Property Get getLOTE() As String
        getLOTE = LOTE
End Property
Public Property Get getFECHA_ENVASADO() As String
        getFECHA_ENVASADO = FECHA_ENVASADO
End Property
Public Property Get getFECHA_CADUCIDAD() As String
        getFECHA_CADUCIDAD = fecha_caducidad
End Property
Public Property Get getCONDICIONES_ALMACENAMIENTO() As String
        getCONDICIONES_ALMACENAMIENTO = CONDICIONES_ALMACENAMIENTO
End Property
Public Property Get getESTADO_EMBALAJE() As String
        getESTADO_EMBALAJE = ESTADO_EMBALAJE
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsPINTURAS_RADMIN_CAR
'***************************************************************
Public Function Carga(oRADMIN As Long, oTIPO As Integer) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PINTURAS_RADMIN_CAR " & _
                   " WHERE RADMIN_ID = " & oRADMIN & _
                   "   AND TIPO = " & oTIPO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                RADMIN_ID = rs("RADMIN_ID")
                tipo = rs("TIPO")
                NOMBRE_BASE = rs("NOMBRE_BASE")
                fabricante = rs("FABRICANTE")
                ESPECIFICACIONES = rs("ESPECIFICACIONES")
                N_LATAS = rs("N_LATAS")
                CANTIDAD_LATA = rs("CANTIDAD_LATA")
                LOTE = rs("LOTE")
                If Not IsNull(rs("FECHA_ENVASADO")) Then
                    FECHA_ENVASADO = rs("FECHA_ENVASADO")
                End If
                If Not IsNull(rs("FECHA_CADUCIDAD")) Then
                    fecha_caducidad = rs("FECHA_CADUCIDAD")
                End If
                CONDICIONES_ALMACENAMIENTO = rs("CONDICIONES_ALMACENAMIENTO")
                ESTADO_EMBALAJE = rs("ESTADO_EMBALAJE")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPINTURAS_RADMIN_CAR)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(RADMIN_ID) FROM PINTURAS_RADMIN_CAR"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                RADMIN_ID = 1
        Else
                RADMIN_ID = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim fEnvasado As String
        If FECHA_ENVASADO = 0 Then
            fEnvasado = "NULL"
        Else
            fEnvasado = "'" & FECHA_ENVASADO & "'"
        End If
        Dim fCaducidad As String
        If fecha_caducidad = 0 Then
            fCaducidad = "NULL"
        Else
            fCaducidad = "'" & fecha_caducidad & "'"
        End If
        consulta = "INSERT INTO PINTURAS_RADMIN_CAR " & _
                   "  VALUES (" & _
                        RADMIN_ID & "," & tipo & "," & _
                        "'" & NOMBRE_BASE & "'" & "," & _
                        "'" & fabricante & "'" & "," & "'" & ESPECIFICACIONES & "'" & "," & "'" & N_LATAS & "'" & "," & _
                        "'" & CANTIDAD_LATA & "'" & "," & "'" & LOTE & "'" & "," & fEnvasado & "," & _
                        fCaducidad & "," & "'" & CONDICIONES_ALMACENAMIENTO & "'" & "," & "'" & ESTADO_EMBALAJE & "'" & _
                ")"
        execute_bd consulta
        Insertar = RADMIN_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPINTURAS_RADMIN_CAR)", vbCritical, Err.Description
End Function
Public Function Modificar(oRADMIN As Long, oTIPO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PINTURAS_RADMIN_CAR SET " & _
                        " NOMBRE_BASE = '" & NOMBRE_BASE & "'," & _
                        " FABRICANTE = '" & fabricante & "'," & _
                        " ESPECIFICACIONES = '" & ESPECIFICACIONES & "'," & _
                        " N_LATAS = '" & N_LATAS & "'," & _
                        " CANTIDAD_LATA = '" & CANTIDAD_LATA & "'," & _
                        " LOTE = '" & LOTE & "'," & _
                        " FECHA_ENVASADO = '" & FECHA_ENVASADO & "'," & _
                        " FECHA_CADUCIDAD = '" & fecha_caducidad & "'," & _
                        " CONDICIONES_ALMACENAMIENTO = '" & CONDICIONES_ALMACENAMIENTO & "'," & _
                        " ESTADO_EMBALAJE = '" & ESTADO_EMBALAJE & "'," & _
                  " WHERE RADMIN_ID = " & oRADMIN & _
                  "   AND TIPO = " & oTIPO
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPINTURAS_RADMIN_CAR)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PINTURAS_RADMIN_CAR " & _
                "    WHERE RADMIN_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPINTURAS_RADMIN_CAR)", vbCritical, Err.Description
End Function
Public Function Listado(fdesde As String, fhasta As String, cCliente As Long, cEstado As Long) As ADODB.Recordset
        Dim consulta As String
        Dim filtro As String
        If cCliente <> 0 Then
            filtro = filtro & " AND A.TIPO = " & cCliente
        End If
        If cEstado <> 0 Then
            filtro = filtro & " AND A.ESTADO_ID = " & cEstado
        End If
            
        consulta = "SELECT A.RADMIN_ID, B.NOMBRE,A.FECHA_RECEPCION, D.DESCRIPCION,E.USUARIO,A.FS " & _
                   "  FROM PINTURAS_RADMIN_CAR A, CLIENTES B, decodificadora D, usuarios E " & _
                   " WHERE FECHA_RECEPCION >= '" & Format(fdesde, "yyyy-mm-dd") & " ' AND FECHA_RECEPCION <= '" & Format(fhasta, "yyyy-mm-dd") & "' " & _
                   "   AND A.TIPO = B.ID_CLIENTE " & _
                   "   AND A.ESTADO_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.PINTURAS_ESTADOS & _
                   "   AND A.USUARIO_ID = E.ID_EMPLEADO " & _
                   " ORDER BY A.RADMIN_ID DESC"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE clsPINTURAS_RADMIN_CAR
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

Public Function ListadoID(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.RADMIN_ID, B.NOMBRE,A.FECHA_RECEPCION, D.DESCRIPCION,E.USUARIO,A.FS " & _
                   "  FROM PINTURAS_RADMIN_CAR A, CLIENTES B, decodificadora D, usuarios E " & _
                   " WHERE A.TIPO = B.ID_CLIENTE " & _
                   "   AND A.ESTADO_ID = D.VALOR AND D.CODIGO = " & DECODIFICADORA.PINTURAS_ESTADOS & _
                   "   AND A.USUARIO_ID = E.ID_EMPLEADO " & _
                   "   AND A.RADMIN_ID = " & ID & _
                   " ORDER BY A.RADMIN_ID DESC"
        Set ListadoID = datos_bd(consulta)
End Function


