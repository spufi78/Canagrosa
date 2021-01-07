VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOfertas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOFERTAS
'***************************************************************
Private ID_OFERTA As Long
Private numero As Long
Private fecha As String
Private CLIENTE_ID As Long
Private ESTADO_ID As Integer
Private Nif As String
Private NOMBRE As String
Private DIRECCION As String
Private POBLACION As String
Private AA As String
Private TELEFONO As String
Private FAX As String
Private EMAIL As String
Private OBRA_DOMICILIO As String
Private OBRA_POBLACION As String
Private OBRA_TIPO As Integer
Private FORMA_PAGO As String
Private AGENTE_ID As Long
Private hora As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOFERTAS
'***************************************************************
Public Property Let setID_OFERTA(ByVal dato As Long)
        ID_OFERTA = dato
End Property
Public Property Let setNUMERO(ByVal dato As Long)
        numero = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setCLIENTE_ID(ByVal dato As Long)
        CLIENTE_ID = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Integer)
        ESTADO_ID = dato
End Property
Public Property Let setNIF(ByVal dato As String)
        Nif = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
        NOMBRE = dato
End Property
Public Property Let setDIRECCION(ByVal dato As String)
        DIRECCION = dato
End Property
Public Property Let setPOBLACION(ByVal dato As String)
        POBLACION = dato
End Property
Public Property Let setAA(ByVal dato As String)
        AA = dato
End Property
Public Property Let setTELEFONO(ByVal dato As String)
        TELEFONO = dato
End Property
Public Property Let setFAX(ByVal dato As String)
        FAX = dato
End Property
Public Property Let setEMAIL(ByVal dato As String)
        EMAIL = dato
End Property
Public Property Let setOBRA_DOMICILIO(ByVal dato As String)
        OBRA_DOMICILIO = dato
End Property
Public Property Let setOBRA_POBLACION(ByVal dato As String)
        OBRA_POBLACION = dato
End Property
Public Property Let setOBRA_TIPO(ByVal dato As Integer)
        OBRA_TIPO = dato
End Property
Public Property Let setFORMA_PAGO(ByVal dato As String)
        FORMA_PAGO = dato
End Property
Public Property Let setAGENTE_ID(ByVal dato As Long)
        AGENTE_ID = dato
End Property
Public Property Let setHORA(ByVal dato As String)
        hora = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOFERTAS
'***************************************************************
Public Property Get getID_OFERTA() As Long
        getID_OFERTA = ID_OFERTA
End Property
Public Property Get getNUMERO() As Long
        getNUMERO = numero
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getCLIENTE_ID() As Long
        getCLIENTE_ID = CLIENTE_ID
End Property
Public Property Get getESTADO_ID() As Integer
        getESTADO_ID = ESTADO_ID
End Property
Public Property Get getNIF() As String
        getNIF = Nif
End Property
Public Property Get getNOMBRE() As String
        getNOMBRE = NOMBRE
End Property
Public Property Get getDIRECCION() As String
        getDIRECCION = DIRECCION
End Property
Public Property Get getPOBLACION() As String
        getPOBLACION = POBLACION
End Property
Public Property Get getAA() As String
        getAA = AA
End Property
Public Property Get getTELEFONO() As String
        getTELEFONO = TELEFONO
End Property
Public Property Get getFAX() As String
        getFAX = FAX
End Property
Public Property Get getEMAIL() As String
        getEMAIL = EMAIL
End Property
Public Property Get getOBRA_DOMICILIO() As String
        getOBRA_DOMICILIO = OBRA_DOMICILIO
End Property
Public Property Get getOBRA_POBLACION() As String
        getOBRA_POBLACION = OBRA_POBLACION
End Property
Public Property Get getOBRA_TIPO() As Integer
        getOBRA_TIPO = OBRA_TIPO
End Property
Public Property Get getFORMA_PAGO() As String
        getFORMA_PAGO = FORMA_PAGO
End Property
Public Property Get getAGENTE_ID() As Long
        getAGENTE_ID = AGENTE_ID
End Property
Public Property Get getHORA() As String
        getHORA = hora
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOFERTAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OFERTAS WHERE ID_OFERTA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_OFERTA = rs("ID_OFERTA")
                numero = rs("NUMERO")
                fecha = rs("FECHA")
                CLIENTE_ID = rs("CLIENTE_ID")
                ESTADO_ID = rs("ESTADO_ID")
                Nif = rs("NIF")
                NOMBRE = rs("NOMBRE")
                DIRECCION = rs("DIRECCION")
                POBLACION = rs("POBLACION")
                AA = rs("AA")
                TELEFONO = rs("TELEFONO")
                FAX = rs("FAX")
                EMAIL = rs("EMAIL")
                OBRA_DOMICILIO = rs("OBRA_DOMICILIO")
                OBRA_POBLACION = rs("OBRA_POBLACION")
                OBRA_TIPO = rs("OBRA_TIPO")
                FORMA_PAGO = rs("FORMA_PAGO")
                AGENTE_ID = rs("AGENTE_ID")
                hora = rs("HORA")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsOfertas)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_OFERTA) FROM OFERTAS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_OFERTA = 1
        Else
                ID_OFERTA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function CrearNUMERO()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(NUMERO) FROM OFERTAS WHERE YEAR(FECHA) = " & Year(fecha)
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                numero = 1
        Else
                numero = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        If ID_OFERTA = 0 Then
            Me.CrearID
        End If
        If numero = 0 Then
            Me.CrearNUMERO
        End If
        consulta = "INSERT INTO OFERTAS " & _
                   "  VALUES (" & _
                        ID_OFERTA & "," & numero & "," & "'" & fecha & "'" & "," & _
                        CLIENTE_ID & "," & ESTADO_ID & ",'" & Nif & "'" & "," & "'" & NOMBRE & "'" & "," & _
                        "'" & DIRECCION & "'" & "," & "'" & POBLACION & "'" & "," & "'" & AA & "'" & "," & _
                        "'" & TELEFONO & "'" & "," & "'" & FAX & "'" & "," & "'" & EMAIL & "'" & "," & _
                        "'" & OBRA_DOMICILIO & "'" & "," & "'" & OBRA_POBLACION & "'" & "," & OBRA_TIPO & ",'" & FORMA_PAGO & "'" & "," & _
                        AGENTE_ID & "," & "'" & hora & "'" & _
                ")"
        execute_bd consulta
        Insertar = ID_OFERTA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsOfertas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OFERTAS SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " CLIENTE_ID = " & CLIENTE_ID & "," & _
                        " ESTADO_ID = " & ESTADO_ID & "," & _
                        " NIF = '" & Nif & "'," & _
                        " NOMBRE = '" & NOMBRE & "'," & _
                        " DIRECCION = '" & DIRECCION & "'," & _
                        " POBLACION = '" & POBLACION & "'," & _
                        " AA = '" & AA & "'," & _
                        " TELEFONO = '" & TELEFONO & "'," & _
                        " FAX = '" & FAX & "'," & _
                        " EMAIL = '" & EMAIL & "'," & _
                        " OBRA_DOMICILIO = '" & OBRA_DOMICILIO & "'," & _
                        " OBRA_POBLACION = '" & OBRA_POBLACION & "'," & _
                        " OBRA_TIPO = " & OBRA_TIPO & "," & _
                        " FORMA_PAGO = '" & FORMA_PAGO & "'," & _
                        " AGENTE_ID = " & AGENTE_ID & "," & _
                        " HORA = '" & hora & "'" & _
                " WHERE ID_OFERTA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsOfertas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM OFERTAS " & _
                "    WHERE ID_OFERTA = " & ID
        execute_bd consulta
        consulta = "DELETE FROM OFERTAS_DETALLE " & _
                "    WHERE OFERTA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsOfertas)", vbCritical, Err.Description
End Function
Public Function Listado(cliente As String, OBRA As String, numero As String, agente As String, anno As Integer, ESTADO As String) As ADODB.Recordset
        Dim consulta As String
        Dim s As String
        If cliente <> "" Then
            s = s & " AND O.NOMBRE LIKE '%" & cliente & "%'"
        End If
        If OBRA <> "" Then
            s = s & " AND O.OBRA_DOMICILIO LIKE '%" & OBRA & "%'"
        End If
        If numero <> "" Then
            If IsNumeric(numero) Then
                s = s & " AND O.NUMERO = " & numero
            End If
        End If
        If agente <> "" Then
            If IsNumeric(agente) Then
                s = s & " AND O.AGENTE_ID = " & agente
            End If
        End If
        If ESTADO <> "" Then
            If IsNumeric(ESTADO) Then
                If ESTADO <> 0 Then
                    s = s & " AND O.ESTADO_ID = " & ESTADO
                End If
            End If
        End If
        
        consulta = "SELECT O.ID_OFERTA, O.NUMERO, O.FECHA,O.NOMBRE, O.OBRA_DOMICILIO, C.NOMBRE, D.DESCRIPCION " & _
                   "  FROM OFERTAS O, COMERCIALES C, DECODIFICADORA D " & _
                   " WHERE O.AGENTE_ID = C.ID_COMERCIAL " & _
                   "   AND YEAR(FECHA) = " & anno & _
                   "   AND D.CODIGO = " & DECODIFICADORA.D_OFERTAS_ESTADOS & _
                   "   AND O.ESTADO_ID = D.VALOR " & _
                   s & _
                   " ORDER BY NUMERO DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_ID(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT O.ID_OFERTA, O.NUMERO, O.FECHA,O.NOMBRE, O.OBRA_DOMICILIO, C.NOMBRE, D.DESCRIPCION " & _
                   "  FROM OFERTAS O, COMERCIALES C, DECODIFICADORA D" & _
                   " WHERE O.AGENTE_ID = C.ID_COMERCIAL " & _
                   "   AND D.CODIGO = " & DECODIFICADORA.D_OFERTAS_ESTADOS & _
                   "   AND O.ESTADO_ID = D.VALOR " & _
                   "   AND O.ID_OFERTA = " & ID
        Set Listado_ID = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_OFERTA,NUMERO FROM OFERTAS ORDER BY NUMERO"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Sub imprimir(OFERTA As Long, impresora As Boolean, logo As Boolean, copias As Integer, Optional pdf As String = "")
    Dim tipo As String
   On Error GoTo Imprimir_Error

    Dim aux As String
'    Dim p1() As String
'    Dim p2() As String
            
'    ReDim p1(1) As String
'    ReDim p2(1) As String
'    p1(1) = "LOGO"
'    If logo Then
'        p2(1) = "S"
'    Else
'        p2(1) = "N"
'    End If
    With frmReport
        .informe = "rptOferta"
'        .ParametrosNombre = p1
'        .ParametrosValores = p2
        .CRITERIO = "{OFERTAS.ID_OFERTA}=" & OFERTA & aux
        If pdf <> "" Then
            .pdf = pdf
        End If
        .imprimir = impresora
        .copias = copias
        .generar
        
        If pdf = "" And impresora = False Then
            .Show 1
        End If
    End With
    Unload frmReport
    
   On Error GoTo 0
   Exit Sub

Imprimir_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Imprimir of Módulo de clase clsDocumentos"
End Sub

Public Function Correo(OFERTA As Long, impresora As Boolean, logo As Boolean, copias As Integer) As Boolean
    
    Dim oOferta As New clsOfertas
   On Error GoTo correo_Error

    With oOferta
    If .Carga(OFERTA) Then
        Dim NOMBRE As String
        NOMBRE = "Oferta " & Year(.getFECHA) & "-" & Format(.getNUMERO, "000000") & ".pdf"
        Dim destino_documento As String
        On Error Resume Next
        MkDir App.Path & "\tmp\"
        destino_documento = App.Path & "\tmp\" & NOMBRE
        If Dir(destino_documento) <> "" Then
            Kill destino_documento
        End If
   On Error GoTo correo_Error
        ' Generamos el pdf
        Me.imprimir OFERTA, impresora, logo, copias, destino_documento
        
        Dim ref As String
        If Dir(destino_documento) = "" Then
            MsgBox "El pdf de la oferta no se ha generado correctamente.", vbInformation, App.Title
            Correo = False
            Exit Function
        End If
        
        ref = "Adjunto Oferta número " & Year(.getFECHA) & "-" & Format(.getNUMERO, "000000") & ".pdf"
        genera_correo_outlook oOferta.getEMAIL, ref, "", destino_documento, 0
    End If
    End With
    Set oOferta = Nothing
    Correo = True

   On Error GoTo 0
   Exit Function

correo_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure correo of Módulo de clase clsDocumentos"
End Function

