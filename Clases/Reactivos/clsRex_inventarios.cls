VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRex_inventarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSREX_INVENTARIOS
'***************************************************************
Private ID_INVENTARIO As Long
Private DESCRIPCION As String
Private fecha As String
Private CENTRO_ID As Long
Private USUARIO_ID As Long
Private ESTADO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSREX_INVENTARIOS
'***************************************************************
Public Property Let setID_INVENTARIO(ByVal dato As Long)
        ID_INVENTARIO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setCENTRO_ID(ByVal dato As Long)
        CENTRO_ID = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setESTADO_ID(ByVal dato As Long)
        ESTADO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSREX_INVENTARIOS
'***************************************************************
Public Property Get getID_INVENTARIO() As Long
        getID_INVENTARIO = ID_INVENTARIO
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getCENTRO_ID() As Long
        getCENTRO_ID = CENTRO_ID
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getESTADO_ID() As Long
        getESTADO_ID = ESTADO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSREX_INVENTARIOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM REX_INVENTARIOS WHERE ID_INVENTARIO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_INVENTARIO = rs("ID_INVENTARIO")
                DESCRIPCION = rs("DESCRIPCION")
                fecha = rs("FECHA")
                CENTRO_ID = rs("CENTRO_ID")
                USUARIO_ID = rs("USUARIO_ID")
                ESTADO_ID = rs("ESTADO_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRex_inventarios)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_INVENTARIO) FROM REX_INVENTARIOS"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_INVENTARIO = 1
        Else
                ID_INVENTARIO = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO REX_INVENTARIOS " & _
                   "  VALUES (" & _
                        ID_INVENTARIO & "," & "'" & DESCRIPCION & "'" & "," & "'" & fecha & "'" & "," & CENTRO_ID & "," & _
                        USUARIO_ID & "," & ESTADO_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_INVENTARIO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRex_inventarios)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE REX_INVENTARIOS SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FECHA = '" & fecha & "'," & _
                        " CENTRO_ID = " & CENTRO_ID & "," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " ESTADO_ID = " & ESTADO_ID & "" & _
                " WHERE ID_INVENTARIO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsRex_inventarios)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM REX_INVENTARIOS " & _
                "    WHERE ID_INVENTARIO = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRex_inventarios)", vbCritical, Err.Description
End Function
Public Function Listado(ldescripcion As String, fdesde As Date, fhasta As Date, USUARIO As String, CENTRO As String) As ADODB.Recordset
        Dim s As String
        If USUARIO <> "" Then
            s = s & " AND A.USUARIO_ID = " & CInt(USUARIO)
        End If
        If CENTRO <> "" Then
            s = s & " AND A.CENTRO_ID = " & CInt(CENTRO)
        End If
        Dim consulta As String
        consulta = "SELECT A.ID_INVENTARIO,A.DESCRIPCION,A.FECHA,C.NOMBRE,U.USUARIO,D.DESCRIPCION " & _
                   "  FROM REX_INVENTARIOS A, usuarios U, decodificadora D,centros C " & _
                   " WHERE A.USUARIO_ID = U.ID_EMPLEADO " & _
                   "   AND D.CODIGO = " & DECODIFICADORA.REX_INVENTARIOS_ESTADOS & _
                   "   AND D.VALOR = A.ESTADO_ID " & _
                   "   AND A.CENTRO_ID = C.ID_CENTRO " & _
                   "   AND A.DESCRIPCION LIKE '%" & ldescripcion & "%'" & _
                   "   AND A.FECHA >= '" & Format(fdesde, "yyyy-mm-dd") & "'" & _
                   "   AND A.FECHA <= '" & Format(fhasta, "yyyy-mm-dd") & "'" & _
                   s & _
                   " ORDER BY A.ID_INVENTARIO DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Sub Imprimir_Listado(ByVal prm_titulo As String, ByVal prm_id_inventario As String)

    Dim objRpt As New frmReport
    Dim arrNom() As String
    Dim arrVal() As String
    Dim strCriterio As String
                    
    ReDim arrNom(2)
    ReDim arrVal(2)

    arrNom(1) = "impresor"
    arrVal(1) = USUARIO.getNOMBRE & " " & USUARIO.getAPELLIDOS
    
    arrNom(2) = "titulo_inventario"
    arrVal(2) = prm_titulo
    
        
    With objRpt
        .iniciar
        .informe = "General\rptListadoInventarioReactivosExternos"
        .criterio = "{rex_inventarios_botes.INVENTARIO_ID} = " & prm_id_inventario
        .imprimir = False
        
        .ParametrosNombre = arrNom
        .ParametrosValores = arrVal

        .generar
        '.Visible = True
        .Show 1
    End With
    Unload objRpt
    Set objRpt = Nothing

End Sub
Public Function Listado_PK(PK As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_INVENTARIO,A.DESCRIPCION,A.FECHA,C.NOMBRE,U.USUARIO,D.DESCRIPCION " & _
                   "  FROM REX_INVENTARIOS A, usuarios U, decodificadora D,centros C " & _
                   " WHERE A.USUARIO_ID = U.ID_EMPLEADO " & _
                   "   AND D.CODIGO = " & DECODIFICADORA.REX_INVENTARIOS_ESTADOS & _
                   "   AND D.VALOR = A.ESTADO_ID " & _
                   "   AND A.CENTRO_ID = C.ID_CENTRO " & _
                   "   AND A.ID_INVENTARIO = " & PK
        Set Listado_PK = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_INVENTARIO,DESCRIPCION FROM REX_INVENTARIOS ORDER BY DESCRIPCION"
        Set Listado_Combo = datos_bd(consulta)
End Function
