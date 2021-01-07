VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSc_paquetes_detalle_generico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSC_PAQUETES_DETALLE_GENERICO
'***************************************************************
Private ID_PAQ_CONCEPTO As Long
Private PAQUETE_ID As Long
Private REFERENCIA As String
Private DESCRIPCION As String
Private unidades As Long
Private DESCUENTO As Long
Private PRECIO As String
Private IMPORTE As String
Private CUSERID As Long
Private FTIMESTP As String
Private EDICION As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSC_PAQUETES_DETALLE_GENERICO
'***************************************************************
Public Property Let setID_PAQ_CONCEPTO(ByVal dato As Long)
        ID_PAQ_CONCEPTO = dato
End Property
Public Property Let setPAQUETE_ID(ByVal dato As Long)
        PAQUETE_ID = dato
End Property
Public Property Let setREFERENCIA(ByVal dato As String)
        REFERENCIA = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setUNIDADES(ByVal dato As Long)
        unidades = dato
End Property
Public Property Let setDESCUENTO(ByVal dato As Long)
        DESCUENTO = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
Public Property Let setIMPORTE(ByVal dato As String)
        IMPORTE = dato
End Property
Public Property Let setCUSERID(ByVal dato As Long)
        CUSERID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
        FTIMESTP = dato
End Property
Public Property Let setEDICION(ByVal dato As Long)
        EDICION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSC_PAQUETES_DETALLE_GENERICO
'***************************************************************
Public Property Get getID_PAQ_CONCEPTO() As Long
        getID_PAQ_CONCEPTO = ID_PAQ_CONCEPTO
End Property
Public Property Get getPAQUETE_ID() As Long
        getPAQUETE_ID = PAQUETE_ID
End Property
Public Property Get getREFERENCIA() As String
        getREFERENCIA = REFERENCIA
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getUNIDADES() As String
        getUNIDADES = unidades
End Property
Public Property Get getDESCUENTO() As String
        getDESCUENTO = DESCUENTO
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
Public Property Get getIMPORTE() As String
        getIMPORTE = IMPORTE
End Property
Public Property Get getCUSERID() As Long
        getCUSERID = CUSERID
End Property
Public Property Get getFTIMESTP() As String
        getFTIMESTP = FTIMESTP
End Property
Public Property Get getEDICION() As Long
        getEDICION = EDICION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSC_PAQUETES_DETALLE_GENERICO
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SC_PAQUETES_DETALLE_GENERICO WHERE ID_PAQ_CONCEPTO = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PAQ_CONCEPTO = rs("ID_PAQ_CONCEPTO")
                PAQUETE_ID = rs("PAQUETE_ID")
                REFERENCIA = rs("REFERENCIA")
                DESCRIPCION = rs("DESCRIPCION")
                unidades = rs("UNIDADES")
                DESCUENTO = rs("DESCUENTO")
                PRECIO = rs("PRECIO")
                IMPORTE = rs("IMPORTE")
                CUSERID = rs("CUSERID")
                FTIMESTP = rs("FTIMESTP")
                EDICION = rs("EDICION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSc_paquetes_detalle_generico)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        CrearID
        consulta = "INSERT INTO SC_PAQUETES_DETALLE_GENERICO " & _
                   "  VALUES (" & ID_PAQ_CONCEPTO & "," & _
                        PAQUETE_ID & ",'" & REFERENCIA & "'," & "'" & DESCRIPCION & "'" & "," & unidades & "," & _
                        DESCUENTO & ",'" & PRECIO & "'" & ",'" & IMPORTE & "'," & USUARIO.getID_EMPLEADO & ", CURRENT_TIMESTAMP, " & EDICION & _
                ",null) ON DUPLICATE KEY UPDATE DESCRIPCION = '" & DESCRIPCION & "', UNIDADES = '" & unidades & "', PRECIO = '" & PRECIO & "', CUSERID = " & USUARIO.getID_EMPLEADO & ", FTIMESTP = CURRENT_TIMESTAMP"
        execute_bd consulta
        Insertar = PAQUETE_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSc_paquetes_detalle_generico)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SC_PAQUETES_DETALLE_GENERICO " & _
                "    WHERE ID_PAQ_CONCEPTO = " & ID & " AND EDICION = " & EDICION
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSc_paquetes_detalle_generico)", vbCritical, Err.Description
End Function
'M1163-I
Public Function Eliminar_Paquete(PK As Long, EDICION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SC_PAQUETES_DETALLE_GENERICO " & _
                "    WHERE PAQUETE_ID = " & PK & " AND EDICION = " & EDICION
        execute_bd consulta
        Eliminar_Paquete = True
        Exit Function
fallo:
        Eliminar_Paquete = False
        MsgBox "Error al Eliminar_Paquete (clsSc_paquetes_detalle_generico)", vbCritical, Err.Description
End Function
'M1163-F

Public Function Listado(stranno As String, strRefCliente As String, strContrata As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SC_PAQUETES_DETALLE_GENERICO ORDER BY ID_PAQ_CONCEPTO"
        Set Listado = datos_bd(consulta)
End Function

'M1147-I
Public Function Eliminar_Concepto(PAQUETE As Long, EDICION As Long, concepto As Long) As Boolean
        On Error GoTo falloEliminar_Concepto
        Dim consulta As String

        consulta = "DELETE FROM SC_PAQUETES_DETALLE_GENERICO " & _
                   " WHERE PAQUETE_ID = " & PAQUETE & _
                   " AND EDICION = " & EDICION & _
                   "   AND REFERENCIA = '" & concepto & "'"
        execute_bd consulta
        
        Eliminar_Concepto = True

        Exit Function
falloEliminar_Concepto:
        Eliminar_Concepto = False
        MsgBox "Error al Eliminar_Concepto (clsSC_Paquetes_Detalle)", vbCritical, Err.Description
End Function

Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    Dim consulta As String
    consulta = "SELECT DISTINCT ID_PAQ_CONCEPTO, CONCAT('(' , REFERENCIA , ')',' - Desc.: ',DESCRIPCION, ' - ', 'Precio: ' , PRECIO) AS LINEA " & _
               "  FROM SC_PAQUETES_DETALLE_GENERICO "
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "SC_PAQUETES_DETALLE_GENERICO"
        .setDESCRIPCION = "Conceptos"
        .setPK = "ID_PAQ_CONCEPTO"
        .setCAMPO = "LINEA"
        .setMUESTRA_DETALLE = True
        .setQUERY = consulta
        Set .FORMULARIO = FORMULARIO
    End With
End Function

Private Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    On Error GoTo fallo
    ' Recuperamos el maximo secuencial de muestra
    consulta = "SELECT MAX(ID_PAQ_CONCEPTO) FROM SC_PAQUETES_DETALLE_GENERICO "

    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_PAQ_CONCEPTO = 1
    Else
        ID_PAQ_CONCEPTO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
    Exit Sub
fallo:
    MsgBox "Error en CrearId : " & Err.Description, vbCritical, App.Title
End Sub

Public Function Listado_conceptos(ID_P As String, EDICION As Long) As ADODB.Recordset
    On Error GoTo fallo
    
    Dim consulta As String
    
    consulta = "SELECT P.ID_PAQUETE, P.CODIGO_SC " & _
               "     , PD.REFERENCIA, PD.DESCRIPCION, PD.UNIDADES, PD.DESCUENTO, PD.PRECIO, PD.IMPORTE " & _
               "  FROM SC_PAQUETES P, SC_PAQUETES_DETALLE_GENERICO PD " & _
               " WHERE P.ID_PAQUETE = PD.PAQUETE_ID " & _
               " AND P.EDICION = PD.EDICION " & _
               "   AND P.ID_PAQUETE = " & ID_P & _
               "   AND P.EDICION = " & EDICION & _
               " ORDER BY PD.ID_PAQ_CONCEPTO "
    Set Listado_conceptos = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox Err.Description, vbCritical, Err.Number
End Function
'M1147-F

