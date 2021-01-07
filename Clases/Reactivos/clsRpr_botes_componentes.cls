VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRpr_botes_componentes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSRPR_BOTES_COMPONENTES
'***************************************************************
Private BOTE_PR_ID As Long
Private BOTE_EX_ID As Long
Private TIPO As String
Private cantidad As String
Private PROCEDIMIENTO As String
Private OBSERVACIONES As String
Private ORDEN As Long
Private EMPLEADO_ID As Long
Private FTIMESTAMP As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSRPR_BOTES_COMPONENTES
'***************************************************************
Public Property Let setBOTE_PR_ID(ByVal dato As Long)
        BOTE_PR_ID = dato
End Property
Public Property Let setBOTE_EX_ID(ByVal dato As Long)
        BOTE_EX_ID = dato
End Property
Public Property Let setTIPO(ByVal dato As String)
        TIPO = dato
End Property
Public Property Let setCANTIDAD(ByVal dato As String)
        cantidad = dato
End Property
Public Property Let setPROCEDIMIENTO(ByVal dato As String)
        PROCEDIMIENTO = dato
End Property
Public Property Let setOBSERVACIONES(ByVal dato As String)
        OBSERVACIONES = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
Public Property Let setFTIMESTAMP(ByVal dato As String)
        FTIMESTAMP = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSRPR_BOTES_COMPONENTES
'***************************************************************
Public Property Get getBOTE_PR_ID() As Long
        getBOTE_PR_ID = BOTE_PR_ID
End Property
Public Property Get getBOTE_EX_ID() As Long
        getBOTE_EX_ID = BOTE_EX_ID
End Property
Public Property Get getTIPO() As String
        getTIPO = TIPO
End Property
Public Property Get getCANTIDAD() As String
        getCANTIDAD = cantidad
End Property
Public Property Get getPROCEDIMIENTO() As String
        getPROCEDIMIENTO = PROCEDIMIENTO
End Property
Public Property Get getOBSERVACIONES() As String
        getOBSERVACIONES = OBSERVACIONES
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Property Get getFTIMESTAMP() As String
        getFTIMESTAMP = FTIMESTAMP
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSRPR_BOTES_COMPONENTES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM RPR_BOTES_COMPONENTES WHERE BOTE_PR_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                BOTE_PR_ID = rs("BOTE_PR_ID")
                BOTE_EX_ID = rs("BOTE_EX_ID")
                TIPO = rs("TIPO")
                cantidad = rs("CANTIDAD")
                PROCEDIMIENTO = rs("PROCEDIMIENTO")
                OBSERVACIONES = rs("OBSERVACIONES")
                ORDEN = rs("ORDEN")
                EMPLEADO_ID = rs("EMPLEADO_ID")
                FTIMESTAMP = rs("FTIMESTAMP")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRpr_botes_componentes)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO RPR_BOTES_COMPONENTES " & _
                   "  VALUES (" & _
                        BOTE_PR_ID & "," & BOTE_EX_ID & "," & "'" & TIPO & "','" & cantidad & "'," & _
                        "'" & PROCEDIMIENTO & "','" & OBSERVACIONES & "'" & "," & ORDEN & "," & usuario.getID_EMPLEADO & ",current_timestamp)"
        execute_bd consulta
        Insertar = BOTE_PR_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRpr_botes_componentes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE RPR_BOTES_COMPONENTES SET " & _
                        " BOTE_EX_ID = " & BOTE_EX_ID & "," & _
                        " CANTIDAD = '" & cantidad & "'," & _
                        " TIPO = '" & TIPO & "'," & _
                        " PROCEDIMIENTO = '" & PROCEDIMIENTO & "'," & _
                        " OBSERVACIONES = '" & OBSERVACIONES & "'," & _
                        " ORDEN = " & ORDEN & "," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "," & _
                        " FTIMESTAMP = '" & FTIMESTAMP & "'" & _
                " WHERE BOTE_PR_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsRpr_botes_componentes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM RPR_BOTES_COMPONENTES " & _
                "    WHERE BOTE_PR_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRpr_botes_componentes)", vbCritical, Err.Description
End Function
Public Function Listado(BOTE As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM RPR_BOTES_COMPONENTES WHERE BOTE_PR_ID = " & BOTE & " ORDER BY ORDEN"
    
'        consulta = "SELECT be.id_bote_ex, " & _
'                   "       rpr.observaciones, " & _
'                   "       be.tipo_bote_ex_id,tr.ID_TIPO_REACTIVO_EX,rpr.cantidad,rpr.procedimiento,pro.nombre " & _
'                   " FROM BOTES_EX be, " & _
'                   "      TIPOS_BOTE_EX tb, " & _
'                   "      proveedores pro, " & _
'                   "      TIPOS_REACTIVO_EX tr, " & _
'                   "      RPR_BOTES_COMPONENTES rpr " & _
'                   " WHERE be.tipo_bote_ex_id = tb.id_tipo_bote_ex " & _
'                   "   AND tb.tipo_reactivo_ex_id = tr.id_tipo_reactivo_ex " & _
'                   "   and tb.proveedor_id = pro.id_proveedor " & _
'                   "   and rpr.bote_pr_id = " & BOTE & _
'                   "   and rpr.bote_ex_id = be.id_bote_ex " & _
'                   " ORDER BY be.id_bote_ex desc"
        
        Set Listado = datos_bd(consulta)
End Function
