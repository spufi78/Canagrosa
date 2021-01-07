VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDocs_pago_conceptos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_CONCEPTO As Long
Private DOC_ID As Long
Private ALBARAN_ID As Long
Private ABONADO As Integer
Private DESCRIPCION As String
Private fecha As String
Private CANTIDAD As String
Private PRECIO As String
Private SUBTOTAL As String
Private dto As String
Private total As String
Private FAMILIA_ID As Integer
Private apartado As Integer

Public Property Let setID_CONCEPTO(ByVal dato As Long)
    ID_CONCEPTO = dato
End Property
Public Property Let setDOC_ID(ByVal dato As Long)
    DOC_ID = dato
End Property
Public Property Let setALBARAN_ID(ByVal dato As Long)
    ALBARAN_ID = dato
End Property
Public Property Let setABONADO(ByVal dato As Integer)
    ABONADO = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = Trim(dato)
End Property
Public Property Let setFECHA(ByVal dato As String)
    fecha = Trim(dato)
End Property
Public Property Let setCANTIDAD(ByVal dato As String)
    CANTIDAD = Trim(dato)
End Property
Public Property Let setPRECIO(ByVal dato As String)
    PRECIO = Trim(dato)
End Property
Public Property Let setSUBTOTAL(ByVal dato As String)
    SUBTOTAL = Trim(dato)
End Property
Public Property Let setDTO(ByVal dato As String)
    dto = Trim(dato)
End Property
Public Property Let setTOTAL(ByVal dato As String)
    total = Trim(dato)
End Property
Public Property Let setFAMILIA_ID(ByVal dato As Integer)
    FAMILIA_ID = Trim(dato)
End Property
Public Property Let setAPARTADO(ByVal dato As Integer)
    apartado = Trim(dato)
End Property



Public Property Get getID_CONCEPTO() As Long
    getID_CONCEPTO = ID_CONCEPTO
End Property
Public Property Get getDOC_ID() As Long
    getDOC_ID = DOC_ID
End Property
Public Property Get getALBARAN_ID() As Long
    getALBARAN_ID = ALBARAN_ID
End Property
Public Property Get getABONADO() As Integer
    getABONADO = ABONADO
End Property
Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFECHA() As String
    getFECHA = fecha
End Property
Public Property Get getCANTIDAD() As String
    getCANTIDAD = CANTIDAD
End Property
Public Property Get getPRECIO() As String
    getPRECIO = PRECIO
End Property
Public Property Get getSUBTOTAL() As String
    getSUBTOTAL = SUBTOTAL
End Property
Public Property Get getDTO() As String
    getDTO = dto
End Property
Public Property Get getTOTAL() As String
    getTOTAL = total
End Property
Public Property Get getFAMILIA_ID() As Integer
    getFAMILIA_ID = FAMILIA_ID
End Property
Public Property Get getAPARTADO() As Integer
    getAPARTADO = apartado
End Property
Public Sub CrearID_CONCEPTO()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT max(ID_CONCEPTO) FROM docs_pago_conceptos"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_CONCEPTO = 1
    Else
        ID_CONCEPTO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Boolean
    Dim consulta As String
    On Error GoTo fallo
    Me.CrearID_CONCEPTO
    consulta = "INSERT INTO docs_pago_conceptos VALUES (" & _
               ID_CONCEPTO & "," & DOC_ID & "," & ALBARAN_ID & "," & ABONADO & ",'" & DESCRIPCION & "','" & _
               fecha & "','" & CANTIDAD & "','" & PRECIO & "','" & SUBTOTAL & "','" & dto & "','" & total & "'," & FAMILIA_ID & "," & apartado & ")"
    execute_bd consulta
    Insertar = True
    Exit Function
fallo:
    Insertar = False
    MsgBox "Error al guardar el documento de pago de conceptos.", vbCritical, Err.Description
End Function

Public Function NumeroConceptosDocumento(num_documento As Integer) As Integer
    On Error GoTo fallo
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT count(*) FROM docs_pago_conceptos WHERE DOC_ID=" & num_documento
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
        NumeroConceptosDocumento = 0
    Else
        NumeroConceptosDocumento = rs.Fields(0)
    End If
    Set rs = Nothing
    Exit Function
fallo:
    NumeroConceptosDocumento = 0
    MsgBox "Error al cargar el numero de conceptos (NumeroConceptosDocumento)", vbCritical, Err.Description
End Function

Public Function ConceptosDocumento(num_documento As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT * FROM docs_pago_conceptos WHERE DOC_ID=" & num_documento & " ORDER BY id_concepto"
    Set ConceptosDocumento = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al cargar los conceptos del documento (ConceptosDocumento)", vbCritical, Err.Description
End Function

Public Function ConceptosListado(ID_DOC As Long) As ADODB.Recordset
    On Error GoTo fallo
    Dim consulta As String
    consulta = "SELECT A.FAMILIA_ID,A.ALBARAN_ID,A.APARTADO,A.FECHA,replace(A.DESCRIPCION,'\r\n',''),B.NOMBRE,A.PRECIO, " & _
               "       A.CANTIDAD,A.SUBTOTAL,A.DTO,A.TOTAL " & _
               "  FROM docs_pago_conceptos A " & _
               "  LEFT JOIN FAMILIAS B ON A.FAMILIA_ID = B.ID_FAMILIA " & _
               " WHERE A.DOC_ID=" & ID_DOC & _
               " ORDER BY id_concepto"
    Set ConceptosListado = datos_bd(consulta)
    Exit Function
fallo:
    MsgBox "Error al cargar los conceptos del documento (ConceptosListado)", vbCritical, Err.Description
End Function

Public Function EliminarConceptos(num_documento As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "DELETE FROM docs_pago_conceptos WHERE DOC_ID=" & num_documento
    execute_bd consulta
    Exit Function
    EliminarConceptos = True
fallo:
    EliminarConceptos = False
    MsgBox "Error al eliminar los conceptos del documento (EliminarConceptos)", vbCritical, Err.Description
End Function

Public Function ImporteDocumento(num_documento As Long) As Currency
    On Error GoTo fallo
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT SUM(TOTAL) FROM docs_pago_conceptos WHERE DOC_ID=" & num_documento
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
        ImporteDocumento = 0
    Else
        ImporteDocumento = rs.Fields(0)
    End If
    Set rs = Nothing
    Exit Function
fallo:
    ImporteDocumento = 0
    MsgBox "Error al calcular el importe del documento de conceptos (ImporteDocumento)", vbCritical, Err.Description
End Function
Public Function abonar_concepto(documento As Long, concepto As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE docs_pago_conceptos SET abonado=1 " & _
               " WHERE DOC_ID=" & documento & " AND ID_CONCEPTO= " & concepto
    execute_bd consulta
    Exit Function
    abonar_concepto = True
fallo:
    abonar_concepto = False
    MsgBox "Error al abonar el concepto.", vbCritical, Err.Description
End Function
