VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDescuentos_documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSDESCUENTOS_DOCUMENTOS
'***************************************************************
Private DESCUENTO_ID As Long
Private APUNTE_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSDESCUENTOS_DOCUMENTOS
'***************************************************************
Public Property Let setDESCUENTO_ID(ByVal dato As Long)
        DESCUENTO_ID = dato
End Property
Public Property Let setAPUNTE_ID(ByVal dato As Long)
        APUNTE_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSDESCUENTOS_DOCUMENTOS
'***************************************************************
Public Property Get getDESCUENTO_ID() As Long
        getDESCUENTO_ID = DESCUENTO_ID
End Property
Public Property Get getAPUNTE_ID() As Long
        getAPUNTE_ID = APUNTE_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSDESCUENTOS_DOCUMENTOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DESCUENTOS_DOCUMENTOS WHERE DESCUENTO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DESCUENTO_ID = rs("DESCUENTO_ID")
                APUNTE_ID = rs("APUNTE_ID")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsDescuentos_documentos)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO DESCUENTOS_DOCUMENTOS " & _
                   "  VALUES (" & _
                        DESCUENTO_ID & "," & APUNTE_ID & _
                ")"
        execute_bd consulta
        Insertar = DESCUENTO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsDescuentos_documentos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DESCUENTOS_DOCUMENTOS SET " & _
                        " APUNTE_ID = " & APUNTE_ID & "" & _
                " WHERE DESCUENTO_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsDescuentos_documentos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.Recordset
        Set rs = datos_bd("select apunte_id from descuentos_documentos where descuento_id = " & ID)
        If rs.RecordCount > 0 Then
            Do
                Dim ORD As New clsRemesas_documentos
                ORD.Carga rs(0)
                If ORD.getDOCUMENTO_ID <> 0 Then
                    execute_bd "UPDATE DOCUMENTOS_RECIBOS SET COBRADO = " & ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_REMESA & " WHERE DOCUMENTO_ID = " & ORD.getDOCUMENTO_ID & " AND VENCIMIENTO = " & ORD.getVENCIMIENTO
                End If
                execute_bd " UPDATE REMESAS_DOCUMENTOS SET SITUACION = 'R' WHERE ID = " & rs(0)
                rs.MoveNext
            Loop Until rs.EOF
        End If
        execute_bd "DELETE FROM DESCUENTOS_DOCUMENTOS WHERE DESCUENTO_ID = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsDescuentos_documentos)", vbCritical, Err.Description
End Function
Public Function Listado(ID_DESCUENTO As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DESCUENTOS_DOCUMENTOS WHERE DESCUENTO_ID = " & ID_DESCUENTO & " ORDER BY DESCUENTO_ID"
        Set Listado = datos_bd(consulta)
End Function
