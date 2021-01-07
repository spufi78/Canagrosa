VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCa_documentos_PNT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCA_DOCUMENTOS_PNT
'***************************************************************
Private DOCUMENTO_ID As Long
Private PNT_ID As Long
Private VINCULADO As Integer
Private ORDEN As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCA_DOCUMENTOS_PNT
'***************************************************************
Public Property Let setDOCUMENTO_ID(ByVal dato As Long)
        DOCUMENTO_ID = dato
End Property
Public Property Let setPNT_ID(ByVal dato As Long)
        PNT_ID = dato
End Property
Public Property Let setVINCULADO(ByVal dato As Integer)
        VINCULADO = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCA_DOCUMENTOS_PNT
'***************************************************************
Public Property Get getDOCUMENTO_ID() As Long
        getDOCUMENTO_ID = DOCUMENTO_ID
End Property
Public Property Get getPNT_ID() As Long
        getPNT_ID = PNT_ID
End Property
Public Property Get getVINCULADO() As Integer
        getVINCULADO = VINCULADO
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCA_DOCUMENTOS_PNT
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM ca_documentos_pnt WHERE DOCUMENTO_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                DOCUMENTO_ID = rs("DOCUMENTO_ID")
                PNT_ID = rs("PNT_ID")
                VINCULADO = rs("VINCULADO")
                ORDEN = rs("ORDEN")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCA_DOCUMENTOS_PNT)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO ca_documentos_pnt " & _
                   "  VALUES (" & _
                        DOCUMENTO_ID & "," & PNT_ID & "," & VINCULADO & "," & ORDEN & _
                ")"
        execute_bd consulta
        Insertar = DOCUMENTO_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCA_DOCUMENTOS_PNT)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long, PNT As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE ca_documentos_pnt SET " & _
                        " VINCULADO = " & VINCULADO & _
                " WHERE DOCUMENTO_ID = " & ID & _
                "   AND PNT_ID = " & PNT
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsCA_DOCUMENTOS_PNT)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long, PNT As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM ca_documentos_pnt " & _
                "    WHERE DOCUMENTO_ID = " & ID & " AND PNT_ID = " & PNT
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCA_DOCUMENTOS_PNT)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.PNT_ID,CONCAT('(',B.CODIGO,') ',B.NOMBRE),B.RUTA,A.VINCULADO " & _
                   "  FROM ca_documentos_pnt A, ca_documentos B  " & _
                   " WHERE A.DOCUMENTO_ID  = " & ID & _
                   "   AND A.PNT_ID = B.ID_DOCUMENTO " & _
                   " ORDER BY ORDEN"
        Set Listado = datos_bd(consulta)
End Function


