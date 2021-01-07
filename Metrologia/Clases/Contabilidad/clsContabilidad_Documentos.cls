VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad_Documentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsContabilidad_Documentos
'***************************************************************
Private Documento As String
Private Debe As String
Private Haber As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsContabilidad_Documentos
'***************************************************************
Public Property Let setDOCUMENTO(ByVal dato As String)
        Documento = dato
End Property
Public Property Let setDEBE(ByVal dato As String)
        Debe = dato
End Property
Public Property Let setHABER(ByVal dato As String)
        Haber = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsContabilidad_Documentos
'***************************************************************
Public Property Get getDOCUMENTO() As String
        getDOCUMENTO = Documento
End Property
Public Property Get getDEBE() As String
        getDEBE = Debe
End Property
Public Property Get getHABER() As String
        getHABER = Haber
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsContabilidad_Documentos
'***************************************************************
Public Function Carga(ID As String) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS WHERE Documento = '" & ID & "'"
        Dim oC As New clsContabilidad
        Set rs = oC.datos_bd_contabilidad(consulta)
        Set oC = Nothing
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                Documento = rs("Documento")
                Debe = rs("Debe")
                Haber = rs("Haber")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsContabilidad_Documentos)", vbCritical, Err.Description
End Function
Public Function Actualizar(DOC As String, IMPORTE As Currency) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        If Me.Carga(DOC) Then
            Debe = Debe + IMPORTE
            Haber = Haber + (-1 * IMPORTE)
            Me.Modificar DOC
        Else
            Documento = DOC
            Debe = IMPORTE
            Haber = -1 * IMPORTE
            Me.Insertar
        End If
        Actualizar = True
        Exit Function
fallo:
        Actualizar = False
        MsgBox "Error al actualizar (clsContabilidad_Documentos)", vbCritical, Err.Description
End Function

Public Function Insertar() As String
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO DOCUMENTOS " & _
                   "  VALUES (" & _
                        "'" & Documento & "'" & "," & "'" & Debe & "'" & "," & "'" & Haber & "'" & _
                ")"
        Dim oC As New clsContabilidad
        oC.execute_bd consulta
        Set oC = Nothing
        Insertar = Documento
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsContabilidad_Documentos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE DOCUMENTOS SET " & _
                        " DEBE = '" & Debe & "'," & _
                        " HABER = '" & Haber & "'" & _
                " WHERE Documento = '" & ID & "'"
        Dim oC As New clsContabilidad
        oC.execute_bd consulta
        Set oC = Nothing
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsContabilidad_Documentos)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As String) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM DOCUMENTOS " & _
                "    WHERE Documento = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsContabilidad_Documentos)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM DOCUMENTOS ORDER BY Documento"
        Set Listado = datos_bd(consulta)
End Function
