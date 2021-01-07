VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsComercial_Comision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCOMERCIALES_COMISIONES
'***************************************************************
Private COMERCIAL_ID As Long
Private DESCUENTO_ID As Long
Private Comision As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCOMERCIALES_COMISIONES
'***************************************************************
Public Property Let setCOMERCIAL_ID(ByVal dato As Long)
        COMERCIAL_ID = dato
End Property
Public Property Let setDESCUENTO_ID(ByVal dato As Long)
        DESCUENTO_ID = dato
End Property
Public Property Let setCOMISION(ByVal dato As String)
        Comision = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCOMERCIALES_COMISIONES
'***************************************************************
Public Property Get getCOMERCIAL_ID() As Long
        getCOMERCIAL_ID = COMERCIAL_ID
End Property
Public Property Get getDESCUENTO_ID() As Long
        getDESCUENTO_ID = DESCUENTO_ID
End Property
Public Property Get getCOMISION() As String
        getCOMISION = Comision
End Property

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO COMERCIALES_COMISIONES " & _
                   "  VALUES (" & _
                        COMERCIAL_ID & "," & DESCUENTO_ID & ",'" & Comision & _
                "')"
        execute_bd consulta
        Insertar = COMERCIAL_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCOMERCIALES_COMISIONES)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM COMERCIALES_COMISIONES " & _
                "    WHERE COMERCIAL_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCOMERCIALES_COMISIONES)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * " & _
                   " FROM COMERCIALES_COMISIONES " & _
                   " WHERE COMERCIAL_ID = " & ID
        Set Listado = datos_bd(consulta)
End Function
Public Function Recuperar_Comision(COMERCIAL As Long, DESCUENTO As Long) As Integer
        Dim consulta As String
        Dim rs As New ADODB.Recordset
   On Error GoTo Recuperar_Comision_Error

        consulta = "SELECT COMISION " & _
                   " FROM COMERCIALES_COMISIONES " & _
                   " WHERE COMERCIAL_ID = " & COMERCIAL & _
                   "   AND DESCUENTO_ID = " & DESCUENTO
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            Recuperar_Comision = 0
        Else
            Recuperar_Comision = rs(0)
        End If

   On Error GoTo 0
   Exit Function

Recuperar_Comision_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure Recuperar_Comision of Módulo de clase clsComercial_Comision"
End Function
