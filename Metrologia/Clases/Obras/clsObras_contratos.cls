VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObras_contratos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSOBRAS_CONTRATOS
'***************************************************************
Private ID_CONTRATO As Long
Private OBRA_ID As Long
Private fecha As String
Private DESCRIPCION As String
Private FINALIZADO As Integer
Private FS As String
Private USUARIO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSOBRAS_CONTRATOS
'***************************************************************
Public Property Let setID_CONTRATO(ByVal dato As Long)
        ID_CONTRATO = dato
End Property
Public Property Let setOBRA_ID(ByVal dato As Long)
        OBRA_ID = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setFINALIZADO(ByVal dato As Integer)
        FINALIZADO = dato
End Property
Public Property Let setFS(ByVal dato As String)
        FS = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSOBRAS_CONTRATOS
'***************************************************************
Public Property Get getID_CONTRATO() As Long
        getID_CONTRATO = ID_CONTRATO
End Property
Public Property Get getOBRA_ID() As Long
        getOBRA_ID = OBRA_ID
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFINALIZADO() As Integer
        getFINALIZADO = FINALIZADO
End Property
Public Property Get getFS() As String
        getFS = FS
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSOBRAS_CONTRATOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM OBRAS_CONTRATOS WHERE ID_CONTRATO = " & ID
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
                Carga = False
        Else
                ID_CONTRATO = RS("ID_CONTRATO")
                OBRA_ID = RS("OBRA_ID")
                fecha = RS("FECHA")
                DESCRIPCION = RS("DESCRIPCION")
                FINALIZADO = RS("FINALIZADO")
                FS = RS("FS")
                USUARIO_ID = RS("USUARIO_ID")
                Carga = True
        End If
        Set RS = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsObras_contratos)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim RS As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_CONTRATO) FROM OBRAS_CONTRATOS"
        Set RS = datos_bd(consulta)
        If IsNull(RS.Fields(0)) Or (RS.EOF And RS.BOF) Then
                ID_CONTRATO = 1
        Else
                ID_CONTRATO = RS.Fields(0) + 1
        End If
        Set RS = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO OBRAS_CONTRATOS " & _
                   "  VALUES (" & _
                        ID_CONTRATO & "," & OBRA_ID & "," & "'" & fecha & "'" & "," & _
                        "'" & DESCRIPCION & "'," & FINALIZADO & ",CURRENT_TIMESTAMP," & USUARIO_ID & _
                ")"
        execute_bd consulta
        Insertar = ID_CONTRATO
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsObras_contratos)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OBRAS_CONTRATOS SET " & _
                        " OBRA_ID = " & OBRA_ID & "," & _
                        " FECHA = '" & fecha & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " FS = '" & FS & "'," & _
                        " USUARIO_ID = " & USUARIO_ID & "" & _
                " WHERE ID_CONTRATO = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsObras_contratos)", vbCritical, Err.Description
End Function
Public Function Modificar_Finalizado(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE OBRAS_CONTRATOS SET " & _
                        " FINALIZADO = 1 " & _
                   " WHERE ID_CONTRATO = " & ID
        execute_bd consulta
        Modificar_Finalizado = True
        Exit Function
fallo:
        Modificar_Finalizado = False
        MsgBox "Error al Modificar_Finalizado (clsObras_contratos)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        execute_bd "DELETE FROM OBRAS_CONTRATOS_DETALLE WHERE CONTRATO_ID = " & ID
        execute_bd "DELETE FROM OBRAS_CONTRATOS_CANTIDADES WHERE CONTRATO_ID = " & ID
        execute_bd "DELETE FROM OBRAS_CONTRATOS_ADJUNTOS WHERE CONTRATO_ID = " & ID
        execute_bd "DELETE FROM OBRAS_CONTRATOS WHERE ID_CONTRATO = " & ID
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsObras_contratos)", vbCritical, Err.Description
End Function
Public Function Listado(OBRA As Long) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT A.ID_CONTRATO,A.FECHA,A.DESCRIPCION,A.FINALIZADO,B.USUARIO " & _
                   "  FROM OBRAS_CONTRATOS A, USUARIOS B " & _
                   " WHERE A.USUARIO_ID = B.ID_EMPLEADO " & _
                   "   AND A.OBRA_ID = " & OBRA & _
                   " ORDER BY A.FECHA DESC"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_CONTRATO,OBRA_ID FROM OBRAS_CONTRATOS ORDER BY OBRA_ID"
        Set Listado_Combo = datos_bd(consulta)
End Function
Public Function ContratosEnVigor(OBRA As Long) As Boolean
        Dim consulta As String
        Dim RS As ADODB.Recordset
        consulta = "SELECT * " & _
                   "  FROM OBRAS_CONTRATOS " & _
                   " WHERE OBRA_ID = " & OBRA & _
                   "   AND FINALIZADO = 0"
        Set RS = datos_bd(consulta)
        If RS.RecordCount > 0 Then
            ContratosEnVigor = True
        Else
            ContratosEnVigor = False
        End If
End Function
