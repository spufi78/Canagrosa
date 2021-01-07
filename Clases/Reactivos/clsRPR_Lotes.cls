VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRPR_Lotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSRPR_BOTES
'***************************************************************
Private ID_LOTE As Long
Private NUMERO_LOTE As String
Private FECHA_CREACION As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSRPR_BOTES
'***************************************************************
Public Property Let setID_LOTE(ByVal dato As Long)
        ID_LOTE = dato
End Property
Public Property Let setNUMERO_LOTE(ByVal dato As String)
        NUMERO_LOTE = dato
End Property
Public Property Let setFECHA_CREACION(ByVal dato As String)
        FECHA_CREACION = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSRPR_BOTES
'***************************************************************
Public Property Get getID_LOTE() As Long
        getID_LOTE = ID_LOTE
End Property
Public Property Get getNUMERO_LOTE() As String
        getNUMERO_LOTE = NUMERO_LOTE
End Property
Public Property Get getFECHA_CREACION() As String
        getFECHA_CREACION = FECHA_CREACION
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSRPR_BOTES
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM RPR_BOTES WHERE ID_BOTE_PR = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_LOTE = rs("ID_LOTE")
                NUMERO_LOTE = rs("NUMERO_LOTE")
                FECHA_CREACION = rs("FECHA_CREACION")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsRpr_Lotes)", vbCritical, Err.Description
End Function
Public Function CrearID()
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT MAX(ID_LOTE) FROM RPR_LOTES"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_LOTE = 1
        Else
                ID_LOTE = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function

Public Function CrearNumeroLote() As String
    Dim oParametro As New clsParametros
    
    If oParametro.Carga(parametros.NUMERO_LOTE, "") Then
        If Right(oParametro.getVALOR, 2) = Right(Year(Date), 2) Then ' Si es el mismo año
            oParametro.setVALOR = (Format(CLng(Left(oParametro.getVALOR, 4)) + 1, "0000")) & "/" & Right(oParametro.getVALOR, 2)
        Else ' si cambia de año se reinicia la cuenta
            oParametro.setVALOR = Format(CLng(1), "0000") & "/" & Right(Year(Date), 2)
        End If
    Else
        oParametro.setVALOR = Format(CLng(1), "0000") & "/" & Right(Year(Date), 2)
    End If
    oParametro.actualizar_valor parametros.NUMERO_LOTE, ""
    NUMERO_LOTE = oParametro.getVALOR
    
End Function

Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        
        Me.CrearID
        consulta = "INSERT INTO RPR_LOTES " & _
                   "  VALUES (" & _
                        ID_LOTE & ",'" & NUMERO_LOTE & "','" & FECHA_CREACION & "' " & _
                        "," & USUARIO.getID_EMPLEADO & ",current_timestamp)"
        execute_bd consulta
        Insertar = ID_LOTE
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsRpr_Lotes)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        consulta = "UPDATE RPR_LOTES " & _
                   "   SET NUMERO_LOTE = '" & NUMERO_LOTE & "'" & _
                   "     , FECHA_CREACION = '" & FECHA_CREACION & "'" & _
                   " WHERE ID_LOTE = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsRpr_Lotes)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        
        ' Se borran todos los botes del lote
        consulta = "DELETE FROM RPR_BOTES " & _
                "    WHERE LOTE_ID = " & ID
        execute_bd consulta
        
        ' Se borra el lote
        consulta = "DELETE FROM RPR_LOTES " & _
                "    WHERE ID_LOTE = " & ID
        execute_bd consulta
        
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsRpr_Lotes)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT * FROM RPR_LOTES ORDER BY NUMERO_LOTE"
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT ID_LOTE, NUMERO_LOTE " & _
                   "  FROM RPR_LOTES " & _
                   " ORDER BY NUMERO_LOTE"
        Set Listado_Combo = datos_bd(consulta)
End Function

Public Function Listado_botes_por_lote(ID As Long) As ADODB.RecordSet
        Dim consulta As String
        
        consulta = "SELECT B.ID_BOTE_PR, B.NUMERO, " & _
                   "       T.CODIGO, T.NOMBRE, " & _
                   "       B.FECHA_FABRICACION, B.FECHA_CADUCIDAD, " & _
                   "       B.VOLUMEN, U.USUARIO " & _
                   "  FROM RPR_BOTES B, RPR_TIPOS T, USUARIOS U, RPR_LOTES L " & _
                   " WHERE B.TIPO_REACTIVO_PR_ID = T.ID_TIPO_REACTIVO_PR " & _
                   "   AND B.EMPLEADO_ID = U.ID_EMPLEADO " & _
                   "   AND B.LOTE_ID = L.ID_LOTE " & _
                   "   AND L.ID_LOTE = " & ID & _
                   " ORDER BY B.NUMERO DESC"
        
        Set Listado_botes_por_lote = datos_bd(consulta)
End Function
