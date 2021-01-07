VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSC_Paquetes_Detalle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSSC_PAQUETES_DETALLE
'***************************************************************
Private PAQUETE_ID As Long
Private DETERMINACION_ID As Long
'M0957-I
Private MUESTRA_ID As Long
Private TIPO_ENSAYO_ID As Long
'M0957-F
Private VALOR_REFERENCIA As String
Private NORMATIVA_APLICABLE As String
Private PRECIO As String
'M1274-I
Private EDICION As Long
'M1274-F
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSSC_PAQUETES_DETALLE
'***************************************************************
Public Property Let setPAQUETE_ID(ByVal dato As Long)
        PAQUETE_ID = dato
End Property
Public Property Let setDETERMINACION_ID(ByVal dato As Long)
        DETERMINACION_ID = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setTIPO_ENSAYO_ID(ByVal dato As Long)
        TIPO_ENSAYO_ID = dato
End Property
Public Property Let setVALOR_REFERENCIA(ByVal dato As String)
        VALOR_REFERENCIA = dato
End Property
Public Property Let setNORMATIVA_APLICABLE(ByVal dato As String)
        NORMATIVA_APLICABLE = dato
End Property
Public Property Let setPRECIO(ByVal dato As String)
        PRECIO = dato
End Property
'M1274-I
Public Property Let setEDICION(ByVal dato As Long)
        EDICION = dato
End Property
'M1274-F
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSSC_PAQUETES_DETALLE
'***************************************************************
Public Property Get getPAQUETE_ID() As Long
        getPAQUETE_ID = PAQUETE_ID
End Property
Public Property Get getDETERMINACION_ID() As Long
        getDETERMINACION_ID = DETERMINACION_ID
End Property
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getTIPO_ENSAYO_ID() As Long
        getTIPO_ENSAYO_ID = TIPO_ENSAYO_ID
End Property
Public Property Get getVALOR_REFERENCIA() As String
        getVALOR_REFERENCIA = VALOR_REFERENCIA
End Property
Public Property Get getNORMATIVA_APLICABLE() As String
        getNORMATIVA_APLICABLE = NORMATIVA_APLICABLE
End Property
Public Property Get getPRECIO() As String
        getPRECIO = PRECIO
End Property
'M1274-I
Public Property Get getEDICION() As Long
        getEDICION = EDICION
End Property
'M1274-F

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSSC_PAQUETES_DETALLE
'***************************************************************
Public Function Carga(ID_P As Long, EDICION As Long, ID_D As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT * " & _
                   "  FROM SC_PAQUETES_DETALLE " & _
                   " WHERE PAQUETE_ID = " & ID_P & _
                   "   AND DETERMINACION_ID = " & ID_D & " AND MUESTRA_ID = 0 AND TIPO_ENSAYO_ID = 0"
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
            Carga = False
        Else
            PAQUETE_ID = RS("PAQUETE_ID")
            DETERMINACION_ID = RS("DETERMINACION_ID")
            VALOR_REFERENCIA = RS("VALOR_REFERENCIA")
            NORMATIVA_APLICABLE = RS("NORMATIVA_APLICABLE")
            PRECIO = RS("PRECIO")
            'M1274-I
            EDICION = RS("EDICION")
            'M1274-F
            Carga = True
        End If
        Set RS = Nothing

        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSC_Paquetes_Detalle)", vbCritical, Err.Description
End Function
'M0957-I
Public Function CargaCE(ID_P As Long, EDICION As Long, ID_M As Long, ID_E As Long) As Boolean
        On Error GoTo fallo
        Dim RS As ADODB.Recordset
        Dim consulta As String
        
        consulta = "SELECT * " & _
                   "  FROM SC_PAQUETES_DETALLE " & _
                   " WHERE PAQUETE_ID = " & ID_P & _
                   "   AND EDICION = " & EDICION & _
                   "   AND DETERMINACION_ID = 0 AND MUESTRA_ID = " & ID_M & " AND TIPO_ENSAYO_ID = " & ID_E
        Set RS = datos_bd(consulta)
        If RS.RecordCount = 0 Then
            CargaCE = False
        Else
            PAQUETE_ID = RS("PAQUETE_ID")
            DETERMINACION_ID = RS("DETERMINACION_ID")
            MUESTRA_ID = RS("MUESTRA_ID")
            TIPO_ENSAYO_ID = RS("TIPO_ENSAYO_ID")
            VALOR_REFERENCIA = RS("VALOR_REFERENCIA")
            NORMATIVA_APLICABLE = RS("NORMATIVA_APLICABLE")
            PRECIO = RS("PRECIO")
            'M1274-I
            EDICION = RS("EDICION")
            'M1274-F
            CargaCE = True
        End If
        Set RS = Nothing

        Exit Function
fallo:
        CargaCE = False
        MsgBox "Error al cargar los datos(clsSC_Paquetes_Detalle)", vbCritical, Err.Description
End Function
'M0957-F
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String

        If Not duplicado Then
            consulta = "INSERT INTO SC_PAQUETES_DETALLE " & _
                       "VALUES (" & _
                                PAQUETE_ID & ", " & DETERMINACION_ID & ", " & MUESTRA_ID & ", " & TIPO_ENSAYO_ID & ", " & _
                                "'" & VALOR_REFERENCIA & "'" & ", " & "'" & NORMATIVA_APLICABLE & "'" & ", " & PRECIO & ", " & _
                                "'" & USUARIO.getUSUARIO & "'" & ", CURRENT_TIMESTAMP," & EDICION & _
                       ")"
            execute_bd consulta
            Insertar = PAQUETE_ID
        End If

        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSC_Paquetes_Detalle)", vbCritical, Err.Description
End Function

Public Function Modificar() As Long
        On Error GoTo fallo
        Dim consulta As String
         
        consulta = "UPDATE  SC_PAQUETES_DETALLE " & _
                       "SET PRECIO = '" & PRECIO & "'" & _
                       " ,CUSERID = '" & USUARIO.getUSUARIO & "'" & _
                       " ,FTIMESTP = CURRENT_TIMESTAMP" & _
                       " WHERE PAQUETE_ID = " & PAQUETE_ID & " AND EDICION = " & EDICION & " AND DETERMINACION_ID = " & DETERMINACION_ID & _
                       " AND MUESTRA_ID = " & MUESTRA_ID & " AND TIPO_ENSAYO_ID = " & TIPO_ENSAYO_ID
        execute_bd consulta
        Modificar = PAQUETE_ID

        Exit Function
fallo:
        Modificar = 0
        MsgBox "Error al Modificar (clsSC_Paquetes_Detalle)", vbCritical, Err.Description
End Function

' se elimina todo el detalle del paquete
Public Function Eliminar(ID_P As Long, EDICION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String

        consulta = "DELETE FROM SC_PAQUETES_DETALLE " & _
                   " WHERE PAQUETE_ID = " & ID_P & " AND EDICION = " & EDICION
        execute_bd consulta
        Eliminar = True

        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSC_Paquetes_Detalle)", vbCritical, Err.Description
End Function
Public Function Eliminar_Determinacion(PAQUETE As Long, EDICION As Long, DETERMINACION As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String

        consulta = "DELETE FROM SC_PAQUETES_DETALLE " & _
                   " WHERE PAQUETE_ID = " & PAQUETE & _
                   " AND EDICION = " & EDICION & _
                   "   AND DETERMINACION_ID = " & DETERMINACION
        execute_bd consulta
        consulta = "UPDATE DETERMINACIONES SET ENVIADO_PAQUETE = 0 " & _
                   " WHERE ID_DETERMINACION = " & DETERMINACION
        execute_bd consulta
        Eliminar_Determinacion = True

        Exit Function
fallo:
        Eliminar_Determinacion = False
        MsgBox "Error al Eliminar_Determinacion (clsSC_Paquetes_Detalle)", vbCritical, Err.Description
End Function

Public Function duplicado() As Boolean
    Dim RS As New ADODB.Recordset
    Dim consulta As String
    
    ' JGM-I No se comprueba para el duplicado el MUESTRA_ID
    
    consulta = "SELECT COUNT(*) AS TOTAL " & _
               "  FROM SC_PAQUETES_DETALLE " & _
               " WHERE PAQUETE_ID = " & PAQUETE_ID & _
               "   AND DETERMINACION_ID = " & DETERMINACION_ID & _
               "   AND MUESTRA_ID = " & MUESTRA_ID & _
               "   AND EDICION = " & EDICION
    Set RS = datos_bd(consulta)
    If RS.Fields("TOTAL") = 0 Then
        duplicado = False
    Else
        MsgBox "Esta determinación ya está en el paquete.", vbExclamation, App.Title
        duplicado = True
    End If
    Set RS = Nothing
End Function

'M0957-I
Public Function Eliminar_Ensayo(PAQUETE As Long, EDICION As Long, ID_MUESTRA As Long, ID_ENSAYO As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String

        consulta = "DELETE FROM SC_PAQUETES_DETALLE " & _
                   " WHERE PAQUETE_ID = " & PAQUETE & " AND EDICION = " & EDICION
        execute_bd consulta
        consulta = "UPDATE CE_RECEPCION SET ENVIADO_PAQUETE = 0 " & _
                   " WHERE MUESTRA_ID = " & ID_MUESTRA & _
                   " AND TIPO_ENSAYO_ID = " & ID_ENSAYO
        execute_bd consulta
        Eliminar_Ensayo = True

        Exit Function
fallo:
        Eliminar_Ensayo = False
        MsgBox "Error al Eliminar_Ensayo (clsSC_Paquetes_Detalle)", vbCritical, Err.Description
End Function
'M0957-F
