VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCe_resultadosX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSCE_RESULTADOS
'***************************************************************
Private RECEPCION_ID As Long
Private MUESTRA_ID As Long
Private ORDEN As Long
Private Resultado As String
Private fecha As String
Private CONFORME As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSCE_RESULTADOS
'***************************************************************
Public Property Let setRECEPCION_ID(ByVal dato As Long)
        RECEPCION_ID = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setORDEN(ByVal dato As Long)
        ORDEN = dato
End Property
Public Property Let setRESULTADO(ByVal dato As String)
        Resultado = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setCONFORME(ByVal dato As Long)
        CONFORME = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSCE_RESULTADOS
'***************************************************************
Public Property Get getRECEPCION_ID() As Long
        getRECEPCION_ID = RECEPCION_ID
End Property
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getORDEN() As Long
        getORDEN = ORDEN
End Property
Public Property Get getRESULTADO() As String
        getRESULTADO = Resultado
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getCONFORME() As Long
        getCONFORME = CONFORME
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSCE_RESULTADOS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_RESULTADOSX  WHERE RECEPCION_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                RECEPCION_ID = rs("RECEPCION_ID")
                MUESTRA_ID = rs("MUESTRA_ID")
                ORDEN = rs("ORDEN")
                Resultado = rs("RESULTADO")
                fecha = rs("FECHA")
                CONFORME = rs("CONFORME")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsCe_resultados)", vbCritical, Err.Description
End Function
Public Function Insertar(RECEPCION As Long, MUESTRA As Long, ORDEN As Long, Resultado As String, CONFORME As Integer) As Long
        On Error GoTo fallo
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        fecha = Format(Date, "YYYY-MM-DD")
        consulta = "SELECT * FROM CE_RESULTADOSX  " & _
                   " WHERE RECEPCION_ID = " & RECEPCION & _
                   "   AND MUESTRA_ID = " & MUESTRA & _
                   "   AND ORDEN = " & ORDEN
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
            consulta = "INSERT INTO CE_RESULTADOSX  " & _
                       "  VALUES (" & _
                            RECEPCION & "," & MUESTRA & "," & ORDEN & "," & _
                            "'" & Resultado & "','" & fecha & "'," & CONFORME & _
                    ")"
        Else
            consulta = "UPDATE CE_RESULTADOSX  " & _
                       "   SET RESULTADO='" & Resultado & "'," & _
                       "       FECHA = '" & fecha & "'," & _
                       "       CONFORME = " & CONFORME & _
                       " WHERE RECEPCION_ID = " & RECEPCION & _
                       "   AND MUESTRA_ID = " & MUESTRA & _
                       "   AND ORDEN = " & ORDEN
        End If
        conn.Execute consulta
        Insertar = RECEPCION
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsCe_resultados)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM CE_RESULTADOSX  " & _
                "    WHERE RECEPCION_ID = " & ID
        conn.Execute consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsCe_resultados)", vbCritical, Err.Description
End Function
Public Function Listado(RECEPCION As Long, MUESTRA As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_RESULTADOSX  " & _
                   " WHERE RECEPCION_ID = " & RECEPCION & _
                   "   AND MUESTRA_ID = " & MUESTRA
        Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Conversion(RECEPCION As Long, MUESTRA As Long) As ADODB.RecordSet
        Dim consulta As String
        consulta = "SELECT * FROM CE_RESULTADOSX  " & _
                   " WHERE RECEPCION_ID = " & RECEPCION & _
                   "   AND MUESTRA_ID = " & MUESTRA & _
                   " ORDER BY ORDEN"
        Set Listado_Conversion = datos_bd(consulta)
End Function
Public Function Cerrar(MUESTRA As Long) As Integer
        Dim consulta As String
        Dim rs As ADODB.RecordSet
        Dim rs2 As ADODB.RecordSet
        consulta = "SELECT * FROM CE_RESULTADOSX  " & _
                   " WHERE MUESTRA_ID = " & MUESTRA
        Set rs = datos_bd(consulta)
        Dim resultados As Integer
        Dim oce_recepcion As New clsCe_recepcionX
        Set rs2 = oce_recepcion.Listado_por_muestra(MUESTRA)
        If rs2.RecordCount > 0 Then
            If IsNumeric(rs2(4)) Then
                resultados = resultados + CInt(rs2(4))
                If rs.RecordCount = resultados Then
                    Cerrar = 1
                Else
                    Cerrar = 0
                End If
            Else
                Cerrar = 0
            End If
        Else
            Cerrar = 0
        End If
        ' Si la muestra es un ECS, verificar que tiene informadas las máquinas
        Dim omuestra As New clsMuestra
        If Cerrar = 1 Then
            If omuestra.CargaMuestra(MUESTRA) Then
                If omuestra.getTIPO_MUESTRA_ID = TIPOS_MUESTRAS.ecs Then
                    If oce_recepcion.Carga_Primera(MUESTRA) Then
                        If oce_recepcion.getMAQUINA = "" Then
                            MsgBox "Para los ECS, debe indicar las máquinas para poder cerrar las muestras.", vbExclamation, App.Title
                            Cerrar = 0
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
End Function
