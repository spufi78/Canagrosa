VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMetrohm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSMETROHM
'***************************************************************
Private ID As Long
Private REGISTRO As String
Private fichero As String
Private TS As String
Private PROCESADO As Long

'***************************************************************
'*   VARIABLES POR ENUMERACION DE LA CLASE CLSMETROHM
'***************************************************************
Public Enum LISTA_CLSMETROHM_COLS
        ENUM_ID = 1
        ENUM_REGISTRO = 2
        ENUM_FICHERO = 3
        ENUM_TS = 4
        ENUM_PROCESADO = 5
End Enum

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSMETROHM
'***************************************************************
Public Property Let setID(ByVal dato As Long)
        ID = dato
End Property
Public Property Let setREGISTRO(ByVal dato As String)
        REGISTRO = dato
End Property
Public Property Let setFICHERO(ByVal dato As String)
        fichero = dato
End Property
Public Property Let setTS(ByVal dato As String)
        TS = dato
End Property
Public Property Let setPROCESADO(ByVal dato As Long)
        PROCESADO = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSMETROHM
'***************************************************************
Public Property Get getID() As Long
        getID = ID
End Property
Public Property Get getREGISTRO() As String
        getREGISTRO = REGISTRO
End Property
Public Property Get getFICHERO() As String
        getFICHERO = fichero
End Property
Public Property Get getTS() As String
        getTS = TS
End Property
Public Property Get getPROCESADO() As Long
        getPROCESADO = PROCESADO
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSMETROHM
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM METROHM WHERE ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID = rs("ID")
                REGISTRO = rs("REGISTRO")
                fichero = rs("FICHERO")
                TS = rs("TS")
                PROCESADO = rs("PROCESADO")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsMetrohm)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO METROHM " & _
                   "  VALUES (" & _
                        ID & "," & "'" & REGISTRO & "'" & "," & "'" & fichero & "'" & "," & _
                        "'" & TS & "'" & "," & PROCESADO & _
                ")"
        execute_bd consulta
        Insertar = ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsMetrohm)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE METROHM SET " & _
                        " REGISTRO = '" & REGISTRO & "'," & _
                        " FICHERO = '" & fichero & "'," & _
                        " TS = '" & TS & "'," & _
                        " PROCESADO = " & PROCESADO & "" & _
                " WHERE ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsMetrohm)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM METROHM " & _
                "    WHERE ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsMetrohm)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM METROHM ORDER BY ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function Pendientes() As Integer
        Dim consulta As String
        Dim rs As ADODB.Recordset
        consulta = "SELECT count(*) FROM METROHM WHERE PROCESADO = 0"
        Set rs = datos_bd(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                Pendientes = 0
        Else
                Pendientes = rs.Fields(0)
        End If
        Set rs = Nothing
End Function
Private Function marcarError(ID As Long, error As String) As Boolean
        execute_bd "update metrohm set error = '" & error & "' where id = " & ID
        marcarError = True
        Exit Function
fallo:
        marcarError = False
End Function
Private Function marcarProcesado(ID As Long)
        execute_bd "update metrohm set procesado = 1 where id = " & ID
        marcarProcesado = True
        Exit Function
fallo:
        marcarProcesado = False
End Function

Public Function procesarListadoPendientes() As Boolean
        Dim consulta As String
        Dim rs As ADODB.Recordset
   On Error GoTo procesarPendientes_Error

        consulta = "SELECT ID FROM METROHM WHERE PROCESADO = 0 ORDER BY ID"
        Set rs = datos_bd(consulta)
        Dim salida As String
        If rs.RecordCount > 0 Then
            Do
                salida = analisis(rs("ID"))
                If salida <> "" Then
                    marcarError rs("ID"), salida
                End If
                marcarProcesado rs("ID")
                DoEvents
                rs.MoveNext
            Loop Until rs.EOF
        End If
        procesarListadoPendientes = True
   On Error GoTo 0
   Exit Function

procesarPendientes_Error:
    procesarListadoPendientes = False
End Function
Public Function analisis(ID As Long) As String
    Dim oM As New clsMetrohm
    Dim error As String
   On Error GoTo analisis_Error
    ' Eliminamos los analisis previos
    Dim oMA As New clsMetrohm_analisis
    oMA.Eliminar ID
    error = ""
    If oM.Carga(ID) Then
        ' Analizo el nombre del fichero
        Dim fichero() As String
        fichero = Split(oM.getFICHERO(), "-")
        ' Buscar la configuración para ese fichero
        Dim oMF As New clsMetrohm_formato
        If oMF.Carga(fichero(0)) Then
            ' Con la configuración de la posición de las columnas, analizamos e insertamos en metrohm_analisis
            Dim parseo() As String
            parseo = Split(oM.getREGISTRO, ";")
            ' Recupero campos del parseo
            Dim muestra As String
            Dim ensayo As String
            Dim fecha As String
            Dim duplicado As Integer
            muestra = Trim(parseo(oMF.getCOL_MUESTRA))
            ' Buscamos el ID de la muestra
            ensayo = parseo(oMF.getCOL_ENSAYO)
            fecha = Replace(parseo(oMF.getCOL_FECHA), "Z", "")
            If UCase(Right(muestra, 1)) = "R" Then
                duplicado = 1
                muestra = Left(muestra, Len(muestra) - 1)
            Else
                duplicado = 0
            End If
            Dim muestraParseo() As String
            muestraParseo = Split(muestra, "-")
            If UBound(muestraParseo) <> 1 Then
                error = "El formato del código de muestra no es correcto -> " & muestra
            Else
                ' Eliminamos si es duplicado
                Dim rs As New ADODB.Recordset
                Dim idMuestra As Long
                Dim c As String
                c = "select id_muestra from muestras a " & _
                    " inner join tipos_muestra b on a.TIPO_MUESTRA_ID = b.ID_TIPO_MUESTRA " & _
                    " Where a.ANNO = Year(CURRENT_date()) and a.ANULADA = 0 " & _
                    "   and b.CODIGO = '" & muestraParseo(0) & "' " & _
                    "   and a.ID_PARTICULAR = " & muestraParseo(1) & _
                    " limit 1"
                Set rs = datos_bd(c)
                If rs.RecordCount = 0 Then
                    error = "No encuentro la muestra -> " & muestra
                Else
                    ' Evaluación de resultados
                    Dim determinaciones() As String
                    determinaciones = Split(oMF.getCOL_DETERMINACIONES, ";")
                    Dim i As Integer
                    For i = LBound(determinaciones) To UBound(determinaciones)
                        Dim oposTipo() As String
                        oposTipo = Split(determinaciones(i), ":")
                        ' Buscar tipo_determinación en la muestra
                        Dim rs2 As ADODB.Recordset
                        c = "select b.TIPO_DETERMINACION_ID from tipos_determinacion a " & _
                            " inner join determinaciones b on a.ID_TIPO_DETERMINACION = b.TIPO_DETERMINACION_ID " & _
                            " inner join tipos_determinacion c on c.ID_TIPO_DETERMINACION = " & oposTipo(1) & _
                            " Where b.MUESTRA_ID = " & rs(0) & " and a.NOMBRE = c.NOMBRE"
                        Set rs2 = datos_bd(c)
                        If rs2.RecordCount = 0 Then
                            error = "No existe la determinacion en la muestra -> " & oposTipo(1)
                        Else
                            Dim RESULTADO As String
                            RESULTADO = parseo(oposTipo(0))
                            With oMA
                                .setID = ID
                                .setMUESTRA_ID = rs(0)
                                .setTIPO_DETERMINACION_ID = rs2(0)
                                .setRESULTADO = RESULTADO
                                .setES_DUPLICADA = duplicado
                                .setFECHA = fecha
                                .Insertar
                            End With
                        End If
                    Next
                End If
            End If
        Else
            error = "No existe la configuración para el fichero -> " & fichero(0)
        End If
    Else
        error = "Error al cargar el ID -> " & ID
    End If
    analisis = error

   On Error GoTo 0
   Exit Function

analisis_Error:
    analisis = "Error " & Err.Number & " (" & Err.Description & ") in procedure analisis of Módulo de clase clsMetrohm"
End Function
'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSMETROHM
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

