VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlanMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPlanMantenimiento
'***************************************************************
Private ID_PLAN_MTO As Long
Private nombre As String
Private DESCRIPCION As String
Private FRECUENCIA_ID As Long
Private FRECUENCIA As String
Private PLANIF_MESES As String
Private PLANIF_DIAS_SEMANA As String
Private PLANIF_DIAS_MES As String
Private ORDEN As Integer
Private CUSERID As String
Private FTIMESTP As String
Private ID_TIPO_PLAN As Long
Private CALENDARIO As String
Private PARAMETROS_FRECUENCIA As String
Private Excepciones As String
Private EVITAR_EXCEPCIONES As Integer
Private arrExcepciones() As Integer
Private indiceExcepciones As Integer
Private mvarobjAcciones As New clsGenericCollection
Private PROTOCOLO_ID As Long
Private PROTOCOLO As String

Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "eq_planes_mantenimiento"
        .setDESCRIPCION = "Planes de Mantenimiento"
        .setPK = "ID_PLAN_MTO"
        .setCAMPO = " concat(Nombre,': ',Descripcion)"
        .setMUESTRA_DETALLE = True
        .setFILTRO = filtro
        Set .FORMULARIO = FORMULARIO
    End With
End Function
Private Sub Class_Initialize()
    mvarobjAcciones.KeyName = "setID_ACCION"
End Sub

Private Sub Class_Terminate()

    Set mvarobjAcciones = Nothing

End Sub


Public Property Get ACCIONES() As clsGenericCollection

    Set ACCIONES = mvarobjAcciones

End Property

Public Property Set ACCIONES(objAcciones As clsGenericCollection)

    Set mvarobjAcciones = objAcciones

End Property

Public Function DevolverCalendarioCodificado(ByRef prmCalendario() As Integer) As String
On Error GoTo DevolverCalendarioCodificado_Error
Dim cont As Integer, strres As String

    strres = ""
    For cont = 1 To 366
        strres = strres & CStr(prmCalendario(cont))
    Next cont
    
    DevolverCalendarioCodificado = strres

On Error GoTo 0
Exit Function
DevolverCalendarioCodificado_Error:
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure DevolverCalendarioCodificado of Módulo de clase clsPlanMantenimiento"
End Function
Public Function DevolverMesesCodificado(ByRef prmMeses() As String) As String

Dim cont As Integer, strres As String

    strres = ""
    For cont = 1 To 12
        strres = strres & CStr(prmMeses(cont))
    Next cont
    
    DevolverMesesCodificado = strres

End Function

Public Function DevolverDiasMesesCodificado(ByRef prmDiasMeses() As String) As String
Dim cont As Integer, strres As String
    strres = ""
    For cont = 1 To 12
        strres = strres & CStr(prmDiasMeses(cont))
    Next cont
    DevolverDiasMesesCodificado = strres
End Function

Public Function DevolverSemanasCodificado(ByRef prmSemanas() As String) As String
Dim cont As Integer, strres As String

    
    strres = ""
    
    For cont = 1 To 7
        strres = strres & CStr(prmSemanas(cont))
    Next cont
    
    DevolverSemanasCodificado = strres
    
End Function

Public Function DevolverExcepcionesCodificado(ByRef prmExcepciones() As Integer, ByVal TotalExcepciones As Integer) As String
Dim cont As Integer, strres As String
Dim arr() As String

    DevolverExcepcionesCodificado = ""
    If TotalExcepciones = 0 Then Exit Function

    
    For cont = 0 To TotalExcepciones - 1
        ReDim Preserve arr(cont)
        arr(cont) = CStr(prmExcepciones(cont))
    Next cont
    
    DevolverExcepcionesCodificado = join(arr, ";")
    
End Function

Public Function generarFechasPlanMto(ByVal prmAno As Integer, Optional ByRef idPlanMto As Long = -1) As clsGenericCollection
   Dim objCol As clsGenericCollection
   Set generarFechasPlanMto = New clsGenericCollection

   If idPlanMto > 0 Then
       Call Carga(idPlanMto)
   End If
        
    Select Case ID_TIPO_PLAN
        Case 0 ' Sin periodicidad
            Exit Function
        Case 1 ' Mensual
            If PLANIF_MESES = String(12, "0") Or PLANIF_DIAS_MES = String(31, "0") Then
               MsgBox "Para el plan Mensual debe señalar al menos 1 Mes y 1 Día", vbInformation, "Generar Plan de Mantenimiento de Equipos"
               Exit Function
            Else
               Set objCol = generarFechasPlanMto_Mensual(prmAno)
            End If
        Case 2 ' Semanal
            If PLANIF_DIAS_SEMANA = String(7, "0") Then
               MsgBox "Para el plan Semanal debe señalar al menos 1 Día de la Semana", vbInformation, "Generar Plan de Mantenimiento de Equipos"
               Exit Function
            Else
               Set objCol = generarFechasPlanMto_Semanal(prmAno)
            End If
        Case 3 ' Diario
            Set objCol = generarFechasPlanMto_diario(prmAno)
        Case 4 ' Calendario
            If CALENDARIO = String(366, "0") Then
               MsgBox "Para el plan Mensual debe señalar al menos 1 Día en el Calendario", vbInformation, "Generar Plan de Mantenimiento de Equipos"
               Exit Function
            Else
               Set objCol = generarFechasPlanMto_Calendario(prmAno)
            End If
    End Select


    Set generarFechasPlanMto = ordenarFechasMantenimiento(objCol)

End Function

Private Function generarFechasPlanMto_Mensual(ByVal prmAno As Integer) As clsGenericCollection
    Dim arrMeses() As Integer
    Dim arrDias() As Integer
    Dim objMto As clsEquipoMantenimiento
    Dim iMes As Integer, iDia As Integer
    Dim dtmFecha As Date
    Dim objCol As New clsGenericCollection
    objCol.KeyName = "setID_MANTENIMIENTO"
    
    arrMeses = SplitString2IntArr(PLANIF_MESES)
    arrDias = SplitString2IntArr(PLANIF_DIAS_MES)
    
    For iMes = 1 To 12
        If arrMeses(iMes) = 1 Then
            For iDia = 1 To 31
                If arrDias(iDia) = 1 Then
                    dtmFecha = DateSerial(prmAno, iMes, iDia)
                    If manejarExcepcion(dtmFecha) Then
                        If Not ComprobarFechaExisteEnColMtos(dtmFecha, objCol) Then
                            Set objMto = New clsEquipoMantenimiento
                            objMto.setFECHA_ACTUAL = Format(dtmFecha, "dd/mm/yyyy")
                            objMto.setESTADO = 0
                            Call objCol.Add(objMto)
                        End If
                    End If
                End If
            Next iDia
        End If
    Next iMes

    Set generarFechasPlanMto_Mensual = objCol
    
End Function

Private Function generarFechasPlanMto_Semanal(ByVal prmAno As Integer) As clsGenericCollection
    Dim arrSemanas() As Integer
    Dim objMto As clsEquipoMantenimiento
    Dim iDia As Integer
    Dim dtmFecha As Date
    Dim objCol As New clsGenericCollection
    objCol.KeyName = "setID_MANTENIMIENTO"
    
    arrSemanas = SplitString2IntArr(PLANIF_DIAS_SEMANA)
    
    
    For iDia = 1 To 366
        dtmFecha = CalcularFechaPorCodigo(iDia, prmAno)
            
        If arrSemanas(Weekday(dtmFecha, vbMonday)) = 1 Then
            ' Entra dentro de la fecha.
            ' comprueba se es alguna excepcion
            If manejarExcepcion(dtmFecha) Then
                If Not ComprobarFechaExisteEnColMtos(dtmFecha, objCol) Then
                    Set objMto = New clsEquipoMantenimiento
                    objMto.setFECHA_ACTUAL = Format(dtmFecha, "dd/mm/yyyy")
                    objMto.setESTADO = 0
                    Call objCol.Add(objMto)
                End If
            End If
        End If
    Next iDia
        
    Set generarFechasPlanMto_Semanal = objCol
End Function

Private Function generarFechasPlanMto_diario(ByVal prmAno As Integer) As clsGenericCollection
    Dim objMto As clsEquipoMantenimiento
    Dim iDia As Integer
    Dim dtmFecha As Date
    Dim objCol As New clsGenericCollection
    Set objCol = New clsGenericCollection
    objCol.KeyName = "setID_MANTENIMIENTO"
    
        
    For iDia = 1 To 366
        dtmFecha = CalcularFechaPorCodigo(iDia, prmAno)
        ' comprueba se es alguna excepcion
        If manejarExcepcion(dtmFecha) Then
            If Not ComprobarFechaExisteEnColMtos(dtmFecha, objCol) Then
                Set objMto = New clsEquipoMantenimiento
                objMto.setFECHA_ACTUAL = Format(dtmFecha, "dd/mm/yyyy")
                objMto.setESTADO = 0
                Call objCol.Add(objMto)
            End If
        End If
    Next iDia
        
    Set generarFechasPlanMto_diario = objCol
End Function

Private Function generarFechasPlanMto_Calendario(ByVal prmAno As Integer) As clsGenericCollection
    Dim objMto As clsEquipoMantenimiento
    Dim arrCalendario() As Integer
    Dim iDia As Integer
    Dim dtmFecha As Date
    Dim objCol As New clsGenericCollection
    objCol.KeyName = "setID_MANTENIMIENTO"
    
    arrCalendario = SplitString2IntArr(CALENDARIO)
    
    For iDia = 1 To 366
        If prmAno Mod 4 <> 0 And iDia = 60 Then iDia = 61
        If arrCalendario(iDia) = 1 Then
            dtmFecha = CalcularFechaPorCodigo(iDia, prmAno)
            
            ' comprueba se es alguna excepcion
            If manejarExcepcion(dtmFecha) Then
                If Not ComprobarFechaExisteEnColMtos(dtmFecha, objCol) Then
                    Set objMto = New clsEquipoMantenimiento
                    objMto.setFECHA_ACTUAL = Format(dtmFecha, "dd/mm/yyyy")
                    objMto.setESTADO = 0
                    Call objCol.Add(objMto)
                End If
            End If
        End If
        
    Next iDia
        
    Set generarFechasPlanMto_Calendario = objCol
End Function

Private Function ComprobarFechaExisteEnColMtos(ByVal dtmFecha As Date, ByRef objCol As clsGenericCollection) As Boolean
Dim objX As clsEquipoMantenimiento

    ComprobarFechaExisteEnColMtos = False
    
    If objCol.Count = 0 Then Exit Function

    ComprobarFechaExisteEnColMtos = True
    
    For Each objX In objCol.Iterator
        If objX.getFECHA_ACTUAL = Format(dtmFecha, "dd/mm/yyyy") Then Exit Function
    Next
    
    ComprobarFechaExisteEnColMtos = False
    
End Function

Public Property Get getEVITAR_EXCEPCIONES() As Integer
    getEVITAR_EXCEPCIONES = EVITAR_EXCEPCIONES
End Property
Private Function manejarExcepcion(ByRef dtmFecha As Date) As Boolean
Dim dtmRes As Date, dtmRes2 As Date
Dim cont As Integer
    dtmRes = dtmFecha
    manejarExcepcion = False
    
    ' comprueba si es sábado o domingo
    If Weekday(dtmFecha, vbMonday) = 6 Then
        If EVITAR_EXCEPCIONES = 1 Then Exit Function
        ' Añade 2 Días
        dtmRes = DateAdd("d", 2, dtmRes)
    ElseIf Weekday(dtmFecha, vbMonday) = 7 Then
        If EVITAR_EXCEPCIONES = 1 Then Exit Function
        ' Añade 2 Días
        dtmRes = DateAdd("d", 1, dtmRes)
    End If
    
    ' De momento, si ya es sábado o Domingo, suma los dias
    ' ahora, si despues de establecer a lunes, resulta que es una excepcion, lo contempla
    If indiceExcepciones = 0 Then
        ' No hay excepciones, continua
        dtmFecha = dtmRes
        manejarExcepcion = True
        Exit Function
    End If
    
    For cont = 0 To indiceExcepciones - 1
        If CalcularCodigoFecha_PorFecha(dtmRes) = arrExcepciones(cont) Then
            If EVITAR_EXCEPCIONES Then Exit Function
            dtmRes2 = DateAdd("d", 1, dtmRes)
            ' Vuelve a comprobar que la nueva fecha no es una excepcion
            ' recursivamente, es decir, va a pasar por todos los
            If manejarExcepcion(dtmRes2) Then dtmRes = dtmRes2
        End If
    Next cont
    
    dtmFecha = dtmRes
    manejarExcepcion = True

End Function



Public Function ordenarFechasMantenimiento(ByRef objMant As clsGenericCollection) As clsGenericCollection
Dim objCol As New clsGenericCollection
Dim obj As clsEquipoMantenimiento
Dim arrFechas() As Long, indiceFechas As Long, cont As Long
Dim blnFin As Boolean

    indiceFechas = 0
    

    objCol.KeyName = objMant.KeyName
    
    ' Primero hace una lista de fechas ordenadas
    For Each obj In objMant.Iterator
        ReDim Preserve arrFechas(indiceFechas)
        arrFechas(indiceFechas) = CLng(CDate(obj.getFECHA_ACTUAL))
        indiceFechas = indiceFechas + 1
    Next
    
    arrFechas = OrdernarValoresArrayLong(arrFechas, 0, indiceFechas - 1)
    
    
    For cont = 0 To indiceFechas - 1
        For Each obj In objMant.Iterator
            If arrFechas(cont) = CLng(CDate(obj.getFECHA_ACTUAL)) Then
                Call objCol.Add(obj)
                                    
                Exit For
            End If
        Next
    Next cont
        
    
    Set ordenarFechasMantenimiento = objCol
    
    
End Function

Public Property Let setEVITAR_EXCEPCIONES(ByVal dato As Integer)
    EVITAR_EXCEPCIONES = dato
End Property






'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsPlanMantenimiento
'***************************************************************

Public Sub setArrMeses(ByRef arrMeses() As Integer)
    PLANIF_MESES = JoinIntArr2String(arrMeses, , 1)
End Sub

Public Sub setArrDiasMeses(ByRef arrDiasMeses() As Integer)
    PLANIF_DIAS_MES = JoinIntArr2String(arrDiasMeses, , 1)
End Sub

Public Sub setArrSemanas(ByRef arrSemanas() As Integer)
    PLANIF_DIAS_SEMANA = JoinIntArr2String(arrSemanas, , 1)
End Sub

Public Sub setArrCalendario(ByRef arrCalendario() As Integer)
    CALENDARIO = JoinIntArr2String(arrCalendario, , 1)
End Sub

Public Sub setArrExcepciones(ByRef prmarrExcepciones() As Integer)
    arrExcepciones = prmarrExcepciones
    indiceExcepciones = UBound(arrExcepciones) + 1
    ' Ahora Genera la cadena
    Excepciones = JoinIntArr2String(arrExcepciones, ";")
End Sub
Public Property Let setPROTOCOLO(ByVal dato As String)
    PROTOCOLO = dato
End Property
Public Property Let setPROTOCOLO_ID(ByVal dato As Long)
    PROTOCOLO_ID = dato
End Property
Public Property Let setID_PLAN_MTO(ByVal dato As Long)
    ID_PLAN_MTO = dato
End Property
Public Property Let setNOMBRE(ByVal dato As String)
    nombre = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
    DESCRIPCION = dato
End Property
Public Property Let setFRECUENCIA_ID(ByVal dato As Long)
    FRECUENCIA_ID = dato
End Property
Public Property Let setFRECUENCIA(ByVal dato As String)
    FRECUENCIA = dato
End Property
Public Property Let setPLANIF_MESES(ByVal dato As String)
    PLANIF_MESES = dato
End Property
Public Property Let setPLANIF_DIAS_SEMANA(ByVal dato As String)
    PLANIF_DIAS_SEMANA = dato
End Property
Public Property Let setPLANIF_DIAS_MES(ByVal dato As String)
    PLANIF_DIAS_MES = dato
End Property
Public Property Let setORDEN(ByVal dato As Integer)
    ORDEN = dato
End Property
Public Property Let setCUSERID(ByVal dato As String)
    CUSERID = dato
End Property
Public Property Let setFTIMESTP(ByVal dato As String)
    FTIMESTP = dato
End Property
Public Property Let setID_TIPO_PLAN(ByVal dato As Long)
    ID_TIPO_PLAN = dato
End Property
Public Property Let setCALENDARIO(ByVal dato As String)
    CALENDARIO = dato
End Property
Public Property Let setExcepciones(ByVal dato As String)
    Excepciones = dato
    Call generarArrayExcepciones
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsPlanMantenimiento
'***************************************************************
Public Property Get getPROTOCOLO_ID() As Long
    getPROTOCOLO_ID = PROTOCOLO_ID
End Property
Public Property Get getPROTOCOLO() As String
    getPROTOCOLO = PROTOCOLO
End Property
Public Property Get getID_PLAN_MTO() As Long
    getID_PLAN_MTO = ID_PLAN_MTO
End Property
Public Property Get getNOMBRE() As String
    getNOMBRE = nombre
End Property
Public Property Get getDESCRIPCION() As String
    getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getFRECUENCIA_ID() As Long
    getFRECUENCIA_ID = FRECUENCIA_ID
End Property
Public Property Get getFRECUENCIA() As String
    getFRECUENCIA = FRECUENCIA
End Property
Public Property Get getPLANIF_MESES() As String
    getPLANIF_MESES = PLANIF_MESES
End Property
Public Property Get getPLANIF_DIAS_SEMANA() As String
    getPLANIF_DIAS_SEMANA = PLANIF_DIAS_SEMANA
End Property
Public Property Get getPLANIF_DIAS_MES() As String
    getPLANIF_DIAS_MES = PLANIF_DIAS_MES
End Property
Public Property Get getORDEN() As Integer
    getORDEN = ORDEN
End Property
Public Property Get getCUSERID() As String
    getCUSERID = CUSERID
End Property
Public Property Get getFTIMESTP() As String
    getFTIMESTP = FTIMESTP
End Property
Public Property Get getID_TIPO_PLAN() As Long
    getID_TIPO_PLAN = ID_TIPO_PLAN
End Property
Public Property Get getCALENDARIO() As String
    getCALENDARIO = CALENDARIO
End Property
Public Property Get getPARAMETROS_FRECUENCIA() As String
    getPARAMETROS_FRECUENCIA = PARAMETROS_FRECUENCIA
End Property

Public Property Get getEXCEPCIONES() As String
    getEXCEPCIONES = Excepciones
End Property



'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsPlanMantenimiento
'***************************************************************
Public Function Carga(ID As Long) As Boolean
    On Error GoTo fallo
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT eq_planes_mantenimiento.*, COALESCE(PER.DESCRIPCION, 'N/A') AS FRECUENCIA, COALESCE(PER.PARAMETROS, '') AS PARAMETROS_FRECUENCIA, COALESCE(CA_DOCUMENTOS.NOMBRE, '') AS PROTOCOLO "
    consulta = consulta & " FROM eq_planes_mantenimiento "
    consulta = consulta & " LEFT OUTER JOIN (SELECT DECODIFICADORA.VALOR AS ID_FRECUENCIA, DECODIFICADORA.DESCRIPCION, DECODIFICADORA.PARAMETROS FROM decodificadora WHERE DECODIFICADORA.CODIGO = " & DECODIFICADORA.EQ_periodicidad & ") AS PER "
    consulta = consulta & " ON eq_planes_mantenimiento.FRECUENCIA_ID = PER.ID_FRECUENCIA"
    consulta = consulta & " LEFT OUTER JOIN ca_documentos on eq_planes_mantenimiento.protocolo_id = ca_documentos.id_documento "
    consulta = consulta & " WHERE ID_PLAN_MTO = " & ID
    
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
        Carga = False
    Else
        ID_PLAN_MTO = rs("ID_PLAN_MTO")
        nombre = rs("NOMBRE")
        DESCRIPCION = rs("DESCRIPCION")
        FRECUENCIA_ID = rs("FRECUENCIA_ID")
        FRECUENCIA = rs("FRECUENCIA")
        PLANIF_MESES = rs("PLANIF_MESES")
        PLANIF_DIAS_SEMANA = rs("PLANIF_DIAS_SEMANA")
        PLANIF_DIAS_MES = rs("PLANIF_DIAS_MES")
        ORDEN = rs("ORDEN")
        CUSERID = rs("CUSERID")
        FTIMESTP = rs("FTIMESTP")
        CALENDARIO = rs("CALENDARIO")
        PROTOCOLO_ID = rs("PROTOCOLO_ID")
        PROTOCOLO = rs("PROTOCOLO")
        ID_TIPO_PLAN = rs("ID_TIPO_PLAN")
        PARAMETROS_FRECUENCIA = rs("PARAMETROS_FRECUENCIA")
        Excepciones = rs("EXCEPCIONES")
        EVITAR_EXCEPCIONES = rs("EVITAR_EXCEPCIONES")
        
        generarArrayExcepciones
        
        Carga = True
    End If
    Set rs = Nothing
    Exit Function
fallo:
    Carga = False
    MsgBox "Error al cargar los datos(clsPlanMantenimiento)", vbCritical, Err.Description
End Function
Public Function CrearID()
    Dim rs As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_PLAN_MTO) FROM eq_planes_mantenimiento"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
        ID_PLAN_MTO = 1
    Else
        ID_PLAN_MTO = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Function
Public Function Insertar() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID
    
    consulta = "INSERT INTO eq_planes_mantenimiento "
    consulta = consulta & "(ID_PLAN_MTO, CUSERID, FTIMESTP)"
    consulta = consulta & "VALUES ("
    consulta = consulta & CStr(ID_PLAN_MTO)
    consulta = consulta & ", " & CStr(USUARIO.getID_EMPLEADO)
    consulta = consulta & ", LOCALTIMESTAMP)"

    
'    consulta = "INSERT INTO eq_planes_mantenimiento " & _
'               "  VALUES (" & _
'            ID_PLAN_MTO & "," & "'" & NOMBRE & "'" & "," & "'" & DESCRIPCION & "'" & "," & _
'            FRECUENCIA_ID & "," & "'" & PLANIF_MESES & "'" & "," & "'" & PLANIF_DIAS_SEMANA & "'" & "," & _
'            "'" & PLANIF_DIAS_MES & "'" & "," & ORDEN & "," & "'" & CUSERID & "'" & "," & _
'            "'" & FTIMESTP & "'" & _
'        ")"

    execute_bd consulta
    Insertar = ID_PLAN_MTO
    
    Me.Modificar ID_PLAN_MTO
    
    Exit Function
fallo:
    Insertar = 0
    MsgBox "Error al insertar (clsPlanMantenimiento)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
    On Error GoTo fallo
    Dim consulta As String
    consulta = "UPDATE eq_planes_mantenimiento SET " & _
            " NOMBRE = '" & nombre & "'," & _
            " DESCRIPCION = '" & DESCRIPCION & "'," & _
            " FRECUENCIA_ID = " & FRECUENCIA_ID & "," & _
            " PLANIF_MESES = '" & PLANIF_MESES & "'," & _
            " PLANIF_DIAS_SEMANA = '" & PLANIF_DIAS_SEMANA & "'," & _
            " PLANIF_DIAS_MES = '" & PLANIF_DIAS_MES & "'," & _
            " CALENDARIO = '" & CALENDARIO & "'," & _
            " ID_TIPO_PLAN = " & ID_TIPO_PLAN & "," & _
            " EXCEPCIONES = '" & Excepciones & "'," & _
            " EVITAR_EXCEPCIONES = " & EVITAR_EXCEPCIONES & "," & _
            " ORDEN = " & ORDEN & "," & _
            " PROTOCOLO_ID = " & PROTOCOLO_ID & "," & _
            " CUSERID = '" & USUARIO.getID_EMPLEADO & "'," & _
            " FTIMESTP = localtimestamp" & _
        " WHERE ID_PLAN_MTO = " & ID
    execute_bd consulta
    Modificar = True
    Exit Function
fallo:
    Modificar = False
    MsgBox "Error al Modificar (clsPlanMantenimiento)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
On Error GoTo fallo
    Dim consulta As String, strListaEq As String, intCont As Integer
    Dim rs As New ADODB.Recordset
    
    
    strListaEq = ""
    
    ' Se comprueba si hay equipos con ese plan asignado
    consulta = "select pm.*, e.nombre from eq_planes_mantenimiento_equipos pm inner join equipos e on pm.equipo_id = e.ID_EQUIPO"

    Set rs = datos_bd(consulta)
    If rs.RecordCount <> 0 Then
        rs.MoveFirst
        intCont = 0
        While Not rs.EOF Or intCont < 10
            strListaEq = strListaEq & vbCrLf & " - " & rs("NOMBRE") & " [ Nº " & CStr(rs("EQUIPO_ID")) & "]"
            intCont = intCont + 1
            rs.MoveNext
        Wend
        If intCont >= 10 Then
            strListaEq = strListaEq & vbCrLf & " - ...(Varios más)"
        End If
        MsgBox "No se puede eliminar el plan de mantenimiento. Está asociado los siguientes Equipos:" & strListaEq, vbExclamation, App.Title
        Eliminar = False
        Exit Function
    End If
        
    ' Se borra el detalle y el plan
    
    Call execute_bd("DELETE FROM eq_planes_mantenimiento_equipos WHERE PLAN_MANTENIMIENTO = " & ID)
    Call execute_bd("DELETE FROM eq_planes_mantenimiento_detalle WHERE PLAN_MANTENIMIENTO = " & ID)
    Call execute_bd("DELETE FROM eq_planes_mantenimiento WHERE PLAN_MANTENIMIENTO = " & ID)
    
    Eliminar = True
    
    Exit Function
fallo:
    Eliminar = False
    MsgBox "Error al Eliminar (clsPlanMantenimiento)", vbCritical, Err.Description
End Function
Public Function Listado(ByVal prmNombre As String, ByVal prmDesc As String, ByVal prmidPeriodicidad As Long) As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT eq_planes_mantenimiento.*, COALESCE(PER.DESCRIPCION, 'N/A') AS FRECUENCIA, COALESCE(PER.PARAMETROS, '') AS PARAMETROS_FRECUENCIA "
    consulta = consulta & "FROM eq_planes_mantenimiento "
    consulta = consulta & " LEFT OUTER JOIN (SELECT DECODIFICADORA.VALOR AS ID_FRECUENCIA, DECODIFICADORA.DESCRIPCION, DECODIFICADORA.PARAMETROS FROM decodificadora WHERE DECODIFICADORA.CODIGO = " & DECODIFICADORA.EQ_periodicidad & ") AS PER "
    consulta = consulta & " ON eq_planes_mantenimiento.FRECUENCIA_ID = PER.ID_FRECUENCIA"
    consulta = consulta & " WHERE eq_planes_mantenimiento.NOMBRE LIKE '%" & prmNombre & "%'"
    consulta = consulta & " AND eq_planes_mantenimiento.DESCRIPCION LIKE '%" & prmDesc & "%'"
    
    If prmidPeriodicidad <> -1 Then
        consulta = consulta & " AND eq_planes_mantenimiento.FRECUENCIA_ID = " & prmidPeriodicidad
    End If
    consulta = consulta & " order by eq_planes_mantenimiento.id_plan_mto"
    
    Set Listado = datos_bd(consulta)
End Function
Public Function Listado_Combo() As ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT ID_PLAN_MTO,NOMBRE FROM eq_planes_mantenimiento ORDER BY NOMBRE"
    Set Listado_Combo = datos_bd(consulta)
End Function


Private Sub generarArrayExcepciones()
    
    Dim cont As Integer
    Dim arr() As String
        
    indiceExcepciones = 0
    If Trim(Excepciones) = "" Or Trim(Excepciones) = "''" Then Exit Sub
    
    arr = Split(Excepciones, ";")
    
    indiceExcepciones = UBound(arr) + 1
    
    For cont = 0 To indiceExcepciones - 1
        ReDim Preserve arrExcepciones(cont)
        arrExcepciones(cont) = Int(arr(cont))
    Next cont
    
End Sub

Public Function getArrayExcepciones() As Integer()
    
    Dim cont As Integer
    Dim arr() As String
        
    indiceExcepciones = 0
    If Trim(Excepciones) = "" Then
        ReDim Preserve arr(0)
        getArrayExcepciones = arrExcepciones
        Exit Function
    End If
    
    arr = Split(Excepciones, ";")
    
    indiceExcepciones = UBound(arr) + 1
    
    For cont = 0 To indiceExcepciones - 1
        ReDim Preserve arrExcepciones(cont)
        arrExcepciones(cont) = Int(arr(cont))
    Next cont
    
    
    getArrayExcepciones = arrExcepciones
    
End Function


Public Function getConsultaListadoGenerico() As String
    
    getConsultaListadoGenerico = "SELECT ID_PLAN_MTO, NOMBRE AS 'Plan', DESCRIPCION FROM eq_planes_mantenimiento WHERE 1=1"

End Function


Public Function devolver_acciones(ID_PLAN As Long) As ADODB.Recordset

Dim sql As String

sql = "select * from eq_planes_acciones where id_accion in (SELECT accion_id FROM eq_planes_mantenimiento_detalle where plan_mto_id = " & CStr(ID_PLAN) & ")"

Set devolver_acciones = datos_bd(sql)
    

End Function

