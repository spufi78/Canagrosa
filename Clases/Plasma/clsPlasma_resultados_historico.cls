VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_resultados_historico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsPlasma_resultados_historico
'***************************************************************
Private MUESTRA_ID As Long
Private EDICION As Long
Private tipo As Long
Private DESIGNACION As String
Private RESULTADO As String
Private CONFORME As Long
Private fecha As String
Private hora As String
Private EMPLEADO_ID As Long
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsPlasma_resultados_historico
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setDESIGNACION(ByVal dato As String)
        DESIGNACION = dato
End Property
Public Property Let setEDICION(ByVal dato As Long)
        EDICION = dato
End Property
Public Property Let setTIPO(ByVal dato As Long)
        tipo = dato
End Property
Public Property Let setRESULTADO(ByVal dato As String)
        RESULTADO = dato
End Property
Public Property Let setCONFORME(ByVal dato As Long)
        CONFORME = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
        fecha = dato
End Property
Public Property Let setHORA(ByVal dato As String)
        hora = dato
End Property
Public Property Let setEMPLEADO_ID(ByVal dato As Long)
        EMPLEADO_ID = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsPlasma_resultados_historico
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getDESIGNACION() As String
        getDESIGNACION = DESIGNACION
End Property
Public Property Get getEDICION() As Long
        getEDICION = EDICION
End Property
Public Property Get getTIPO() As Long
        getTIPO = tipo
End Property
Public Property Get getRESULTADO() As String
        getRESULTADO = RESULTADO
End Property
Public Property Get getCONFORME() As Long
        getCONFORME = CONFORME
End Property
Public Property Get getFECHA() As String
        getFECHA = fecha
End Property
Public Property Get getHORA() As String
        getHORA = hora
End Property
Public Property Get getEMPLEADO_ID() As Long
        getEMPLEADO_ID = EMPLEADO_ID
End Property
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Dim rs As ADODB.Recordset
        ' Varificamos antes de insertar, si el resultado del campo es
        ' distinto del anterior
        CONSULTA = "SELECT RESULTADO,CONFORME " & _
                   "  FROM plasma_resultados_historico " & _
                   " WHERE MUESTRA_ID = " & MUESTRA_ID & _
                   "   AND EDICION = " & EDICION & _
                   "   AND TIPO = " & tipo & _
                   "   AND DESIGNACION = '" & DESIGNACION & "'"
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount > 0 Then
            If rs(0) = RESULTADO And rs(1) = CONFORME Then
                Insertar = MUESTRA_ID
                Exit Function
            End If
        End If
        ' Insertamos el resultado
        Me.setEMPLEADO_ID = usuario.getID_EMPLEADO
        CONSULTA = "REPLACE INTO plasma_resultados_historico " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & EDICION & "," & tipo & ",'" & DESIGNACION & "'," & _
                        "'" & RESULTADO & "'" & "," & _
                        CONFORME & ",CURRENT_DATE,CURRENT_TIME," & EMPLEADO_ID & _
                ")"
        execute_bd CONSULTA
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlasma_resultados_historico)", vbCritical, Err.Description
End Function
Public Function InsertarConversion() As Long
        On Error GoTo fallo
        Dim CONSULTA As String
        Dim rs As ADODB.Recordset
        ' Varificamos antes de insertar, si el resultado del campo es
        ' distinto del anterior
        CONSULTA = "SELECT RESULTADO,CONFORME " & _
                   "  FROM plasma_resultados_historico " & _
                   " WHERE MUESTRA_ID = " & MUESTRA_ID & _
                   "   AND TIPO = " & tipo & _
                   "   AND DESIGNACION = '" & DESIGNACION & "'"
        Set rs = datos_bd(CONSULTA)
        If rs.RecordCount > 0 Then
            If rs(0) = RESULTADO And rs(1) = CONFORME Then
                InsertarConversion = MUESTRA_ID
                Exit Function
            End If
        End If
        ' Insertamos el resultado
        CONSULTA = "INSERT INTO plasma_resultados_historico " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & EDICION & "," & tipo & ",'" & DESIGNACION & "'," & _
                        "'" & RESULTADO & "'" & "," & _
                        CONFORME & ",'" & Format(fecha, "yyyy-mm-dd") & "','" & Format(hora, "hh:mm:ss") & "'," & EMPLEADO_ID & _
                ")"
        execute_bd CONSULTA
        InsertarConversion = MUESTRA_ID
        Exit Function
fallo:
        InsertarConversion = 0
        MsgBox "Error al InsertarConversion (clsPlasma_resultados_historico)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "DELETE FROM plasma_resultados_historico " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd CONSULTA
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlasma_resultados_historico)", vbCritical, Err.Description
End Function
Public Function Listado(ID As Long) As ADODB.Recordset
        Dim CONSULTA As String
        CONSULTA = "SELECT b.usuario as ANALISTA, a.* FROM plasma_resultados_historico a,usuarios b " & _
                   " WHERE MUESTRA_ID = " & ID & _
                   "   and a.empleado_id = b.id_empleado " & _
                   " ORDER BY TIPO,DESIGNACION,EDICION"
        Set Listado = datos_bd(CONSULTA)
End Function
Public Function ModificarVida(lMUESTRA_ID As Long, lEDICION As Long, lTIPO As Long, lDESIGNACION As String) As Boolean
        On Error GoTo fallo
        Dim CONSULTA As String
        CONSULTA = "UPDATE plasma_resultados_historico SET " & _
                        " FECHA = '" & fecha & "'," & _
                        " HORA = '" & hora & "'," & _
                        " EMPLEADO_ID = " & EMPLEADO_ID & "" & _
                " WHERE MUESTRA_ID = " & lMUESTRA_ID & _
                "   AND EDICION = " & lEDICION & _
                "   AND TIPO = " & lTIPO & _
                "   AND DESIGNACION = '" & lDESIGNACION & "'"
        execute_bd CONSULTA
        ModificarVida = True
        Exit Function
fallo:
        ModificarVida = False
        MsgBox "Error al ModificarVida (clsPlasma_resultados_historico)", vbCritical, Err.Description
End Function

Public Function generar(MUESTRA_ID As Long)
    Dim oPP As New clsPlasma_procesos
    Dim oPR As New clsPlasma_resultados
    Dim oPF As New clsPlasma_ficha
    Dim OPReception As New clsPlasma_recepcion
    Dim oPFE As New clsPlasma_ficha_estructura
    Dim oPRH As New clsPlasma_resultados_historico
    Dim rs As ADODB.Recordset
    Dim rsFicha As ADODB.Recordset
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra MUESTRA_ID
    OPReception.Carga MUESTRA_ID
            
            oPP.Carga OPReception.getPROCESO_ID
            oPR.Carga MUESTRA_ID, 1
            
            With oPRH
                .setMUESTRA_ID = MUESTRA_ID
                .setEDICION = oMuestra.getULT_EDICION_IMP + 1
                .setTIPO = 1
                ' MICROESTRUCTURA
                If oPP.getBOND_MICROESTRUCTURA = 1 Then
                    oPF.Carga oPP.getBOND_COAT_FICHA_ID
                    If oPF.getMETCO <> "N/A" Then
                        .setDESIGNACION = "METALLOGRAPHIC EXAMINATION"
                        .setRESULTADO = ""
                        .setCONFORME = 1
                        If oPR.getMICROESTRUCTURA1 = 2 Or oPR.getMICROESTRUCTURA2 = 2 Or oPR.getMICROESTRUCTURA3 = 2 Or oPR.getMICROESTRUCTURA4 = 2 Or oPR.getMICROESTRUCTURA5 = 2 Or oPR.getMICROESTRUCTURA6 = 2 Then
                            .setCONFORME = 2
                        Else
                            Set rsFicha = oPFE.ListadoCompleto(oPP.getBOND_COAT_FICHA_ID)
                            If rsFicha.RecordCount > 0 Then
                                Do
                                    If rsFicha("REQUIREMENT") <> "" And rsFicha("REQUIREMENT") <> "N/A" Then
                                        Select Case rsFicha("ENSAYO_ID")
                                        Case 1
                                            If oPR.getMICROESTRUCTURA1 = 0 Then .setCONFORME = 0
                                        Case 2
                                            If oPR.getMICROESTRUCTURA2 = 0 Then .setCONFORME = 0
                                        Case 3
                                            If oPR.getMICROESTRUCTURA3 = 0 Then .setCONFORME = 0
                                        Case 4
                                            If oPR.getMICROESTRUCTURA4 = 0 Then .setCONFORME = 0
                                        Case 5
                                            If oPR.getMICROESTRUCTURA5 = 0 Then .setCONFORME = 0
                                        Case 6
                                            If oPR.getMICROESTRUCTURA6 = 0 Then .setCONFORME = 0
                                        End Select
                                    End If
                                    rsFicha.MoveNext
                                Loop Until rsFicha.EOF
                            End If
                        End If
                        .Insertar
                    End If
                End If
                ' TRACCION
                If oPR.getTRACCION_RES <> "" Then
                    .setDESIGNACION = "TENSILE STRENGTH"
                    .setRESULTADO = oPR.getTRACCION_RES
                    .setCONFORME = oPR.getTRACCION_PASS
                    .Insertar
                End If
                ' MACRO
                If oPR.getMACRO_DUREZA_RES <> "" Then
                    .setDESIGNACION = "MACRO HARDNESS"
                    .setRESULTADO = oPR.getMACRO_DUREZA_RES
                    .setCONFORME = oPR.getMACRO_DUREZA_PASS
                    .Insertar
                End If
                ' MICRO
                If oPR.getMICRO_DUREZA_RES <> "" Then
                    .setDESIGNACION = "MICRO HARDNESS"
                    .setRESULTADO = oPR.getMICRO_DUREZA_RES
                    .setCONFORME = oPR.getMICRO_DUREZA_PASS
                    .Insertar
                End If
                ' ESPESOR
                If oPR.getESPESOR_RES <> "" Then
                    .setDESIGNACION = "THICKNESS"
                    .setRESULTADO = oPR.getESPESOR_RES
                    .setCONFORME = oPR.getESPESOR_PASS
                    .Insertar
                End If
            End With
            oPR.Carga MUESTRA_ID, 2
            With oPRH
                .setMUESTRA_ID = MUESTRA_ID
                .setEDICION = oMuestra.getULT_EDICION_IMP + 1
                .setTIPO = 2
                'MICROESTRUCTURA
                If oPP.getTOP_MICROESTRUCTURA = 1 Then
                    oPF.Carga oPP.getTOP_COAT_FICHA_ID
                    If oPF.getMETCO <> "N/A" Then
                        .setDESIGNACION = "METALLOGRAPHIC EXAMINATION"
                        .setRESULTADO = ""
                        .setCONFORME = 1
                        If oPR.getMICROESTRUCTURA1 = 2 Or oPR.getMICROESTRUCTURA2 = 2 Or oPR.getMICROESTRUCTURA3 = 2 Or oPR.getMICROESTRUCTURA4 = 2 Or oPR.getMICROESTRUCTURA5 = 2 Or oPR.getMICROESTRUCTURA6 = 2 Then
                            .setCONFORME = 2
                        Else
                            Set rsFicha = oPFE.ListadoCompleto(oPP.getTOP_COAT_FICHA_ID)
                            If rsFicha.RecordCount > 0 Then
                                Do
                                    If rsFicha("REQUIREMENT") <> "" And rsFicha("REQUIREMENT") <> "N/A" Then
                                        Select Case rsFicha("ENSAYO_ID")
                                        Case 1
                                            If oPR.getMICROESTRUCTURA1 = 0 Then .setCONFORME = 0
                                        Case 2
                                            If oPR.getMICROESTRUCTURA2 = 0 Then .setCONFORME = 0
                                        Case 3
                                            If oPR.getMICROESTRUCTURA3 = 0 Then .setCONFORME = 0
                                        Case 4
                                            If oPR.getMICROESTRUCTURA4 = 0 Then .setCONFORME = 0
                                        Case 5
                                            If oPR.getMICROESTRUCTURA5 = 0 Then .setCONFORME = 0
                                        Case 6
                                            If oPR.getMICROESTRUCTURA6 = 0 Then .setCONFORME = 0
                                        End Select
                                    End If
                                    rsFicha.MoveNext
                                Loop Until rsFicha.EOF
                            End If
                        End If
                        .Insertar
                    End If
                End If
                ' TRACCION
                If oPR.getTRACCION_RES <> "" Then
                    .setDESIGNACION = "TENSILE STRENGTH"
                    .setRESULTADO = oPR.getTRACCION_RES
                    .setCONFORME = oPR.getTRACCION_PASS
                    .Insertar
                End If
                ' MACRO
                If oPR.getMACRO_DUREZA_RES <> "" Then
                    .setDESIGNACION = "MACRO HARDNESS"
                    .setRESULTADO = oPR.getMACRO_DUREZA_RES
                    .setCONFORME = oPR.getMACRO_DUREZA_PASS
                    .Insertar
                End If
                ' MICRO
                If oPR.getMICRO_DUREZA_RES <> "" Then
                    .setDESIGNACION = "MICRO HARDNESS"
                    .setRESULTADO = oPR.getMICRO_DUREZA_RES
                    .setCONFORME = oPR.getMICRO_DUREZA_PASS
                    .Insertar
                End If
                ' ESPESOR
                If oPR.getESPESOR_RES <> "" Then
                    .setDESIGNACION = "THICKNESS"
                    .setRESULTADO = oPR.getESPESOR_RES
                    .setCONFORME = oPR.getESPESOR_PASS
                    .Insertar
                End If
            End With
        
End Function
Public Function generar_dureza(MUESTRA_ID As Long)
    Dim oMuestra As New clsMuestra
    Dim opd As New clsPlasma_dureza
    Dim rs As ADODB.Recordset
    oMuestra.CargaMuestra MUESTRA_ID
    opd.Carga MUESTRA_ID
            
    With Me
        .setMUESTRA_ID = MUESTRA_ID
        .setEDICION = oMuestra.getULT_EDICION_IMP + 1
        .setTIPO = 0
        .setDESIGNACION = opd.getIDENTIFICATION
        .setRESULTADO = opd.getAVERAGE
        .setCONFORME = opd.getPASS
        .Insertar
    End With
End Function

