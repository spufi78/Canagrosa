VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlasma_recepcion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSPLASMA_RECEPCION
'***************************************************************
Private MUESTRA_ID As Long
Private PROCESO_ID As Long
Private CUSTOMER_ID As Long
Private SPECIMEN_ID As String
Private NTYPE As Long
Private PN As String
Private sn As String
Private PRODUCT_TYPE As String
Private PRODUCT_SN As String
Private MODULE_SN As String
Private BOND_CONTROL_SPECIFICATION As String
Private TOP_CONTROL_SPECIFICATION As String
Private RESULT As Integer

Private MP As Integer
Private MP_FECHA As String
Private MP_USUARIO_ID As Integer
Private MP_PASS As Integer
Private MACRO_DUREZA_T1 As String
Private MACRO_DUREZA_T2 As String
Private MICRO_DUREZA_T1 As String
Private MICRO_DUREZA_T2 As String


'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSPLASMA_RECEPCION
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setPROCESO_ID(ByVal dato As Long)
        PROCESO_ID = dato
End Property
Public Property Let setCUSTOMER_ID(ByVal dato As Long)
        CUSTOMER_ID = dato
End Property
Public Property Let setSPECIMEN_ID(ByVal dato As String)
        SPECIMEN_ID = dato
End Property
Public Property Let setNTYPE(ByVal dato As Long)
        NTYPE = dato
End Property
Public Property Let setPN(ByVal dato As String)
        PN = dato
End Property
Public Property Let setSN(ByVal dato As String)
        sn = dato
End Property
Public Property Let setPRODUCT_TYPE(ByVal dato As String)
        PRODUCT_TYPE = dato
End Property
Public Property Let setPRODUCT_SN(ByVal dato As String)
        PRODUCT_SN = dato
End Property
Public Property Let setMODULE_SN(ByVal dato As String)
        MODULE_SN = dato
End Property
Public Property Let setBOND_CONTROL_SPECIFICATION(ByVal dato As String)
        BOND_CONTROL_SPECIFICATION = dato
End Property
Public Property Let setTOP_CONTROL_SPECIFICATION(ByVal dato As String)
        TOP_CONTROL_SPECIFICATION = dato
End Property
Public Property Let setRESULT(ByVal dato As Integer)
        RESULT = dato
End Property
Public Property Let setMP(ByVal dato As Integer)
        MP = dato
End Property
Public Property Let setMP_FECHA(ByVal dato As String)
        MP_FECHA = dato
End Property
Public Property Let setMP_USUARIO_ID(ByVal dato As Integer)
        MP_USUARIO_ID = dato
End Property
Public Property Let setMP_PASS(ByVal dato As Integer)
        MP_PASS = dato
End Property
Public Property Let setMACRO_DUREZA_T1(ByVal dato As String)
        MACRO_DUREZA_T1 = dato
End Property
Public Property Let setMACRO_DUREZA_T2(ByVal dato As String)
        MACRO_DUREZA_T2 = dato
End Property
Public Property Let setMICRO_DUREZA_T1(ByVal dato As String)
        MICRO_DUREZA_T1 = dato
End Property
Public Property Let setMICRO_DUREZA_T2(ByVal dato As String)
        MICRO_DUREZA_T2 = dato
End Property

'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSPLASMA_RECEPCION
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getPROCESO_ID() As Long
        getPROCESO_ID = PROCESO_ID
End Property
Public Property Get getCUSTOMER_ID() As Long
        getCUSTOMER_ID = CUSTOMER_ID
End Property
Public Property Get getSPECIMEN_ID() As String
        getSPECIMEN_ID = SPECIMEN_ID
End Property
Public Property Get getNTYPE() As Long
        getNTYPE = NTYPE
End Property
Public Property Get getPN() As String
        getPN = PN
End Property
Public Property Get getSN() As String
        getSN = sn
End Property
Public Property Get getPRODUCT_TYPE() As String
        getPRODUCT_TYPE = PRODUCT_TYPE
End Property
Public Property Get getPRODUCT_SN() As String
        getPRODUCT_SN = PRODUCT_SN
End Property
Public Property Get getMODULE_SN() As String
        getMODULE_SN = MODULE_SN
End Property
Public Property Get getBOND_CONTROL_SPECIFICATION() As String
        getBOND_CONTROL_SPECIFICATION = BOND_CONTROL_SPECIFICATION
End Property
Public Property Get getTOP_CONTROL_SPECIFICATION() As String
        getTOP_CONTROL_SPECIFICATION = TOP_CONTROL_SPECIFICATION
End Property
Public Property Get getRESULT() As Integer
        getRESULT = RESULT
End Property

Public Property Get getMP() As Integer
        getMP = MP
End Property
Public Property Get getMP_FECHA() As String
        getMP_FECHA = MP_FECHA
End Property
Public Property Get getMP_USUARIO_ID() As Integer
        getMP_USUARIO_ID = MP_USUARIO_ID
End Property
Public Property Get getMP_PASS() As Integer
        getMP_PASS = MP_PASS
End Property

Public Property Get getMACRO_DUREZA_T1() As String
        getMACRO_DUREZA_T1 = MACRO_DUREZA_T1
End Property
Public Property Get getMACRO_DUREZA_T2() As String
        getMACRO_DUREZA_T2 = MACRO_DUREZA_T2
End Property
Public Property Get getMICRO_DUREZA_T1() As String
        getMICRO_DUREZA_T1 = MICRO_DUREZA_T1
End Property
Public Property Get getMICRO_DUREZA_T2() As String
        getMICRO_DUREZA_T2 = MICRO_DUREZA_T2
End Property

'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSPLASMA_RECEPCION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_RECEPCION WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                PROCESO_ID = rs("PROCESO_ID")
                CUSTOMER_ID = rs("CUSTOMER_ID")
                SPECIMEN_ID = rs("SPECIMEN_ID")
                NTYPE = rs("NTYPE")
                PN = rs("PN")
                sn = rs("SN")
                PRODUCT_TYPE = rs("PRODUCT_TYPE")
                PRODUCT_SN = rs("PRODUCT_SN")
                MODULE_SN = rs("MODULE_SN")
                BOND_CONTROL_SPECIFICATION = rs("BOND_CONTROL_SPECIFICATION")
                TOP_CONTROL_SPECIFICATION = rs("TOP_CONTROL_SPECIFICATION")
                RESULT = rs("RESULT")
                
                MP = rs("MP")
                If Not IsNull(rs("MP_FECHA")) Then
                    MP_FECHA = rs("MP_FECHA")
                Else
                    MP_FECHA = ""
                End If
                MP_USUARIO_ID = rs("MP_USUARIO_ID")
                If Not IsNull(rs("MP_PASS")) Then
                    MP_PASS = rs("MP_PASS")
                Else
                    MP_PASS = 0
                End If
                MACRO_DUREZA_T1 = rs("MACRO_DUREZA_T1")
                MACRO_DUREZA_T2 = rs("MACRO_DUREZA_T2")
                MICRO_DUREZA_T1 = rs("MICRO_DUREZA_T1")
                MICRO_DUREZA_T2 = rs("MICRO_DUREZA_T2")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsPlasma_recepcion)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO PLASMA_RECEPCION " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & PROCESO_ID & "," & CUSTOMER_ID & "," & _
                        "'" & SPECIMEN_ID & "'," & NTYPE & ",'" & PN & "'" & "," & "'" & sn & "'" & "," & _
                        "'" & PRODUCT_TYPE & "'" & "," & "'" & PRODUCT_SN & "'" & "," & "'" & MODULE_SN & "'" & "," & _
                        "'" & BOND_CONTROL_SPECIFICATION & "'," & _
                        "'" & TOP_CONTROL_SPECIFICATION & "',1," & _
                        MP & "," & MP_FECHA & "," & MP_USUARIO_ID & "," & MP_PASS & "," & _
                        "'" & MACRO_DUREZA_T1 & "','" & MACRO_DUREZA_T2 & "'," & _
                        "'" & MICRO_DUREZA_T1 & "','" & MICRO_DUREZA_T2 & "'" & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsPlasma_recepcion)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PLASMA_RECEPCION SET " & _
                        " PROCESO_ID = " & PROCESO_ID & "," & _
                        " CUSTOMER_ID = " & CUSTOMER_ID & "," & _
                        " SPECIMEN_ID = '" & SPECIMEN_ID & "'," & _
                        " NTYPE = " & NTYPE & "," & _
                        " PN = '" & PN & "'," & _
                        " SN = '" & sn & "'," & _
                        " PRODUCT_TYPE = '" & PRODUCT_TYPE & "'," & _
                        " PRODUCT_SN = '" & PRODUCT_SN & "'," & _
                        " MODULE_SN = '" & MODULE_SN & "'," & _
                        " MP = " & MP & "," & _
                        " MP_FECHA = " & MP_FECHA & "," & _
                        " MP_USUARIO_ID = " & MP_USUARIO_ID & "," & _
                        " MP_PASS = " & MP_PASS & _
                        " ,MACRO_DUREZA_T1 = '" & MACRO_DUREZA_T1 & "'" & _
                        " ,MACRO_DUREZA_T2 = '" & MACRO_DUREZA_T2 & "'" & _
                        " ,MICRO_DUREZA_T1 = '" & MICRO_DUREZA_T1 & "'" & _
                        " ,MICRO_DUREZA_T2 = '" & MICRO_DUREZA_T2 & "'" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsPlasma_recepcion)", vbCritical, Err.Description
End Function
Public Function ModificarCopia(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PLASMA_RECEPCION SET " & _
                        " MP = " & MP & _
                        " ,MP_FECHA = '" & MP_FECHA & "'" & _
                        " ,MP_USUARIO_ID = " & MP_USUARIO_ID & _
                        " ,MP_PASS = " & MP_PASS & _
                        " ,MACRO_DUREZA_T1 = '" & MACRO_DUREZA_T1 & "'" & _
                        " ,MACRO_DUREZA_T2 = '" & MACRO_DUREZA_T2 & "'" & _
                        " ,MICRO_DUREZA_T1 = '" & MICRO_DUREZA_T1 & "'" & _
                        " ,MICRO_DUREZA_T2 = '" & MICRO_DUREZA_T2 & "'" & _
                        " ,RESULT = '" & RESULT & "'" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        ModificarCopia = True
        Exit Function
fallo:
        ModificarCopia = False
        MsgBox "Error al Modificar (clsPlasma_recepcion)", vbCritical, Err.Description
End Function

Public Function ModificarResultado(ID As Long, resultado As Integer) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE PLASMA_RECEPCION " & _
                   "   SET RESULT = " & resultado & _
                   " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        ModificarResultado = True
        Exit Function
fallo:
        ModificarResultado = False
        MsgBox "Error al ModificarResultado (clsPlasma_recepcion)", vbCritical, Err.Description
End Function

Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM PLASMA_RECEPCION " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsPlasma_recepcion)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM PLASMA_RECEPCION ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function
Public Function informarRequirement(MUESTRA_ID As Long)
    Dim oPR As New clsPlasma_recepcion
    Dim oPP As New clsPlasma_procesos
    Dim oPF As New clsPlasma_ficha
    Dim oPFE As New clsPlasma_ficha_estructura
    Dim consulta As String
   On Error GoTo informarRequirement_Error

    If oPR.Carga(MUESTRA_ID) Then
            If oPP.Carga(oPR.getPROCESO_ID) Then
                If oPF.Carga(oPP.getBOND_COAT_FICHA_ID) Then
                    consulta = ""
                    If oPFE.Carga(oPF.getID_FICHA, 1) Then
                        consulta = consulta & " microestructura1_r = '" & oPFE.getREQUIREMENT & "',"
                    End If
                    If oPFE.Carga(oPF.getID_FICHA, 2) Then
                        consulta = consulta & " microestructura2_r = '" & oPFE.getREQUIREMENT & "',"
                    End If
                    If oPFE.Carga(oPF.getID_FICHA, 3) Then
                        consulta = consulta & " microestructura3_r = '" & oPFE.getREQUIREMENT & "',"
                    End If
                    If oPFE.Carga(oPF.getID_FICHA, 4) Then
                        consulta = consulta & " microestructura4_r = '" & oPFE.getREQUIREMENT & "',"
                    End If
                    If oPFE.Carga(oPF.getID_FICHA, 5) Then
                        consulta = consulta & " microestructura5_r = '" & oPFE.getREQUIREMENT & "',"
                    End If
                    If oPFE.Carga(oPF.getID_FICHA, 6) Then
                        consulta = consulta & " microestructura6_r = '" & oPFE.getREQUIREMENT & "'"
                    End If
                    If consulta <> "" Then
                        consulta = "UPDATE plasma_resultados SET " & consulta & " where muestra_id = " & MUESTRA_ID & " and tipo = 1"
                        execute_bd consulta
                    End If
                End If
                If oPF.Carga(oPP.getTOP_COAT_FICHA_ID) Then
                    consulta = ""
                    If oPFE.Carga(oPF.getID_FICHA, 1) Then
                        consulta = consulta & " microestructura1_r = '" & oPFE.getREQUIREMENT & "',"
                    End If
                    If oPFE.Carga(oPF.getID_FICHA, 2) Then
                        consulta = consulta & " microestructura2_r = '" & oPFE.getREQUIREMENT & "',"
                    End If
                    If oPFE.Carga(oPF.getID_FICHA, 3) Then
                        consulta = consulta & " microestructura3_r = '" & oPFE.getREQUIREMENT & "',"
                    End If
                    If oPFE.Carga(oPF.getID_FICHA, 4) Then
                        consulta = consulta & " microestructura4_r = '" & oPFE.getREQUIREMENT & "',"
                    End If
                    If oPFE.Carga(oPF.getID_FICHA, 5) Then
                        consulta = consulta & " microestructura5_r = '" & oPFE.getREQUIREMENT & "',"
                    End If
                    If oPFE.Carga(oPF.getID_FICHA, 6) Then
                        consulta = consulta & " microestructura6_r = '" & oPFE.getREQUIREMENT & "'"
                    End If
                    If consulta <> "" Then
                        consulta = "UPDATE plasma_resultados SET " & consulta & " where muestra_id = " & MUESTRA_ID & " and tipo = 2"
                        execute_bd consulta
                    End If
                End If
            End If
    End If

   On Error GoTo 0
   Exit Function

informarRequirement_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure informarRequirement of Módulo de clase clsPlasma_recepcion"
End Function
Public Function informarEquiposRecepcion(MUESTRA_ID As Long, PROCESO_ID As Long)
    Dim consulta As String
   On Error GoTo informarEquiposRecepcion_Error

'    CONSULTA = " insert into plasma_equipos " & _
'               " select distinct " & MUESTRA_ID & ",equipo_id,equipo_id,0,en_informe from plasma_ensayos_equipos aaa," & _
'               "  ( " & _
'               " select group_concat(aa.c1) c1 from ( " & _
'               " select concat(microestructura,IF(traccion<>0,concat(',',traccion),''),IF(macro_dureza<>0,concat(',',macro_dureza),''),IF(micro_dureza<>0,concat(',',micro_dureza),''),IF(espesor<>0,concat(',',espesor),'')) as c1 from plasma_ficha " & _
'               " where id_ficha = (select bond_coat_ficha_id from plasma_procesos where id_proceso = " & PROCESO_ID & ") " & _
'               "    or id_ficha = (select top_coat_ficha_id from plasma_procesos where id_proceso = " & PROCESO_ID & ") " & _
'               "  ) aa " & _
'               " ) bbb " & _
'               " where find_in_set(aaa.ENSAYO_ID,bbb.c1) " & _
'               " order by equipo_id "
               
    consulta = " REPLACE into plasma_equipos " & _
               " select distinct " & MUESTRA_ID & ",@rownum:=@rownum+1 as orden,equipo_id,0,en_informe from (SELECT @rownum:=0) r, plasma_ensayos_equipos aaa," & _
               "  ( " & _
               " select group_concat(aa.c1) c1 from ( " & _
                " SELECT CONCAT(IF(b.BOND_MICROESTRUCTURA=1,a.microestructura,''), " & _
                "              IF(b.BOND_TRACCION=1 and a.traccion<>0, CONCAT(',',a.traccion),''), " & _
                "                  IF(b.BOND_MACRO_DUREZA=1 and a.macro_dureza<>0, CONCAT(',',a.macro_dureza),''), " & _
                "                  IF(b.BOND_MICRO_DUREZA=1 and a.micro_dureza<>0, CONCAT(',',a.micro_dureza),''), " & _
                "                  IF(b.BOND_ESPESOR=1 and  a.espesor<>0, CONCAT(',',a.espesor),''), IF(a.shorea<>0, CONCAT(',',a.shorea),'')) AS c1 " & _
                "  FROM plasma_ficha a, plasma_procesos b " & _
                " where a.ID_FICHA = b.BOND_COAT_FICHA_ID " & _
                "   AND b.ID_PROCESO = " & PROCESO_ID & _
                " Union " & _
                " SELECT CONCAT(IF(b.TOP_MICROESTRUCTURA=1,a.microestructura,''), " & _
                "              IF(b.TOP_TRACCION=1 and a.traccion<>0, CONCAT(',',a.traccion),''), " & _
                "                  IF(b.TOP_MACRO_DUREZA=1 and a.macro_dureza<>0, CONCAT(',',a.macro_dureza),''), " & _
                "                  IF(b.TOP_MICRO_DUREZA=1 and a.micro_dureza<>0, CONCAT(',',a.micro_dureza),''), " & _
                "                  IF(b.TOP_ESPESOR=1 and  a.espesor<>0, CONCAT(',',a.espesor),'')) AS c1 " & _
                "  FROM plasma_ficha a, plasma_procesos b " & _
                " where a.ID_FICHA = b.TOP_COAT_FICHA_ID " & _
                "   AND b.ID_PROCESO = " & PROCESO_ID & _
                   "  ) aa " & _
               " ) bbb " & _
               " where find_in_set(aaa.ENSAYO_ID,bbb.c1) " & _
               " order by equipo_id "
    
    execute_bd consulta

   On Error GoTo 0
   Exit Function

informarEquiposRecepcion_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure informarEquiposRecepcion of Módulo de clase clsPlasma_recepcion"

End Function
Public Function informarControlSpecification(MUESTRA_ID As Long)
    Dim oPR As New clsPlasma_recepcion
   On Error GoTo informarControlSpecification_Error

    If oPR.Carga(MUESTRA_ID) Then
        Dim oPP As New clsPlasma_procesos
'        Dim oPF As New clsPlasma_ficha
'        Dim oPE As New clsPlasma_ensayos
        Dim bond As String
        Dim top As String
'        Dim texto As String
        oPP.Carga oPR.getPROCESO_ID
        bond = oPP.getBOND_SPECIFICATION
        top = oPP.getTOP_SPECIFICATION
'        ' BOND
'        oPF.Carga oPP.getBOND_COAT_FICHA_ID
'        ' BOND->Microestructura
'        If oPP.getBOND_MICROESTRUCTURA = 1 Then
'            If oPF.getMICROESTRUCTURA <> 0 Then
'                oPE.Carga oPF.getMICROESTRUCTURA
'                If oPE.getREQUIREMENT <> "" Then
'                    If texto <> "" Then texto = texto & " / "
'                    texto = texto & oPE.getREQUIREMENT
'                End If
'            End If
'        End If
'        ' BOND->Traccion
'        If oPP.getBOND_TRACCION = 1 Then
'            If oPF.getTRACCION <> 0 Then
'                oPE.Carga oPF.getTRACCION
'                If oPE.getREQUIREMENT <> "" Then
'                    If texto <> "" Then texto = texto & " / "
'                    texto = texto & oPE.getREQUIREMENT
'                End If
'            End If
'        End If
'        ' BOND->MacroDureza
'        If oPP.getBOND_MACRO_DUREZA = 1 Then
'            If oPF.getMACRO_DUREZA <> 0 Then
'                oPE.Carga oPF.getMACRO_DUREZA
'                If oPE.getREQUIREMENT <> "" Then
'                    If texto <> "" Then texto = texto & " / "
'                    texto = texto & oPE.getREQUIREMENT
'                End If
'            End If
'        End If
'        ' BOND->MicroDureza
'        If oPP.getBOND_MICRO_DUREZA = 1 Then
'            If oPF.getMICRO_DUREZA <> 0 Then
'                oPE.Carga oPF.getMICRO_DUREZA
'                If oPE.getREQUIREMENT <> "" Then
'                    If texto <> "" Then texto = texto & " / "
'                    texto = texto & oPE.getREQUIREMENT
'                End If
'            End If
'        End If
'        bond = texto
'        texto = ""
'        ' TOP
'        oPF.Carga oPP.getTOP_COAT_FICHA_ID
'        ' TOP->Microestructura
'        If oPP.getTOP_MICROESTRUCTURA = 1 Then
'            If oPF.getMICROESTRUCTURA <> 0 Then
'                oPE.Carga oPF.getMICROESTRUCTURA
'                If oPE.getREQUIREMENT <> "" Then
'                    If texto <> "" Then texto = texto & " / "
'                    texto = texto & oPE.getREQUIREMENT
'                End If
'            End If
'        End If
'        ' TOP->Traccion
'        If oPP.getTOP_TRACCION = 1 Then
'            If oPF.getTRACCION <> 0 Then
'                oPE.Carga oPF.getTRACCION
'                If oPE.getREQUIREMENT <> "" Then
'                    If texto <> "" Then texto = texto & " / "
'                    texto = texto & oPE.getREQUIREMENT
'                End If
'            End If
'        End If
'        ' TOP->MacroDureza
'        If oPP.getTOP_MACRO_DUREZA = 1 Then
'            If oPF.getMACRO_DUREZA <> 0 Then
'                oPE.Carga oPF.getMACRO_DUREZA
'                If oPE.getREQUIREMENT <> "" Then
'                    If texto <> "" Then texto = texto & " / "
'                    texto = texto & oPE.getREQUIREMENT
'                End If
'            End If
'        End If
'        ' TOP->MicroDureza
'        If oPP.getTOP_MICRO_DUREZA = 1 Then
'            If oPF.getMICRO_DUREZA <> 0 Then
'                oPE.Carga oPF.getMICRO_DUREZA
'                If oPE.getREQUIREMENT <> "" Then
'                    If texto <> "" Then texto = texto & " / "
'                    texto = texto & oPE.getREQUIREMENT
'                End If
'            End If
'        End If
'        top = texto
'        Set oPP = Nothing
'        Set oPF = Nothing
        Dim consulta As String
        consulta = "UPDATE PLASMA_RECEPCION " & _
                   "   SET BOND_CONTROL_SPECIFICATION = '" & bond & "'," & _
                   "       TOP_CONTROL_SPECIFICATION = '" & top & "'" & _
                   " WHERE MUESTRA_ID = " & MUESTRA_ID
        execute_bd consulta
    End If

   On Error GoTo 0
   Exit Function

informarControlSpecification_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure informarControlSpecification of Módulo de clase clsPlasma_recepcion"

End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSPLASMA_RECEPCION
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function

