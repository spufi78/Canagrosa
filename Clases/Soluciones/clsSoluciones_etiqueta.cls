VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSoluciones_etiqueta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsSoluciones_etiqueta
'***************************************************************
Private ID_ETIQUETA As Long
Private DESCRIPCION As String
Private SUBTITULO As String
Private FRASES_PEQ As String
Private FRASES_MED As String
Private FRASES_GRA As String
Private COMPONENTES As String
Private ADVERTENCIA As String
Private PIC1 As Long
Private PIC2 As Long
Private PIC3 As Long
Private PIC4 As Long
Private PIC5 As Long
Private PIC6 As Long
Private PIC7 As Long
Private PIC8 As Long
Private PIC9 As Long

'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsSoluciones_etiqueta
'***************************************************************
Public Property Let setID_ETIQUETA(ByVal dato As Long)
        ID_ETIQUETA = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setSUBTITULO(ByVal dato As String)
        SUBTITULO = dato
End Property
Public Property Let setFRASES_PEQ(ByVal dato As String)
        FRASES_PEQ = dato
End Property
Public Property Let setFRASES_MED(ByVal dato As String)
        FRASES_MED = dato
End Property
Public Property Let setFRASES_GRA(ByVal dato As String)
        FRASES_GRA = dato
End Property
Public Property Let setCOMPONENTES(ByVal dato As String)
        COMPONENTES = dato
End Property
Public Property Let setADVERTENCIA(ByVal dato As String)
        ADVERTENCIA = dato
End Property
Public Property Let setPIC1(ByVal dato As Long)
        PIC1 = dato
End Property
Public Property Let setPIC2(ByVal dato As Long)
        PIC2 = dato
End Property
Public Property Let setPIC3(ByVal dato As Long)
        PIC3 = dato
End Property
Public Property Let setPIC4(ByVal dato As Long)
        PIC4 = dato
End Property
Public Property Let setPIC5(ByVal dato As Long)
        PIC5 = dato
End Property
Public Property Let setPIC6(ByVal dato As Long)
        PIC6 = dato
End Property
Public Property Let setPIC7(ByVal dato As Long)
        PIC7 = dato
End Property
Public Property Let setPIC8(ByVal dato As Long)
        PIC8 = dato
End Property
Public Property Let setPIC9(ByVal dato As Long)
        PIC9 = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsSoluciones_etiqueta
'***************************************************************
Public Property Get getID_ETIQUETA() As Long
        getID_ETIQUETA = ID_ETIQUETA
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getSUBTITULO() As String
        getSUBTITULO = SUBTITULO
End Property
Public Property Get getFRASES_PEQ() As String
        getFRASES_PEQ = FRASES_PEQ
End Property
Public Property Get getFRASES_MED() As String
        getFRASES_MED = FRASES_MED
End Property
Public Property Get getFRASES_GRA() As String
        getFRASES_GRA = FRASES_GRA
End Property
Public Property Get getCOMPONENTES() As String
        getCOMPONENTES = COMPONENTES
End Property
Public Property Get getADVERTENCIA() As String
        getADVERTENCIA = ADVERTENCIA
End Property
Public Property Get getPIC1() As Long
        getPIC1 = PIC1
End Property
Public Property Get getPIC2() As Long
        getPIC2 = PIC2
End Property
Public Property Get getPIC3() As Long
        getPIC3 = PIC3
End Property
Public Property Get getPIC4() As Long
        getPIC4 = PIC4
End Property
Public Property Get getPIC5() As Long
        getPIC5 = PIC5
End Property
Public Property Get getPIC6() As Long
        getPIC6 = PIC6
End Property
Public Property Get getPIC7() As Long
        getPIC7 = PIC7
End Property
Public Property Get getPIC8() As Long
        getPIC8 = PIC8
End Property
Public Property Get getPIC9() As Long
        getPIC9 = PIC9
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE clsSoluciones_etiqueta
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SOLUCIONES_ETIQUETA WHERE ID_ETIQUETA = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_ETIQUETA = rs("ID_ETIQUETA")
                DESCRIPCION = rs("DESCRIPCION")
                SUBTITULO = rs("SUBTITULO")
                FRASES_PEQ = rs("FRASES_PEQ")
                FRASES_MED = rs("FRASES_MED")
                FRASES_GRA = rs("FRASES_GRA")
                COMPONENTES = rs("COMPONENTES")
                ADVERTENCIA = rs("ADVERTENCIA")
                PIC1 = rs("PIC1")
                PIC2 = rs("PIC2")
                PIC3 = rs("PIC3")
                PIC4 = rs("PIC4")
                PIC5 = rs("PIC5")
                PIC6 = rs("PIC6")
                PIC7 = rs("PIC7")
                PIC8 = rs("PIC8")
                PIC9 = rs("PIC9")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsSoluciones_etiqueta)", vbCritical, Err.Description
End Function
Public Sub CrearID()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_ETIQUETA) FROM SOLUCIONES_ETIQUETA"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_ETIQUETA = 1
    Else
        ID_ETIQUETA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO SOLUCIONES_ETIQUETA " & _
                   "  VALUES (" & _
                        ID_ETIQUETA & ",'" & DESCRIPCION & "','" & SUBTITULO & "','" & FRASES_PEQ & "'" & "," & "'" & FRASES_MED & "'" & "," & _
                        "'" & FRASES_GRA & "','" & COMPONENTES & "','" & ADVERTENCIA & "'," & PIC1 & "," & PIC2 & "," & _
                        PIC3 & "," & PIC4 & "," & PIC5 & "," & PIC6 & "," & PIC7 & "," & PIC8 & "," & PIC9 & _
                ")"
        execute_bd consulta
        Insertar = ID_ETIQUETA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsSoluciones_etiqueta)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE SOLUCIONES_ETIQUETA SET " & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " SUBTITULO = '" & SUBTITULO & "'," & _
                        " FRASES_PEQ = '" & FRASES_PEQ & "'," & _
                        " FRASES_MED = '" & FRASES_MED & "'," & _
                        " FRASES_GRA = '" & FRASES_GRA & "'," & _
                        " COMPONENTES = '" & COMPONENTES & "'," & _
                        " ADVERTENCIA = '" & ADVERTENCIA & "'," & _
                        " PIC1 = " & PIC1 & "," & _
                        " PIC2 = " & PIC2 & "," & _
                        " PIC3 = " & PIC3 & "," & _
                        " PIC4 = " & PIC4 & "," & _
                        " PIC5 = " & PIC5 & "," & _
                        " PIC6 = " & PIC6 & "," & _
                        " PIC7 = " & PIC7 & "," & _
                        " PIC8 = " & PIC8 & "," & _
                        " PIC9 = " & PIC9 & "" & _
                " WHERE ID_ETIQUETA = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsSoluciones_etiqueta)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM SOLUCIONES_ETIQUETA " & _
                "    WHERE ID_ETIQUETA = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsSoluciones_etiqueta)", vbCritical, Err.Description
End Function
Public Function Listado(DESCRIPCION As String) As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM SOLUCIONES_ETIQUETA WHERE DESCRIPCION LIKE '%" & DESCRIPCION & "%' ORDER BY DESCRIPCION"
        Set Listado = datos_bd(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE clsSoluciones_etiqueta
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd(consulta)
End Function
Public Function llenar_combo(conn As ADODB.Connection, combo As miCombo, FK As Long, FORMULARIO As Form, filtro As String)
  
    Dim consulta As String
    consulta = "SELECT ID_ETIQUETA, CONCAT(DESCRIPCION, '  (',SUBTITULO,')') " & _
               "  FROM soluciones_etiqueta " & _
               filtro
    With combo
        .setCONN = conn
        .setFK_CAMPO = ""
        .setFK_VALOR = FK
        .setTABLA = "SOLUCIONES_ETIQUETA"
        .setDESCRIPCION = "Etiquetas"
        .setPK = "ID_ETIQUETA"
        .setCAMPO = " CONCAT(DESCRIPCION, '  (',SUBTITULO,')')"
        .setQUERY = consulta
        .setMUESTRA_DETALLE = False
        Set .FORMULARIO = FORMULARIO
    End With
    
End Function
Public Function ImprimirEtiquetas(PK As Long, ETIQUETAS As Long) As Boolean
    Dim rs As ADODB.Recordset
    Dim cont As Long
    Dim datos As String ' donde guarda los xml de los tres tamaños
    
    Dim oMuestra As New clsMuestra
    oMuestra.CargaMuestra PK

    Dim oMe As New clsMuestras_soluciones
    If oMe.Carga(PK) = False Then Exit Function
    If oMe.getTAMANO <> 2 Then
        frmPosicionPegatina.Show 1
    Else
        pegatina = 1
    End If
    For cont = 1 To pegatina - 1
        datos = datos & "<etiqueta>"
        datos = datos & "<PRODUCTO></PRODUCTO>"
        datos = datos & "<SUBTITULO></SUBTITULO>"
        datos = datos & "<FECHA_FABRICACION></FECHA_FABRICACION>"
        datos = datos & "<FECHA_CADUCIDAD></FECHA_CADUCIDAD>"
        datos = datos & "<SP></SP>"
        datos = datos & "<CENTRO_ID></CENTRO_ID>"
        datos = datos & "<MOSTRAR>0</MOSTRAR>"
        datos = datos & "<FRASES_PEQ></FRASES_PEQ>"
        datos = datos & "<FRASES_MED></FRASES_MED>"
        datos = datos & "<FRASES_GRA></FRASES_GRA>"
        datos = datos & "<COMPONENTES></COMPONENTES>"
        datos = datos & "<ADVERTENCIA></ADVERTENCIA>"
        datos = datos & "<PIC1></PIC1>"
        datos = datos & "<PIC2></PIC2>"
        datos = datos & "<PIC3></PIC3>"
        datos = datos & "<PIC4></PIC4>"
        datos = datos & "<PIC5></PIC5>"
        datos = datos & "</etiqueta>"
    Next cont

    Dim oSe As New clsSoluciones_etiqueta
    oSe.Carga oMe.getETIQUETA_ID

    ' Pictogramas
    Dim ubicacionPic As String
    ubicacionPic = ReadINI(App.Path & "\config.ini", "Otros", "Pictogramas")
    Dim pos As Integer
    Dim pictogramas(1 To 5) As String
    pos = 1
    If oSe.getPIC1 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC1.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC2 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC2.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC3 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC3.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC4 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC4.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC5 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC5.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC6 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC6.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC7 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC7.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC8 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC8.jpg"
        pos = pos + 1
    End If
    If oSe.getPIC9 = 1 And pos <= 5 Then
        pictogramas(pos) = ubicacionPic & "\" & "PIC9.jpg"
        pos = pos + 1
    End If


    For cont = 1 To ETIQUETAS
        datos = datos & "<etiqueta>"
        datos = datos & "<PRODUCTO>" & oSe.getDESCRIPCION & "</PRODUCTO>"
        datos = datos & "<SUBTITULO>" & oSe.getSUBTITULO & "</SUBTITULO>"
        datos = datos & "<FECHA_FABRICACION>" & Format(oMe.getFECHA_FABRICACION, "dd/mm/yyyy") & "</FECHA_FABRICACION>"
        datos = datos & "<FECHA_CADUCIDAD>" & Format(oMe.getFECHA_CADUCIDAD, "dd/mm/yyyy") & "</FECHA_CADUCIDAD>"
        datos = datos & "<SP>" & oMe.getSP & "</SP>"
        datos = datos & "<CENTRO_ID>" & oMuestra.getCENTRO_ID & "</CENTRO_ID>"
        datos = datos & "<MOSTRAR>1</MOSTRAR>"
        datos = datos & "<FRASES_PEQ>" & formatearFrase(oSe.getFRASES_PEQ) & "</FRASES_PEQ>"
        datos = datos & "<FRASES_MED>" & formatearFrase(oSe.getFRASES_MED) & "</FRASES_MED>"
        datos = datos & "<FRASES_GRA>" & formatearFrase(oSe.getFRASES_GRA) & "</FRASES_GRA>"
        datos = datos & "<COMPONENTES>" & formatearFrase(oSe.getCOMPONENTES) & "</COMPONENTES>"
        datos = datos & "<ADVERTENCIA>" & formatearFrase(oSe.getADVERTENCIA) & "</ADVERTENCIA>"
        
        Dim i As Integer
        For i = 1 To 5
            datos = datos & "<PIC" & i & ">" & pictogramas(i) & "</PIC" & i & ">"
        Next
        
        datos = datos & "</etiqueta>"
    Next
    
    If Trim(datos) <> "" Then
       Dim tx As TextStream
        Set tx = gFSO.CreateTextFile((ReadINI(App.Path & "\config.ini", "Documentos", "Reportes") & "\soluciones\datos.xml"), True)
        tx.WriteLine Replace("<?xml version='1.0' encoding='ISO-8859-1'?>", "'", Chr(34))
        tx.WriteLine "<lista_etiquetas>"
        tx.WriteLine datos
        tx.WriteLine "</lista_etiquetas>"
        tx.Close

        ' AHORA MUESTRA EL INFORME
        Dim objRpt As New frmReport

        With objRpt
            .iniciar
            .informe = "Soluciones\rptEtiquetasSoluciones_t" & oMe.getTAMANO + 1
            '.CRITERIO = ""
            .imprimir = False
            .generar
            '.Visible = True
            .Show 1
        End With

        Unload objRpt
        Set objRpt = Nothing
        Set tx = Nothing

    End If


    Set rs = Nothing


End Function
Public Function formatearFrase(FRASE As String) As String
    Dim s As String
    s = FRASE
    s = Replace(s, "<negrita>", "&lt;b&gt;")
    s = Replace(s, "</negrita>", "&lt;/b&gt;")
    s = Replace(s, vbCrLf, "&lt;br&gt;")
    formatearFrase = Trim(s)
End Function
