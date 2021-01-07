VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEquipos_plantillas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSEQUIPOS_PLANTILLAS
'***************************************************************
Private ID_PLANTILLA As Long
Private TIPO_EQUIPO_ID As Long
Private EDICION As Long
Private FILE_NAME As String
Private FILE_TYPE As String
Private DESCRIPCION As String
Private vigor As Long
Private MOTIVO As String
Private FILE As String
Private USUARIO_ID As Long
Private TS As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSEQUIPOS_PLANTILLAS
'***************************************************************
Public Property Let setID_PLANTILLA(ByVal dato As Long)
        ID_PLANTILLA = dato
End Property
Public Property Let setTIPO_EQUIPO_ID(ByVal dato As Long)
        TIPO_EQUIPO_ID = dato
End Property
Public Property Let setEDICION(ByVal dato As Long)
        EDICION = dato
End Property
Public Property Let setFILE_NAME(ByVal dato As String)
        FILE_NAME = dato
End Property
Public Property Let setFILE_TYPE(ByVal dato As String)
        FILE_TYPE = dato
End Property
Public Property Let setDESCRIPCION(ByVal dato As String)
        DESCRIPCION = dato
End Property
Public Property Let setVIGOR(ByVal dato As Long)
        vigor = dato
End Property
Public Property Let setMOTIVO(ByVal dato As String)
        MOTIVO = dato
End Property
Public Property Let setFILE(ByVal dato As String)
        FILE = dato
End Property
Public Property Let setUSUARIO_ID(ByVal dato As Long)
        USUARIO_ID = dato
End Property
Public Property Let setTS(ByVal dato As String)
        TS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSEQUIPOS_PLANTILLAS
'***************************************************************
Public Property Get getID_PLANTILLA() As Long
        getID_PLANTILLA = ID_PLANTILLA
End Property
Public Property Get getTIPO_EQUIPO_ID() As Long
        getTIPO_EQUIPO_ID = TIPO_EQUIPO_ID
End Property
Public Property Get getEDICION() As Long
        getEDICION = EDICION
End Property
Public Property Get getFILE_NAME() As String
        getFILE_NAME = FILE_NAME
End Property
Public Property Get getFILE_TYPE() As String
        getFILE_TYPE = FILE_TYPE
End Property
Public Property Get getDESCRIPCION() As String
        getDESCRIPCION = DESCRIPCION
End Property
Public Property Get getVIGOR() As Long
        getVIGOR = vigor
End Property
Public Property Get getMOTIVO() As String
        getMOTIVO = MOTIVO
End Property
Public Property Get getFILE() As String
        getFILE = FILE
End Property
Public Property Get getUSUARIO_ID() As Long
        getUSUARIO_ID = USUARIO_ID
End Property
Public Property Get getTS() As String
        getTS = TS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSEQUIPOS_PLANTILLAS
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EQUIPOS_PLANTILLAS WHERE ID_PLANTILLA = " & ID
        Set rs = datos_bd_doc(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                ID_PLANTILLA = rs("ID_PLANTILLA")
                TIPO_EQUIPO_ID = rs("TIPO_EQUIPO_ID")
                EDICION = rs("EDICION")
                FILE_NAME = rs("FILE_NAME")
                FILE_TYPE = rs("FILE_TYPE")
                DESCRIPCION = rs("DESCRIPCION")
                vigor = rs("VIGOR")
                MOTIVO = rs("MOTIVO")
                FILE = rs("FILE")
                USUARIO_ID = rs("USUARIO_ID")
                TS = rs("TS")
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsEquipos_plantillas)", vbCritical, Err.Description
End Function
Public Function CargaVigorPorTipo(TIPO_ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PLANTILLA FROM EQUIPOS_PLANTILLAS WHERE TIPO_EQUIPO_ID = " & TIPO_ID & " AND VIGOR = 1"
        Set rs = datos_bd_doc(consulta)
        CargaVigorPorTipo = False
        If rs.RecordCount > 0 Then
            CargaVigorPorTipo = Me.Carga(rs(0))
        End If
        Set rs = Nothing
        Exit Function
fallo:
        CargaVigorPorTipo = False
        MsgBox "Error al cargar los datos(clsEquipos_plantillas)", vbCritical, Err.Description
End Function
Public Function GeneraCertificado(TIPO_ID As Long, fichero As String, Abrir As Boolean) As String
   On Error GoTo GeneraCertificado_Error
        
        Dim fichero_local As String
        If Me.CargaVigorPorTipo(TIPO_ID) Then
            Dim consulta As String
            Dim rs As ADODB.Recordset
            consulta = "SELECT FILE FROM equipos_plantillas where ID_PLANTILLA = " & Me.getID_PLANTILLA
            Set rs = datos_bd_doc(consulta)
            If rs.RecordCount > 0 Then
                Dim mystream As New ADODB.Stream
                mystream.Type = adTypeBinary
                mystream.Open
                mystream.Write rs(0)
                On Error Resume Next
                MkDir App.Path & "\" & RUTA_CERTIFICADOS & "\"
                fichero_local = App.Path & "\" & RUTA_CERTIFICADOS & "\" & fichero
                mystream.SaveToFile fichero_local, adSaveCreateOverWrite
                mystream.Close
                
                If Abrir Then
                    Dim iret As Long
                    iret = ShellExecute(0, vbNullString, fichero_local, vbNullString, App.Path, 1)
                End If
            Else
                MsgBox "No existe la plantilla asociada.", vbCritical, App.Title
            End If
        Else
            MsgBox "No existe la plantilla asociada a ese tipo de equipo.", vbCritical, App.Title
        End If
        GeneraCertificado = fichero_local

   On Error GoTo 0
   Exit Function

GeneraCertificado_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure GeneraCertificado of Módulo de clase clsEquipos_plantillas"
End Function
Public Function cumplimentarExcel(fichero_local As String, ID As Long) As Boolean
    Dim XLA As excel.Application
    Dim XLW As excel.Workbook
    Dim XLS As excel.Worksheet
    
   On Error GoTo cumplimentarExcel_Error

    Set XLA = New excel.Application
    Set XLW = XLA.Workbooks.Open(fichero_local)
    Set XLS = XLW.Worksheets(1)
    
    XLS.Unprotect PASS_XLSX
    
    Dim consulta As String
    Dim rs As ADODB.Recordset
    consulta = "select campo,col,fila from certificator_campos"
    
    Dim campos() As String
    Dim filas() As Integer
    Dim columnas() As Integer
    Dim i As Integer
    i = 0
    Set rs = datos_bd_metrologia(consulta)
    If rs.RecordCount > 0 Then
        ReDim campos(rs.RecordCount - 1)
        ReDim filas(rs.RecordCount - 1)
        ReDim columnas(rs.RecordCount - 1)
        Do
            campos(i) = rs(0)
            columnas(i) = rs(1)
            filas(i) = rs(2)
            i = i + 1
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    consulta = "select * from certificator_excel where id = " & ID
    Set rs = datos_bd_metrologia(consulta)
    If rs.RecordCount > 0 Then
        Do
            Dim pos As Integer
            Dim j As Integer
            For j = 0 To rs.Fields.Count - 1
                For i = LBound(campos) To UBound(campos)
                    If UCase(campos(i)) = UCase(rs.Fields(j).Name) Then
                        If UCase(campos(i)) = "CERTIFICADO" Then
                            XLS.Cells(filas(i), columnas(i)) = "C-" & rs.Fields(j) & "/" & Year(Date)
                        Else
                            XLS.Cells(filas(i), columnas(i)) = rs.Fields(j)
                        End If
                    End If
                Next
            Next
            rs.MoveNext
        Loop Until rs.EOF
    End If
    
    XLS.Protect PASS_XLSX
    
    XLW.Save
    XLA.visible = True
    Set XLS = Nothing
    Set XLW = Nothing
    Set XLA = Nothing

   On Error GoTo 0
   Exit Function

cumplimentarExcel_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cumplimentarExcel of Módulo de clase clsEquipos_plantillas"
End Function

Public Function CrearID()
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT MAX(ID_PLANTILLA) FROM EQUIPOS_PLANTILLAS"
        Set rs = datos_bd_doc(consulta)
        If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then
                ID_PLANTILLA = 1
        Else
                ID_PLANTILLA = rs.Fields(0) + 1
        End If
        Set rs = Nothing
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        Me.CrearID
        consulta = "INSERT INTO EQUIPOS_PLANTILLAS " & _
                   "  VALUES (" & _
                        ID_PLANTILLA & "," & TIPO_EQUIPO_ID & "," & EDICION & "," & _
                        "'" & FILE_NAME & "'" & "," & "'" & FILE_TYPE & "'" & "," & "'" & DESCRIPCION & "'" & "," & _
                        vigor & "," & "'" & MOTIVO & "'" & "," & "'" & FILE & "'" & "," & _
                        USUARIO_ID & "," & "'" & TS & "'" & _
                ")"
        execute_bd_doc consulta
        Insertar = ID_PLANTILLA
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsEquipos_plantillas)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE EQUIPOS_PLANTILLAS SET " & _
                        " TIPO_EQUIPO_ID = " & TIPO_EQUIPO_ID & "," & _
                        " EDICION = " & EDICION & "," & _
                        " FILE_NAME = '" & FILE_NAME & "'," & _
                        " FILE_TYPE = '" & FILE_TYPE & "'," & _
                        " DESCRIPCION = '" & DESCRIPCION & "'," & _
                        " VIGOR = " & vigor & "," & _
                        " MOTIVO = '" & MOTIVO & "'," & _
                        " FILE = '" & FILE & "'," & _
                        " USUARIO_ID = " & USUARIO_ID & "," & _
                        " TS = '" & TS & "'" & _
                " WHERE ID_PLANTILLA = " & ID
        execute_bd_doc consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsEquipos_plantillas)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM EQUIPOS_PLANTILLAS " & _
                "    WHERE ID_PLANTILLA = " & ID
        execute_bd_doc consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsEquipos_plantillas)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM EQUIPOS_PLANTILLAS ORDER BY ID_PLANTILLA"
        Set Listado = datos_bd_doc(consulta)
End Function

'***************************************************************
'*   TABLAS HEREDERAS DADA UNA TABLA DE LA CLASE CLSEQUIPOS_PLANTILLAS
'***************************************************************
Public Function ListadoHerencia() As ADODB.Recordset
        Dim consulta As String
                consulta = "SELECT RC.TABLE_NAME as HERENCIA, RC.CONSTRAINT_NAME as CLAVE FROM INFORMATION_SCHEMA.REFERENTIAL_CONSTRAINTS RC INNER JOIN INFORMATION_SCHEMA.KEY_COLUMN_USAGE  KCU ON RC.CONSTRAINT_NAME = KCU.CONSTRAINT_NAME;"
        Set ListadoHerencia = datos_bd_doc(consulta)
End Function

Public Function Listado_Combo() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT ID_PLANTILLA,TIPO_EQUIPO_ID FROM EQUIPOS_PLANTILLAS ORDER BY TIPO_EQUIPO_ID"
        Set Listado_Combo = datos_bd_doc(consulta)
End Function


