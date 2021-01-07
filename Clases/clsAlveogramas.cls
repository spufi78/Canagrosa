VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlveogramas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ID_ALVEOGRAMA As Long
Private MUESTRA_ID As Long
Private DETERMINACION_ID As Long
Private TEMPERATURA As String
Private HUMEDAD_AIRE As String
Private HUMEDAD_HARINA As String
Private INDICE_DEGRADACION As String
Private fecha As String
Public Property Let setID_ALVEOGRAMA(ByVal dato As Long)
    ID_ALVEOGRAMA = dato
End Property
Public Property Let setMUESTRA_ID(ByVal dato As Long)
    MUESTRA_ID = dato
End Property
Public Property Let setDETERMINACION_ID(ByVal dato As Long)
    DETERMINACION_ID = dato
End Property
Public Property Let setTEMPERATURA(ByVal dato As String)
    TEMPERATURA = dato
End Property
Public Property Let setHUMEDAD_AIRE(ByVal dato As String)
    HUMEDAD_AIRE = dato
End Property
Public Property Let setHUMEDAD_HARINA(ByVal dato As String)
    HUMEDAD_HARINA = dato
End Property
Public Property Let setINDICE_DEGRADACION(ByVal dato As String)
    INDICE_DEGRADACION = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
    fecha = Trim(dato)
End Property
Public Property Get getID_ALVEOGRAMA() As Long
    getID_ALVEOGRAMA = ID_ALVEOGRAMA
End Property
Public Property Get getMUESTRA_ID() As Long
    getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getDETERMINACION_ID() As Long
    getDETERMINACION_ID = DETERMINACION_ID
End Property
Public Property Get getTEMPERATURA() As String
    getTEMPERATURA = TEMPERATURA
End Property
Public Property Get getHUMEDAD_AIRE() As String
    getHUMEDAD_AIRE = HUMEDAD_AIRE
End Property
Public Property Get getHUMEDAD_HARINA() As String
    getHUMEDAD_HARINA = HUMEDAD_HARINA
End Property
Public Property Get getINDICE_DEGRADACION() As String
    getINDICE_DEGRADACION = INDICE_DEGRADACION
End Property
Public Property Get getFECHA() As String
    getFECHA = fecha
End Property
Public Function InsertarAlveograma() As Long
    On Error GoTo fallo
    Dim consulta As String
    Me.CrearID_ALVEOGRAMA
    consulta = "Insert into alveogramas " & _
               " values(" & _
               ID_ALVEOGRAMA & "," & MUESTRA_ID & "," & DETERMINACION_ID & ",'" & _
               TEMPERATURA & "','" & HUMEDAD_AIRE & "','" & HUMEDAD_HARINA & "','" & INDICE_DEGRADACION & "','" & _
               fecha & "')"
    execute_bd consulta
    InsertarAlveograma = ID_ALVEOGRAMA
    Exit Function
fallo:
    InsertarAlveograma = 0
    MsgBox "Error al insertar el Alveograma (InsertarAlveograma)", vbCritical, Err.Description
End Function

Public Sub CrearID_ALVEOGRAMA()
    Dim rs As New ADODB.Recordset
    Dim consulta As String
    consulta = "SELECT MAX(ID_ALVEOGRAMA) FROM ALVEOGRAMAS"
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ID_ALVEOGRAMA = 1
    Else
        ID_ALVEOGRAMA = rs.Fields(0) + 1
    End If
    Set rs = Nothing
End Sub
Public Function CargarAlveograma(id As Long) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT * FROM ALVEOGRAMAS WHERE ID_ALVEOGRAMA=" & id
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
         CargarAlveograma = False
    Else
        ID_ALVEOGRAMA = rs.Fields("ID_ALVEOGRAMA")
        MUESTRA_ID = rs.Fields("muestra_id")
        DETERMINACION_ID = rs.Fields("DETERMINACION_ID")
        If Not IsNull(rs.Fields("TEMPERATURA")) Then
            TEMPERATURA = rs.Fields("TEMPERATURA")
        Else
            TEMPERATURA = 0
        End If
        If Not IsNull(rs.Fields("HUMEDAD_AIRE")) Then
            HUMEDAD_AIRE = rs.Fields("HUMEDAD_AIRE")
        Else
            HUMEDAD_AIRE = 0
        End If
        If Not IsNull(rs.Fields("HUMEDAD_HARINA")) Then
            HUMEDAD_HARINA = rs.Fields("HUMEDAD_HARINA")
        Else
            HUMEDAD_HARINA = 0
        End If
        If Not IsNull(rs.Fields("INDICE_DEGRADACION")) Then
            INDICE_DEGRADACION = rs.Fields("INDICE_DEGRADACION")
        Else
            INDICE_DEGRADACION = 0
        End If
        If Not IsNull(rs.Fields("FECHA")) Then
            fecha = rs.Fields("FECHA")
        Else
            fecha = 0
        End If
    End If
    Set rs = Nothing
    CargarAlveograma = True
    Exit Function
fallo:
    CargarAlveograma = False
    MsgBox "Error al cargar el Alveograma (CargarAlveograma)", vbCritical, Err.Description
End Function
Public Function ComprobarAlveograma(muestra As Long, deter As Long) As Long
    Dim consulta As String
    Dim rs As New ADODB.Recordset
    consulta = "SELECT ID_ALVEOGRAMA FROM ALVEOGRAMAS WHERE MUESTRA_ID=" & muestra & " AND DETERMINACION_ID=" & deter
    Set rs = datos_bd(consulta)
    If IsNull(rs.Fields(0)) Or (rs.EOF And rs.BOF) Then  'si es nulo No se recupero ninguno
        ComprobarAlveograma = 0
    Else
        ComprobarAlveograma = rs("ID_ALVEOGRAMA")
    End If
    Set rs = Nothing
    Exit Function
fallo:
    ComprobarAlveograma = 0
    MsgBox "Error al comprobar el alveograma (ComprobarAlveograma)", vbCritical, Err.Description
End Function

Public Function ModificarAlveograma(id As Long)
    On Error GoTo fallo
    Dim consulta As String
    consulta = "update alveogramas set " & _
               " TEMPERATURA = " & TEMPERATURA & "," & _
               " HUMEDAD_AIRE = " & HUMEDAD_AIRE & "," & _
               " HUMEDAD_HARINA = " & HUMEDAD_HARINA & "," & _
               " INDICE_DEGRADACION = '" & INDICE_DEGRADACION & "'," & _
               " FECHA = '" & Format(Date, "yyyy-mm-dd") & "'" & _
               " WHERE ID_ALVEOGRAMA = " & id
    execute_bd consulta
    Exit Function
fallo:
    MsgBox "Error al modificar el Alveograma (ModificarAlveograma)", vbCritical, Err.Description
End Function


