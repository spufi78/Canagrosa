VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlveograma_valores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ALVEOGRAMA_ID As Long
Private TENACIDAD As String
Private EXTENSIBILIDAD As String
Private s As String
Private w As String
Private g As String
Private DE_REPOSO As Integer
Private CURVA_1 As String
Private CURVA_2 As String
Private CURVA_3 As String
Private CURVA_MEDIA As String

Public Property Let setALVEOGRAMA_ID(ByVal dato As Long)
    ALVEOGRAMA_ID = dato
End Property
Public Property Let setTENACIDAD(ByVal dato As String)
    TENACIDAD = dato
End Property
Public Property Let setEXTENSIBILIDAD(ByVal dato As String)
    EXTENSIBILIDAD = dato
End Property
Public Property Let setS(ByVal dato As String)
    s = dato
End Property
Public Property Let setW(ByVal dato As String)
    w = dato
End Property
Public Property Let setG(ByVal dato As String)
    g = Trim(dato)
End Property
Public Property Let setDE_REPOSO(ByVal dato As Integer)
    DE_REPOSO = dato
End Property
Public Property Let setCURVA_1(ByVal dato As String)
    CURVA_1 = Trim(dato)
End Property
Public Property Let setCURVA_2(ByVal dato As String)
    CURVA_2 = Trim(dato)
End Property
Public Property Let setCURVA_3(ByVal dato As String)
    CURVA_3 = Trim(dato)
End Property
Public Property Let setCURVA_MEDIA(ByVal dato As String)
    CURVA_MEDIA = Trim(dato)
End Property
Public Property Get getALVEOGRAMA_ID() As Long
    getALVEOGRAMA_ID = ALVEOGRAMA_ID
End Property
Public Property Get getTENACIDAD() As String
    getTENACIDAD = TENACIDAD
End Property
Public Property Get getEXTENSIBILIDAD() As String
    getEXTENSIBILIDAD = EXTENSIBILIDAD
End Property
Public Property Get getS() As String
    getS = s
End Property
Public Property Get getW() As String
    getW = w
End Property
Public Property Get getG() As String
    getG = g
End Property
Public Property Get getDE_REPOSO() As Integer
    getDE_REPOSO = DE_REPOSO
End Property
Public Property Get getCURVA_1() As String
    getCURVA_1 = CURVA_1
End Property
Public Property Get getCURVA_2() As String
    getCURVA_2 = CURVA_2
End Property
Public Property Get getCURVA_3() As String
    getCURVA_3 = CURVA_3
End Property
Public Property Get getCURVA_MEDIA() As String
    getCURVA_MEDIA = CURVA_MEDIA
End Property
Public Function InsertarAlveogramaValores() As Long
    On Error GoTo fallo
    Dim consulta As String
    consulta = "Insert into alveograma_valores " & _
               " values(" & _
               ALVEOGRAMA_ID & ",'" & _
               TENACIDAD & "','" & EXTENSIBILIDAD & "','" & s & "','" & w & "','" & _
               g & "'," & DE_REPOSO & ",'" & CURVA_1 & "','" & CURVA_2 & "','" & _
               CURVA_3 & "','" & CURVA_MEDIA & "')"
    execute_bd consulta
    InsertarAlveogramaValores = ALVEOGRAMA_ID
    Exit Function
fallo:
    InsertarAlveogramaValores = 0
    MsgBox "Error al insertar el Alveograma_Valores (InsertarAlveogramaValores)", vbCritical, Err.Description
End Function
Public Function CargarAlveogramaValores(ID As Long, reposo As Integer) As Boolean
    Dim consulta As String
    Dim rs As New ADODB.RecordSet
    consulta = "SELECT * FROM ALVEOGRAMA_valores WHERE ALVEOGRAMA_ID=" & ID & " AND DE_REPOSO = " & reposo
    Set rs = datos_bd(consulta)
    If rs.RecordCount = 0 Then
         CargarAlveogramaValores = False
    Else
        ALVEOGRAMA_ID = rs.Fields("ALVEOGRAMA_ID")
        If Not IsNull(rs.Fields("TENACIDAD")) Then
            TENACIDAD = rs.Fields("TENACIDAD")
        Else
            TENACIDAD = 0
        End If
        If Not IsNull(rs.Fields("EXTENSIBILIDAD")) Then
            EXTENSIBILIDAD = rs.Fields("EXTENSIBILIDAD")
        Else
            EXTENSIBILIDAD = 0
        End If
        If Not IsNull(rs.Fields("S")) Then
            s = rs.Fields("S")
        Else
            s = 0
        End If
        If Not IsNull(rs.Fields("W")) Then
            w = rs.Fields("W")
        Else
            w = 0
        End If
        If Not IsNull(rs.Fields("G")) Then
            g = rs.Fields("G")
        Else
            g = 0
        End If
        DE_REPOSO = rs.Fields("DE_REPOSO")
        CURVA_1 = rs.Fields("CURVA_1")
        CURVA_2 = rs.Fields("CURVA_2")
        CURVA_3 = rs.Fields("CURVA_3")
        CURVA_MEDIA = rs.Fields("CURVA_MEDIA")
    End If
    Set rs = Nothing
    CargarAlveogramaValores = True
    Exit Function
fallo:
    CargarAlveogramaValores = False
    MsgBox "Error al cargar el Alveograma_Valores (CargarAlveogramaValores)", vbCritical, Err.Description
End Function
Public Sub ModificarAlveogramaValores(ID As Long, reposo As Integer)
    On Error GoTo fallo
    Dim consulta As String
    consulta = "select * from alveograma_valores where alveograma_id = " & ID & " and de_reposo = " & reposo
    Dim rs As ADODB.RecordSet
    Set rs = datos_bd(consulta)
    If rs.RecordCount > 0 Then
        consulta = "update alveograma_valores set " & _
                   " TENACIDAD = " & TENACIDAD & "," & _
                   " EXTENSIBILIDAD = " & EXTENSIBILIDAD & "," & _
                   " S = " & s & "," & _
                   " W = " & w & "," & _
                   " G = " & g & "," & _
                   " CURVA_1 = '" & CURVA_1 & "'," & _
                   " CURVA_2 = '" & CURVA_2 & "'," & _
                   " CURVA_3 = '" & CURVA_2 & "'," & _
                   " CURVA_MEDIA = '" & CURVA_MEDIA & "'" & _
                   " WHERE alveograma_id = " & ID & _
                   "   and de_reposo = " & reposo
        execute_bd consulta
    Else
        Me.setALVEOGRAMA_ID = ID
        Me.setDE_REPOSO = reposo
        Me.InsertarAlveogramaValores
    End If
    Exit Sub
fallo:
    MsgBox "Error al modificar los Alveograma_Valores", vbCritical, Err.Description
End Sub


