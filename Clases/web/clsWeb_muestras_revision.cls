VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsWeb_muestras_revision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE CLSWEB_MUESTRAS_REVISION
'***************************************************************
Private MUESTRA_ID As Long
Private CONFORME As Long
Private CONFORME_TEXTO As String
Private CONFORME_USUARIO_ID As Long
Private CONFORME_FS As String
Private RECARGA As Long
Private RECARGA_TEXTO As String
Private RECARGA_USUARIO_ID As Long
Private RECARGA_FS As String
Private PARADA As Long
Private PARADA_TEXTO As String
Private PARADA_USUARIO_ID As Long
Private PARADA_FS As String
Private REV_RECARGA As Long
Private REV_RECARGA_TEXTO As String
Private REV_RECARGA_USUARIO_ID As Long
Private REV_RECARGA_FS As String
'***************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE CLSWEB_MUESTRAS_REVISION
'***************************************************************
Public Property Let setMUESTRA_ID(ByVal dato As Long)
        MUESTRA_ID = dato
End Property
Public Property Let setCONFORME(ByVal dato As Long)
        CONFORME = dato
End Property
Public Property Let setCONFORME_TEXTO(ByVal dato As String)
        CONFORME_TEXTO = dato
End Property
Public Property Let setCONFORME_USUARIO_ID(ByVal dato As Long)
        CONFORME_USUARIO_ID = dato
End Property
Public Property Let setCONFORME_FS(ByVal dato As String)
        CONFORME_FS = dato
End Property
Public Property Let setRECARGA(ByVal dato As Long)
        RECARGA = dato
End Property
Public Property Let setRECARGA_TEXTO(ByVal dato As String)
        RECARGA_TEXTO = dato
End Property
Public Property Let setRECARGA_USUARIO_ID(ByVal dato As Long)
        RECARGA_USUARIO_ID = dato
End Property
Public Property Let setRECARGA_FS(ByVal dato As String)
        RECARGA_FS = dato
End Property
Public Property Let setPARADA(ByVal dato As Long)
        PARADA = dato
End Property
Public Property Let setPARADA_TEXTO(ByVal dato As String)
        PARADA_TEXTO = dato
End Property
Public Property Let setPARADA_USUARIO_ID(ByVal dato As Long)
        PARADA_USUARIO_ID = dato
End Property
Public Property Let setPARADA_FS(ByVal dato As String)
        PARADA_FS = dato
End Property
Public Property Let setREV_RECARGA(ByVal dato As Long)
        REV_RECARGA = dato
End Property
Public Property Let setREV_RECARGA_TEXTO(ByVal dato As String)
        REV_RECARGA_TEXTO = dato
End Property
Public Property Let setREV_RECARGA_USUARIO_ID(ByVal dato As Long)
        REV_RECARGA_USUARIO_ID = dato
End Property
Public Property Let setREV_RECARGA_FS(ByVal dato As String)
        REV_RECARGA_FS = dato
End Property
'***************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE CLSWEB_MUESTRAS_REVISION
'***************************************************************
Public Property Get getMUESTRA_ID() As Long
        getMUESTRA_ID = MUESTRA_ID
End Property
Public Property Get getCONFORME() As Long
        getCONFORME = CONFORME
End Property
Public Property Get getCONFORME_TEXTO() As String
        getCONFORME_TEXTO = CONFORME_TEXTO
End Property
Public Property Get getCONFORME_USUARIO_ID() As Long
        getCONFORME_USUARIO_ID = CONFORME_USUARIO_ID
End Property
Public Property Get getCONFORME_FS() As String
        getCONFORME_FS = CONFORME_FS
End Property
Public Property Get getRECARGA() As Long
        getRECARGA = RECARGA
End Property
Public Property Get getRECARGA_TEXTO() As String
        getRECARGA_TEXTO = RECARGA_TEXTO
End Property
Public Property Get getRECARGA_USUARIO_ID() As Long
        getRECARGA_USUARIO_ID = RECARGA_USUARIO_ID
End Property
Public Property Get getRECARGA_FS() As String
        getRECARGA_FS = RECARGA_FS
End Property
Public Property Get getPARADA() As Long
        getPARADA = PARADA
End Property
Public Property Get getPARADA_TEXTO() As String
        getPARADA_TEXTO = PARADA_TEXTO
End Property
Public Property Get getPARADA_USUARIO_ID() As Long
        getPARADA_USUARIO_ID = PARADA_USUARIO_ID
End Property
Public Property Get getPARADA_FS() As String
        getPARADA_FS = PARADA_FS
End Property
Public Property Get getREV_RECARGA() As Long
        getREV_RECARGA = REV_RECARGA
End Property
Public Property Get getREV_RECARGA_TEXTO() As String
        getREV_RECARGA_TEXTO = REV_RECARGA_TEXTO
End Property
Public Property Get getREV_RECARGA_USUARIO_ID() As Long
        getREV_RECARGA_USUARIO_ID = REV_RECARGA_USUARIO_ID
End Property
Public Property Get getREV_RECARGA_FS() As String
        getREV_RECARGA_FS = REV_RECARGA_FS
End Property
'***************************************************************
'*   PROCEDIMIENTOS Y FUNCIONES DE LA CLASE CLSWEB_MUESTRAS_REVISION
'***************************************************************
Public Function Carga(ID As Long) As Boolean
        On Error GoTo fallo
        Dim rs As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM WEB_MUESTRAS_REVISION WHERE MUESTRA_ID = " & ID
        Set rs = datos_bd(consulta)
        If rs.RecordCount = 0 Then
                Carga = False
        Else
                MUESTRA_ID = rs("MUESTRA_ID")
                CONFORME = rs("CONFORME")
                CONFORME_TEXTO = rs("CONFORME_TEXTO")
                CONFORME_USUARIO_ID = rs("CONFORME_USUARIO_ID")
                CONFORME_FS = rs("CONFORME_FS")
                RECARGA = rs("RECARGA")
                RECARGA_TEXTO = rs("RECARGA_TEXTO")
                RECARGA_USUARIO_ID = rs("RECARGA_USUARIO_ID")
                RECARGA_FS = rs("RECARGA_FS")
                PARADA = rs("PARADA")
                PARADA_TEXTO = rs("PARADA_TEXTO")
                PARADA_USUARIO_ID = rs("PARADA_USUARIO_ID")
                PARADA_FS = rs("PARADA_FS")
                REV_RECARGA = rs("REV_RECARGA")
                If Not IsNull(rs("REV_RECARGA_TEXTO")) Then
                    REV_RECARGA_TEXTO = rs("REV_RECARGA_TEXTO")
                End If
                If Not IsNull(rs("REV_RECARGA_USUARIO_ID")) Then
                    REV_RECARGA_USUARIO_ID = rs("REV_RECARGA_USUARIO_ID")
                Else
                    REV_RECARGA_USUARIO_ID = 0
                End If
                If Not IsNull(rs("REV_RECARGA_FS")) Then
                    REV_RECARGA_FS = rs("REV_RECARGA_FS")
                Else
                    REV_RECARGA_FS = 0
                End If
                Carga = True
        End If
        Set rs = Nothing
        Exit Function
fallo:
        Carga = False
        MsgBox "Error al cargar los datos(clsWeb_muestras_revision)", vbCritical, Err.Description
End Function
Public Function Insertar() As Long
        On Error GoTo fallo
        Dim consulta As String
        consulta = "INSERT INTO WEB_MUESTRAS_REVISION " & _
                   "  VALUES (" & _
                        MUESTRA_ID & "," & CONFORME & "," & "'" & CONFORME_TEXTO & "'" & "," & _
                        CONFORME_USUARIO_ID & "," & "'" & CONFORME_FS & "'" & "," & RECARGA & "," & _
                        "'" & RECARGA_TEXTO & "'" & "," & RECARGA_USUARIO_ID & "," & "'" & RECARGA_FS & "'" & "," & _
                        PARADA & "," & "'" & PARADA_TEXTO & "'" & "," & PARADA_USUARIO_ID & "," & _
                        "'" & PARADA_FS & "'" & "," & REV_RECARGA & "," & "'" & REV_RECARGA_TEXTO & "'" & "," & _
                        REV_RECARGA_USUARIO_ID & "," & "'" & REV_RECARGA_FS & "'" & _
                ")"
        execute_bd consulta
        Insertar = MUESTRA_ID
        Exit Function
fallo:
        Insertar = 0
        MsgBox "Error al insertar (clsWeb_muestras_revision)", vbCritical, Err.Description
End Function
Public Function Modificar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "UPDATE WEB_MUESTRAS_REVISION SET " & _
                        " CONFORME = " & CONFORME & "," & _
                        " CONFORME_TEXTO = '" & CONFORME_TEXTO & "'," & _
                        " CONFORME_USUARIO_ID = " & CONFORME_USUARIO_ID & "," & _
                        " CONFORME_FS = '" & CONFORME_FS & "'," & _
                        " RECARGA = " & RECARGA & "," & _
                        " RECARGA_TEXTO = '" & RECARGA_TEXTO & "'," & _
                        " RECARGA_USUARIO_ID = " & RECARGA_USUARIO_ID & "," & _
                        " RECARGA_FS = '" & RECARGA_FS & "'," & _
                        " PARADA = " & PARADA & "," & _
                        " PARADA_TEXTO = '" & PARADA_TEXTO & "'," & _
                        " PARADA_USUARIO_ID = " & PARADA_USUARIO_ID & "," & _
                        " PARADA_FS = '" & PARADA_FS & "'," & _
                        " REV_RECARGA = " & REV_RECARGA & "," & _
                        " REV_RECARGA_TEXTO = '" & REV_RECARGA_TEXTO & "'," & _
                        " REV_RECARGA_USUARIO_ID = " & REV_RECARGA_USUARIO_ID & "," & _
                        " REV_RECARGA_FS = '" & REV_RECARGA_FS & "'" & _
                " WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Modificar = True
        Exit Function
fallo:
        Modificar = False
        MsgBox "Error al Modificar (clsWeb_muestras_revision)", vbCritical, Err.Description
End Function
Public Function Eliminar(ID As Long) As Boolean
        On Error GoTo fallo
        Dim consulta As String
        consulta = "DELETE FROM WEB_MUESTRAS_REVISION " & _
                "    WHERE MUESTRA_ID = " & ID
        execute_bd consulta
        Eliminar = True
        Exit Function
fallo:
        Eliminar = False
        MsgBox "Error al Eliminar (clsWeb_muestras_revision)", vbCritical, Err.Description
End Function
Public Function Listado() As ADODB.Recordset
        Dim consulta As String
        consulta = "SELECT * FROM WEB_MUESTRAS_REVISION ORDER BY MUESTRA_ID"
        Set Listado = datos_bd(consulta)
End Function
