VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad_Cuentas_BM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsContabilidad_Cuentas_BM
'***************************************************************
Private Cuenta As String
Private Titulo As String
Private Domicilio As String
Private CodPostal As String
Private DNICIF As String
Private Telefono As String
Private Localidad As String
'*******************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsContabilidad_Cuentas_BM
'*******************************************************************

Public Property Let setCuenta(ByVal dato As String)
    Cuenta = dato
End Property
Public Property Let setTitulo(ByVal dato As String)
    Titulo = dato
End Property
Public Property Let setDomicilio(ByVal dato As String)
    Domicilio = dato
End Property
Public Property Let setCodPostal(ByVal dato As String)
    CodPostal = dato
End Property
Public Property Let setDNICIF(ByVal dato As String)
    DNICIF = dato
End Property
Public Property Let setTelefono(ByVal dato As String)
    Telefono = dato
End Property
Public Property Let setLocalidad(ByVal dato As String)
    Localidad = dato
End Property
'******************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsContabilidad_Cuentas_BM
'******************************************************************
Public Property Get getCuenta() As String
        getCuenta = Cuenta
End Property
Public Property Get getTitulo() As String
        getTitulo = Titulo
End Property
Public Property Get getDomicilio() As String
        getDomicilio = Domicilio
End Property
Public Property Get getCodPostal() As String
        getCodPostal = CodPostal
End Property
Public Property Get getDNICIF() As String
        getDNICIF = DNICIF
End Property
Public Property Get getTelefono() As String
        getTelefono = Telefono
End Property
Public Property Get getLocalidad() As String
        getLocalidad = Localidad
End Property

Public Function generarCuenta(ruta As String) As Boolean
    Dim apunteContable As String
    
    generarCuenta = False
    On Error GoTo errorFichero
    Open ruta & "\cuentas.dat" For Append As #1
    
    apunteContable = Format(getCuenta, String(12, "@")) & _
                     Format(getTitulo, String(30, "@")) & _
                     Format(getDomicilio, String(30, "@")) & _
                     Format(getCodPostal, String(5, "@")) & _
                     Format(getDNICIF, String(14, "@")) & _
                     Format(getTelefono, String(10, "@")) & _
                     Format(getLocalidad, String(30, "@"))
                        
    Print #1, apunteContable
    Close
    generarCuenta = True
    Exit Function
errorFichero:
    generarCuenta = False
    Close
    Exit Function
End Function

