VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContabilidad_Apuntes_BM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************
'*   VARIABLES DE INSTANCIA DE LA CLASE clsContabilidad_Apuntes_BM
'***************************************************************
Private Empresa As String
Private tipo As String
Private Asiento As String
Private fecha As String
Private apunte As String
Private Cargo As String
Private Abono As String
Private auxiliar As String
Private factura As String
Private FechaFact As String
Private Neto As String
Private base As String
Private iva As String
Private ImporteIVA As String
Private Recargo As String
Private ImpreCargo As String
Private Auto As String
Private Extenso As String
Private Base2 As String
Private IVA2 As String
Private ImporteIV2 As String
Private Recargo2 As String
Private Imprecarg2 As String
Private Base3 As String
Private IVA3 As String
Private ImporteIV3 As String
Private Recargo3 As String
Private Imprecarg3 As String
Private Libre As String
Private Nif As String
Private Sujeto As String
Private BaseRet As String
Private PorcRet As String
Private Retencion As String
Private Arrendam As String

'*******************************************************************
'*   PROPIEDADES DE ESCRITURA DE LA CLASE clsContabilidad_apuntes_BM
'*******************************************************************

Public Property Let setEmpresa(ByVal dato As String)
    Empresa = dato
End Property
Public Property Let setTipo(ByVal dato As String)
    tipo = dato
End Property
Public Property Let setAsiento(ByVal dato As String)
    Asiento = dato
End Property
Public Property Let setFECHA(ByVal dato As String)
    fecha = dato
End Property
Public Property Let setApunte(ByVal dato As String)
    apunte = dato
End Property
Public Property Let setCargo(ByVal dato As String)
    Cargo = dato
End Property
Public Property Let setAbono(ByVal dato As String)
    Abono = dato
End Property
Public Property Let setAuxiliar(ByVal dato As String)
    auxiliar = dato
End Property
Public Property Let setFactura(ByVal dato As String)
    factura = dato
End Property
Public Property Let setFechaFact(ByVal dato As String)
    FechaFact = dato
End Property
Public Property Let setNeto(ByVal dato As String)
    Neto = dato
End Property
Public Property Let setBase(ByVal dato As String)
    base = dato
End Property
Public Property Let setIVA(ByVal dato As String)
    iva = dato
End Property
Public Property Let setImporteIVA(ByVal dato As String)
    ImporteIVA = dato
End Property
Public Property Let setRecargo(ByVal dato As String)
    Recargo = dato
End Property
Public Property Let setImpreCargo(ByVal dato As String)
    ImpreCargo = dato
End Property
Public Property Let setAuto(ByVal dato As String)
    Auto = dato
End Property
Public Property Let setExtenso(ByVal dato As String)
    Extenso = Left(dato, 30)
End Property
Public Property Let setBase2(ByVal dato As String)
    Base2 = dato
End Property
Public Property Let setIVA2(ByVal dato As String)
    IVA2 = dato
End Property
Public Property Let setImporteIV2(ByVal dato As String)
    ImporteIV2 = dato
End Property
Public Property Let setRecargo2(ByVal dato As String)
    Recargo2 = dato
End Property
Public Property Let setImprecarg2(ByVal dato As String)
    Imprecarg2 = dato
End Property
Public Property Let setBase3(ByVal dato As String)
    Base3 = dato
End Property
Public Property Let setIVA3(ByVal dato As String)
    IVA3 = dato
End Property
Public Property Let setImporteIV3(ByVal dato As String)
    ImporteIV3 = dato
End Property
Public Property Let setRecargo3(ByVal dato As String)
    Recargo3 = dato
End Property
Public Property Let setImprecarg3(ByVal dato As String)
    Imprecarg3 = dato
End Property
Public Property Let setLibre(ByVal dato As String)
    Libre = dato
End Property
Public Property Let setNIF(ByVal dato As String)
    Nif = dato
End Property
Public Property Let setSujeto(ByVal dato As String)
    Sujeto = dato
End Property
Public Property Let setBaseRet(ByVal dato As String)
    BaseRet = dato
End Property
Public Property Let setPorcRet(ByVal dato As String)
    PorcRet = dato
End Property
Public Property Let setRetencion(ByVal dato As String)
    Retencion = dato
End Property
Public Property Let setArrendam(ByVal dato As String)
    Arrendam = dato
End Property


'******************************************************************
'*   PROPIEDADES DE LECTURA DE LA CLASE clsContabilidad_Apuntes_BM
'******************************************************************

Public Property Get getEmpresa() As String
    getEmpresa = Empresa
End Property
Public Property Get getTipo() As String
    getTipo = tipo
End Property
Public Property Get getAsiento() As String
    getAsiento = Asiento
End Property
Public Property Get getFECHA() As String
    getFECHA = fecha
End Property
Public Property Get getApunte() As String
    getApunte = apunte
End Property
Public Property Get getCargo() As String
    getCargo = Cargo
End Property
Public Property Get getAbono() As String
    getAbono = Abono
End Property
Public Property Get getAuxiliar() As String
    getAuxiliar = auxiliar
End Property
Public Property Get getFactura() As String
    getFactura = factura
End Property
Public Property Get getFechaFact() As String
    getFechaFact = FechaFact
End Property
Public Property Get getNeto() As String
    getNeto = Neto
End Property
Public Property Get getBase() As String
    getBase = base
End Property
Public Property Get getIVA() As String
    getIVA = iva
End Property
Public Property Get getImporteIVA() As String
    getImporteIVA = ImporteIVA
End Property
Public Property Get getRecargo() As String
    getRecargo = Recargo
End Property
Public Property Get getImpreCargo() As String
    getImpreCargo = ImpreCargo
End Property
Public Property Get getAuto() As String
    getAuto = Auto
End Property
Public Property Get getExtenso() As String
    getExtenso = Extenso
End Property
Public Property Get getBase2() As String
    getBase2 = Base2
End Property
Public Property Get getIVA2() As String
    getIVA2 = IVA2
End Property
Public Property Get getImporteIV2() As String
    getImporteIV2 = ImporteIV2
End Property
Public Property Get getRecargo2() As String
    getRecargo2 = Recargo2
End Property
Public Property Get getImprecarg2() As String
    getImprecarg2 = Imprecarg2
End Property
Public Property Get getBase3() As String
    getBase3 = Base3
End Property
Public Property Get getIVA3() As String
    getIVA3 = IVA3
End Property
Public Property Get getImporteIV3() As String
    getImporteIV3 = ImporteIV3
End Property
Public Property Get getRecargo3() As String
    getRecargo3 = Recargo3
End Property
Public Property Get getImprecarg3() As String
    getImprecarg3 = Imprecarg3
End Property
Public Property Get getLibre() As String
    getLibre = Libre
End Property
Public Property Get getNIF() As String
    getNIF = Nif
End Property
Public Property Get getSujeto() As String
    getSujeto = Sujeto
End Property
Public Property Get getBaseRet() As String
    getBaseRet = BaseRet
End Property
Public Property Get getPorcRet() As String
    getPorcRet = PorcRet
End Property
Public Property Get getRetencion() As String
    getRetencion = Retencion
End Property
Public Property Get getArrendam() As String
    getArrendam = Arrendam
End Property

Public Function generarApunte(ruta As String) As Boolean
    Dim apunteContable As String
    
    generarApunte = False
    On Error GoTo errorFichero
    Open ruta & "\apuntes.dat" For Append As #1
    
    apunteContable = Format(getEmpresa, String(2, "@")) & getTipo & _
                    Format(getAsiento, String(2, "@")) & Format(getFECHA, "DDMMYY") & _
                    getApunte & Format(getCargo, String(12, "@")) & Format(getAbono, String(12, "@")) & _
                    Format(getAuxiliar, String(8, "@")) & Format(getFactura, String(10, "@")) & Format(getFechaFact, "DDMMYY") & _
                    Format(Format(getNeto, "0.00"), String(12, "@")) & Format(Format(getBase, "0.00"), String(12, "@")) & _
                    Format(getIVA, String(4, "@")) & Format(Format(getImporteIVA, "0.00"), String(12, "@")) & _
                    Format(getRecargo, String(4, "@")) & Format(Format(getImpreCargo, "0.00"), String(12, "@")) & _
                    getAuto & Format(getExtenso, "!" & String(30, "@")) & Format(Format(getBase2, "0.00"), String(12, "@")) & _
                    Format(getIVA2, String(4, "@")) & Format(Format(getImporteIV2, "0.00"), String(12, "@")) & _
                    Format(getRecargo2, String(4, "@")) & _
                    Format(Format(getImprecarg2, "0.00"), String(12, "@")) & _
                    Format(Format(getBase3, "0.00"), String(12, "@")) & _
                    Format(getIVA3, String(4, "@")) & Format(Format(getImporteIV3, "0.00"), String(12, "@")) & Format(getRecargo3, String(4, "@")) & Format(Format(getImprecarg3, "0.00"), String(12, "@")) & _
                    getLibre & Format(getNIF, String(14, "@")) & getSujeto & _
                    Format(Format(getBaseRet, "0.00"), String(12, "@")) & _
                    Format(Format(getPorcRet, "0.00"), String(5, "@")) & _
                    Format(Format(getRetencion, "0.00"), String(12, "@")) & getArrendam
                        
    Print #1, apunteContable
    Close
    generarApunte = True
    Exit Function
errorFichero:
    generarApunte = False
    Close
    Exit Function
End Function
