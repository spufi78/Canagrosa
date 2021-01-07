VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmConversion 
   Caption         =   "Conversor de datos"
   ClientHeight    =   9150
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11775
   Icon            =   "frmConversion.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   11775
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "Conta"
      Height          =   555
      Left            =   2610
      TabIndex        =   31
      Top             =   8370
      Width           =   1455
   End
   Begin VB.CheckBox chkac 
      Caption         =   "Actualizar"
      Height          =   285
      Left            =   1035
      TabIndex        =   30
      Top             =   8910
      Width           =   1725
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Descuentos"
      Height          =   555
      Left            =   1035
      TabIndex        =   29
      Top             =   8370
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Fecha Cobro"
      Height          =   555
      Left            =   8820
      TabIndex        =   28
      Top             =   8460
      Width           =   2040
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Documentos Recibos Pendientes de contabilizar"
      Height          =   285
      Index           =   19
      Left            =   90
      TabIndex        =   27
      Top             =   5625
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Liquidaciones"
      Height          =   285
      Index           =   18
      Left            =   90
      TabIndex        =   26
      Top             =   5280
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Descuentos"
      Height          =   285
      Index           =   17
      Left            =   90
      TabIndex        =   25
      Top             =   4980
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Efectos (Remesas_Documentos)"
      Height          =   285
      Index           =   12
      Left            =   90
      TabIndex        =   24
      Top             =   4680
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Remesas"
      Height          =   285
      Index           =   11
      Left            =   90
      TabIndex        =   23
      Top             =   4380
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Ofertas"
      Height          =   285
      Index           =   9
      Left            =   90
      TabIndex        =   22
      Top             =   4080
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Usuarios"
      Height          =   285
      Index           =   8
      Left            =   90
      TabIndex        =   21
      Top             =   3810
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Vehiculos"
      Height          =   285
      Index           =   16
      Left            =   90
      TabIndex        =   20
      Top             =   3510
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Facturas"
      Height          =   285
      Index           =   15
      Left            =   90
      TabIndex        =   19
      Top             =   3240
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Agentes"
      Height          =   285
      Index           =   14
      Left            =   90
      TabIndex        =   18
      Top             =   2370
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Descuentos"
      Height          =   285
      Index           =   13
      Left            =   90
      TabIndex        =   17
      Top             =   2070
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Tarifas"
      Height          =   285
      Index           =   7
      Left            =   90
      TabIndex        =   16
      Top             =   1800
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Tarifas Portes"
      Height          =   285
      Index           =   6
      Left            =   90
      TabIndex        =   15
      Top             =   1260
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Obras"
      Height          =   285
      Index           =   3
      Left            =   90
      TabIndex        =   14
      Top             =   960
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Formas de Pago"
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   12
      Top             =   90
      Width           =   4305
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1020
      TabIndex        =   10
      Top             =   7560
      Visible         =   0   'False
      Width           =   1230
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1020
      TabIndex        =   8
      Top             =   7260
      Width           =   10725
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Articulos"
      Height          =   285
      Index           =   10
      Left            =   90
      TabIndex        =   7
      Top             =   1530
      Width           =   4305
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   405
      Left            =   90
      TabIndex        =   6
      Top             =   8010
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   714
      _Version        =   393216
      Appearance      =   1
      Min             =   1
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Albaranes_Detalle"
      Height          =   285
      Index           =   5
      Left            =   90
      TabIndex        =   5
      Top             =   2970
      Width           =   4305
   End
   Begin VB.TextBox Text1 
      Height          =   7125
      Left            =   4470
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      Top             =   60
      Width           =   7275
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Albaranes"
      Height          =   285
      Index           =   4
      Left            =   90
      TabIndex        =   3
      Top             =   2670
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Proveedores"
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   2
      Top             =   675
      Width           =   4305
   End
   Begin VB.CheckBox chkt 
      Caption         =   "Clientes"
      Height          =   285
      Index           =   0
      Left            =   90
      TabIndex        =   1
      Top             =   390
      Width           =   4305
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Convertir"
      Height          =   525
      Left            =   5160
      TabIndex        =   0
      Top             =   8505
      Width           =   1725
   End
   Begin VB.Label lbll 
      Alignment       =   2  'Center
      Height          =   285
      Left            =   4050
      TabIndex        =   13
      Top             =   7650
      Width           =   3660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Ruta"
      Height          =   195
      Left            =   90
      TabIndex        =   11
      Top             =   7590
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "B.D. Antigua"
      Height          =   195
      Left            =   90
      TabIndex        =   9
      Top             =   7290
      Width           =   900
   End
End
Attribute VB_Name = "frmConversion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public BD_ANTIGUA As String
Public ruta As String
Public rs_antiguo As ADODB.Recordset
Public rs_antiguo2 As ADODB.Recordset
Public rs As ADODB.Recordset
Public consulta As String
Public conn_antigua As ADODB.Connection

Private Sub Command1_Click()
    Me.MousePointer = 11
'    obras_destino_facturas
'    obras_telefonos
    documentos_servido
 '   usuarios_conversion
 '   descuentos
 '   formasPago
 '   clientes
 '   proveedores
'    agentes
'     obras
'    tarifasPortes
'    ARTICULOS
'    tarifas
'    albaranes
'    albaranes_detalle
'    facturas
'    vehiculos
'    ofertas
 '   remesas
 '   efectos
 '   descuentos
 '   liquidaciones
 '   documentos_recibos
  
    
    Me.MousePointer = 0
    MsgBox "Proceso finalizado.", vbInformation, App.Title
End Sub
Public Sub ARTICULOS()
    Dim oArticulo As New clsArticulos
    execute_bd "DELETE FROM ARTICULOS"
    execute_bd "DELETE FROM ARTICULOS_TIPOS"
    execute_bd "INSERT INTO ARTICULOS_TIPOS VALUES (0,'Sin Especificar')"
    consulta_antigua "select * from articulos order by coart"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Do
            With oArticulo
                .setID_ARTICULO = rs_antiguo("coart")
                .setTIPO_ARTICULO_ID = 0
                .setDESCRIPCION = ""
                If Not IsNull(rs_antiguo("descripcion")) Then
                    .setDESCRIPCION = Trim(rs_antiguo("descripcion"))
                End If
                .setPROVEEDOR_ID = 0
                If Not IsNull(rs_antiguo("costo")) Then
                    .setPRECIO_COMPRA = moneda_bd(rs_antiguo("costo"))
                Else
                    .setPRECIO_COMPRA = moneda_bd("0")
                End If
                .setSTOCK = 0
                .setCOMENTARIO = ""
                If Not IsNull(rs_antiguo("detalle")) Then
                    .setCOMENTARIO = Trim(rs_antiguo("detalle"))
                End If
                .setPESO = moneda_bd("0")
                If Not IsNull(rs_antiguo("PESO")) Then
                    .setPESO = moneda_bd(Trim(rs_antiguo("PESO")))
                End If
                If Not IsNull(rs_antiguo("COMI")) Then
                    .setCOMISION = moneda_bd(Trim(rs_antiguo("COMI")))
                Else
                    .setCOMISION = moneda_bd("0")
                End If
                .setINE = 0
                If Not IsNull(rs_antiguo("EINE")) Then
                    .setINE = 1
                End If
                If Not IsNull(rs_antiguo("CORE")) Then
                    .setINE_CODIGO = Trim(rs_antiguo("CORE"))
                Else
                    .setINE_CODIGO = 0
                End If
                .Insertar
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(10).Value = Checked
End Sub
Public Sub clientes()
    Dim ocliente As New clsCliente
    execute_bd "DELETE FROM PROVINCIAS"
    execute_bd "DELETE FROM MUNICIPIOS"
    execute_bd "DELETE FROM CLIENTES"
    execute_bd "DELETE FROM CLIENTES_DIRECCIONES"
    execute_bd "DELETE FROM AGENDA"
    consulta_antigua "select * from clientes order by cobra"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    Dim oProv As New clsProvincias
    Dim oMun As New clsMunicipios
    
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Dim DIRECCION As String
        Do
            With ocliente
                .setID_CLIENTE = rs_antiguo("cobra")
                If IsNull(rs_antiguo("NIF")) Then
                    .setCIF = ""
                Else
                    .setCIF = rs_antiguo("NIF")
                End If
                If IsNull(rs_antiguo("NOMBRE")) Then
                    .setNOMBRE = ""
                Else
                    .setNOMBRE = UCase(Replace(rs_antiguo("NOMBRE"), "'", "`"))
                End If
                DIRECCION = ""
                If Not IsNull(rs_antiguo("sg")) Then
                    DIRECCION = rs_antiguo("sg") & " "
                End If
                If Not IsNull(rs_antiguo("DIRECCION")) Then
                    DIRECCION = DIRECCION & rs_antiguo("DIRECCION") & " "
                End If
                If Not IsNull(rs_antiguo("numero")) Then
                    DIRECCION = DIRECCION & rs_antiguo("numero") & " "
                End If
                .setDIRECCION = Trim(DIRECCION)
                
                If rs_antiguo("cp") = "" Then
                    .setCP = 0
                Else
                    If IsNumeric(rs_antiguo("cp")) Then
                        .setCP = rs_antiguo("CP")
                    Else
                        .setCP = 0
                    End If
                End If
                ' PROVINCIA
                Dim PROVINCIA As Integer
                PROVINCIA = 0
                If IsNull(rs_antiguo("PROVINCIA")) Then
                    .setPROVINCIA_ID = 0
                Else
                    If Trim(rs_antiguo("PROVINCIA")) = "" Then
                        .setPROVINCIA_ID = 0
                    Else
                        C = "SELECT * FROM PROVINCIAS WHERE NOMBRE LIKE '%" & rs_antiguo("PROVINCIA") & "%'"
                        Set rs = datos_bd(C)
                        If rs.RecordCount > 0 Then
                            .setPROVINCIA_ID = rs("ID_PROVINCIA")
                        Else
                            oProv.setNOMBRE = rs_antiguo("PROVINCIA")
                            .setPROVINCIA_ID = oProv.Insertar
                        End If
                        PROVINCIA = .getPROVINCIA_ID
                    End If
                End If
                ' MUNICIPIO
                If IsNull(rs_antiguo("POBLACION")) Then
                    .setMUNICIPIO_ID = 0
                Else
                    If Trim(rs_antiguo("POBLACION")) = "" Then
                        .setMUNICIPIO_ID = 0
                    Else
                        C = "SELECT * FROM MUNICIPIOS WHERE NOMBRE LIKE '%" & rs_antiguo("POBLACION") & "%'"
                        Set rs = datos_bd(C)
                        If rs.RecordCount > 0 Then
                            .setMUNICIPIO_ID = rs("ID_MUNICIPIO")
                        Else
                            oMun.setPROVINCIA_ID = PROVINCIA
                            oMun.setNOMBRE = rs_antiguo("POBLACION")
                            .setMUNICIPIO_ID = oMun.Insertar
                        End If
                    End If
                End If
                
                .setIVA = ReadINI(App.Path & "\config.ini", "parametros", "iva")
                .setCOPIAS_FACTURA = ReadINI(App.Path & "\config.ini", "parametros", "Copias_facturas")
                ' TELEFONOS
                If IsNull(rs_antiguo("TELEFONOS")) Then
                    .setTELEFONO = ""
                Else
                    .setTELEFONO = Trim(rs_antiguo("TELEFONOS"))
                End If
                If IsNull(rs_antiguo("FAX")) Then
                    .setFAX = ""
                Else
                    .setFAX = Trim(rs_antiguo("FAX"))
                End If
                If IsNull(rs_antiguo("contabilidad")) Then
                    .setCCONTABLE = ""
                Else
                    .setCCONTABLE = Trim(rs_antiguo("contabilidad"))
                End If
                
                
                .setMOVIL = ""
                
                If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                    .setRAZON = ""
                Else
                    If IsNull(rs_antiguo("contacto")) Then
                        .setRAZON = ""
                    Else
                        .setRAZON = Trim(rs_antiguo("contacto"))
                    End If
                End If
                If IsNull(rs_antiguo("email")) Then
                    .setEMAIL = ""
                Else
                    .setEMAIL = Trim(rs_antiguo("email"))
                End If
                If IsNull(rs_antiguo("obse")) Then
                    .setOBSERVACIONES = ""
                Else
                    .setOBSERVACIONES = Trim(rs_antiguo("obse"))
                End If
                ' Forma de pago
                .setFORMA_PAGO = 0
                If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                    .setFORMA_PAGO = 0
                    .setRIESGO = moneda_bd("0")
                    .setRIESGO_REAL = moneda_bd("0")
                Else
                    If IsNull(rs_antiguo("FORPA")) Then
                        .setFORMA_PAGO = 0
                    Else
                        If rs_antiguo("FORPA") = "" Then
                            .setFORMA_PAGO = 0
                        Else
                            C = "SELECT * FROM FORMA_PAGO WHERE NOMBRE LIKE '%" & rs_antiguo("FORPA") & "%'"
                            Set rs = datos_bd(C)
                            If rs.RecordCount > 0 Then
                                .setFORMA_PAGO = rs("ID_FORMA_PAGO")
                            Else
                                Dim oFp As New clsForma_pago
                                oFp.setNOMBRE = rs_antiguo("FORPA")
                                .setFORMA_PAGO = oFp.Insertar
                                Text1 = Text1 & vbNewLine & "Error, no existe la FP: " & rs_antiguo("FORPA") & ". Cliente : " & rs_antiguo("COBRA") & "/" & rs_antiguo("NOMBRE")
                            End If
                        End If
                    End If
                    ' RIESGO
                    If IsNull(rs_antiguo("RIESGO")) Then
                        .setRIESGO = moneda_bd("0")
                    Else
                        .setRIESGO = moneda_bd(rs_antiguo("RIESGO"))
                    End If
                    If IsNull(rs_antiguo("RIESGOR")) Then
                        .setRIESGO_REAL = moneda_bd("0")
                    Else
                        .setRIESGO_REAL = moneda_bd(rs_antiguo("RIESGOR"))
                    End If
                End If

                
                .insertar_cliente
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(0).Value = Checked
End Sub
Public Sub proveedores()
    Dim oProveedor As New clsProveedor
    execute_bd "DELETE FROM PROVEEDORES"
    consulta_antigua "select * from proveedores order by cuenta"
    execute_bd "INSERT INTO PROVEEDORES (ID_PROVEEDOR,NOMBRE,CIF) VALUES (0,'Sin especificar','')"
    Dim rs As ADODB.Recordset
    Dim oProv As New clsProvincias
    Dim oMun As New clsMunicipios
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Do
            With oProveedor
                .setID_PROVEEDOR = rs_antiguo("CUENTA")
                If Not IsNull(rs_antiguo("NIF")) Then
                    .setCIF = rs_antiguo("NIF")
                Else
                    .setCIF = ""
                End If
                .setNOMBRE = Replace(rs_antiguo("NOMBRE"), "'", "`")
                If Not IsNull(rs_antiguo("ACTIVIDAD")) Then
                    .setACTIVIDAD = Replace(rs_antiguo("ACTIVIDAD"), "'", "`")
                Else
                    .setACTIVIDAD = ""
                End If
                If Not IsNull(rs_antiguo("CONTACTO")) Then
                    .setRESPONSABLE = rs_antiguo("CONTACTO")
                Else
                    .setRESPONSABLE = ""
                End If
                ' Dirección
                DIRECCION = ""
                If Not IsNull(rs_antiguo("cv1")) Then
                    DIRECCION = rs_antiguo("cv1") & " "
                End If
                If Not IsNull(rs_antiguo("DIRECCION")) Then
                    DIRECCION = DIRECCION & Replace(rs_antiguo("DIRECCION"), "'", "`") & " "
                End If
                If Not IsNull(rs_antiguo("numero")) Then
                    DIRECCION = DIRECCION & rs_antiguo("numero") & " "
                End If
                If Not IsNull(rs_antiguo("bloque")) Then
                    DIRECCION = DIRECCION & rs_antiguo("bloque") & " "
                End If
                If Not IsNull(rs_antiguo("piso")) Then
                    DIRECCION = DIRECCION & rs_antiguo("piso") & " "
                End If
                If Not IsNull(rs_antiguo("puerta")) Then
                    DIRECCION = DIRECCION & rs_antiguo("puerta") & " "
                End If
                .setDIRECCION = Trim(DIRECCION)
                .setCP = 0
                If rs_antiguo("CP1") <> "" Then
                    If Not IsNull(rs_antiguo("CP1")) Then
                        If IsNumeric(rs_antiguo("CP1")) Then
                            .setCP = rs_antiguo("CP1")
                        End If
                    End If
                End If
                ' PROVINCIA
                Dim PROVINCIA As Integer
                PROVINCIA = 0
                If IsNull(rs_antiguo("PROVINCIA")) Then
                    .setPROVINCIA_ID = 0
                Else
                    If Trim(rs_antiguo("PROVINCIA")) = "" Then
                        .setPROVINCIA_ID = 0
                    Else
                        C = "SELECT * FROM PROVINCIAS WHERE NOMBRE LIKE '%" & Replace(rs_antiguo("PROVINCIA"), "'", "`") & "%'"
                        Set rs = datos_bd(C)
                        If rs.RecordCount > 0 Then
                            .setPROVINCIA_ID = rs("ID_PROVINCIA")
                        Else
                            oProv.setNOMBRE = Replace(rs_antiguo("PROVINCIA"), "'", "`")
                            .setPROVINCIA_ID = oProv.Insertar
                        End If
                        PROVINCIA = .getPROVINCIA_ID
                    End If
                End If
                ' MUNICIPIO
                If IsNull(rs_antiguo("POBLACION")) Then
                    .setMUNICIPIO_ID = 0
                Else
                    If Trim(rs_antiguo("POBLACION")) = "" Then
                        .setMUNICIPIO_ID = 0
                    Else
                        C = "SELECT * FROM MUNICIPIOS WHERE NOMBRE LIKE '%" & Replace(rs_antiguo("POBLACION"), "'", "`") & "%'"
                        Set rs = datos_bd(C)
                        If rs.RecordCount > 0 Then
                            .setMUNICIPIO_ID = rs("ID_MUNICIPIO")
                        Else
                            oMun.setPROVINCIA_ID = PROVINCIA
                            oMun.setNOMBRE = Replace(rs_antiguo("POBLACION"), "'", "`")
                            .setMUNICIPIO_ID = oMun.Insertar
                        End If
                    End If
                End If
                ' Dirección NOTIFICACIONES
                DIRECCION = ""
                If Not IsNull(rs_antiguo("cvN")) Then
                    DIRECCION = rs_antiguo("cvN") & " "
                End If
                If Not IsNull(rs_antiguo("DIRECCIONN")) Then
                    DIRECCION = DIRECCION & Replace(rs_antiguo("DIRECCIONN"), "'", "`") & " "
                End If
                If Not IsNull(rs_antiguo("numeroN")) Then
                    DIRECCION = DIRECCION & rs_antiguo("numeroN") & " "
                End If
                If Not IsNull(rs_antiguo("bloqueN")) Then
                    DIRECCION = DIRECCION & rs_antiguo("bloqueN") & " "
                End If
                If Not IsNull(rs_antiguo("pisoN")) Then
                    DIRECCION = DIRECCION & rs_antiguo("pisoN") & " "
                End If
                If Not IsNull(rs_antiguo("puertaN")) Then
                    DIRECCION = DIRECCION & rs_antiguo("puertaN") & " "
                End If
                .setDIRECCIONN = Trim(DIRECCION)
                .setCPN = 0
                If rs_antiguo("CPN") <> "" Then
                    If Not IsNull(rs_antiguo("CPN")) Then
                        If IsNumeric(rs_antiguo("CPN")) Then
                            .setCPN = rs_antiguo("CPN")
                        End If
                    End If
                End If
                ' PROVINCIA
                PROVINCIA = 0
                If IsNull(rs_antiguo("PROVINCIAN")) Then
                    .setPROVINCIAN_ID = 0
                Else
                    If Trim(rs_antiguo("PROVINCIAN")) = "" Then
                        .setPROVINCIAN_ID = 0
                    Else
                        C = "SELECT * FROM PROVINCIAS WHERE NOMBRE LIKE '%" & Replace(rs_antiguo("PROVINCIAN"), "'", "`") & "%'"
                        Set rs = datos_bd(C)
                        If rs.RecordCount > 0 Then
                            .setPROVINCIAN_ID = rs("ID_PROVINCIA")
                        Else
                            oProv.setNOMBRE = Replace(rs_antiguo("PROVINCIAN"), "'", "`")
                            .setPROVINCIAN_ID = oProv.Insertar
                        End If
                        PROVINCIA = .getPROVINCIAN_ID
                    End If
                End If
                ' MUNICIPIO
                If IsNull(rs_antiguo("POBLACIONN")) Then
                    .setMUNICIPION_ID = 0
                Else
                    If Trim(rs_antiguo("POBLACIONN")) = "" Then
                        .setMUNICIPION_ID = 0
                    Else
                        C = "SELECT * FROM MUNICIPIOS WHERE NOMBRE LIKE '%" & Replace(rs_antiguo("POBLACIONN"), "'", "`") & "%'"
                        Set rs = datos_bd(C)
                        If rs.RecordCount > 0 Then
                            .setMUNICIPION_ID = rs("ID_MUNICIPIO")
                        Else
                            oMun.setPROVINCIA_ID = PROVINCIA
                            oMun.setNOMBRE = Replace(rs_antiguo("POBLACION"), "'", "`")
                            .setMUNICIPION_ID = oMun.Insertar
                        End If
                    End If
                End If
                
                If IsNull(rs_antiguo("TLM")) Then
                    .setMOVIL = ""
                Else
                    .setMOVIL = rs_antiguo("TLM")
                End If
                If IsNull(rs_antiguo("TLF")) Then
                    .setTELEFONO = ""
                Else
                    .setTELEFONO = rs_antiguo("TLF")
                End If
                If IsNull(rs_antiguo("fax")) Then
                    .setFAX = ""
                Else
                    .setFAX = Left(rs_antiguo("FAX"), 30)
                End If
                If IsNull(rs_antiguo("EMAIL")) Then
                    .setEMAIL = ""
                Else
                    .setEMAIL = Trim(rs_antiguo("EMAIL"))
                End If
                
                .setBANCO = ""
                .setCCC = "____-____-__-__________"
                .setFORMA_PAGO = 0
                
                If IsNull(rs_antiguo("OBSERVACIONES")) Then
                    .setOBSERVACIONES = ""
                Else
                    .setOBSERVACIONES = rs_antiguo("OBSERVACIONES")
                End If
                .Insertar
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(1).Value = Checked
End Sub
Public Sub agentes()
    execute_bd "DELETE FROM COMERCIALES"
    consulta_antigua "select * from agentes order by codigo"
    Dim rs As ADODB.Recordset
    Dim oComercial As New clsComercial
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Do
            With oComercial
                .setID_COMERCIAL = rs_antiguo("codigo")
                .setNOMBRE = rs_antiguo("nombre")
                If IsNull(rs_antiguo("dni")) Then
                    .setCIF = ""
                Else
                    .setCIF = rs_antiguo("dni")
                End If
                .setMOVIL = rs_antiguo("telefono")
                .setCOMISION = moneda_bd(rs_antiguo("pcomi"))
                .Insertar
                
'                Dim oCC As New clsComercial_Comision
'                oCC.setCOMERCIAL_ID = rs_antiguo("codigo")
'                oCC.setCOMISION = rs_antiguo("pcomi")
'                oCC.setDESCUENTO_ID = 0
'                oCC.Insertar
'                Set oCC = Nothing
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(14).Value = Checked
End Sub

Public Sub albaranes()
    Dim oDOC As New clsDocumentos
    execute_bd "DELETE FROM DOCUMENTOS"
    execute_bd "DELETE FROM DOCUMENTOS_ANULADOS"
    consulta_antigua "select * from albaranes order by clalb"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Dim DIRECCION As String
        Do
            With oDOC
                .setID_DOCUMENTO = rs_antiguo("clalb")
                .setTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.ALBARAN
                .setANNO = Year(rs_antiguo("fecha"))
                .setNUMERO = rs_antiguo("clalb")
                .setOBRA_ID = rs_antiguo("OBRA")
                .setFECHA = Format(rs_antiguo("fecha"), "yyyy-mm-dd")
                .setTOTAL = moneda_bd(rs_antiguo("importe"))
                .setPORTES = moneda_bd(rs_antiguo("PORTES"))
                If rs_antiguo("situacion") = "P" Then
                    .setFACTURADO = 0
                ElseIf rs_antiguo("situacion") = "F" Then
                    .setFACTURADO = 1
                Else
                    Text1 = Text1 & vbNewLine & "Error, no existe la situacion: " & rs_antiguo("situacion") & ". DOCUMENTO : " & rs_antiguo("clalb")
                End If
                If IsNull(rs_antiguo("OBSERVACIONES")) Then
                    .setOBSERVACIONES = ""
                Else
                    .setOBSERVACIONES = Trim(rs_antiguo("OBSERVACIONES"))
                End If
                .setTARIFA_ID = rs_antiguo("tarifa")
                .setSERVIDO = rs_antiguo("SERVIDO")
                If rs_antiguo("VALORACION") = "S" Then
                    .setVALORACION = 1
                Else
                    .setVALORACION = 0
                End If
                
                .setMATRICULA = ""
                If Not IsNumeric(rs_antiguo("vehiculo")) Then
                    .setVEHICULO_ID = 1
                    .setNIF = ""
                    .setREMOLQUE = ""
                Else
                    If CInt(rs_antiguo("vehiculo")) = 1 Then
                        .setVEHICULO_ID = 1
                        If Not IsNull(rs_antiguo("matricula")) Then
                            .setMATRICULA = rs_antiguo("matricula")
                        End If
                        If Not IsNull(rs_antiguo("nifveh")) Then
                            .setNIF = rs_antiguo("nifveh")
                        End If
                        .setREMOLQUE = ""
                    Else
                        .setVEHICULO_ID = rs_antiguo("vehiculo")
                        .setNIF = ""
                        .setREMOLQUE = ""
                    End If
                End If
                If IsNull(rs_antiguo("hora")) Then
                    .setHORA = "00:00"
                Else
                    .setHORA = rs_antiguo("hora")
                End If
                If IsNull(rs_antiguo("bultos")) Then
                    .setBULTOS = 0
                Else
                    .setBULTOS = rs_antiguo("bultos")
                End If
                If IsNull(rs_antiguo("PA")) Then
                    .setPESO = 0
                Else
                    .setPESO = rs_antiguo("PA")
                End If
                ' Usuario
                If rs_antiguo("operador") = "LORENZO" Then
                    .setUSUARIO_ID = 5
                ElseIf rs_antiguo("operador") = "JUAN JOSE" Then
                    .setUSUARIO_ID = 4
                ElseIf rs_antiguo("operador") = "SORAYA" Then
                    .setUSUARIO_ID = 8
                Else
                    Text1 = Text1 & vbNewLine & " No existe el usuario del documento : " & rs_antiguo("clalb")
                End If
                .Insertar
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(4).Value = Checked

End Sub
Public Sub facturas()
    Dim oDOC As New clsDocumentos
    Dim C As String
    execute_bd "delete from DOCUMENTOS_RECIBOS"
    execute_bd "DELETE FROM DOCUMENTOS_COBROS"
    C = "select ffactura,nfactura,obra,tarifa,servido,sum(importe),sum(portes) " & _
        "  from albaranes " & _
        " WHERE situacion='F' " & _
        " group by ffactura,nfactura,obra,tarifa,servido "
    consulta_antigua C
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    Dim ID As Long
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Dim DIRECCION As String
        Do
            With oDOC
                .setID_DOCUMENTO = 0
                .setTIPO_DOCUMENTO_ID = ENUM_TIPOS_DOCUMENTOS.factura
                .setANNO = Year(rs_antiguo("ffactura"))
                If IsNull(rs_antiguo("nfactura")) Then
                    .setNUMERO = 0
                Else
                    .setNUMERO = rs_antiguo("nfactura")
                End If
                .setOBRA_ID = rs_antiguo("OBRA")
                .setFECHA = Format(rs_antiguo("ffactura"), "yyyy-mm-dd")
                .setTOTAL = moneda_bd(rs_antiguo(5))
                .setPORTES = moneda_bd(rs_antiguo(6))
                .setFACTURADO = 1
                .setOBSERVACIONES = ""
                .setTARIFA_ID = rs_antiguo("tarifa")
                .setSERVIDO = rs_antiguo("SERVIDO")
                
                ' Nuevos
                .setVEHICULO_ID = 1
                .setMATRICULA = ""
                .setUSUARIO_ID = 1
                .setHORA = "00:00"
                .setPESO = 0
                .setBULTOS = 0
                .setESTADO_ID = 1
                
                ID = .Insertar
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                ' INFORMAR LOS ALBARANES CON EL NUMERO DE FACTURA
                C = "select * from albaranes where nfactura = '" & rs_antiguo("nfactura") & "'"
                consulta_antigua2 C
                If rs_antiguo2.RecordCount > 0 Then
                    Do
                        execute_bd "UPDATE DOCUMENTOS SET DOCUMENTO_ID_REL = " & ID & " WHERE ID_DOCUMENTO = " & rs_antiguo2("clalb")
                        rs_antiguo2.MoveNext
                    Loop Until rs_antiguo2.EOF
                End If
                rs_antiguo2.Close
                ' VERIFICAR SI ESTA COBRADA Y FORMA DE PAGO (DOCUMENTOS_COBROS)
                C = "select * from [diario facturas] where nfactura = '" & rs_antiguo("nfactura") & "' AND ANNO='11'"
                consulta_antigua2 C
                Dim rss As ADODB.Recordset
                If rs_antiguo2.RecordCount > 0 Then
                    Do
                        Dim FP As Integer
                        If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                            FP = 0
                        Else
                            C = "SELECT * FROM forma_pago WHERE NOMBRE ='" & rs_antiguo2("forpa") & "'"
                            Set rss = datos_bd(C)
                            If rss.RecordCount > 0 Then
                                execute_bd "UPDATE DOCUMENTOS SET FP_ID = " & rss("ID_FORMA_PAGO") & " WHERE ID_DOCUMENTO =  " & ID
                                FP = rss("ID_FORMA_PAGO")
                            Else
                                Text1 = Text1 & vbNewLine & "No encuento FORMA DE PAGO. FACTURA : " & rs_antiguo2("NFACTURA") & " FORPA = " & rs_antiguo2("FORPA")
                                FP = 0
                            End If
                        End If
                        If Not IsNull(rs_antiguo2("fechacobro")) Then
                            Dim oDC As New clsDocumentos_cobros
                            oDC.setDOCUMENTO_ID = ID
                            oDC.setVENCIMIENTO = 1
                            oDC.setDATOS = "COBRADA"
                            oDC.setEMPLEADO_ID = 1
                            oDC.setFECHA = Format(rs_antiguo2("fechacobro"), "yyyy-mm-dd")
                            oDC.setFP_ID = FP
                            oDC.setHORA = "00:00"
                            oDC.setOBSERVACIONES = "Conversión programa antiguo."
                            oDC.Insertar
                            Set oDC = Nothing
                            .modificar_estado ID, ENUM_DOCUMENTOS_ESTADOS.DOCUMENTOS_ESTADOS_COBRADA
                        End If
                        
                        ' Contabilizado
                        If rs_antiguo2("actualizado") = "S" Then
                            oDOC.Contabilizar ID
                        End If
                        rs_antiguo2.MoveNext
                    Loop Until rs_antiguo2.EOF
                End If
                rs_antiguo2.Close
                
                
                ' LIQUIDACION
                
                
                ' DETALLE
                C = "select l.articulo,l.cantidad,l.concepto,l.millar,[l.importe linea],l.portes,a.fecha,a.clalb " & _
                    "  from lineas l, albaranes a " & _
                    " where l.albaran = a.clalb " & _
                    "   and a.nfactura = '" & rs_antiguo(1) & "'" & _
                    " order by l.albaran, clavelinea desc"
                consulta_antigua2 C
                Dim oDD As New clsDocumentos_detalle
                Dim orden As Integer
                orden = 0
                If rs_antiguo2.RecordCount > 0 Then
                    Do
                        oDD.setDOCUMENTO_ID = ID
                        oDD.setORDEN = orden
                        orden = orden + 1
                        If Not IsNumeric(rs_antiguo2("articulo")) Then
                            oDD.setARTICULO_ID = 0
                        Else
                            oDD.setARTICULO_ID = rs_antiguo2("articulo")
                        End If
                        oDD.setCANTIDAD = rs_antiguo2("cantidad")
                        If IsNull(rs_antiguo2("concepto")) Then
                            oDD.setDESCRIPCION = ""
                        Else
                            oDD.setDESCRIPCION = rs_antiguo2("concepto")
                        End If
                        oDD.setPRECIO = moneda_bd(rs_antiguo2("millar"))
                        oDD.setTOTAL = moneda_bd(rs_antiguo2(4))
                        oDD.setPORTES = moneda_bd(rs_antiguo2("portes"))
                        
                        oDD.setFECHA_ALBARAN = rs_antiguo2("FECHA")
                        oDD.setNUMERO_ALBARAN = rs_antiguo2("CLALB")
                        oDD.Insertar
                        rs_antiguo2.MoveNext
                    Loop Until rs_antiguo2.EOF
                End If
                rs_antiguo2.Close
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(15).Value = Checked

End Sub

Public Sub albaranes_detalle()
    Dim oDOC As New clsDocumentos_detalle
    execute_bd "DELETE FROM DOCUMENTOS_DETALLE"
'    consulta_antigua "select * from lineas where albaran = '001346' order by albaran,clavelinea"
    consulta_antigua "select * from lineas order by albaran,clavelinea desc"
    Dim rs As ADODB.Recordset
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Do
            With oDOC
                .setDOCUMENTO_ID = rs_antiguo("albaran")
                If Not IsNumeric(rs_antiguo("articulo")) Then
                    .setARTICULO_ID = 0
                    .setORDEN = 0
                Else
                    .setARTICULO_ID = rs_antiguo("articulo")
                    .setORDEN = rs_antiguo("articulo")
                End If
                .setCANTIDAD = rs_antiguo("cantidad")
                If IsNull(rs_antiguo("concepto")) Then
                    .setDESCRIPCION = ""
                Else
                    .setDESCRIPCION = rs_antiguo("concepto")
                End If
                .setPRECIO = moneda_bd(rs_antiguo("millar"))
                .setTOTAL = moneda_bd(rs_antiguo("importe linea"))
                .setPORTES = moneda_bd(rs_antiguo("portes"))
                .setFECHA_ALBARAN = ""
                .setNUMERO_ALBARAN = 0
                .Insertar
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(5).Value = Checked
End Sub

Private Sub Command2_Click()
    Dim rs As ADODB.Recordset
    Dim oDOC As New clsDocumentos
    Set rs = datos_bd("SELECT * FROM DOCUMENTOS_COBROS")
    If rs.RecordCount > 0 Then
        Do
            oDOC.COBRADO rs("DOCUMENTO_ID"), rs("FECHA")
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    MsgBox "OK"
End Sub

Private Sub Command3_Click()
    Dim rs As ADODB.Recordset
    Dim C As String
    C = "select * from descuentos_documentos where descuento_id >= 824 order by descuento_id,apunte_id"
    Set rs = datos_bd(C)
    Dim oC As New clsContabilidad
    oC.Conectar
    If rs.RecordCount > 0 Then
        Do
            Dim oRE As New clsRemesas_documentos
            oRE.Carga rs("apunte_id")
            Dim oD As New clsDocumentos
            oD.Carga oRE.getDOCUMENTO_ID
            
            DES_APUNTE = "Desc. " & Format(rs("DESCUENTO_ID"), "00000") & _
                         " Rem. " & Format(oRE.getREMESA_ID, "000000") & " /" & Format(rs("apunte_id"), "000000") & _
                         " Fac. FACTURA " & oD.getNUMERO & "/" & Right(Year(oD.getFECHA), 4) & " " & "Vto. " & Format(oRE.getFECHA_VENCIMIENTO, "dd/mm/yy")

            DES_APUNTE_BUENO = "Desc. " & Format(rs("DESCUENTO_ID"), "00000") & _
                               " Rem. " & Format(oRE.getREMESA_ID, "000000") & " /" & Format(rs("apunte_id"), "000000") & _
                               " Fac. " & oRE.getDESCRIPCION & " " & "Vto. " & Format(oRE.getFECHA_VENCIMIENTO, "dd/mm/yy")

            DES_PARTE = "Desc. " & Format(rs("DESCUENTO_ID"), "00000") & _
                        " Rem. " & Format(oRE.getREMESA_ID, "000000") & " /" & Format(rs("apunte_id"), "000000") & _
                        " Fac. "
            If DES_APUNTE <> DES_APUNTE_BUENO Then
                Text1 = Text1 & "---------------" & vbNewLine
                Text1 = Text1 & DES_APUNTE & vbNewLine
                Text1 = Text1 & DES_APUNTE_BUENO & vbNewLine
'            End If
'                Text1 = Text1 & "---------------" & vbNewLine
'            Else
'                Text1 = Text1 & "OK" & vbNewLine
                Dim rs2 As ADODB.Recordset
                Set rs2 = oC.datos_bd_contabilidad("SELECT * FROM MOVIMIENTOS WHERE DESCRIPCION_APUNTE LIKE '" & DES_PARTE & "%'")
'                Set RS2 = oC.datos_bd_contabilidad("SELECT * FROM MOVIMIENTOS WHERE APUNTE = '" & Format(rs("apunte_id"), "000000") & "'")
                If rs2.RecordCount > 0 Then
                    Do
                        Text1 = Text1 & rs2("DESCRIPCION_APUNTE") & " -> " & rs2("APUNTE") & vbNewLine
                        If chkac.Value = Checked Then
                            oC.execute_bd "UPDATE MOVIMIENTOS SET DESCRIPCION_APUNTE = '" & DES_APUNTE_BUENO & "' WHERE APUNTE = '" & Format(rs2("APUNTE"), "000000") & "'"
                        End If
                        rs2.MoveNext
                    Loop Until rs2.EOF
                Else
                    Text1 = Text1 & "MALOOOOO" & vbNewLine
                End If
            End If
            rs.MoveNext
        Loop Until rs.EOF
    End If
    oC.DesConectar
    Set rs = Nothing
End Sub

Private Sub Command4_Click()
    Dim rs As ADODB.Recordset
    Dim C As String
    C = "select * from documentos where id_documento in (1884,1891,1892,1893,1898,1899,1905,1906,1913,1914,1915,1926,1928,1930,1937,1940,1941,1942,1944,1948,1949,1950,1951,1957,1965,1966,2015,2057,2116) order by id_documento"
    Set rs = datos_bd(C)
    Dim oC As New clsContabilidad
    oC.Conectar
    If rs.RecordCount > 0 Then
        Do
            DES_PARTE = rs("NUMERO") & "/11"
                Dim rs2 As ADODB.Recordset
                Set rs2 = oC.datos_bd_contabilidad("SELECT * FROM MOVIMIENTOS WHERE DESCRIPCION_APUNTE LIKE 'NUESTRA FACTURA Nº 000" & DES_PARTE & "%' and TIPO_APUNTE='D'")
'                Set RS2 = oC.datos_bd_contabilidad("SELECT * FROM MOVIMIENTOS WHERE APUNTE = '" & Format(rs("apunte_id"), "000000") & "'")
                If rs2.RecordCount > 0 Then
                    Do
                        Text1 = Text1 & rs2("DESCRIPCION_APUNTE") & "->" & rs2("IMPORTE_APUNTE") & " -> " & rs("TOTAL") & vbNewLine
                        If chkac.Value = Checked Then
                            oC.execute_bd "UPDATE MOVIMIENTOS SET IMPORTE_APUNTE = '" & rs("TOTAL") & "' WHERE APUNTE = '" & Format(rs2("APUNTE"), "000000") & "'"
                        End If
                        rs2.MoveNext
                    Loop Until rs2.EOF
                Else
                    Text1 = Text1 & "MALOOOOO" & vbNewLine
                End If
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    oC.DesConectar
    Set rs = Nothing

End Sub

Private Sub Form_Load()
    log (Me.Name)
'    BD_ANTIGUA = ReadINI(App.Path + "\config.ini", "server", "bd_conversion")
    BD_ANTIGUA = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ReadINI(App.Path + "\config.ini", "server", "bd_conversion")
    Text2 = ReadINI(App.Path + "\config.ini", "server", "bd_conversion")
'    Text3 = ruta
    Set conn_antigua = New ADODB.Connection
    Set rs_antiguo = New ADODB.Recordset
    Set rs_antiguo2 = New ADODB.Recordset
    conn_antigua.ConnectionString = BD_ANTIGUA
    conn_antigua.Open
End Sub
Private Function consulta_antigua(query As String)
    rs_antiguo.ActiveConnection = conn_antigua
    rs_antiguo.CursorLocation = adUseClient
    rs_antiguo.CursorType = adOpenForwardOnly
    rs_antiguo.LockType = adLockReadOnly
    rs_antiguo.Open query
End Function
Private Function consulta_antigua2(query As String)
    rs_antiguo2.ActiveConnection = conn_antigua
    rs_antiguo2.CursorLocation = adUseClient
    rs_antiguo2.CursorType = adOpenForwardOnly
    rs_antiguo2.LockType = adLockReadOnly
    rs_antiguo2.Open query
End Function

Private Sub Form_Unload(Cancel As Integer)
    conn_antigua.Close
End Sub
Public Sub usuarios_conversion()
    Dim oUsuario As New ClsUsuario
    execute_bd "DELETE FROM usuarios where id_empleado <> 1"
    consulta_antigua "select * from autorizados order by nombre"
    Dim rs As ADODB.Recordset
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Do
            With oUsuario
                .setID_EMPLEADO = 0
                .setNOMBRE = UCase(rs_antiguo("NOMBRE"))
                .setUSUARIO = UCase(rs_antiguo("NOMBRE"))
                .setPASSWORD = rs_antiguo("CLAVE")
                .setPER_1 = 1
                .setPER_2 = 1
                .setPER_3 = 1
                .setPER_4 = 1
                .setPER_5 = 1
                .setPER_6 = 1
                .setPER_7 = 1
                .Insertar
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(8).Value = Checked
End Sub

Public Sub formasPago()
    Dim oFp As New clsForma_pago
    execute_bd "DELETE FROM FORMA_PAGO"
'    execute_bd "INSERT INTO FORMA_PAGO VALUES (0,'A CONVENIR',0,0,0,0,0,'');"
    consulta_antigua "select * from [formas de pago] order by cofpa"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Dim DIRECCION As String
        Do
            With oFp
'                .setID_FORMA_PAGO = rs_antiguo("cofpa")
                .setNOMBRE = rs_antiguo("defpa")
                .Insertar
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(2).Value = Checked
End Sub
Public Sub vehiculos()
    Dim oVehiculo As New clsVehiculos
    execute_bd "DELETE FROM VEHICULOS"
    consulta_antigua "select * from [vehiculos] order by covei"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Dim DIRECCION As String
        Do
            With oVehiculo
                .setID_VEHICULO = rs_antiguo("covei")
                If IsNull(rs_antiguo("descve")) Then
                    .setNOMBRE = ""
                Else
                    .setNOMBRE = rs_antiguo("descve")
                End If
                If IsNull(rs_antiguo("matri")) Then
                    .setMATRICULA = ""
                Else
                    .setMATRICULA = rs_antiguo("matri")
                End If
                If IsNull(rs_antiguo("nifve")) Then
                    .setNIF = ""
                Else
                    .setNIF = rs_antiguo("nifve")
                End If
                If IsNull(rs_antiguo("remolque")) Then
                    .setREMOLQUE = ""
                Else
                    .setREMOLQUE = rs_antiguo("remolque")
                End If
                .Insertar
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(16).Value = Checked
End Sub
Public Sub remesas()
    Dim oRemesa As New clsRemesas
    execute_bd "DELETE FROM REMESAS"
    consulta_antigua "select * from [remesas] order by remesa"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Do
            With oRemesa
                .setID_REMESA = rs_antiguo("remesa")
                If IsDate(rs_antiguo("fecha")) Then
                    .setFECHA = Format(rs_antiguo("fecha"), "yyyy-mm-dd")
                Else
                    .setFECHA = "0000-00-00"
                End If
                If IsNull(rs_antiguo("banco")) Then
                    .setBANCO_ID = 0
                Else
                    .setBANCO_ID = rs_antiguo("banco")
                End If
                .setUSUARIO_ID = 1
                .Insertar
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(11).Value = Checked
End Sub
Public Sub efectos()
    Dim ORD As New clsRemesas_documentos
    execute_bd "DELETE FROM REMESAS_DOCUMENTOS"
    consulta_antigua "select * from [EFECT] order by clave"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Do
            With ORD
                .setID = rs_antiguo("clave")
                .setREMESA_ID = rs_antiguo("remesa")
                .setDOCUMENTO_ID = 0
                .setVENCIMIENTO = 0
                .setCLIENTE_ID = rs_antiguo("cliente")
                If IsDate(rs_antiguo("vencimiento")) Then
                    .setFECHA_VENCIMIENTO = Format(rs_antiguo("vencimiento"), "yyyy-mm-dd")
                Else
                    .setFECHA_VENCIMIENTO = "0000-00-00"
                End If
                .setDESCRIPCION = rs_antiguo("factura")
                If IsNull(rs_antiguo("importe")) Then
                    .setIMPORTE = moneda_bd("0")
                Else
                    .setIMPORTE = moneda_bd(rs_antiguo("importe"))
                End If
                .setSITUACION = rs_antiguo("situacion")
                .setCONTA = rs_antiguo("conta")
                
                .Insertar
                
                
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(12).Value = Checked
End Sub
Public Sub documentos_recibos()
    Dim oDR As New clsDocumentos_Recibos
    execute_bd "DELETE FROM DOCUMENTOS_RECIBOS"
    consulta_antigua "select * from [EFECT] WHERE CONTA = 'N' order by clave asc"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Dim factura As String
        Do
            With oDR
                ' Buscar factura asociada
                des = rs_antiguo("FACTURA")
                pos = InStr(1, des, "/")
                factura = ""
                For i = pos - 1 To 1 Step -1
                    If Not IsNumeric(Mid(des, i, 1)) Then
                        Exit For
                    End If
                    factura = Mid(des, i, 1) + factura
                Next
'                MsgBox rs_antiguo("FACTURA") & " --> " & factura
                Set rs = datos_bd("select id_documento from documentos where numero = " & factura & " and tipo_documento_id = 2")
                If rs.RecordCount > 0 Then
                    ' Actualizar numero de factura del efecto
                    execute_bd "UPDATE remesas_documentos SET VENCIMIENTO=1, DOCUMENTO_ID = " & rs(0) & " WHERE ID = " & rs_antiguo("CLAVE")
                ' Insertar documento recibo
                    .setDOCUMENTO_ID = rs(0)
                    .setFECHA = Format(rs_antiguo("vencimiento"), "yyyy-mm-dd")
                    .setIMPORTE = moneda_bd(rs_antiguo("importe") / 1.18)
                    
                    
                    .setVENCIMIENTO = .CalcularNumeroVencimiento(rs(0))
                    If rs_antiguo("CLDESCU") <> 0 Then
                        .setCOBRADO = ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_DESCUENTO
                    Else
                        .setCOBRADO = ENUM_EFECTOS_ESTADOS.EFECTOS_ESTADOS_REMESA
                    End If
                    .setCONTABILIZADO = 0
                    .Insertar
                    
     '               execute_bd "DELETE FROM DOCUMENTOS_COBROS WHERE DOCUMENTO_ID = " & rs(0)
                End If
                
                
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(19).Value = Checked
End Sub

Public Sub descuentos()
    Dim oDTO As New clsDescuentos
    execute_bd "DELETE FROM DESCUENTOS"
    execute_bd "DELETE FROM DESCUENTOS_DOCUMENTOS"
    
    consulta_antigua "SELECT DISTINCT CLDESCU,FDESCU,BANCO,CONDE FROM EFECT;"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    Dim oDD As New clsDescuentos_documentos
    
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Do
            With oDTO
                If IsNumeric(rs_antiguo(0)) Then
                    .setID_DESCUENTO = rs_antiguo(0)
                    If IsDate(rs_antiguo(1)) Then
                        .setFECHA = Format(rs_antiguo(1), "yyyy-mm-dd")
                    Else
                        .setFECHA = "0000-00-00"
                    End If
                    .setBANCO_ID = rs_antiguo(2)
                    .setUSUARIO_ID = 1
                    .Insertar
                    
                    If rs_antiguo(3) = "S" Then
                        .Contabilizar rs_antiguo(0)
                    End If
                    ' APUNTES
                    consulta_antigua2 "SELECT CLAVE FROM EFECT WHERE CLDESCU = '" & rs_antiguo(0) & "' ORDER BY CLAVE"
                    If rs_antiguo2.RecordCount > 0 Then
                        Do
                            oDD.setDESCUENTO_ID = rs_antiguo(0)
                            oDD.setAPUNTE_ID = rs_antiguo2(0)
                            oDD.Insertar
                            rs_antiguo2.MoveNext
                        Loop Until rs_antiguo2.EOF
                    End If
                    rs_antiguo2.Close
                End If
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(17).Value = Checked
End Sub
Private Sub liquidaciones()
    execute_bd "DELETE FROM LIQUIDACION"
    execute_bd "DELETE FROM LIQUIDACION_DOCUMENTOS"
    Dim C As String
    C = "select agente,fechaliq,min(fechacobro),max(fechacobro) " & _
        " from [diario facturas] " & _
        " where fechaliq <> 0-0-0 AND ANNO='11'" & _
        " group by agente,fechaliq"

    consulta_antigua C
    Dim rs As ADODB.Recordset
    Dim oL As New clsLiquidacion
    Dim oLD As New clsLiquidacion_documentos
    Dim liquidacion As Long
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Do
            With oL
                If Not IsNull(rs_antiguo(0)) Then
                    .setDESCRIPCION = "LIQUIDACION FECHA : " & Format(rs_antiguo(1), "dd-mm-yyyy")
                    .setFDESDE = Format(rs_antiguo(2), "yyyy-mm-dd")
                    .setFHASTA = Format(rs_antiguo(3), "yyyy-mm-dd")
                    .setFLIQUIDACION = Format(rs_antiguo(1), "yyyy-mm-dd")
                    .setAGENTE_ID = rs_antiguo(0)
                    .setUSUARIO_ID = 1
                    liquidacion = .Insertar
                    ' Documentos
                    C = "select nfactura,comision " & _
                        " from [diario facturas] " & _
                        " where fechaliq = #" & Format(rs_antiguo(1), "mm-dd-yyyy") & "#" & _
                        "   and agente = '" & rs_antiguo(0) & "'" & _
                        "   AND ANNO='11'"
                        
                    consulta_antigua2 C
                    If rs_antiguo2.RecordCount > 0 Then
                        Do
                            Set rs = datos_bd("select * from documentos where numero = " & rs_antiguo2(0) & " and tipo_documento_id = " & ENUM_TIPOS_DOCUMENTOS.factura)
                            If rs.RecordCount > 0 Then
                            
                                oLD.setDOCUMENTO_ID = rs("ID_DOCUMENTO")
                                oLD.setLIQUIDACION_ID = liquidacion
                                oLD.setCOMISION = moneda_bd(rs_antiguo2("comision"))
                                oLD.Insertar
                            Else
                                
                                Text1.Text = Text1.Text & vbNewLine & " No existe la factura para liquidacion. Numero : " & rs_antiguo(2)
                                
                            End If
                            rs_antiguo2.MoveNext
                        Loop Until rs_antiguo2.EOF
                    End If
                    rs_antiguo2.Close
                End If
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(18).Value = Checked
End Sub

Public Sub obras()
    Dim oObra As New clsObras
    execute_bd "DELETE FROM OBRAS"
    consulta_antigua "select * from obras order by cobra"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    Dim oProv As New clsProvincias
    Dim oMun As New clsMunicipios
    
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Dim DIRECCION As String
        Do
            With oObra
                .setID_OBRA = rs_antiguo("cobra")
                ' Cliente
                .setCLIENTE_ID = 0
                If Not IsNull(rs_antiguo("cocli")) Then
                    If Trim(rs_antiguo("cocli")) <> "" Then
                        C = "SELECT * FROM CLIENTES WHERE ID_CLIENTE =" & rs_antiguo("COCLI")
                        Set rs = datos_bd(C)
                        If rs.RecordCount > 0 Then
                            .setCLIENTE_ID = rs("ID_CLIENTE")
                        Else
                            Text1 = Text1 & vbNewLine & "No encuento cliente. Obra : " & rs_antiguo("OBRA")
                        End If
                    End If
                End If
                If IsNull(rs_antiguo("obra")) Then
                    .setNOMBRE = ""
                Else
                    .setNOMBRE = UCase(Replace(rs_antiguo("obra"), "'", "`"))
                End If
                DIRECCION = ""
                If Not IsNull(rs_antiguo("destino factura")) Then
                    DIRECCION = Trim(rs_antiguo("destino factura")) & " "
                End If
                If Not IsNull(rs_antiguo("ndt")) Then
                    DIRECCION = DIRECCION & Trim(rs_antiguo("ndt")) & " "
                End If
                .setDIRECCION = Trim(DIRECCION)
                
                If rs_antiguo("cpdes") = "" Then
                    .setCP = 0
                Else
                    If IsNumeric(rs_antiguo("cpdes")) Then
                        .setCP = rs_antiguo("cpdes")
                    Else
                        .setCP = 0
                    End If
                End If
                ' PROVINCIA
                Dim PROVINCIA As Integer
                PROVINCIA = 0
                If IsNull(rs_antiguo("PROVDES")) Then
                    .setPROVINCIA_ID = 0
                Else
                    If Trim(rs_antiguo("PROVDES")) = "" Then
                        .setPROVINCIA_ID = 0
                    Else
                        C = "SELECT * FROM PROVINCIAS WHERE NOMBRE LIKE '%" & rs_antiguo("PROVDES") & "%'"
                        Set rs = datos_bd(C)
                        If rs.RecordCount > 0 Then
                            .setPROVINCIA_ID = rs("ID_PROVINCIA")
                        Else
                            oProv.setNOMBRE = rs_antiguo("PROVDES")
                            .setPROVINCIA_ID = oProv.Insertar
                        End If
                        PROVINCIA = .getPROVINCIA_ID
                    End If
                End If
                ' MUNICIPIO
                If IsNull(rs_antiguo("POBLADES")) Then
                    .setMUNICIPIO_ID = 0
                Else
                    If Trim(rs_antiguo("POBLADES")) = "" Then
                        .setMUNICIPIO_ID = 0
                    Else
                        C = "SELECT * FROM MUNICIPIOS WHERE NOMBRE LIKE '%" & rs_antiguo("POBLADES") & "%'"
                        Set rs = datos_bd(C)
                        If rs.RecordCount > 0 Then
                            .setMUNICIPIO_ID = rs("ID_MUNICIPIO")
                        Else
                            oMun.setPROVINCIA_ID = PROVINCIA
                            oMun.setNOMBRE = rs_antiguo("POBLADES")
                            .setMUNICIPIO_ID = oMun.Insertar
                        End If
                    End If
                End If
                If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                    .setTELEFONO = ""
                Else
                    If IsNull(rs_antiguo("teleobra")) Then
                        .setTELEFONO = ""
                    Else
                        .setTELEFONO = Trim(rs_antiguo("teleobra"))
                    End If
                End If
                If IsNull(rs_antiguo("cfp")) Then
                    .setFORMA_PAGO_ID = ""
                Else
                    .setFORMA_PAGO_ID = rs_antiguo("cfp")
                End If
                If IsNull(rs_antiguo("contra")) Then
                    .setCONTRAPARTIDA = 0
                Else
                    If IsNumeric(rs_antiguo("contra")) Then
                        .setCONTRAPARTIDA = rs_antiguo("contra")
                    Else
                        .setCONTRAPARTIDA = 0
                    End If
                End If
                If IsNull(rs_antiguo("email")) Then
                    .setEMAIL = ""
                Else
                    .setEMAIL = Trim(rs_antiguo("email"))
                End If
                If IsNull(rs_antiguo("banco")) Then
                    .setBANCO = ""
                Else
                    .setBANCO = Trim(rs_antiguo("banco"))
                End If
                If IsNull(rs_antiguo("ccbanco")) Then
                    .setCCC = "____-____-__-__________"
                Else
                    If Trim(rs_antiguo("ccbanco")) = "" Then
                        .setCCC = "____-____-__-__________"
                    Else
                        .setCCC = Format(Trim(rs_antiguo("ccbanco")), "0000-0000-00-0000000000")
                    End If
                End If
                If IsNull(rs_antiguo("dirban")) Then
                    .setBANCO_DIRECCION = ""
                Else
                    .setBANCO_DIRECCION = Trim(rs_antiguo("dirban"))
                End If
                .setFECHA_OFERTA = "0000-00-00"
                If Not IsNull(rs_antiguo("FOFERTA")) Then
                    If IsDate(rs_antiguo("FOFERTA")) Then
                        .setFECHA_OFERTA = Format(rs_antiguo("FOFERTA"), "YYYY-MM-DD")
                    End If
                End If
                .setTARIFA_PORTE_ID = CInt(rs_antiguo("tarpor"))
                If rs_antiguo("tifac") = "S" Then
                    .setTIPO_FACTURACION = 1
                Else
                    .setTIPO_FACTURACION = 2
                End If
                ' ALMACEN/OBRA
                If UCase(rs_antiguo("alob")) = "A" Then
                    .setTIPO_OBRA_ID = 1
                Else
                    .setTIPO_OBRA_ID = 2
                End If
                ' tipo de impresion
                If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                    .setTIPO_IMPORTE_ID = 1
                Else
                    If IsNull(rs_antiguo("ti")) Then
                        .setTIPO_IMPORTE_ID = ""
                    Else
                        .setTIPO_IMPORTE_ID = rs_antiguo("ti")
                    End If
                End If
                ' AGENTE
                .setCOMERCIAL_ID = rs_antiguo("AGENTE")
                .setTIPO_IVA = rs_antiguo("TIVA")
                If frmMenu.StatusBar1.Panels(3) = "Server: " & IP_RESPALDO Then
                    .setLIBRO = 0
                Else
                    .setLIBRO = rs_antiguo("LIBRO")
                End If
                .setDESCUENTO = moneda_bd(rs_antiguo("DTO"))

'                execute_bd "UPDATE CLIENTES SET COMERCIAL_ID = " & rs_antiguo("AGENTE") & " WHERE ID_CLIENTE = " & rs_antiguo("COCLI")
                .Insertar
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(3).Value = Checked
End Sub
Public Sub obras_destino_facturas()
    Dim oObra As New clsObras
    consulta_antigua "select * from obras order by cobra"
    Dim rs As ADODB.Recordset
    
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Dim RF As String
        Dim DF As String
        Do
            RF = ""
            DF = ""
            If Not IsNull(rs_antiguo("reffa")) Then
                RF = rs_antiguo("REFFA")
            End If
            If Not IsNull(rs_antiguo("DESTINO FACTURA")) Then
                DF = rs_antiguo("DESTINO FACTURA")
            End If
            If Not IsNull(rs_antiguo("NDT")) Then
                DF = DF + " " + rs_antiguo("NDT")
            End If
            execute_bd "UPDATE OBRAS SET REFERENCIA_FACTURA = '" & RF & "', DESTINO_FACTURA = '" & DF & "' WHERE ID_OBRA = " & rs_antiguo("COBRA")
            If pb.Max > pb.Value Then
                pb.Value = pb.Value + 1
            End If
            lbll = pb.Value & " de " & pb.Max
            DoEvents
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(3).Value = Checked
End Sub

Public Sub tarifasPortes()
    Dim oTP As New clsTarifas_portes
    Dim oTPA As New clsTarifas_portes_articulos
    execute_bd "DELETE FROM TARIFAS_PORTES"
    execute_bd "DELETE FROM TARIFAS_PORTES_ARTICULOS"
    consulta_antigua "select * from [tarifas portes] order by ctapor"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Dim DIRECCION As String
        Do
            C = "select * from tarifas_portes where id_tarifa_porte = " & Left(rs_antiguo("ctapor"), 2)
            Set rs = datos_bd(C)
            If rs.RecordCount = 0 Then
                With oTP
                    .setID_TARIFA_PORTE = Left(rs_antiguo("ctapor"), 2)
                    .setDESCRIPCION = "TARIFA PORTE " & Left(rs_antiguo("ctapor"), 2)
                    .Insertar
                End With
            End If
            With oTPA
                .setTARIFA_PORTE_ID = Left(rs_antiguo("ctapor"), 2)
                .setARTICULO_ID = Right(rs_antiguo("ctapor"), 3)
                .setPRECIO = moneda_bd(Trim(rs_antiguo("dtapor")))
                .Insertar
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(6).Value = Checked
End Sub
Public Sub tarifas()
    Dim oTAR As New clsTarifas
    execute_bd "DELETE FROM TARIFAS"
    consulta_antigua "select * from tarifas order by ctari"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Do
            With oTAR
                If Not IsNull(rs_antiguo("obra")) Then
                    .setOBRA_ID = rs_antiguo("obra")
                    If Not IsNull(rs_antiguo("arti")) Then
                        .setARTICULO_ID = rs_antiguo("arti")
                    End If
                    If IsNull(rs_antiguo("tfabr")) Then
                        .setPRECIO_FABRICA = moneda_bd("0")
                    Else
                        .setPRECIO_FABRICA = moneda_bd(rs_antiguo("tfabr"))
                    End If
                    If IsNull(rs_antiguo("tobra")) Then
                        .setPRECIO_OBRA = moneda_bd("0")
                    Else
                        .setPRECIO_OBRA = moneda_bd(rs_antiguo("tobra"))
                    End If
                    .Insertar
                End If
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(7).Value = Checked
End Sub
Public Sub ofertas()
    Dim oO As New clsOfertas
    Dim oOD As New clsOfertas_detalle
    execute_bd "DELETE FROM OFERTAS"
    execute_bd "DELETE FROM OFERTAS_DETALLE"
    consulta_antigua "select * from [diario OFERTASn] order by noferta"
    Dim Cuenta As String
    Dim rs As ADODB.Recordset
    Dim ID As Long
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Do
            ' Oferta
            With oO
                .setID_OFERTA = 0
                .setNUMERO = Right(rs_antiguo("noferta"), 5)
                .setFECHA = Format(rs_antiguo("foferta"), "yyyy-mm-dd")
                If IsNull(rs_antiguo("clien")) Then
                    ' Buscar Cliente
                    .setCLIENTE_ID = 0
                Else
                    If rs_antiguo("clien") = "" Then
                        .setCLIENTE_ID = 0
                    Else
                        .setCLIENTE_ID = rs_antiguo("clien")
                    End If
                End If
                ' Cliente
                If IsNull(rs_antiguo("nif")) Then
                    .setNIF = ""
                Else
                    .setNIF = Trim(rs_antiguo("nif"))
                End If
                If IsNull(rs_antiguo("nombre")) Then
                    .setNOMBRE = ""
                Else
                    .setNOMBRE = Trim(rs_antiguo("nombre"))
                End If
                If IsNull(rs_antiguo("direccion")) Then
                    .setDIRECCION = ""
                Else
                    .setDIRECCION = Trim(rs_antiguo("direccion"))
                End If
                If IsNull(rs_antiguo("poblacion")) Then
                    .setPOBLACION = ""
                Else
                    .setPOBLACION = Trim(rs_antiguo("poblacion"))
                End If
                
                If IsNull(rs_antiguo("contacto")) Then
                    .setAA = ""
                Else
                    .setAA = Trim(rs_antiguo("contacto"))
                End If
                If IsNull(rs_antiguo("telefono")) Then
                    .setTELEFONO = ""
                Else
                    .setTELEFONO = Trim(rs_antiguo("telefono"))
                End If
                If IsNull(rs_antiguo("fax")) Then
                    .setFAX = ""
                Else
                    .setFAX = Trim(rs_antiguo("fax"))
                End If
                If IsNull(rs_antiguo("email")) Then
                    .setEMAIL = ""
                Else
                    .setEMAIL = Trim(rs_antiguo("email"))
                End If
                If IsNull(rs_antiguo("obra")) Then
                    .setOBRA_DOMICILIO = ""
                Else
                    .setOBRA_DOMICILIO = Trim(rs_antiguo("obra"))
                End If
                If IsNull(rs_antiguo("pobra")) Then
                    .setOBRA_POBLACION = ""
                Else
                    .setOBRA_POBLACION = Trim(rs_antiguo("pobra"))
                End If
                If IsNull(rs_antiguo("agente")) Then
                    .setAGENTE_ID = 0
                Else
                    If rs_antiguo("agente") = "" Then
                        .setAGENTE_ID = 0
                    Else
                        .setAGENTE_ID = rs_antiguo("agente")
                    End If
                End If
                .setHORA = "00:00"
            
                ' Forma de pago
                If IsNull(rs_antiguo("fp")) Then
                    .setFORMA_PAGO = ""
                Else
                    .setFORMA_PAGO = rs_antiguo("fp")
                End If
                ID = .Insertar
                If ID > 0 Then
                        ' DETALLE
                        oOD.setOFERTA_ID = ID
                        
                        Dim j As Integer
                        For j = 1 To 30
                            oOD.setORDEN = j
'                            MsgBox Format(j, "00") & "C"
                            If Not IsNull(rs_antiguo(Format(j, "00") & "C")) Then
                                oOD.setMATERIAL_ID = Right(rs_antiguo(Format(j, "00") & "C"), 3)
                                If IsNull(rs_antiguo(Format(j, "00") & "F")) Then
                                    oOD.setPRECIO_FABRICA = moneda_bd("0")
                                Else
                                    oOD.setPRECIO_FABRICA = moneda_bd(rs_antiguo(Format(j, "00") & "F"))
                                End If
                                If IsNull(rs_antiguo(Format(j, "00") & "O")) Then
                                    oOD.setPRECIO_OBRA = moneda_bd("0")
                                Else
                                    oOD.setPRECIO_OBRA = moneda_bd(rs_antiguo(Format(j, "00") & "O"))
                                End If
                                oOD.Insertar
                            End If
                        Next
                        
'                        ' RAS25
'                        oOD.setORDEN = 1
'                        oOD.setMATERIAL_ID = 1
'                        If IsNull(rs_antiguo("ras25f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("ras25f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("ras25o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("ras25o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' TA4
'                        oOD.setORDEN = 2
'                        oOD.setMATERIAL_ID = 2
'                        If IsNull(rs_antiguo("ta4f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("ta4f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("ta4o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("ta4o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' DH5
'                        oOD.setORDEN = 3
'                        oOD.setMATERIAL_ID = 3
'                        If IsNull(rs_antiguo("dh5f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("dh5f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("dh5o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("dh5o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' DH7
'                        oOD.setORDEN = 4
'                        oOD.setMATERIAL_ID = 4
'                        If IsNull(rs_antiguo("dh7f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("dh7f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("dh7o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("dh7o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' DH8
'                        oOD.setORDEN = 5
'                        oOD.setMATERIAL_ID = 5
'                        If IsNull(rs_antiguo("dh8f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("dh8f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("dh8o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("dh8o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' DH8R
'                        oOD.setORDEN = 6
'                        oOD.setMATERIAL_ID = 6
'                        If IsNull(rs_antiguo("dh8rf")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("dh8rf"), 30))
'                        End If
'                        If IsNull(rs_antiguo("dh8ro")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("dh8ro"), 30))
'                        End If
'                        oOD.Insertar
'                        ' DH9
'                        oOD.setORDEN = 7
'                        oOD.setMATERIAL_ID = 7
'                        If IsNull(rs_antiguo("dh9f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("dh9f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("dh9o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("dh9o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' TH10
'                        oOD.setORDEN = 8
'                        oOD.setMATERIAL_ID = 8
'                        If IsNull(rs_antiguo("th10f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("th10f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("th10o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("th10o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' TP7
'                        oOD.setORDEN = 9
'                        oOD.setMATERIAL_ID = 9
'                        If IsNull(rs_antiguo("tp7f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("tp7f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("tp7o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("tp7o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' TP9
'                        oOD.setORDEN = 10
'                        oOD.setMATERIAL_ID = 10
'                        If IsNull(rs_antiguo("tp9f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("tp9f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("tp9o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("tp9o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' TP10
'                        oOD.setORDEN = 11
'                        oOD.setMATERIAL_ID = 11
'                        If IsNull(rs_antiguo("tp10f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("tp10f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("tp10o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("tp10o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' A3T
'                        oOD.setORDEN = 12
'                        oOD.setMATERIAL_ID = 12
'                        If IsNull(rs_antiguo("a3tf")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("a3tf"), 30))
'                        End If
'                        If IsNull(rs_antiguo("a3to")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("a3to"), 30))
'                        End If
'                        oOD.Insertar
'                        ' FA7
'                        oOD.setORDEN = 13
'                        oOD.setMATERIAL_ID = 13
'                        If IsNull(rs_antiguo("fa7f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("fa7f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("fa7o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("fa7o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' FA10
'                        oOD.setORDEN = 14
'                        oOD.setMATERIAL_ID = 14
'                        If IsNull(rs_antiguo("fa10f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("fa10f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("fa10o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("fa10o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' DH530
'                        oOD.setORDEN = 15
'                        oOD.setMATERIAL_ID = 15
'                        If IsNull(rs_antiguo("dh530f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("dh530f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("dh530o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("dh530o"), 30))
'                        End If
'                        oOD.Insertar
'                        ' DH730
'                        oOD.setORDEN = 16
'                        oOD.setMATERIAL_ID = 16
'                        If IsNull(rs_antiguo("dh730f")) Then
'                            oOD.setPRECIO_FABRICA = ""
'                        Else
'                            oOD.setPRECIO_FABRICA = Trim(Left(rs_antiguo("dh730f"), 30))
'                        End If
'                        If IsNull(rs_antiguo("dh730o")) Then
'                            oOD.setPRECIO_OBRA = ""
'                        Else
'                            oOD.setPRECIO_OBRA = Trim(Left(rs_antiguo("dh730o"), 30))
'                        End If
'                        oOD.Insertar
                Else
                    Text1 = Text1 & "ERROR AL INSERTAR LA OFERTA"
                End If
                If pb.Max > pb.Value Then
                    pb.Value = pb.Value + 1
                End If
                lbll = pb.Value & " de " & pb.Max
                DoEvents
            End With
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(6).Value = Checked
End Sub

Public Sub obras_telefonos()
    Dim oObra As New clsObras
    consulta_antigua "select * from obras order by cobra"
    Dim rs As ADODB.Recordset
    
    If rs_antiguo.RecordCount > 0 Then
        pb.Min = 0
        pb.Max = rs_antiguo.RecordCount
        pb.Value = 0
        Dim RF As String
        Dim DF As String
        Do
            If Not IsNull(rs_antiguo("telefonos")) Then
                If Trim(rs_antiguo("telefonos")) <> "" Then
                    Text1 = Text1 & vbNewLine & "UPDATE OBRAS SET TELEFONO = '" & Trim(rs_antiguo("TELEFONOS")) & "' WHERE ID_OBRA = " & rs_antiguo("COBRA") & ";"
                End If
            End If
            If pb.Max > pb.Value Then
                pb.Value = pb.Value + 1
            End If
            lbll = pb.Value & " de " & pb.Max
            DoEvents
            rs_antiguo.MoveNext
        Loop Until rs_antiguo.EOF
    End If
    rs_antiguo.Close
    chkt(3).Value = Checked
End Sub

Public Sub documentos_servido()
    Dim rs As ADODB.Recordset
    Dim rs2 As ADODB.Recordset
    Set rs = datos_bd("SELECT DISTINCT A.ID_DOCUMENTO, B.NUMERO_ALBARAN FROM DOCUMENTOS A, DOCUMENTOS_DETALLE B WHERE A.TIPO_DOCUMENTO_ID = 2 AND A.ID_DOCUMENTO = B.DOCUMENTO_ID ORDER BY A.ID_DOCUMENTO")
    If rs.RecordCount > 0 Then
        Do
            C = "SELECT * FROM DOCUMENTOS WHERE NUMERO = " & rs(1) & " AND TIPO_DOCUMENTO_ID = 1 "
            Set rs2 = datos_bd(C)
            If rs2.RecordCount > 0 Then
                execute_bd "UPDATE DOCUMENTOS_DETALLE SET SERVIDO = '" & rs2("SERVIDO") & "' WHERE DOCUMENTO_ID = " & rs(0) & " AND NUMERO_ALBARAN = " & rs(1)
            End If
            
            rs.MoveNext
        Loop Until rs.EOF
    End If
    Set rs = Nothing
    MsgBox "ok"
    
End Sub


