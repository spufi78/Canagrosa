VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMSOExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Public Function ActualizarGraficasExcel(ByRef prmWB As Excel.Workbook, ByVal prmFilaIni As Integer, ByVal prmFilaFin As Integer) As Excel.Workbook

    ' Los nombres de las hojas
    Dim Hoja1 As String, Hoja2 As String
    Dim oSerie As Excel.Series, oChart As Excel.ChartObject
    Dim oSheet As Excel.Worksheet
    Dim lst_Chart As String, lst_Serie As String
    
    lst_Chart = ""
    lst_Serie = ""
    
    ' Guarda los nombres de las hojas
    Hoja1 = prmWB.Worksheets(1).Name
    Hoja2 = prmWB.Worksheets(2).Name
    Set oSheet = prmWB.Worksheets(2)
    
    
    On Error GoTo ActualizarGraficasExcel_Error

            contChart = 1
            Do While contChart > 0
                On Error Resume Next
                Set oChart = oSheet.ChartObjects(contChart)
                
                If lst_Chart = oChart.Name Then
                    If Err.Number <> 0 Then contSerie = 0
                    On Error GoTo ActualizarGraficasExcel_Error
                
                    contChart = 0 ' para salir del bucle
                Else
                    On Error GoTo ActualizarGraficasExcel_Error
                    ' trabajamos con el Chart
                    ' Guardamos el nombre como último tratado
                    lst_Chart = oChart.Name '(Tip para Excel, por obligacion hay que hacerlo así)
                    
                    contSerie = 1
                    Do While contSerie > 0
                        ' Recogemos la Serie
                        On Error Resume Next
                        Set oSerie = oChart.Chart.SeriesCollection(contSerie)
                        
                        
                        If lst_Serie = oSerie.Name Then
                            If Err.Number <> 0 Then contSerie = 0
                            On Error GoTo ActualizarGraficasExcel_Error
                            contSerie = 0 ' para salir del bucle
                        Else
                            On Error GoTo ActualizarGraficasExcel_Error
                            ' trabajamos con la serie
                            ' guardamos el nombre de la serie como última con la que hemos tratado (Tip para Excel, por obligacion hay que hacerlo así)
                            lst_Serie = oSerie.Name
                            
                            oSerie.Formula = modificar_formula(oSerie.Formula, prmFilaIni, prmFilaFin, Hoja1, Hoja2)
                            contSerie = contSerie + 1
                        End If
                    Loop
                    contChart = contChart + 1
                End If
            Loop
    
    Set ActualizarGraficasExcel = prmWB

Exit Function

ActualizarGraficasExcel_Error:
    Log "Se ha producido un error al actualizar una grafica excel: Estos son los datos: " & vbCrLf & Err.Number & " " & Err.Description & vbCrLf & "Libro: " & oSheet.Parent.Name & "(Reportado por : " & USUARIO.getUSUARIO & ")"
'    Call Enviar_Mail_CDO("informatica@canagrosa", "[ERROR] Actualicion Gráficas Excel", "Se ha producido un error al actualizar una grafica excel: Estos son los datos: " & vbCrLf & Err.Number & " " & Err.Description & vbCrLf & "Libro: " & oSheet.Parent.Name & "(Reportado por : " & USUARIO.getUSUARIO & ")", vbNullString)
    Set ActualizarGraficasExcel = Nothing
    
End Function


Private Function modificar_formula(ByRef prmFormula As String, ByVal prmFilaIni As Integer, ByVal prmFilaFin As Integer, ByVal Hoja1 As String, ByVal Hoja2 As String) As String

    
    Dim valores As String
    Dim fechas As String
    Dim parte1 As String, parte4 As String
    Dim strCad As String
    Dim pA(1 To 2) As String
    Dim pB(1 To 2) As String
    Dim col(1 To 2) As String
    
    
    modificar_formula = prmFormula
    
    
    ' Referencia de la formula
    ' =SERIE("TURCO  NCLT",Hoja1!$B$51:$B$92,Hoja1!$C$51:$C$92,1)
    
    If InStr(1, prmFormula, Hoja2) > 0 Then
        Exit Function
    End If
    
    ' Comienza a Desglosar
    
    parte1 = Split(prmFormula, ",")(0)
    fechas = Split(prmFormula, ",")(1)
    valores = Split(prmFormula, ",")(2)
    parte4 = Split(prmFormula, ",")(3)
    
    ' Reestablece el Rango para FECHAS (Eje X)
    strCad = fechas
    pA(1) = Split(strCad, "!")(0)
    pA(2) = Split(strCad, "!")(1)
    pB(1) = Split(pA(2), ":")(0)
    pB(2) = Split(pA(2), ":")(1)
    col(1) = Split(pB(1), "$")(1)
    col(2) = Split(pB(2), "$")(1)
    strCad = pA(1) & "!$" & col(1) & "$" & CStr(prmFilaIni) & ":$" & col(2) & "$" & CStr(prmFilaFin)
    fechas = strCad
    
    ' Reestablece el Rango para VALORES
    strCad = valores
    pA(1) = Split(strCad, "!")(0)
    pA(2) = Split(strCad, "!")(1)
    pB(1) = Split(pA(2), ":")(0)
    pB(2) = Split(pA(2), ":")(1)
    col(1) = Split(pB(1), "$")(1)
    col(2) = Split(pB(2), "$")(1)
    strCad = pA(1) & "!$" & col(1) & "$" & CStr(prmFilaIni) & ":$" & col(2) & "$" & CStr(prmFilaFin)
    valores = strCad
    
    ' Composicion final
    strCad = parte1 & "," & fechas & "," & valores & "," & parte4
    
    modificar_formula = strCad
End Function


