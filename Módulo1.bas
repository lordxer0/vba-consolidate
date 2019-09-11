Attribute VB_Name = "Módulo1"
Sub consolidadoporCC()
  
  Dim hojaDeArany As Worksheet
    Set hojaDeArany = Sheets("aranysport")
    
    Dim hojaDeTaller As Worksheet
    Set hojaDeTaller = Sheets("areadetrabajo")
    'definimos la hoja que se evaluara
    
    Dim hojaDeOperaciones As Worksheet
    Set hojaDeOperaciones = Sheets("operaciones")
    'definimos la hoja donde se haran los calculos

    Dim cuentasAEvaluar As Range
    Set cuentasAEvaluar = Sheets("base").Range("G:G")
    'en la hoja de base es donde se tienen los valores inicales

    Dim CCAEvaluar As Range
    Set CCAEvaluar = Sheets("base").Range("I:I")
    'en la hoja de base es donde se tienen los valores inicales


'estos los valores de esa cuenta en todo el documento
    
    Dim subTotalDebito As Long
    Dim subTotalCredito As Long
    Dim subTotalSaldo As Long

    Dim cuentaEvaluada As String
    
    Dim limpieza As Range
    Set limpieza = Sheets("operaciones").Range("p1")
    
    Dim ccEvaluado As String
    ccEvaluado = CCAEvaluar.Cells((Sheets("base").Range("M2").Value)).Value
    'aqui tomamos el valor correspondiente a el cc desde el combobox
    
    ' creamos lo hoja de el centro de costo asociado
    If ccEvaluado <> "" Then
        On Error Resume Next
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = ccEvaluado
    End If
    
    
    For i = 1 To cuentasAEvaluar.Application.WorksheetFunction.CountA(cuentasAEvaluar)
        
        'limpiamos la hoja para evitar errores
        hojaDeOperaciones.Cells.Clear
        
        'vamos a recorer el arreglo de cuentas y hacer un filtrado por cuenta para sacar la sumatoria de valores
       
        cuentaEvaluada = cuentasAEvaluar.Cells(i, 1)
        
            hojaDeArany.UsedRange.AutoFilter 4, cuentaEvaluada
            hojaDeArany.UsedRange.AutoFilter 5, ccEvaluado
            
            hojaDeArany.UsedRange.Copy hojaDeOperaciones.Range("A1")
            
            hojaDeArany.AutoFilterMode = False
            
        'como los datos estan en texto los comvertimos a numero limpiando los valores de caracteres no imprimibles
            
            For h = 2 To Application.WorksheetFunction.CountA(hojaDeOperaciones.Range("D:D"))
            
                If hojaDeOperaciones.Cells(h, 11) <> "" Then
                    limpieza = hojaDeOperaciones.Cells(h, 11)
                    hojaDeOperaciones.Cells(h, 11) = Application.WorksheetFunction.Clean(limpieza)
                End If
                If hojaDeOperaciones.Cells(h, 12) <> "" Then
                    limpieza = hojaDeOperaciones.Cells(h, 12)
                    hojaDeOperaciones.Cells(h, 12) = Application.WorksheetFunction.Clean(limpieza)
                End If
                If hojaDeOperaciones.Cells(h, 13) <> "" Then
                    limpieza = hojaDeOperaciones.Cells(h, 13)
                    hojaDeOperaciones.Cells(h, 13) = Application.WorksheetFunction.Clean(limpieza)
                End If
                
            Next h
            
            'con los datos en un atabla aparte entonces vamos a hacer las sumas correspondientes para los rangos
            
            subTotalDebito = Application.WorksheetFunction.Sum(hojaDeOperaciones.Range("K2:K2000"))
            
            subTotalCredito = Application.WorksheetFunction.Sum(hojaDeOperaciones.Range("L2:L2000"))
            
            subTotalSaldo = Application.WorksheetFunction.Sum(hojaDeOperaciones.Range("M2:M2000"))
        
        
        '-----------------------------------------------------------------------------
        'ahora hacemos consolidado de la hoja de taller
        
         'limpiamos la hoja para evitar errores
         hojaDeOperaciones.Cells.Clear
         
         'vamos a recorer el arreglo de cuentas y hacer un filtrado por cuenta para sacar la sumatoria de valores
        
         cuentaEvaluada = cuentasAEvaluar.Cells(i, 1)
         
             hojaDeTaller.UsedRange.AutoFilter 4, cuentaEvaluada
             
             hojaDeTaller.UsedRange.Copy hojaDeOperaciones.Range("A1")
             
             hojaDeTaller.AutoFilterMode = False
             
         'como los datos estan en texto los comvertimos a numero limpiando los valores de caracteres no imprimibles
             
                For h = 2 To Application.WorksheetFunction.CountA(hojaDeOperaciones.Range("D:D"))
             
                 If hojaDeOperaciones.Cells(h, 11) <> "" Then
                     limpieza = hojaDeOperaciones.Cells(h, 11)
                     hojaDeOperaciones.Cells(h, 11) = Application.WorksheetFunction.Clean(limpieza)
                 End If
                 If hojaDeOperaciones.Cells(h, 12) <> "" Then
                     limpieza = hojaDeOperaciones.Cells(h, 12)
                     hojaDeOperaciones.Cells(h, 12) = Application.WorksheetFunction.Clean(limpieza)
                 End If
                 If hojaDeOperaciones.Cells(h, 13) <> "" Then
                     limpieza = hojaDeOperaciones.Cells(h, 13)
                     hojaDeOperaciones.Cells(h, 13) = Application.WorksheetFunction.Clean(limpieza)
                 End If
                 
                Next h
             
             'con los datos en un atabla aparte entonces vamos a hacer las sumas correspondientes para los rangos
             
             subTotalDebito = Application.WorksheetFunction.Sum(hojaDeOperaciones.Range("K2:K2000"))
             
             subTotalCredito = Application.WorksheetFunction.Sum(hojaDeOperaciones.Range("L2:L2000"))
             
             subTotalSaldo = Application.WorksheetFunction.Sum(hojaDeOperaciones.Range("M2:M2000"))
             
        '---------------------------------------------------------------------
         
         'ahora ya con las sumas lo organizamos en la tabla general para sacar el reporte general de datos
         
             Sheets(ccEvaluado).Cells(i + 1, 1) = cuentaEvaluada
             
             If subTotalDebito = 0 Then
                Sheets(ccEvaluado).Cells(i + 1, 3) = ""
             Else
                Sheets(ccEvaluado).Cells(i + 1, 3) = subTotalDebito
             End If
             
             If subTotalCredito = 0 Then
                Sheets(ccEvaluado).Cells(i + 1, 4) = ""
             Else
                Sheets(ccEvaluado).Cells(i + 1, 4) = subTotalCredito
             End If
             
             If subTotalSaldo = 0 Then
                Sheets(ccEvaluado).Cells(i + 1, 5) = ""
             Else
                Sheets(ccEvaluado).Cells(i + 1, 5) = subTotalSaldo
             End If
             
             
             
            
    Next i

End Sub

