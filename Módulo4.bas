Attribute VB_Name = "Módulo4"
Sub consolidadoporCCompleto()
  
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
    Dim NombreCuentaEvaluada As String
    
    Dim limpieza As Range
    Set limpieza = Sheets("operaciones").Range("p1")
    
    Dim ccEvaluado As String
    ccEvaluado = Sheets("base").Range("M4").Value
    'aqui tomamos el valor correspondiente a el cc desde el combobox
    
    ' creamos lo hoja de el centro de costo asociado
    If ccEvaluado <> "" Then
        On Error Resume Next
        Worksheets.Add(After:=Worksheets(Worksheets.Count)).Name = ccEvaluado
    End If
    
    Dim u As Integer
    Dim suma As Double
    
    'variable de control para no espacios es blanco
    u = 2
    
    
    Dim valores() As String
    Dim limite As Integer
    limite = 3
    
    ReDim valores(limite) As String
        
    For p = 0 To limite
        valores(p) = (ccEvaluado * 10) + p + 1
    Next p
    
    
    For i = 1 To cuentasAEvaluar.Application.WorksheetFunction.CountA(cuentasAEvaluar)
        
        'limpiamos la hoja para evitar errores
        hojaDeOperaciones.Cells.Clear
        
        'vamos a recorer el arreglo de cuentas y hacer un filtrado por cuenta para sacar la sumatoria de valores
       
        cuentaEvaluada = cuentasAEvaluar.Cells(i, 1)
        
            hojaDeArany.UsedRange.AutoFilter 5, valores(), xlFilterValues
            
            hojaDeArany.UsedRange.AutoFilter 4, cuentaEvaluada
            
            hojaDeArany.UsedRange.Copy hojaDeOperaciones.Range("A1")
            
            hojaDeArany.AutoFilterMode = False
        
        'por si los datos no esta en el formato correcto
        
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
        
        NombreCuentaEvaluada = hojaDeOperaciones.Range("F2")
        
        
        '-----------------------------------------------------------------------------
        'ahora hacemos consolidado de la hoja de taller
        
         'limpiamos la hoja para evitar errores
         hojaDeOperaciones.Cells.Clear
         
         'vamos a recorer el arreglo de cuentas y hacer un filtrado por cuenta para sacar la sumatoria de valores
        
         cuentaEvaluada = cuentasAEvaluar.Cells(i, 1)
            
             hojaDeTaller.UsedRange.AutoFilter 5, valores(), xlFilterValues
             
             hojaDeTaller.UsedRange.AutoFilter 4, cuentaEvaluada
             
             hojaDeTaller.UsedRange.Copy hojaDeOperaciones.Range("A1")
             
             hojaDeTaller.AutoFilterMode = False
             
        'por si los datos no esta en el formato correcto
        
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
         
         subTotalDebito = subTotalDebito + Application.WorksheetFunction.Sum(hojaDeOperaciones.Range("K2:K2000"))
         
         subTotalCredito = subTotalCredito + Application.WorksheetFunction.Sum(hojaDeOperaciones.Range("L2:L2000"))
         
         subTotalSaldo = subTotalSaldo + Application.WorksheetFunction.Sum(hojaDeOperaciones.Range("M2:M2000"))
         
        If NombreCuentaEvaluada = "" Then
        
            NombreCuentaEvaluada = hojaDeOperaciones.Range("F2")
        
        End If
        
        '---------------------------------------------------------------------
         
        'ahora ya con las sumas lo organizamos en la tabla general para sacar el reporte general de datos
         
        'en caso de no tener valores en la columnas entonces borramos esa cuenta
        suma = 0
        
        If subTotalDebito = 0 Then
         suma = suma + 1
        End If
        
        If subTotalCredito = 0 Then
         suma = suma + 1
        End If
        
        If subTotalSaldo = 0 Then
         suma = suma + 1
        End If
         
        
        If suma <> 3 Then
        
             Sheets(ccEvaluado).Cells(u + 1, 1) = cuentaEvaluada
             
             Sheets(ccEvaluado).Cells(u + 1, 2) = NombreCuentaEvaluada
             
             If subTotalDebito = 0 Then
                Sheets(ccEvaluado).Cells(u + 1, 3) = ""
             Else
                Sheets(ccEvaluado).Cells(u + 1, 3) = subTotalDebito
             End If
             
             If subTotalCredito = 0 Then
                Sheets(ccEvaluado).Cells(u + 1, 4) = ""
             Else
                Sheets(ccEvaluado).Cells(u + 1, 4) = subTotalCredito
             End If
             
             If subTotalSaldo = 0 Then
                Sheets(ccEvaluado).Cells(u + 1, 5) = ""
             Else
                Sheets(ccEvaluado).Cells(u + 1, 5) = subTotalSaldo
             End If
            
            u = u + 1
            
        End If
            
             
            
    Next i
    Sheets(ccEvaluado).Cells(2, 1) = "Cuenta - Cód."
    Sheets(ccEvaluado).Cells(2, 2) = "Cuenta - Nombre de la cuenta NIIF"
    Sheets(ccEvaluado).Cells(2, 3) = "debito"
    Sheets(ccEvaluado).Cells(2, 4) = "credito"
    Sheets(ccEvaluado).Cells(2, 5) = "saldo"
    
    'convertimos el formato a dinero para mejor lectura
    Sheets(ccEvaluado).Range("C:C").Style = "Currency"
    Sheets(ccEvaluado).Range("D:D").Style = "Currency"
    Sheets(ccEvaluado).Range("E:E").Style = "Currency"
    
    
    Sheets(ccEvaluado).Range("A1").EntireRow.Insert
    
    'fecha inicio registros
    Sheets(ccEvaluado).Range("A1") = "fecha inicio"
    Sheets(ccEvaluado).Range("B1") = hojaDeTaller.Range("A2")
    Sheets(ccEvaluado).Range("B1").NumberFormat = "yyyy-mm-dd"
    
    'fecho fin registros
    Sheets(ccEvaluado).Range("D1") = "fecha inicio"
    Sheets(ccEvaluado).Range("E1") = hojaDeTaller.Cells((hojaDeTaller.Application.WorksheetFunction.CountA(hojaDeTaller.Range("A:A"))), 1)
    Sheets(ccEvaluado).Range("E1").NumberFormat = "yyyy-mm-dd"
    

End Sub


