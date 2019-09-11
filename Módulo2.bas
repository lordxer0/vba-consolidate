Attribute VB_Name = "Módulo2"
Sub obtenerCuentasyCC()
'bloque de cc imputado

'limpiamos las columnas para evitar errores

Sheets("base").Range("A:A").Clear
Sheets("base").Range("B:B").Clear

Sheets("base").Range("D:D").Clear
Sheets("base").Range("E:E").Clear

Sheets("base").Range("G:G").Clear
Sheets("base").Range("I:I").Clear

    Dim CCIdatosArany As Range
    'defino CCIdatosArany como rango para los nombre
    
    Set CCIdatosArany = Sheets("base").Range("A:A")
        
        Sheets("aranysport").Range("E:E").Copy (Sheets("base").Range("A:A"))
        
        Sheets("base").Range("A:A").RemoveDuplicates (1)
        'removemos los duplicados
        Sheets("base").Range("A2:A500").Sort Key1:=Sheets("base").Range("A1"), Order1:=xlAscending
        'organizamos acendente
           
    
    Dim CCIdatosTaller As Range
    'defino CCIdatosTaller como rango para los nombre
    
    Set CCIdatosTaller = Sheets("base").Range("B:B")
    
        Sheets("areadetrabajo").Range("E:E").Copy (Sheets("base").Range("B:B"))
        
        Sheets("base").Range("B:B").RemoveDuplicates Columns:=1
        'removemos los duplicados
        Sheets("base").Range("B2:B100").Sort Key1:=Sheets("base").Range("B1"), Order1:=xlAscending
        'organizamos acendente
        
       
    
        
    Dim CCIdatos As Range
    'creamos el rango para la columna de datos ccimputado
    
    Set CCIdatos = Sheets("base").Range("I1:I200")
    'establesemos la columna CCImputado en donde el otro boton funciona
    
    Dim contador As Integer
    'llevamos un contador para saber donde acaba una columna de datos y empieza la otra
        
        'rellenamos la columna con los datos que corresponden a los CCimputados de ambos docs
        For i = 1 To CCIdatosArany.Application.WorksheetFunction.CountA(CCIdatosArany)
        
            If CCIdatosArany.Cells(i, 1) <> "" Then
                contador = i
                CCIdatos.Cells(i, 1) = CCIdatosArany.Cells(i, 1)
                'rellenamos hasta el primer corte
            End If
            
        Next i
        
        For j = contador + 1 To (contador + CCIdatosTaller.Application.WorksheetFunction.CountA(CCIdatosTaller))
        
            If CCIdatosTaller.Cells((j - contador), 1) <> "" Then
                CCIdatos.Cells(j, 1) = CCIdatosTaller.Cells((j - contador), 1)
                'rellenamos hasta el segundo corte
            End If
            
        Next j
        
    Sheets("base").Range("I:I").RemoveDuplicates Columns:=1
    'removemos los duplicados
    Sheets("base").Range("I2:I500").Sort Key1:=Sheets("base").Range("I1"), Order1:=xlAscending
    'organizamos acendente

'termina bloque cc imputado
    
'---------------

'bloque de cuentas

    Dim CuentadatosArany As Range
    'defino CuentadatosArany como rango para los nombre
    
    Set CuentadatosArany = Sheets("base").Range("D:D")
    
        Sheets("aranysport").Range("D:D").Copy (Sheets("base").Range("D:D"))
        
        Sheets("base").Range("D:D").RemoveDuplicates Columns:=1
        'removemos los duplicados
    
        Sheets("base").Range("D2:D500").Sort Key1:=Sheets("base").Range("D1"), Order1:=xlAscending
        'organizamos acendente
      


    Dim CuentadatosTaller As Range
    'defino CuentadatosTaller como rango para los nombre
    
    Set CuentadatosTaller = Sheets("base").Range("E:E")
    

        Sheets("areadetrabajo").Range("D:D").Copy (Sheets("base").Range("E:E"))
        
        Sheets("base").Range("E:E").RemoveDuplicates Columns:=1
        'removemos los duplicados
    
        Sheets("base").Range("E2:E500").Sort Key1:=Sheets("base").Range("E1"), Order1:=xlAscending
        'organizamos acendente
            
          
        
    Dim CuentaDatos As Range
    'creamos el rango para la columna de datos ccimputado
    
    Set CuentaDatos = Sheets("base").Range("G1:G500")
    'establesemos la columna CCImputado en donde el otro boton funciona
    
    Dim contador2 As Integer
    'llevamos un contador para saber donde acaba una columna de datos y empieza la otra
        
        'rellenamos la columna con los datos que corresponden a los CCimputados de ambos docs
        For w = 1 To CuentadatosArany.Application.WorksheetFunction.CountA(CuentadatosArany)
        
            If CuentadatosArany.Cells(w, 1) <> "" Then
                contador2 = w
                CuentaDatos.Cells(w, 1) = CuentadatosArany.Cells(w, 1)
                'rellenamos hasta el primer corte
            End If
            
        Next w
        
        For j = contador2 + 1 To (contador2 + CuentadatosTaller.Application.WorksheetFunction.CountA(CuentadatosTaller))
        
            If CuentadatosTaller.Cells((j - contador2), 1) <> "" Then
                CuentaDatos.Cells(j, 1) = CuentadatosTaller.Cells((j - contador2), 1)
                'rellenamos hasta el segundo corte
            End If
             
        Next j
        
    Sheets("base").Range("G:G").RemoveDuplicates Columns:=1
    'removemos los duplicados
    Sheets("base").Range("G2:G500").Sort Key1:=Sheets("base").Range("G1"), Order1:=xlAscending
    'organizamos acendente
    
'termina bloque de cuentas


End Sub
