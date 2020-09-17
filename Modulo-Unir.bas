Attribute VB_Name = "Modulo1"
Sub Unir()
  'El macro se debe ejecutar teniendo seleccionado el workbook (de una Hoja) que tiene la tabla principal a la que se
  'copiar� la data.
  'La data se copia desde tablas en distintos archivos a la tabla principal. La data se posiciona en la tabla principal
  'bas�ndose en una columna que tienen en com�n todas las tablas con la tabla principal.
  'la tabla principal tiene una columna con nombres. Las tablas externas tienen una columna con
  'algunos de esos nombres cada una y data en varias columnas a la derecha de cada nombre. Por eso se busca jalar la data
  'a la tabla principal que tiene todos los nombres pero nada de la data.
  
  
  'Al comienzo se pide al usario que selecciones los archivos de excel de donde se sacar� la data.
  'Cada archivo de excel debe ser un workbook de una sola hoja. En la hoja debe haber una tabla con la data a jalar.
  With Application.FileDialog(msoFileDialogFilePicker)
        'Makes sure the user can select only one file
        .AllowMultiSelect = True
        .Title = "Escoga los archivos con la data a extraer (Reportes del sistema):"
        'Filter to just the following types of files to narrow down selection options
        .Filters.Add "Excel Files", "*.xlsx; *.xlsm; *.xls; *.xlsb", 1
        'Show the dialog box
        .Show
        
        'Store in archivos variable
        Dim myCount As Integer
            
        
        myCount = .SelectedItems.Count
        If myCount > 0 Then
            ReDim archivos(1 To myCount)
            Dim idx As Long
            For idx = 1 To myCount
                archivos(idx) = .SelectedItems(idx)
                
            Next
        End If
        
        'cada elemento del array archivo() es uno de los workbooks seleccionados
        'archivos es el array que contiene todos los workbooks o archivos de Excel (.xls)
        'Se trabaja suponiendo que cada workbook tiene una Hoja o Sheet
        Dim i As Variant
        Dim j As Variant
        
        Dim lastRow As Long
        Dim archivoWb As Workbook
        
        Dim colIArchivo As Variant
        Dim colK As Variant, colL As Variant, colM As Variant, colN As Variant, colO As Variant
        Dim colP As Variant, colQ As Variant, colR As Variant, colS As Variant, colT As Variant, colU As Variant
        Dim colV As Variant, colW As Variant, colX As Variant, colY As Variant, colZ As Variant, colAA As Variant
        Dim colAB As Variant, colAC As Variant, colAD As Variant
        Dim cellArchivo As Variant
        
        'actualWorkbook hace referencia al Workbook(de una sola hoja) donde est� la tabla principal a la que se
        'copiar� la informaci�n de las tablas que se encuentran en cada workbook del array archivo()
        Dim actualWorkbook As Workbook
        Set actualWb = ActiveWorkbook
        lastRowActual = Cells(Rows.Count, "L").End(xlUp).Row '//Se halla el final de la tabla principal'
        Dim cellActual As Variant
        
        Dim MatchRow As Long
                 
        'Le quito los espacios finales e iniciales a los nombres del archivo principal(actual)
        i = 1
        For Each cellActual In Range(Cells(5, "L"), Cells(lastRowActual, "L"))
                 
        actualWb.Sheets(1).Cells((j + 4), "L") = Trim(cellActual)
                 
        j = j + 1
        Next cellActual
                 
        'itero en cada workbook del array archivos(), uno por uno y voy pasando la data a la tabla principal basandome
        'en si la data de una columna de las tablas externas coincide con la tabla principal.
        For i = 1 To UBound(archivos)
            'Se encuentra la �ltima fila de la tabla en el archivo(i) que tiene la data a copiar a la tabla principal.
            Set archivoWb = Application.Workbooks.Open(archivos(i))
            archivoWb.Sheets(1).Activate
            lastRowArchivo = Cells(Rows.Count, "I").End(xlUp).Row
            j = 1
            
            ReDim colIArchivo(1 To (lastRowArchivo - 20))
            ReDim colK(1 To (lastRowArchivo - 20))
            ReDim colL(1 To (lastRowArchivo - 20))
            ReDim colM(1 To (lastRowArchivo - 20))
            ReDim colN(1 To (lastRowArchivo - 20))
            ReDim colO(1 To (lastRowArchivo - 20))
            ReDim colP(1 To (lastRowArchivo - 20))
            ReDim colQ(1 To (lastRowArchivo - 20))
            ReDim colR(1 To (lastRowArchivo - 20))
            ReDim colS(1 To (lastRowArchivo - 20))
            ReDim colT(1 To (lastRowArchivo - 20))
            ReDim colU(1 To (lastRowArchivo - 20))
            ReDim colV(1 To (lastRowArchivo - 20))
            ReDim colW(1 To (lastRowArchivo - 20))
            ReDim colX(1 To (lastRowArchivo - 20))
            ReDim colY(1 To (lastRowArchivo - 20))
            ReDim colZ(1 To (lastRowArchivo - 20))
            ReDim colAA(1 To (lastRowArchivo - 20))
            ReDim colAB(1 To (lastRowArchivo - 20))
            ReDim colAC(1 To (lastRowArchivo - 20))
            ReDim colAD(1 To (lastRowArchivo - 20))
            
            'Guardo en varios array toda la tabla del archivo(i) de donde quiero jalar data.
            For Each cellArchivo In Range(Cells(21, "I"), Cells(lastRowArchivo, "I"))
                colIArchivo(j) = cellArchivo
                colK(j) = Sheets(1).Cells((j + 20), "K")
                colL(j) = Sheets(1).Cells((j + 20), "L")
                colM(j) = Sheets(1).Cells((j + 20), "M")
                colN(j) = Sheets(1).Cells((j + 20), "N")
                colO(j) = Sheets(1).Cells((j + 20), "O")
                colP(j) = Sheets(1).Cells((j + 20), "P")
                colQ(j) = Sheets(1).Cells((j + 20), "Q")
                colR(j) = Sheets(1).Cells((j + 20), "R")
                colS(j) = Sheets(1).Cells((j + 20), "S")
                colT(j) = Sheets(1).Cells((j + 20), "T")
                colU(j) = Sheets(1).Cells((j + 20), "U")
                colV(j) = Sheets(1).Cells((j + 20), "V")
                colW(j) = Sheets(1).Cells((j + 20), "W")
                colX(j) = Sheets(1).Cells((j + 20), "X")
                colY(j) = Sheets(1).Cells((j + 20), "Y")
                colZ(j) = Sheets(1).Cells((j + 20), "Z")
                colAA(j) = Sheets(1).Cells((j + 20), "AA")
                colAB(j) = Sheets(1).Cells((j + 20), "AB")
                colAC(j) = Sheets(1).Cells((j + 20), "AC")
                colAD(j) = Sheets(1).Cells((j + 20), "AD")
                j = j + 1
            Next cellArchivo
            
            j = 1
            
            actualWb.Sheets(1).Activate
            
            'Voy a iterar sobre la columna en com�nde la tabla principal del archivo principal(actual)
            'las tablas externas tienen esa columna pero no todas las filas de esa columna est�n en la columna de la tabla principal
            'Se itera para agregar la data de la tabla del archivo(i) que ahora est� en varios arrays
            For Each cellActual In Range(Cells(5, "L"), Cells(lastRowActual, "L"))
                    
                'uso Match para hallar la coincidencia entre elementos de la columna de la
                'tabla externa (que ahora est�n en el array colIArchivo(j)) con los
                'elementos en la columna de la tabla principal.
                On Error Resume Next 'match throws an error if nothing matched
                MatchRow = 0
                MatchRow = Application.Match(cellActual, colIArchivo, 0)
                On Error GoTo 0
                
                'Cuando se encuentra que un elemento del array colIArchivo(j)) coincide con un fila de la columna con nombres
                'de la tabla principal, se copia toda la data correspondiente a ese nombre a la derecha del nombre en la tabla principal.
                'La data se jala de los arrays. Se usa offset porque se empieza a colocar la data a la derecha del nombre
                'pero desde 14 columnas a la derecha,
                If MatchRow <> 0 Then
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 14) = colK(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 15) = colL(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 16) = colM(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 17) = colN(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 18) = colO(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 19) = colP(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 20) = colQ(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 21) = colR(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 22) = colS(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 23) = colT(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 24) = colU(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 25) = colV(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 26) = colW(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 27) = colX(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 28) = colY(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 29) = colZ(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 30) = colAA(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 31) = colAB(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 32) = colAC(MatchRow)
                    actualWb.Sheets(1).Cells((j + 4), "L").Offset(0, 33) = colAD(MatchRow)
                End If
                     
                j = j + 1
                
            Next cellActual
         
            archivoWb.Close
                    
        Next i
        
        
    End With
   
    
End Sub

