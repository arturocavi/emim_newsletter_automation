Attribute VB_Name = "AUT_GRAF_EMIM_24_02"
' Variables globales
Global Grafica_HIS_POT
Global grafica_rank_pot
Global Grafica_HIS_HOR
Global grafica_rank_hor
Global Grafica_HIS_VP
Global Grafica_RANK_VP
Sub Macro_inicio()
Attribute Macro_inicio.VB_ProcData.VB_Invoke_Func = "a\n14"
'Ctrl a
Call WorksheetLoop
Call Generación_Word

End Sub

Sub WorksheetLoop()

        Dim WS_Count As Integer
        Dim I As Integer

         ' Define WS_Count como el número de hojas en el libro activo
        WS_Count = ActiveWorkbook.Worksheets.Count

         ' Empieza el loop
        For I = 1 To WS_Count
            
            Worksheets(I).Select
            nombre = Worksheets(I).Name
            
            If InStr(1, nombre, "HIS POT", vbBinaryCompare) = 1 Then
                Call Macro_linea_promedio_HIS_POT
            ElseIf InStr(1, nombre, "HIS HOR", vbBinaryCompare) = 1 Then
                Call Macro_linea_promedio_HIS_HOR
            ElseIf InStr(1, nombre, "HIS VP", vbBinaryCompare) = 1 Then
                Call Macro_linea_promedio_HIS_VP
            ElseIf InStr(1, nombre, "RANK HOR", vbBinaryCompare) = 1 Then
                Call Macro_Graficas_Nacional_RANK_HOR
            ElseIf InStr(1, nombre, "RANK POT", vbBinaryCompare) = 1 Then
                Call Macro_Graficas_Nacional_RANK_POT
            ElseIf InStr(1, nombre, "RANK VP", vbBinaryCompare) = 1 Then
                Call Macro_Graficas_Nacional_RANK_VP
            End If

        Next I

End Sub
Sub Macro_Graficas_Nacional_RANK_HOR()
'RANK
' Ranking por entidad federativa con valor nacional, anual
'

Dim Grafica As ChartObject

fila = 4

Do While Cells(fila, 1) = ""
    fila = fila + 1
Loop

ultimo = fila
Do While Cells(ultimo, 1) <> ""
    ultimo = ultimo + 1
Loop
'

'
Range("B" & (fila + 1) & ":B" & (ultimo - 1)).NumberFormat = "0.0"
hoja = ActiveSheet.Name

'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'

Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Sort Key1:=Range("B" & (fila + 1)), Order1:=xlAscending

'
jalisco = fila
nacional = fila

Do While Cells(jalisco, 1) <> "Jalisco"
    jalisco = jalisco + 1
Loop
jalisco = jalisco - fila

Do While Cells(nacional, 1) <> "Nacional"
    nacional = nacional + 1
Loop
nacional = nacional - fila
'
Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Select
'
Set grafica_rank_hor = ActiveSheet.ChartObjects.Add(Left:=4 * 48, Width:=468.1, Top:=60, Height:=448.5)

With grafica_rank_hor.Chart
    .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC RANK AUT.crtx") ' UBICACIÓN PERSONAL
    .SetSourceData Source:=Range("A" & (fila + 1) & ":B" & (ultimo - 1))
    .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    .SeriesCollection(1).Points(nacional).Format.Fill.ForeColor.RGB = RGB(149, 104, 43)
End With
'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'
End Sub
Sub Macro_Graficas_Nacional_RANK_POT()
'RANK
' Ranking por entidad federativa con valor nacional, anual
'

Dim Grafica As ChartObject

fila = 4

Do While Cells(fila, 1) = ""
    fila = fila + 1
Loop

ultimo = fila
Do While Cells(ultimo, 1) <> ""
    ultimo = ultimo + 1
Loop
'

'
Range("B" & (fila + 1) & ":B" & (ultimo - 1)).NumberFormat = "0.0"
hoja = ActiveSheet.Name

'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'

Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Sort Key1:=Range("B" & (fila + 1)), Order1:=xlAscending

'
jalisco = fila
nacional = fila

Do While Cells(jalisco, 1) <> "Jalisco"
    jalisco = jalisco + 1
Loop
jalisco = jalisco - fila

Do While Cells(nacional, 1) <> "Nacional"
    nacional = nacional + 1
Loop
nacional = nacional - fila
'
Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Select
'
Set grafica_rank_pot = ActiveSheet.ChartObjects.Add(Left:=4 * 48, Width:=468.1, Top:=60, Height:=448.5)

With grafica_rank_pot.Chart
    .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC RANK AUT.crtx") ' UBICACIÓN PERSONAL
    .SetSourceData Source:=Range("A" & (fila + 1) & ":B" & (ultimo - 1))
    .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    .SeriesCollection(1).Points(nacional).Format.Fill.ForeColor.RGB = RGB(149, 104, 43)
End With
'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'
End Sub
Sub Macro_Graficas_Nacional_RANK_VP()
'RANK
' Ranking por entidad federativa con valor nacional, anual
'

Dim Grafica As ChartObject

fila = 4

Do While Cells(fila, 1) = ""
    fila = fila + 1
Loop

ultimo = fila
Do While Cells(ultimo, 1) <> ""
    ultimo = ultimo + 1
Loop
'

'
Range("B" & (fila + 1) & ":B" & (ultimo - 1)).NumberFormat = "0.0"
hoja = ActiveSheet.Name

'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'

Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Sort Key1:=Range("B" & (fila + 1)), Order1:=xlAscending

'
jalisco = fila
nacional = fila

Do While Cells(jalisco, 1) <> "Jalisco"
    jalisco = jalisco + 1
Loop
jalisco = jalisco - fila

Do While Cells(nacional, 1) <> "Nacional"
    nacional = nacional + 1
Loop
nacional = nacional - fila
'
Range("A" & (fila + 1) & ":B" & (ultimo - 1)).Select
'
Set Grafica_RANK_VP = ActiveSheet.ChartObjects.Add(Left:=4 * 48, Width:=468.1, Top:=60, Height:=448.5)

With Grafica_RANK_VP.Chart
    .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC RANK AUT.crtx") ' UBICACIÓN PERSONAL
    .SetSourceData Source:=Range("A" & (fila + 1) & ":B" & (ultimo - 1))
    .SeriesCollection(1).Points(jalisco).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    .SeriesCollection(1).Points(nacional).Format.Fill.ForeColor.RGB = RGB(149, 104, 43)
End With
'
Range("A" & fila & ":B" & (ultimo - 1)).AutoFilter
'
End Sub
Sub Macro_linea_promedio_HIS_POT()
Attribute Macro_linea_promedio_HIS_POT.VB_ProcData.VB_Invoke_Func = "p\n14"
'HIS & VAR
'Gráfica de Barras de Históricos Mensuales  con Linea de Promedio de Últimos 12 Meses
'Ctrl + p

nombre = ActiveSheet.Name

If InStr(1, nombre, "HIS", vbBinaryCompare) = 1 Then
    Range("C6:D6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "#,##0"

    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 2) <> ""
        fin = fin + 1
    Loop
    
    Range("A" & (inicio) & ":D" & (fin - 1)).Select

    Set Grafica_HIS_POT = ActiveSheet.ChartObjects.Add(Left:=5 * 48, Width:=468.1, Top:=60, Height:=250)
    
    With Grafica_HIS_POT.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC HIS AUT.crtx") ' UBICACIÓN PERSONAL
        .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
        For k = 1 To (fin - 1)
            If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
            End If
        Next k
        .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    End With
End If

End Sub
Sub Macro_linea_promedio_HIS_HOR()
'HIS & VAR
'Gráfica de Barras de Históricos Mensuales  con Linea de Promedio de Últimos 12 Meses
'Ctrl + p

nombre = ActiveSheet.Name

If InStr(1, nombre, "HIS", vbBinaryCompare) = 1 Then
    Range("C6:D6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "#,##0"

    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 2) <> ""
        fin = fin + 1
    Loop
    
    Range("A" & (inicio) & ":D" & (fin - 1)).Select

    Set Grafica_HIS_HOR = ActiveSheet.ChartObjects.Add(Left:=5 * 48, Width:=468.1, Top:=60, Height:=250)
    
    With Grafica_HIS_HOR.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC HIS AUT.crtx") ' UBICACIÓN PERSONAL
        .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
        For k = 1 To (fin - 1)
            If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
            End If
        Next k
        .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    End With
End If
Grafica_HIS_HOR.Activate
ActiveChart.Axes(xlValue).Select
ActiveChart.Axes(xlValue).MinimumScale = 40000



End Sub
Sub Macro_linea_promedio_HIS_VP()
'HIS & VAR
'Gráfica de Barras de Históricos Mensuales  con Linea de Promedio de Últimos 12 Meses
'Ctrl + p

nombre = ActiveSheet.Name

If InStr(1, nombre, "HIS", vbBinaryCompare) = 1 Then
    Range("C6:D6").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.NumberFormat = "#,##0"

    inicio = 5
    fin = inicio
    
    Do While Cells(fin, 2) <> ""
        fin = fin + 1
    Loop
    
    Range("A" & (inicio) & ":D" & (fin - 1)).Select

    Set Grafica_HIS_VP = ActiveSheet.ChartObjects.Add(Left:=5 * 48, Width:=468.1, Top:=60, Height:=250)
    
    With Grafica_HIS_VP.Chart
        .ApplyChartTemplate ("C:\Users\arturo.carrillo\AppData\Roaming\Microsoft\Plantillas\Charts\ENEC HIS AUT.crtx") ' UBICACIÓN PERSONAL
        .SetSourceData Source:=Range("A" & (inicio) & ":D" & (fin - 1))
        For k = 1 To (fin - 1)
            If (k Mod 12) = ((fin - 1 - inicio) Mod 12) Then
                .SeriesCollection(1).Points(k).Format.Fill.ForeColor.RGB = RGB(124, 135, 142)
            End If
        Next k
        .SeriesCollection(1).Points(fin - 1 - inicio).Format.Fill.ForeColor.RGB = RGB(251, 187, 39)
    End With
End If
Grafica_HIS_VP.Activate
ActiveChart.Axes(xlValue).Select
ActiveChart.Axes(xlValue).MinimumScale = 45000
End Sub
Sub Generación_Word()

' Nombre  y ubicación de la plantilla
plantilla = "C:\Users\arturo.carrillo\Documents\EMIM\AUT\PLANTILLA.dotx" ' UBICACIÓN PERSONAL

' Creamos el nuevo archivo word usando la plantilla
Set aplicacion = CreateObject("Word.Application")
aplicacion.Visible = True

Set documento = aplicacion.Documents.Add(Template:=plantilla, NewTemplate:=False, DocumentType:=0)

' Cambiamos la fecha del encabezado
diahoy = Format(Day(Now), "00")
meshoy = Format(Month(Now), "00")
añohoy = Year(Now)
If Month(Now) = 1 Then
    meshoypal = "enero"
    mesbas = Format(11, "00")
    mesbaspal = "noviembre"
    añobas = Year(Now) - 1
ElseIf Month(Now) = 2 Then
    meshoypal = "febrero"
    mesbas = Format(12, "00")
    mesbaspal = "diciembre"
    añobas = Year(Now) - 1
ElseIf Month(Now) = 3 Then
    meshoypal = "marzo"
    mesbas = Format(1, "00")
    mesbaspal = "enero"
    añobas = Year(Now)
ElseIf Month(Now) = 4 Then
    meshoypal = "abril"
    mesbas = Format(2, "00")
    mesbaspal = "febrero"
    añobas = Year(Now)
ElseIf Month(Now) = 5 Then
    meshoypal = "mayo"
    mesbas = Format(3, "00")
    mesbaspal = "marzo"
    añobas = Year(Now)
ElseIf Month(Now) = 6 Then
    meshoypal = "junio"
    mesbas = Format(4, "00")
    mesbaspal = "abril"
    añobas = Year(Now)
ElseIf Month(Now) = 7 Then
    meshoypal = "julio"
    mesbas = Format(5, "00")
    mesbaspal = "mayo"
    añobas = Year(Now)
ElseIf Month(Now) = 8 Then
    meshoypal = "agosto"
    mesbas = Format(6, "00")
    mesbaspal = "junio"
    añobas = Year(Now)
ElseIf Month(Now) = 9 Then
    meshoypal = "septiembre"
    mesbas = Format(7, "00")
    mesbaspal = "julio"
    añobas = Year(Now)
ElseIf Month(Now) = 10 Then
    meshoypal = "octubre"
    mesbas = Format(8, "00")
    mesbaspal = "agosto"
    añobas = Year(Now)
ElseIf Month(Now) = 11 Then
    meshoypal = "noviembre"
    mesbas = Format(9, "00")
    mesbaspal = "septiembre"
    añobas = Year(Now)
ElseIf Month(Now) = 12 Then
    meshoypal = "diciembre"
    mesbas = Format(10, "00")
    mesbaspal = "octubre"
    añobas = Year(Now)
End If

'FECHAS MANUALES
'diahoy = InputBox("Ingresa el día de hoy en formato de número a dos dígitos (ej. 23):")'
'meshoy = InputBox("Ingresa el mes de hoy en formato de número a dos dígitos (ej. 10):")
'añohoy = InputBox("Ingresa el año de hoy en formato de número a cuatro dígitos (ej. 2019):")
'meshoypal = InputBox("Ingresa el mes de hoy en formato de palabra en minúsculas (ej. octubre):")
'mesbas = InputBox("Ingresa el mes de la última base de datos del INEGI (dos meses atrás) en formato de número a dos dígitos (ej. 08):")
'mesbaspal = InputBox("Ingresa el mes de la última base de datos del INEGI (dos meses atrás) en formato de palabra en minúsculas (ej. agosto):")
'añobas = InputBox("Ingresa el año de la última base de datos del INEGI (dos meses atrás) en formato de número a cuatro dígitos (ej. 2019):")


' Cambiamos los espaciados del boletín
With documento.Content
    .Style = "Espaciado principal"
End With

' Insertar título del boletín
documento.Content.insertparagraphafter

With documento.Content
    .InsertAfter Hoja7.Cells(2, 1).Value ' Título del boletín [Paragraph(2)]
    .insertparagraphafter
End With

With documento.Paragraphs(2).Range
    .Style = "Título 1"
End With

' Insertar párrafo de texto MES
With documento.Content
    .InsertAfter Hoja7.Cells(5, 2).Value ' Texto MES [Paragraph(4)]
    .insertparagraphafter
End With

With documento.Paragraphs(3).Range
    .Style = "Normal"
End With

' Insertar título de gráfica MES
With documento.Content
    .InsertAfter Hoja7.Cells(6, 2).Value ' Título de gráfica MES [Paragraph(5)]
    .insertparagraphafter
End With

With documento.Paragraphs(4).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica MES
Grafica_HIS_POT.Chart.ChartArea.Copy
documento.Paragraphs(5).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG con información de INEGI. EMIM." ' Nota [Paragraph(7)]
    .insertparagraphafter
End With

With documento.Paragraphs(6).Range
    .Style = "Fuentes"
End With

' Insertar fuente
With documento.Content
    .InsertAfter Hoja7.Cells(8, 2).Value ' Nota [Paragraph(8)]
    .insertparagraphafter
End With

With documento.Paragraphs(7).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(8).Range.InsertBreak Type:=7 'wdSectionBreakNextPage
documento.Content.insertparagraphafter

' Insertar texto de la gráfica Grafica_RANK_POT
With documento.Content
    .InsertAfter Hoja7.Cells(11, 2).Value ' Texto de la gráfica HIS [Paragraph(11)]
    .insertparagraphafter
End With

With documento.Paragraphs(10).Range
    .Style = "Normal"
End With

' Insertar título de gráfica Grafica_RANK_POT
With documento.Content
    .InsertAfter Hoja7.Cells(12, 2).Value ' Título de gráfica HIS [Paragraph(12)]
    .insertparagraphafter
End With

With documento.Paragraphs(11).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica Grafica_RANK_POT
grafica_rank_pot.Chart.ChartArea.Copy
documento.Paragraphs(12).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG, con información de INEGI. EMIM." ' Nota [Paragraph(14)]
    .insertparagraphafter
End With

With documento.Paragraphs(13).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja7.Cells(14, 2).Value ' Nota [Paragraph(15)]
    .insertparagraphafter
End With

With documento.Paragraphs(14).Range
    .Style = "Fuentes"
End With

' Insertar salto de página
documento.Paragraphs(15).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter





' Insertar texto de la gráfica Grafica_HIS_HOR
With documento.Content
    .InsertAfter Hoja7.Cells(17, 2).Value ' Texto de la gráfica RANKDIS [Paragraph(18)]
    .insertparagraphafter
End With

With documento.Paragraphs(17).Range
    .Style = "Normal"
End With

' Insertar título de gráfica Grafica_HIS_HOR
With documento.Content
    .InsertAfter Hoja7.Cells(18, 2).Value ' Título de gráfica RANKDIS [Paragraph(19)]
    .insertparagraphafter
End With

With documento.Paragraphs(18).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica Grafica_HIS_HOR
Grafica_HIS_HOR.Chart.ChartArea.Copy
documento.Paragraphs(19).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG, con información de INEGI. EMIM." ' Nota [Paragraph(21)]
    .insertparagraphafter
End With

With documento.Paragraphs(20).Range
    .Style = "Fuentes"
End With

' Insertar fuente
With documento.Content
    .InsertAfter Hoja7.Cells(20, 2).Value ' Nota [Paragraph(22)]
    .insertparagraphafter
End With

With documento.Paragraphs(21).Range
    .Style = "Fuentes"
End With

' Insertar salto de página
documento.Paragraphs(22).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter

' Insertar texto de la gráfica Grafica_RANK_HOR
With documento.Content
    .InsertAfter Hoja7.Cells(23, 2).Value ' Texto de la gráfica RANKVPP [Paragraph(25)]
    .insertparagraphafter
End With

With documento.Paragraphs(24).Range
    .Style = "Normal"
End With

' Insertar título de gráfica Grafica_RANK_HOR
With documento.Content
    .InsertAfter Hoja7.Cells(24, 2).Value ' Título de la gráfica RANKVPP [Paragraph(26)]
    .insertparagraphafter
End With

With documento.Paragraphs(25).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica RANKVPP
grafica_rank_hor.Chart.ChartArea.Copy
documento.Paragraphs(26).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG, con información de INEGI. EMIM." ' Nota [Paragraph(28)]
    .insertparagraphafter
End With

With documento.Paragraphs(27).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja7.Cells(26, 2).Value ' Nota [Paragraph(29)]
    .insertparagraphafter
End With

With documento.Paragraphs(28).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(29).Range.InsertBreak Type:=7
documento.Content.insertparagraphafter

' Insertar texto de la gráfica Grafica_HIS_VP
With documento.Content
    .InsertAfter Hoja7.Cells(29, 2).Value ' Texto de la gráfica RANKVMEN [Paragraph(32)]
    .insertparagraphafter
End With

With documento.Paragraphs(31).Range
    .Style = "Normal"
End With

' Insertar título de gráfica Grafica_HIS_VP
With documento.Content
    .InsertAfter Hoja7.Cells(30, 2).Value ' Título de la gráfica RANKVMEN [Paragraph(33)]
    .insertparagraphafter
End With

With documento.Paragraphs(32).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica Grafica_HIS_VP
Grafica_HIS_VP.Chart.ChartArea.Copy
documento.Paragraphs(33).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG, con información de INEGI. EMIM." ' Nota [Paragraph(35)]
    .insertparagraphafter
End With

With documento.Paragraphs(34).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja7.Cells(32, 2).Value ' Nota [Paragraph(36)]
    .insertparagraphafter
End With

With documento.Paragraphs(35).Range
    .Style = "Fuentes"
End With
' Insertar salto de página
documento.Paragraphs(36).Range.InsertBreak Type:=7




' Insertar texto de la gráfica Grafica_RANK_VP
With documento.Content
    .InsertAfter Hoja7.Cells(35, 2).Value ' Texto de la gráfica [Paragraph(38)]
    .insertparagraphafter
End With

With documento.Paragraphs(37).Range
    .Style = "Normal"
End With

' Insertar título de gráfica Grafica_RANK_VP
With documento.Content
    .InsertAfter Hoja7.Cells(36, 2).Value ' Título de la gráfica [Paragraph(39)]
    .insertparagraphafter
End With

With documento.Paragraphs(38).Range
    .Style = "Figura - titulos"
End With

' Pasar gráfica Grafica_RANK_VP
Grafica_RANK_VP.Chart.ChartArea.Copy
documento.Paragraphs(39).Range.Paste
documento.Content.insertparagraphafter

' Insertar fuente
With documento.Content
    .InsertAfter "Fuente: IIEG, con información de INEGI. EMIM." ' Nota [Paragraph(41)]
    .insertparagraphafter
End With

With documento.Paragraphs(40).Range
    .Style = "Fuentes"
End With

' Insertar nota
With documento.Content
    .InsertAfter Hoja7.Cells(38, 2).Value ' Nota [Paragraph(42)]
End With

With documento.Paragraphs(41).Range
    .Style = "Fuentes"
End With



' Cambiar la fecha de realización
Set cuadrofecha = documento.Sections(1).Headers(1).Shapes.AddTextbox(msoTextOrientationHorizontal, _
                  340, 35 - 7, 240, 70 / 2)
                  ' wdHeaderFooterPrimary = 1
cuadrofecha.TextFrame.TextRange.Text = "Ficha informativa, " & diahoy & " de " & meshoypal & " de " & añohoy
cuadrofecha.TextFrame.TextRange.Font.Color = RGB(98, 113, 120)
cuadrofecha.TextFrame.TextRange.Font.Underline = wdUnderlineSingle
cuadrofecha.TextFrame.TextRange.ParagraphFormat.Alignment = wdAlignParagraphRight
cuadrofecha.Fill.ForeColor = RGB(255, 255, 255)
cuadrofecha.Line.ForeColor = RGB(255, 255, 255)

' Guardar el documento
documento.SaveAs "C:\Users\arturo.carrillo\Documents\EMIM\" & añobas & " " & mesbas & "\Ficha informativa Encuesta Mensual de la Industria Manufacturera (EMIM), " & mesbaspal & " " & añobas & "-" & añohoy & meshoy & diahoy & ".docx" ' UBICACIÓN PERSONAL

End Sub




