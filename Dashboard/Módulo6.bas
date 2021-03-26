Attribute VB_Name = "M�dulo6"
Sub diferenca_dia()
Attribute diferenca_dia.VB_ProcData.VB_Invoke_Func = " \n14"
'
' diferenca_dia Macro
'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    Sheets("TabelaDin").Select
    Range("G3").Select
    ActiveWorkbook.ShowPivotTableFieldList = True
    ActiveSheet.PivotTables("Tabela din�mica1").ClearTable
    With ActiveSheet.PivotTables("Tabela din�mica1").PivotFields("dia")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabela din�mica1").PivotFields("dia").AutoGroup
    ActiveSheet.PivotTables("Tabela din�mica1").PivotFields("Meses").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("Tabela din�mica1").AddDataField ActiveSheet. _
        PivotTables("Tabela din�mica1").PivotFields("valor"), "Soma de valor", xlSum
    ActiveSheet.PivotTables("Tabela din�mica1").AddDataField ActiveSheet. _
        PivotTables("Tabela din�mica1").PivotFields("Esperado por dia "), _
        "Soma de Esperado por dia ", xlSum
    With ActiveSheet.PivotTables("Tabela din�mica1").PivotFields( _
        "Soma de Esperado por dia ")
        .Caption = "M�dia de Esperado por dia "
        .Function = xlAverage
    End With
    ActiveSheet.PivotTables("Tabela din�mica1").AddDataField ActiveSheet. _
        PivotTables("Tabela din�mica1").PivotFields("Diferenca dia"), _
        "Soma de Diferenca dia", xlSum
    With ActiveSheet.PivotTables("Tabela din�mica1").PivotFields( _
        "Soma de Diferenca dia")
        .Caption = "M�dia de Diferenca dia"
        .Function = xlAverage
    End With
    Sheets("Dashboard").Select
    ActiveSheet.ChartObjects("Gr�fico 1").Activate
    ActiveChart.PlotArea.Select
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(2).AxisGroup = 1
    ActiveChart.FullSeriesCollection(3).ChartType = xlLine
    ActiveChart.FullSeriesCollection(3).AxisGroup = 1
Application.DisplayAlerts = True
Application.ScreenUpdating = True
Range("A1").Select
Exit Sub

End Sub
Sub diferenca_mes()
Attribute diferenca_mes.VB_ProcData.VB_Invoke_Func = " \n14"
'
' diferenca_mes Macro
'

'
Application.ScreenUpdating = False
Application.DisplayAlerts = False
    Sheets("TabelaDin").Select
    Range("G3").Select
    ActiveSheet.PivotTables("Tabela din�mica1").ClearTable
    With ActiveSheet.PivotTables("Tabela din�mica1").PivotFields("dia")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabela din�mica1").PivotFields("dia").AutoGroup
    ActiveSheet.PivotTables("Tabela din�mica1").PivotFields("dia").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("Tabela din�mica1").AddDataField ActiveSheet. _
        PivotTables("Tabela din�mica1").PivotFields("Esperado por m�s"), _
        "Soma de Esperado por m�s", xlSum
    ActiveSheet.PivotTables("Tabela din�mica1").PivotFields( _
        "Soma de Esperado por m�s").Orientation = xlHidden
    ActiveSheet.PivotTables("Tabela din�mica1").AddDataField ActiveSheet. _
        PivotTables("Tabela din�mica1").PivotFields("valor"), "Soma de valor", xlSum
    ActiveSheet.PivotTables("Tabela din�mica1").AddDataField ActiveSheet. _
        PivotTables("Tabela din�mica1").PivotFields("Esperado por m�s"), _
        "Soma de Esperado por m�s", xlSum
    With ActiveSheet.PivotTables("Tabela din�mica1").PivotFields( _
        "Soma de Esperado por m�s")
        .Caption = "M�dia de Esperado por m�s"
        .Function = xlAverage
    End With
    ActiveSheet.PivotTables("Tabela din�mica1").AddDataField ActiveSheet. _
        PivotTables("Tabela din�mica1").PivotFields("Diferenca m�s"), _
        "Soma de Diferenca m�s", xlSum
    With ActiveSheet.PivotTables("Tabela din�mica1").PivotFields( _
        "Soma de Diferenca m�s")
        .Caption = "M�dia de Diferenca m�s"
        .Function = xlAverage
    End With
    ActiveWorkbook.ShowPivotTableFieldList = False
    Sheets("Dashboard").Select
    ActiveSheet.ChartObjects("Gr�fico 1").Activate
    ActiveChart.ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(1).AxisGroup = 1
    ActiveChart.FullSeriesCollection(2).ChartType = xlColumnClustered
    ActiveChart.FullSeriesCollection(2).AxisGroup = 1
    ActiveChart.FullSeriesCollection(3).ChartType = xlLine
    ActiveChart.FullSeriesCollection(3).AxisGroup = 1
Application.ScreenUpdating = True
Application.DisplayAlerts = True
Range("A1").Select
End Sub
