Attribute VB_Name = "Módulo6"
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
    ActiveSheet.PivotTables("Tabela dinâmica1").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("dia")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("dia").AutoGroup
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("Meses").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("Tabela dinâmica1").AddDataField ActiveSheet. _
        PivotTables("Tabela dinâmica1").PivotFields("valor"), "Soma de valor", xlSum
    ActiveSheet.PivotTables("Tabela dinâmica1").AddDataField ActiveSheet. _
        PivotTables("Tabela dinâmica1").PivotFields("Esperado por dia "), _
        "Soma de Esperado por dia ", xlSum
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields( _
        "Soma de Esperado por dia ")
        .Caption = "Média de Esperado por dia "
        .Function = xlAverage
    End With
    ActiveSheet.PivotTables("Tabela dinâmica1").AddDataField ActiveSheet. _
        PivotTables("Tabela dinâmica1").PivotFields("Diferenca dia"), _
        "Soma de Diferenca dia", xlSum
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields( _
        "Soma de Diferenca dia")
        .Caption = "Média de Diferenca dia"
        .Function = xlAverage
    End With
    Sheets("Dashboard").Select
    ActiveSheet.ChartObjects("Gráfico 1").Activate
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
    ActiveSheet.PivotTables("Tabela dinâmica1").ClearTable
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("dia")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("dia").AutoGroup
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields("dia").Orientation = _
        xlHidden
    ActiveSheet.PivotTables("Tabela dinâmica1").AddDataField ActiveSheet. _
        PivotTables("Tabela dinâmica1").PivotFields("Esperado por mês"), _
        "Soma de Esperado por mês", xlSum
    ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields( _
        "Soma de Esperado por mês").Orientation = xlHidden
    ActiveSheet.PivotTables("Tabela dinâmica1").AddDataField ActiveSheet. _
        PivotTables("Tabela dinâmica1").PivotFields("valor"), "Soma de valor", xlSum
    ActiveSheet.PivotTables("Tabela dinâmica1").AddDataField ActiveSheet. _
        PivotTables("Tabela dinâmica1").PivotFields("Esperado por mês"), _
        "Soma de Esperado por mês", xlSum
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields( _
        "Soma de Esperado por mês")
        .Caption = "Média de Esperado por mês"
        .Function = xlAverage
    End With
    ActiveSheet.PivotTables("Tabela dinâmica1").AddDataField ActiveSheet. _
        PivotTables("Tabela dinâmica1").PivotFields("Diferenca mês"), _
        "Soma de Diferenca mês", xlSum
    With ActiveSheet.PivotTables("Tabela dinâmica1").PivotFields( _
        "Soma de Diferenca mês")
        .Caption = "Média de Diferenca mês"
        .Function = xlAverage
    End With
    ActiveWorkbook.ShowPivotTableFieldList = False
    Sheets("Dashboard").Select
    ActiveSheet.ChartObjects("Gráfico 1").Activate
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
