Imports NUnit.Framework

Namespace Core

    Public MustInherit Class ChartTest

        Private sut As Application
        Private workbook As Workbook
        Private sheet As Worksheet

        <SetUp()> Public Sub SetUp()
            sut = New Application
            workbook = sut.Workbooks.Add
            sheet = workbook.Sheets.Add()
        End Sub

        <TearDown()> Public Sub TearDown()
            If sut Is Nothing Then
                Return
            End If
            sut.Dispose()
        End Sub

        'Public Class LocationTest : Inherits ChartTest

        '    <Test()> Public Sub Cellsの値をRangeと比較できる()
        '        Dim chart As Chart = sut.Charts.Add
        '        chart.Location(Core.Chart.XlChartLocation.xlLocationAsObject, name:="ほげ")

        '        Assert.That(chart.Name, [Is].EqualTo("ほげ"))
        '    End Sub

        '    '<Test()> Public Sub Cellsの値をRangeと比較できる2()
        '    '    workbook.Sheets.Item(0).Cells(2, 0).Value = "aiueo"

        '    '    Assert.That(workbook.Sheets.Item(0).Range("A3").Value, [Is].EqualTo("aiueo"))
        '    'End Sub

        '    '<Test()> Public Sub Hoge()
        '    '    Dim start As Range = workbook.Sheets.Item(0).Cells(1, 1)

        '    '    Dim target As Range = workbook.Sheets.Item(0).Range(start, start)
        '    '    target.Value = "xyz"
        '    '    Assert.That(workbook.Sheets.Item(0).Range("B2").Value, [Is].EqualTo("xyz"))
        '    'End Sub

        'End Class

        Public Class ExcelObjectたちTest : Inherits ChartTest

            <Test()> Public Sub SeriesCollectionが閉じられること()
                Dim value As SeriesCollection = sheet.ChartObjects.Add(2, 3, 4, 5).Chart.SeriesCollection

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

        Public Class PropertyたちTest : Inherits ChartTest

            <Test()> Public Sub HasLegend_(<Values(True, False)> ByVal value As Boolean)
                Dim chart As Chart = sut.Charts.Add
                chart.HasLegend = value
                Assert.That(chart.HasLegend, [Is].EqualTo(value))
            End Sub

        End Class

    End Class
End Namespace
