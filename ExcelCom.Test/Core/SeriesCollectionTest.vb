Imports NUnit.Framework

Namespace Core

    Public MustInherit Class SeriesCollectionTest

        Private sut As Application
        Private workbook As Workbook
        Private sheet As Worksheet

        <SetUp()> Public Sub SetUp()
            sut = New Application
            workbook = sut.Workbooks.Add
            sheet = workbook.Sheets.Add
        End Sub

        <TearDown()> Public Sub TearDown()
            If sut Is Nothing Then
                Return
            End If
            sut.Dispose()
        End Sub

        Public Class プロパティたちTest : Inherits SeriesCollectionTest

            <Test()> Public Sub Count_最初は0()
                Assert.That(sheet.ChartObjects.Add(2, 3, 4, 5).Chart.SeriesCollection.Count, [Is].EqualTo(0))
            End Sub

            '<Test()> Public Sub Count_追加すれば増える()
            '    sheet.SeriesCollection.Add(1, 2, 3, 4)
            '    sheet.SeriesCollection.Add(5, 6, 7, 8)
            '    Assert.That(sheet.SeriesCollection.Count, [Is].EqualTo(2))
            'End Sub

        End Class

        'Public Class Item_Test : Inherits SeriesCollectionTest

        '    <Test()> Public Sub AddしたChartObjectとName指定と同じインスタンスである()
        '        Dim ChartObject As WorkChartObject = workbook.SeriesCollection.Add
        '        Assert.That(workbook.SeriesCollection(ChartObject.Name), [Is].SameAs(ChartObject))
        '    End Sub

        '    <Test()> Public Sub Nameとindexとで同じインスタンスである()
        '        Dim ChartObject As WorkChartObject = workbook.SeriesCollection.Add
        '        Assert.That(workbook.SeriesCollection(ChartObject.Name), [Is].SameAs(workbook.SeriesCollection(0)))
        '    End Sub

        'End Class

        Public Class ExcelObjectたちTest : Inherits SeriesCollectionTest

            '<Test()> Public Sub AddのSeriesが閉じられること()
            '    Dim value As Series = sheet.ChartObjects.Add(2, 3, 4, 5).Chart.SeriesCollection.Add

            '    sut.Dispose()
            '    sut = Nothing

            '    TestUtil.AssertNotExistsExcelPropcess()
            'End Sub

            <Test()> Public Sub NewSeriesのSeriesが閉じられること()
                Dim value As Series = sheet.ChartObjects.Add(2, 3, 4, 5).Chart.SeriesCollection.NewSeries

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

    End Class
End Namespace