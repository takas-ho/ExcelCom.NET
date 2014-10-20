Imports NUnit.Framework

Namespace Core

    Public MustInherit Class ChartObjectsTest

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

        Public Class プロパティたちTest : Inherits ChartObjectsTest

            <Test()> Public Sub Count_最初は0()
                Assert.That(sheet.ChartObjects.Count, [Is].EqualTo(0))
            End Sub

            <Test()> Public Sub Count_追加すれば増える()
                sheet.ChartObjects.Add(1, 2, 3, 4)
                sheet.ChartObjects.Add(5, 6, 7, 8)
                Assert.That(sheet.ChartObjects.Count, [Is].EqualTo(2))
            End Sub

        End Class

        'Public Class Item_Test : Inherits ChartObjectsTest

        '    <Test()> Public Sub AddしたChartObjectとName指定と同じインスタンスである()
        '        Dim ChartObject As WorkChartObject = workbook.ChartObjects.Add
        '        Assert.That(workbook.ChartObjects(ChartObject.Name), [Is].SameAs(ChartObject))
        '    End Sub

        '    <Test()> Public Sub Nameとindexとで同じインスタンスである()
        '        Dim ChartObject As WorkChartObject = workbook.ChartObjects.Add
        '        Assert.That(workbook.ChartObjects(ChartObject.Name), [Is].SameAs(workbook.ChartObjects(0)))
        '    End Sub

        'End Class

        Public Class ExcelObjectたちTest : Inherits ChartObjectsTest

            <Test()> Public Sub ChartObjectが閉じられること()
                Dim chartObject As ChartObject = sheet.ChartObjects.Add(2, 3, 4, 5)

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

    End Class
End Namespace