Imports NUnit.Framework

Namespace Core

    Public MustInherit Class BordersTest

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

        Public Class ExcelObjectたちTest : Inherits BordersTest

            '<Test()> Public Sub Lineが閉じられること()
            '    Dim Border As Border = workbook.Sheets(0).Borders.AddLine(0, 10, 20, 30)

            '    sut.Dispose()
            '    sut = Nothing

            '    TestUtil.AssertNotExistsExcelPropcess()
            'End Sub

        End Class

        Public Class PropertyたちTest : Inherits BordersTest

            <Test()> Public Sub Count_最初は0超()
                Assert.That(sheet.Cells(3, 4).Borders.Count, [Is].GreaterThan(0))
            End Sub

        End Class

    End Class
End Namespace