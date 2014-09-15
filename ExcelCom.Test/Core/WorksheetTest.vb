Imports NUnit.Framework

Namespace Core

    Public MustInherit Class WorksheetTest

        Private sut As Application

        <SetUp()> Public Sub SetUp()
            sut = New Application
        End Sub

        <TearDown()> Public Sub TearDown()
            If sut Is Nothing Then
                Return
            End If
            sut.Dispose()
        End Sub

        Public Class CellsTest : Inherits WorksheetTest

            <Test()> Public Sub Cellsの値をRangeと比較できる()
                Dim workbook As Workbook = sut.Workbooks.Add
                workbook.Sheets.Item(0).Cells(0, 1).Value = "abc"

                Assert.That(workbook.Sheets.Item(0).Range("B1").Value, [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub Cellsの値をRangeと比較できる2()
                Dim workbook As Workbook = sut.Workbooks.Add
                workbook.Sheets.Item(0).Cells(2, 0).Value = "aiueo"

                Assert.That(workbook.Sheets.Item(0).Range("A3").Value, [Is].EqualTo("aiueo"))
            End Sub

            <Test()> Public Sub Hoge()
                Dim workbook As Workbook = sut.Workbooks.Add
                Dim start As Range = workbook.Sheets.Item(0).Cells(1, 1)

                Dim target As Range = workbook.Sheets.Item(0).Range(start, start)
                target.Value = "xyz"
                Assert.That(workbook.Sheets.Item(0).Range("B2").Value, [Is].EqualTo("xyz"))
            End Sub

        End Class

        Public Class ExcelObjectたちTest : Inherits WorksheetTest

            <Test()> Public Sub Cellsが閉じられること()
                Dim workbook As Workbook = sut.Workbooks.Add
                Dim cells As Range = workbook.Sheets.Item(0).Cells(0, 0)

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

    End Class
End Namespace