Imports NUnit.Framework

Namespace Core

    Public MustInherit Class RangeTest

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

        Public Class HogeTest : Inherits RangeTest

            <Test()> Public Sub Hoge()
                Dim workbook As Workbook = sut.Workbooks.Add
                workbook.Sheets.Item(0).Range("A1").Cells(0, 0).Value = "aiueo"
                Assert.That(workbook.Sheets.Item(0).Cells(0, 0).Value, [Is].EqualTo("aiueo"))
            End Sub

            <Test()> Public Sub Hoge2()
                Dim workbook As Workbook = sut.Workbooks.Add
                workbook.Sheets.Item(0).Range("D2").Cells(0, 0).Value = "xyz"
                Assert.That(workbook.Sheets.Item(0).Cells(1, 3).Value, [Is].EqualTo("xyz"))
            End Sub

        End Class

        Public Class ExcelObjectたちTest : Inherits RangeTest

            <Test()> Public Sub Columnsが閉じられること()
                Dim workbook As Workbook = sut.Workbooks.Add
                Dim columns As Range = workbook.Sheets.Item(0).Range("A1").Columns

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub Rowsが閉じられること()
                Dim workbook As Workbook = sut.Workbooks.Add
                Dim rows As Range = workbook.Sheets.Item(0).Range("A1").Rows

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

    End Class
End Namespace