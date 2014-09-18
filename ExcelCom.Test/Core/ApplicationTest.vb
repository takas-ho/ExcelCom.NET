Imports NUnit.Framework

Namespace Core

    Public MustInherit Class ApplicationTest

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

        Public Class ExcelObjectたちTest : Inherits ApplicationTest

            <Test()> Public Sub Workbooksが閉じられること()
                Dim workbooks As Workbooks = sut.Workbooks

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub ActiveSheetが閉じられること()
                sut.Workbooks.Add()
                Dim activeSheet As Worksheet = sut.ActiveSheet

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub ActiveWorkbookが閉じられること()
                sut.Workbooks.Add()
                Dim activeWorkbook As Workbook = sut.ActiveWorkbook

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

        Public Class PropertyたちTest : Inherits ApplicationTest

            <Test()> Public Sub Calculation(<Values(Application.XlCalculation.xlCalculationAutomatic, Application.XlCalculation.xlCalculationManual, Application.XlCalculation.xlCalculationSemiautomatic)> ByVal value As Application.XlCalculation)
                Dim book As Workbook = sut.Workbooks.Add
                sut.Calculation = value
                Assert.That(sut.Calculation, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub DisplayAlerts(<Values(True, False)> ByVal value As Boolean)
                sut.DisplayAlerts = value
                Assert.That(sut.DisplayAlerts, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub ScreenUpdating(<Values(True, False)> ByVal value As Boolean)
                sut.ScreenUpdating = value
                Assert.That(sut.ScreenUpdating, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub Visible(<Values(True, False)> ByVal value As Boolean)
                sut.Visible = value
                Assert.That(sut.Visible, [Is].EqualTo(value))
            End Sub

        End Class

        Public Class ActiveSheetTest : Inherits ApplicationTest

            <Test()> Public Sub Bookを開いてないならnullを返す()
                Dim sheet As Worksheet = sut.ActiveSheet
                Assert.That(sheet, [Is].Null)
            End Sub

            <Test()> Public Sub Bookを開いてれば_有効値を返す()
                sut.Workbooks.Add()
                Dim sheet As Worksheet = sut.ActiveSheet
                Assert.That(sheet, [Is].Not.Null)
            End Sub

            <Test()> Public Sub Sheetsのインスタンスと同じである()
                Dim book As Workbook = sut.Workbooks.Add()
                Dim sheet As Worksheet = sut.ActiveSheet
                Assert.That(sheet, [Is].SameAs(book.Sheets(0)))
            End Sub

            <Test()> Public Sub Cellsに書込める()
                Dim workbook As Workbook = sut.Workbooks.Add()
                sut.ActiveSheet.Cells(1, 1).Value = "abc"
                Assert.That(workbook.Sheets(0).Cells(1, 1).Value, [Is].EqualTo("abc"))
            End Sub

        End Class

        Public Class ActiveWorkbookTest : Inherits ApplicationTest

            <Test()> Public Sub Bookを開いてないならnullを返す()
                Dim book As Workbook = sut.ActiveWorkbook
                Assert.That(book, [Is].Null)
            End Sub

            <Test()> Public Sub Bookを開いてれば_有効値を返す()
                sut.Workbooks.Add()
                Dim book As Workbook = sut.ActiveWorkbook
                Assert.That(book, [Is].Not.Null)
            End Sub

            <Test()> Public Sub Workbooksのインスタンスと同一である()
                sut.Workbooks.Add()
                Dim book As Workbook = sut.ActiveWorkbook
                Assert.That(book, [Is].SameAs(sut.Workbooks(0)))
            End Sub

            <Test()> Public Sub Addしたインスタンスと同一である()
                Dim book1 As Workbook = sut.Workbooks.Add()
                Dim book2 As Workbook = sut.Workbooks.Add()

                Dim activeBook As Workbook = sut.ActiveWorkbook

                Assert.That(activeBook, [Is].SameAs(book2))
            End Sub

        End Class

    End Class
End Namespace