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

            <Test()> Public Sub Windowsが閉じられること()
                Dim windows As Windows = sut.Windows

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub ActiveWindowが閉じられること()
                sut.Workbooks.Add()
                Dim window As Window = sut.ActiveWindow

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub Chartsが閉じられること()
                sut.Workbooks.Add()
                Dim value As Charts = sut.Charts

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

            <Test()> Public Sub CutCopyMode_初期状態はFalse()
                Dim book As Workbook = sut.Workbooks.Add
                Assert.That(sut.CutCopyMode, [Is].EqualTo(Application.XlCutCopyMode.False))
            End Sub

            <Test()> Public Sub CutCopyMode_CopyしたらXlCopy()
                Dim book As Workbook = sut.Workbooks.Add
                book.Sheets(0).Cells(0, 0).Copy()
                Assert.That(sut.CutCopyMode, [Is].EqualTo(Application.XlCutCopyMode.xlCopy))
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

            <Test()> Public Sub Version()
                ' Excel2003だと、"11.0"
                Assert.That(sut.Version, [Is].Not.Empty)
            End Sub

            <Test()> Public Sub Value()
                Assert.That(sut.Value, [Is].EqualTo("Microsoft Excel"))
            End Sub

            <Test()> Public Sub StandardFontSize()
                ' Excelの初期フォントサイズ 10とか11とか
                Assert.That(sut.StandardFontSize, [Is].GreaterThan(0.0R))
            End Sub

            <Test()> Public Sub StandardFont()
                Assert.That(sut.StandardFont, [Is].Not.Empty)
            End Sub

            <Test()> Public Sub StartupPath()
                Assert.That(sut.StartupPath, [Is].Not.Empty)
            End Sub

            <Test()> Public Sub UserName()
                Assert.That(sut.UserName, [Is].Not.Empty, "ツール | オプション | 全般タブ のユーザー名 ex.'山田 太郎'")
            End Sub

            <Test()> Public Sub ProductCode()
                Assert.That(sut.ProductCode, [Is].Not.Empty, "こんなやつ→{90110411-6000-11D3-8CFE-0150048383C9}")
            End Sub

            <Test()> Public Sub Caption()
                Assert.That(sut.Caption, [Is].EqualTo("Microsoft Excel"))
            End Sub

            <Test()> Public Sub Name()
                Assert.That(sut.Name, [Is].EqualTo("Microsoft Excel"))
            End Sub

        End Class

        Public Class ActiveCellTest : Inherits ApplicationTest

            <Test()> Public Sub Bookを開いてないならnullを返す()
                Dim cell As Range = sut.ActiveCell
                Assert.That(cell, [Is].Null)
            End Sub

            <Test()> Public Sub Bookを開いてれば_有効値を返す()
                sut.Workbooks.Add()
                Dim cell As Range = sut.ActiveCell
                Assert.That(cell, [Is].Not.Null)
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

        Public Class ActiveWindowTest : Inherits ApplicationTest

            <Test()> Public Sub Bookを開いてないならnullを返す()
                Dim window As Window = sut.ActiveWindow
                Assert.That(window, [Is].Null)
            End Sub

            <Test()> Public Sub Bookを開いてれば_有効値を返す()
                sut.Workbooks.Add()
                Dim window As Window = sut.ActiveWindow
                Assert.That(window, [Is].Not.Null)
            End Sub

            <Test()> Public Sub Windowsのインスタンスと同一である()
                sut.Workbooks.Add()
                Dim window As Window = sut.ActiveWindow
                Assert.That(window, [Is].SameAs(sut.Windows(0)))
            End Sub

        End Class

    End Class
End Namespace