Imports NUnit.Framework

Namespace Core

    Public MustInherit Class RangeTest

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

        Public Class CellsTest : Inherits RangeTest

            <Test()> Public Sub A1のCells00をCells00と比較できる()
                workbook.Sheets.Item(0).Range("A1").Cells(0, 0).Value = "aiueo"
                Assert.That(workbook.Sheets.Item(0).Cells(0, 0).Value, [Is].EqualTo("aiueo"))
            End Sub

            <Test()> Public Sub D2のCells00をCells13と比較できる()
                workbook.Sheets.Item(0).Range("D2").Cells(0, 0).Value = "xyz"
                Assert.That(workbook.Sheets.Item(0).Cells(1, 3).Value, [Is].EqualTo("xyz"))
            End Sub

        End Class

        Public Class FindTest : Inherits RangeTest

            <Test()> Public Sub 見つけたセルを返す(<Values(0, 4, 100)> ByVal row As Integer, <Values(0, 34, 100)> ByVal column As Integer)
                sheet.Cells(row, column).Value = "りんご"
                Dim result As Range = sheet.Cells.Find("りんご")

                Assert.That(result, [Is].Not.Null)
                Assert.That(result.Row, [Is].EqualTo(row))
                Assert.That(result.Column, [Is].EqualTo(column))
            End Sub

            <Test()> Public Sub 見つからなければnull()
                sheet.Cells(3, 4).Value = "りんご"
                Dim result As Range = sheet.Cells.Find("ばなな")

                Assert.That(result, [Is].Null)
            End Sub

        End Class

        Public Class InsertTest : Inherits RangeTest

            <Test()> Public Sub 列を挿入できる()
                workbook.Sheets(0).Cells(0, 1).Value = "a01"
                workbook.Sheets(0).Cells(0, 2).Value = "b02"

                workbook.Sheets(0).Columns(2).Insert()

                Assert.That(workbook.Sheets(0).Cells(0, 1).Value, [Is].EqualTo("a01"))
                Assert.That(workbook.Sheets(0).Cells(0, 3).Value, [Is].EqualTo("b02"), "列1と列2の間に挿入したから")
            End Sub

            <Test()> Public Sub 行を挿入できる()
                workbook.Sheets(0).Cells(2, 0).Value = "a20"
                workbook.Sheets(0).Cells(3, 0).Value = "b30"

                workbook.Sheets(0).Rows(3).Insert()

                Assert.That(workbook.Sheets(0).Cells(2, 0).Value, [Is].EqualTo("a20"))
                Assert.That(workbook.Sheets(0).Cells(4, 0).Value, [Is].EqualTo("b30"), "行2と行3列の間に挿入したから")
            End Sub

            <Test()> Public Sub コピー列を挿入できる()
                workbook.Sheets(0).Cells(0, 1).Value = "a01"
                workbook.Sheets(0).Cells(0, 2).Value = "b02"

                workbook.Sheets(0).Columns(2).Copy()
                workbook.Sheets(0).Columns(1).Insert()

                Assert.That(workbook.Sheets(0).Cells(0, 1).Value, [Is].EqualTo("b02"), "列2だった値が挿入された")
                Assert.That(workbook.Sheets(0).Cells(0, 2).Value, [Is].EqualTo("a01"))
                Assert.That(workbook.Sheets(0).Cells(0, 3).Value, [Is].EqualTo("b02"))
            End Sub

        End Class

        Public Class DeleteTest : Inherits RangeTest

            <Test()> Public Sub 列を削除できる()
                workbook.Sheets(0).Cells(0, 1).Value = "a01"
                workbook.Sheets(0).Cells(0, 2).Value = "b02"
                workbook.Sheets(0).Cells(0, 3).Value = "c03"

                workbook.Sheets(0).Columns(2).Delete()

                Assert.That(workbook.Sheets(0).Cells(0, 1).Value, [Is].EqualTo("a01"))
                Assert.That(workbook.Sheets(0).Cells(0, 2).Value, [Is].EqualTo("c03"), "b02の列が削除された")
            End Sub

            <Test()> Public Sub 行を削除できる()
                workbook.Sheets(0).Cells(2, 0).Value = "a20"
                workbook.Sheets(0).Cells(3, 0).Value = "b30"
                workbook.Sheets(0).Cells(4, 0).Value = "c40"

                workbook.Sheets(0).Rows(3).Delete()

                Assert.That(workbook.Sheets(0).Cells(2, 0).Value, [Is].EqualTo("a20"))
                Assert.That(workbook.Sheets(0).Cells(3, 0).Value, [Is].EqualTo("c40"), "b30の行が削除された")
            End Sub

        End Class

        Public Class ExcelObjectたちTest : Inherits RangeTest

            <Test()> Public Sub Columnsが閉じられること()
                Dim columns As Range = workbook.Sheets.Item(0).Range("A1").Columns

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub Rowsが閉じられること()
                Dim rows As Range = workbook.Sheets.Item(0).Range("A1").Rows

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub Cellsが閉じられること()
                Dim cells As Range = workbook.Sheets.Item(0).Range("A1").Cells

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub Itemが閉じられること()
                Dim item As Range = workbook.Sheets.Item(0).Range("A1").Item(0, 0)

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

        Public Class PropertyたちTest : Inherits RangeTest

            <Test()> Public Sub NumberFormatLocal_Valueへ文字列10を設定すると_Double値になる()

                workbook.Sheets(0).Cells(0, 0).Value = "10"

                Assert.That(workbook.Sheets(0).Cells(0, 0).Value, [Is].EqualTo(10D))
            End Sub

            <Test()> Public Sub NumberFormatLocal_Valueへ文字列10を設定すると_Double値になる_がNumberFormatLocal_at_にすれば文字列になる()

                workbook.Sheets(0).Cells(0, 0).NumberFormatLocal = "@"
                workbook.Sheets(0).Cells(0, 0).Value = "10"

                Assert.That(workbook.Sheets(0).Cells(0, 0).Value, [Is].EqualTo("10"))
            End Sub

            <Test()> Public Sub Column_(<Values(4, 23)> ByVal column As Integer)
                Assert.That(workbook.Sheets(0).Cells(0, column).Column, [Is].EqualTo(column))
            End Sub

            <Test()> Public Sub Row_(<Values(2, 34)> ByVal row As Integer)
                Assert.That(workbook.Sheets(0).Cells(row, 0).Row, [Is].EqualTo(row))
            End Sub

            <Test()> Public Sub ColumnWidth_(<Values(2, 34)> ByVal column As Integer)
                sheet.Columns(column).ColumnWidth = 10.4R
                Assert.That(sheet.Columns(column).ColumnWidth, [Is].EqualTo(10.43R), "10.4は10.43になる. 精度の問題")
            End Sub

        End Class

    End Class
End Namespace