﻿Imports NUnit.Framework
Imports System.Reflection

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

        Public Class AutoFilterTest : Inherits RangeTest

            <Test()> Public Sub AutoFilter呼出ができる()
                sheet.Cells(0, 0).Value = "3"
                sheet.Cells(0, 0).AutoFilter(columnIndex:=0, criteria1:="6")
            End Sub

            <Test()> Public Sub 値がないとエラーになる()
                Try
                    sheet.Cells(0, 0).AutoFilter(columnIndex:=0, criteria1:="6")
                    Assert.Fail()
                Catch expected As TargetInvocationException
                    Assert.That(expected.InnerException.Message, [Is].EqualTo("Range クラスの AutoFilter メソッドが失敗しました。"))
                End Try
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

        Public Class AutoFitTest : Inherits RangeTest

            <Test()> Public Sub aiueoの文字列をAutoFitしたら_初期幅より狭くなる()
                Dim baseWidth As Double = sheet.Columns(2).ColumnWidth
                sheet.Cells(2, 2).Value = "aiueo"

                sheet.Columns(2).AutoFit()

                Assert.That(sheet.Columns(2).ColumnWidth, [Is].LessThan(baseWidth))
            End Sub

        End Class

        Public Class SpecialCellsTest : Inherits RangeTest

            <Test()> Public Sub xlCellTypeLastCell_データ入力最終セルを返す()
                sheet.Cells(2, 2).Value = "aaa"
                sheet.Cells(3, 3).Value = "bbb"
                Dim range As Range = sheet.Cells.SpecialCells(range.XlCellType.xlCellTypeLastCell)

                Assert.That(range.Row, [Is].EqualTo(3))
            End Sub

        End Class

        Public Class MergeTest : Inherits RangeTest

            <Test()> Public Sub Mergeでセル結合できる()
                sheet.Range("A1:C3").Merge()
                Assert.That(sheet.Range("A1:C3").MergeCells, [Is].True)
                Assert.That(sheet.Range("A1").MergeCells, [Is].True)
                Assert.That(sheet.Range("B2").MergeCells, [Is].True)
                Assert.That(sheet.Range("C3").MergeCells, [Is].True)
            End Sub

            <Test()> Public Sub UnMergeでセル結合を解除する()
                sheet.Range("A1:C3").Merge()
                sheet.Range("A1:C3").UnMerge()
                Assert.That(sheet.Range("A1:C3").MergeCells, [Is].False)
            End Sub

            <Test()> Public Sub UnMergeでセル結合を解除する_任意の一つのセルだけですべて解除される()
                sheet.Range("A1:C3").Merge()
                sheet.Range("A1").UnMerge()
                Assert.That(sheet.Range("A1").MergeCells, [Is].False)
                Assert.That(sheet.Range("B2").MergeCells, [Is].False)
                Assert.That(sheet.Range("C3").MergeCells, [Is].False)
            End Sub

            <Test()> Public Sub MergeAreaでセル結合範囲を取得できる_任意の一つのセルだけで取得できる()
                sheet.Range("A1:C3").Merge()
                Dim actual As Range = sheet.Range("A1").MergeArea
                Assert.That(actual.Columns.Count, [Is].EqualTo(3))
                Assert.That(actual.Rows.Count, [Is].EqualTo(3))
            End Sub

        End Class

        Public Class BorderAroundTest : Inherits RangeTest

            <Test()> Public Sub 設定できたらtrueになる()
                Assert.That(sheet.Range("A1:C3").BorderAround(Weight:=Range.XlBorderWeight.xlMedium), [Is].True)
            End Sub

        End Class

        Public Class ClearContentsTest : Inherits RangeTest

            <Test()> Public Sub Clear_数式文字列装飾をクリア(<Values("aaaa", "=D4")> ByVal value As String)
                sheet.Cells(2, 2).Value = value
                Assert.That(sheet.Range("A1:C3").Clear, [Is].True)
                Assert.That(sheet.Cells(2, 2).Value, [Is].Null)
            End Sub

            <Test()> Public Sub ClearComments_数式文字列装飾をクリア(<Values("aaaa")> ByVal value As String)
                sheet.Cells(2, 2).AddComment(value)
                sheet.Cells(2, 2).ClearComments()
                Assert.That(sheet.Cells(2, 2).Comment, [Is].Null)
            End Sub

            <Test()> Public Sub ClearContents_数式文字列をクリア(<Values("aaaa", "=D4")> ByVal value As String)
                sheet.Cells(2, 2).Value = value
                Assert.That(sheet.Range("A1:C3").ClearContents, [Is].True)
                Assert.That(sheet.Cells(2, 2).Value, [Is].Null)
            End Sub

            <Test()> Public Sub ClearFormats_書式のクリア()
                sheet.Cells(2, 2).Value = "aaaa"
                sheet.Cells(2, 2).Font.Bold = True
                Assert.That(sheet.Range("A1:C3").ClearFormats, [Is].True)
                Assert.That(sheet.Cells(2, 2).Value, [Is].EqualTo("aaaa"))
                Assert.That(sheet.Cells(2, 3).Font.Bold, [Is].False, "書式がクリアされる")
            End Sub

        End Class

        Public Class CalculateTest : Inherits RangeTest

            <Test()> Public Sub Calculateの実行テスト()
                sheet.Range("A1:C3").Calculate()
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

            <Test()> Public Sub Fontが閉じられること()
                Dim item As Font = sheet.Cells(2, 3).Font

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub Interiorが閉じられること()
                Dim item As Interior = sheet.Cells(2, 3).Interior

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub Bordersが閉じられること()
                Dim item As Borders = sheet.Cells(2, 3).Borders

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub AddCommentが閉じられること()
                Dim item As Comment = sheet.Cells(2, 3).AddComment("aiueo")

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub SpecialCellsのRangeが閉じられること()
                sheet.Cells(2, 3).Value = "aaa"
                Dim range As Range = sheet.Cells.SpecialCells(range.XlCellType.xlCellTypeLastCell)

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub MergeAreaのRangeが閉じられること()
                sheet.Range("A1:B5").Merge()
                Dim value As Range = sheet.Range("A2").MergeArea

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub EntireColumnのRangeが閉じられること()
                Dim value As Range = sheet.Range("C5").EntireColumn()

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub EntireRowのRangeが閉じられること()
                Dim value As Range = sheet.Range("C5").EntireRow()

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub AddCommentのCommentが閉じられること()
                Dim value As Comment = sheet.Cells(2, 3).AddComment("aiueo")

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub CommentのCommentが閉じられること()
                sheet.Cells(2, 3).AddComment("aiueo")
                Dim value As Comment = sheet.Cells(2, 3).Comment

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

            <Test()> Public Sub Value2_日付のシリアル値を取得できる()

                sheet.Cells(0, 0).Value = "2014/1/1"

                Assert.That(sheet.Cells(0, 0).Value2, [Is].EqualTo(41640.0R))
            End Sub

            <Test()> Public Sub Column_(<Values(4, 23)> ByVal column As Integer)
                Assert.That(workbook.Sheets(0).Cells(0, column).Column, [Is].EqualTo(column))
            End Sub

            <Test()> Public Sub Row_(<Values(2, 34)> ByVal row As Integer)
                Assert.That(workbook.Sheets(0).Cells(row, 0).Row, [Is].EqualTo(row))
            End Sub

            <Test()> Public Sub ColumnWidth_(<Values(2, 34)> ByVal column As Integer)
                sheet.Columns(column).ColumnWidth = 10.4R
                Assert.That(sheet.Columns(column).ColumnWidth, [Is].EqualTo(10.43R), "10.4は10.43になる. Excel内部の問題")
            End Sub

            <Test()> Public Sub RowHeight_(<Values(2, 34)> ByVal column As Integer)
                sheet.Rows(column).RowHeight = 12.3R
                Assert.That(sheet.Rows(column).RowHeight, [Is].EqualTo(12.5R), "12.3は12.5になる. Excel内部の問題")
            End Sub

            <Test()> Public Sub HorizontalAlignment_(<Values(Range.XlHAlign.xlHAlignCenter, Range.XlHAlign.xlHAlignCenterAcrossSelection, Range.XlHAlign.xlHAlignDistributed, Range.XlHAlign.xlHAlignFill, Range.XlHAlign.xlHAlignGeneral, Range.XlHAlign.xlHAlignJustify, Range.XlHAlign.xlHAlignLeft, Range.XlHAlign.xlHAlignRight)> ByVal value As Range.XlHAlign)
                sheet.Cells(1, 2).HorizontalAlignment = value
                Assert.That(sheet.Cells(1, 2).HorizontalAlignment, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub VerticalAlignment_(<Values(Range.XlVAlign.xlVAlignBottom, Range.XlVAlign.xlVAlignCenter, Range.XlVAlign.xlVAlignDistributed, Range.XlVAlign.xlVAlignJustify, Range.XlVAlign.xlVAlignTop)> ByVal value As Range.XlVAlign)
                sheet.Cells(1, 2).VerticalAlignment = value
                Assert.That(sheet.Cells(1, 2).VerticalAlignment, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub Wrap_(<Values(False, True)> ByVal value As Boolean)
                sheet.Cells(1, 1).WrapText = value
                Assert.That(sheet.Cells(1, 1).WrapText, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub ShrinkToFit_(<Values(False, True)> ByVal value As Boolean)
                sheet.Cells(1, 1).ShrinkToFit = value
                Assert.That(sheet.Cells(1, 1).ShrinkToFit, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub MergeCells_(<Values(False, True)> ByVal value As Boolean)
                sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, 3)).MergeCells = value
                Assert.That(sheet.Range(sheet.Cells(1, 1), sheet.Cells(1, 3)).MergeCells, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub Formula_(<Values("=A1", "=SUM(A1:G1)")> ByVal value As String)
                sheet.Cells(1, 1).Formula = value
                Assert.That(sheet.Cells(1, 1).Formula, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub FormulaLocal_(<Values("=A1", "=SUM(A1:G1)")> ByVal value As String)
                sheet.Cells(1, 1).FormulaLocal = value
                Assert.That(sheet.Cells(1, 1).FormulaLocal, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub FormulaR1C1_(<Values("=R[3]C", "=R[1]C[2]")> ByVal value As String)
                sheet.Cells(1, 1).FormulaR1C1 = value
                Assert.That(sheet.Cells(1, 1).FormulaR1C1, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub FormulaR1C1Local_(<Values("=R[3]C", "=R[1]C[2]")> ByVal value As String)
                sheet.Cells(1, 1).FormulaR1C1Local = value
                Assert.That(sheet.Cells(1, 1).FormulaR1C1Local, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub Orientation_(<Values(Range.XlOrientation.xlDownward, Range.XlOrientation.xlHorizontal, Range.XlOrientation.xlUpward, Range.XlOrientation.xlVertical)> ByVal value As Range.XlOrientation)
                sheet.Cells(1, 1).Orientation = value
                Assert.That(sheet.Cells(1, 1).Orientation, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub Hidden_(<Values(False, True)> ByVal value As Boolean)
                sheet.Rows(3).Hidden = value
                Assert.That(sheet.Rows(3).Hidden, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub HasFormula_()
                sheet.Cells(3, 3).Formula = "=C3+D3"
                Assert.That(sheet.Cells(3, 3).HasFormula, [Is].True)
                Assert.That(sheet.Cells(1, 1).HasFormula, [Is].False)
            End Sub

            <Test()> Public Sub Width_()
                Const EPSILON As Double = 0.001
                If Math.Abs(sut.StandardFontSize - 10.0R) < EPSILON Then
                    Assert.That(sheet.Cells(3, 3).Width, [Is].EqualTo(48.0R))
                ElseIf Math.Abs(sut.StandardFontSize - 11.0R) < EPSILON Then
                    Assert.That(sheet.Cells(3, 3).Width, [Is].EqualTo(54.0R))
                Else
                    Assert.Fail("フォントサイズ" & sut.StandardFontSize & " は未対応")
                End If
            End Sub

        End Class

    End Class
End Namespace