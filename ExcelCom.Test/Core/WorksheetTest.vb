﻿Imports NUnit.Framework
Imports System.Reflection

Namespace Core

    Public MustInherit Class WorksheetTest

        Private sut As Application
        Private workbook As Workbook

        <SetUp()> Public Sub SetUp()
            sut = New Application
            workbook = sut.Workbooks.Add
        End Sub

        <TearDown()> Public Sub TearDown()
            If sut Is Nothing Then
                Return
            End If
            sut.Dispose()
        End Sub

        Public Class CellsTest : Inherits WorksheetTest

            <Test()> Public Sub Cellsの値をRangeと比較できる()
                workbook.Sheets.Item(0).Cells(0, 1).Value = "abc"

                Assert.That(workbook.Sheets.Item(0).Range("B1").Value, [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub Cellsの値をRangeと比較できる2()
                workbook.Sheets.Item(0).Cells(2, 0).Value = "aiueo"

                Assert.That(workbook.Sheets.Item(0).Range("A3").Value, [Is].EqualTo("aiueo"))
            End Sub

            <Test()> Public Sub Hoge()
                Dim start As Range = workbook.Sheets.Item(0).Cells(1, 1)

                Dim target As Range = workbook.Sheets.Item(0).Range(start, start)
                target.Value = "xyz"
                Assert.That(workbook.Sheets.Item(0).Range("B2").Value, [Is].EqualTo("xyz"))
            End Sub

        End Class

        Public Class CopyTest : Inherits WorksheetTest

            <Test()> Public Sub 引数一つなら_そのシートの手前にコピー挿入する()
                workbook.Sheets(0).Cells(2, 3).Value = "opq"
                Dim sheet2 As Worksheet = workbook.Sheets.Add
                Dim sheet As Worksheet = workbook.Sheets.Add
                sheet2.Cells(2, 3).Value = "xyz"
                sheet.Cells(2, 3).Value = "abc"
                Assert.That(workbook.Sheets(0).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(1).Cells(2, 3).Value, [Is].EqualTo("xyz"))
                Assert.That(workbook.Sheets(2).Cells(2, 3).Value, [Is].EqualTo("opq"))

                sheet.Copy(sheet2)

                Assert.That(workbook.Sheets(0).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(1).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(2).Cells(2, 3).Value, [Is].EqualTo("xyz"))
                Assert.That(workbook.Sheets(3).Cells(2, 3).Value, [Is].EqualTo("opq"))
            End Sub

            <Test()> Public Sub After引数なら_そのシートの後ろにコピー挿入する()
                workbook.Sheets(0).Cells(2, 3).Value = "opq"
                Dim sheet2 As Worksheet = workbook.Sheets.Add
                Dim sheet As Worksheet = workbook.Sheets.Add
                sheet2.Cells(2, 3).Value = "xyz"
                sheet.Cells(2, 3).Value = "abc"
                Assert.That(workbook.Sheets(0).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(1).Cells(2, 3).Value, [Is].EqualTo("xyz"))
                Assert.That(workbook.Sheets(2).Cells(2, 3).Value, [Is].EqualTo("opq"))

                sheet.Copy(after:=sheet2)

                Assert.That(workbook.Sheets(0).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(1).Cells(2, 3).Value, [Is].EqualTo("xyz"))
                Assert.That(workbook.Sheets(2).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(3).Cells(2, 3).Value, [Is].EqualTo("opq"))
            End Sub

            <Test()> Public Sub After引数なら_そのシートの後ろにコピー挿入する_境界()
                workbook.Sheets(0).Cells(2, 3).Value = "opq"
                Dim sheet2 As Worksheet = workbook.Sheets.Add
                Dim sheet As Worksheet = workbook.Sheets.Add
                sheet2.Cells(2, 3).Value = "xyz"
                sheet.Cells(2, 3).Value = "abc"
                Assert.That(workbook.Sheets(0).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(1).Cells(2, 3).Value, [Is].EqualTo("xyz"))
                Assert.That(workbook.Sheets(2).Cells(2, 3).Value, [Is].EqualTo("opq"))

                sheet.Copy(after:=workbook.Sheets(2))

                Assert.That(workbook.Sheets(0).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(1).Cells(2, 3).Value, [Is].EqualTo("xyz"))
                Assert.That(workbook.Sheets(2).Cells(2, 3).Value, [Is].EqualTo("opq"))
                Assert.That(workbook.Sheets(3).Cells(2, 3).Value, [Is].EqualTo("abc"))
            End Sub

            <Test()> Public Sub 引数なしなら_新Bookにコピーされる()
                Dim sheet2 As Worksheet = workbook.Sheets.Add
                Dim sheet As Worksheet = workbook.Sheets.Add
                sheet2.Cells(2, 3).Value = "xyz"
                sheet.Cells(2, 3).Value = "abc"

                sheet.Copy()

                Assert.That(sut.ActiveWorkbook, [Is].Not.SameAs(workbook), "新bookになる")
                Assert.That(sut.ActiveWorkbook.Sheets.Count, [Is].EqualTo(1), "コピーしたシートだけ")
                Assert.That(sut.ActiveWorkbook.Sheets(0).Cells(2, 3).Value, [Is].EqualTo("abc"))
            End Sub

        End Class

        Public Class MoveTest : Inherits WorksheetTest

            <Test()> Public Sub 引数一つなら_そのシートの手前に移動する()
                Dim sheet3 As Worksheet = workbook.Sheets(0)
                Dim sheet2 As Worksheet = workbook.Sheets.Add
                Dim sheet As Worksheet = workbook.Sheets.Add
                sheet3.Cells(2, 3).Value = "opq"
                sheet2.Cells(2, 3).Value = "xyz"
                sheet.Cells(2, 3).Value = "abc"
                Assert.That(workbook.Sheets(0).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(1).Cells(2, 3).Value, [Is].EqualTo("xyz"))
                Assert.That(workbook.Sheets(2).Cells(2, 3).Value, [Is].EqualTo("opq"))

                sheet3.Move(sheet2)

                Assert.That(workbook.Sheets(0).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(1).Cells(2, 3).Value, [Is].EqualTo("opq"))
                Assert.That(workbook.Sheets(2).Cells(2, 3).Value, [Is].EqualTo("xyz"))

                Assert.That(workbook.Sheets(0), [Is].SameAs(sheet))
                Assert.That(workbook.Sheets(1), [Is].SameAs(sheet3))
                Assert.That(workbook.Sheets(2), [Is].SameAs(sheet2))
            End Sub

            <Test()> Public Sub After引数なら_そのシートの後ろに移動する()
                Dim sheet3 As Worksheet = workbook.Sheets(0)
                Dim sheet2 As Worksheet = workbook.Sheets.Add
                Dim sheet As Worksheet = workbook.Sheets.Add
                sheet3.Cells(2, 3).Value = "opq"
                sheet2.Cells(2, 3).Value = "xyz"
                sheet.Cells(2, 3).Value = "abc"
                Assert.That(workbook.Sheets(0).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(1).Cells(2, 3).Value, [Is].EqualTo("xyz"))
                Assert.That(workbook.Sheets(2).Cells(2, 3).Value, [Is].EqualTo("opq"))

                sheet.Move(after:=sheet2)

                Assert.That(workbook.Sheets(0).Cells(2, 3).Value, [Is].EqualTo("xyz"))
                Assert.That(workbook.Sheets(1).Cells(2, 3).Value, [Is].EqualTo("abc"))
                Assert.That(workbook.Sheets(2).Cells(2, 3).Value, [Is].EqualTo("opq"))

                Assert.That(workbook.Sheets(0), [Is].SameAs(sheet2))
                Assert.That(workbook.Sheets(1), [Is].SameAs(sheet))
                Assert.That(workbook.Sheets(2), [Is].SameAs(sheet3))
            End Sub

        End Class

        Public Class ProtectTest : Inherits WorksheetTest

            <Test()> Public Sub 保護すると変更は出来ない()
                Dim sheet As Worksheet = workbook.Sheets.Add
                sheet.Cells(2, 2).Value = "a"
                sheet.Protect()
                Try
                    sheet.Cells(2, 2).Value = "b"
                    Assert.Fail()
                Catch ex As TargetInvocationException
                    Assert.That(ex.InnerException.Message, [Is].StringContaining("変更しようとしているセルまたはグラフは保護されているため"))
                End Try
            End Sub

            <Test()> Public Sub 保護解除すれば_無事に変更できる()
                Dim sheet As Worksheet = workbook.Sheets.Add
                sheet.Cells(2, 2).Value = "a"
                sheet.Protect()
                sheet.Unprotect()

                sheet.Cells(2, 2).Value = "b"

                Assert.That(sheet.Cells(2, 2).Value, [Is].EqualTo("b"))
            End Sub

        End Class

        Public Class CalculateTest : Inherits WorksheetTest

            <Test()> Public Sub Calculateの実行テスト()
                Dim sheet As Worksheet = workbook.Sheets.add
                sheet.Calculate()
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

            <Test()> Public Sub Shapesが閉じられること()
                Dim shapes As Shapes = workbook.Sheets(0).Shapes

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub ChartObjectsが閉じられること()
                Dim value As ChartObjects = workbook.Sheets(0).ChartObjects

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

        Public Class PropertyたちTest : Inherits WorksheetTest

            <Test()> Public Sub Visible(<Values(Worksheet.XlSheetVisibility.xlSheetHidden, Worksheet.XlSheetVisibility.xlSheetVeryHidden, Worksheet.XlSheetVisibility.xlSheetVisible)> ByVal value As Worksheet.XlSheetVisibility)
                Dim sheet1 As Worksheet = workbook.Sheets.Add()
                Dim sheet2 As Worksheet = workbook.Sheets.Add(after:=sheet1)
                sheet2.Select()

                workbook.Sheets(0).Visible = value
                Assert.That(workbook.Sheets(0).Visible, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub CodeName()
                Dim sheet1 As Worksheet = workbook.Sheets.Add()
                Assert.That(sheet1.CodeName, [Is].EqualTo(""), "追加したシートのコード名は空文字")
            End Sub

            <Test()> Public Sub DisplayPageBreaks(<Values(True, False)> ByVal value As Boolean)
                Dim sheet1 As Worksheet = workbook.Sheets.Add()
                sheet1.DisplayPageBreaks = value
                Assert.That(sheet1.DisplayPageBreaks, [Is].EqualTo(value))
            End Sub

            <Test()> Public Sub ProtectContents()
                Dim sheet1 As Worksheet = workbook.Sheets.Add()
                Assert.That(sheet1.ProtectContents, [Is].False, "追加直後はfalse")
            End Sub

            <Test()> Public Sub ProtectDrawingObjects()
                Dim sheet1 As Worksheet = workbook.Sheets.Add()
                Assert.That(sheet1.ProtectDrawingObjects, [Is].False, "追加直後はfalse")
            End Sub

            <Test()> Public Sub FilterMode()
                Dim sheet1 As Worksheet = workbook.Sheets.Add()
                Assert.That(sheet1.FilterMode, [Is].False, "追加直後はfalse")
            End Sub

            <Test()> Public Sub ProtectionMode()
                Dim sheet1 As Worksheet = workbook.Sheets.Add()
                Assert.That(sheet1.ProtectionMode, [Is].False, "追加直後はfalse")
            End Sub

        End Class

    End Class
End Namespace