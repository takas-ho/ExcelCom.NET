﻿Imports NUnit.Framework

Namespace Core

    Public MustInherit Class ShapesTest

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

        Public Class AddLineTest : Inherits ShapesTest

            <Test()> Public Sub Addすれば1増える()
                Dim baseCount As Integer = workbook.Sheets(0).Shapes.Count
                workbook.Sheets(0).Shapes.AddLine(10, 20, 30, 40)
                Assert.That(workbook.Sheets.Count, [Is].EqualTo(baseCount + 1))
            End Sub

            <Test()> Public Sub Addしたら最後に追加される_同じインスタンスである()
                Dim line1 As Shape = workbook.Sheets(0).Shapes.AddLine(10, 20, 30, 40)
                Dim line2 As Shape = workbook.Sheets(0).Shapes.AddLine(50, 60, 70, 80)
                Assert.That(workbook.Sheets(0).Shapes(0).Name, [Is].EqualTo(line1.Name))
                Assert.That(workbook.Sheets(0).Shapes(1).Name, [Is].EqualTo(line2.Name), "2番目に追加したら2番目にある")
            End Sub

        End Class

        Public Class Item_Test : Inherits ShapesTest

            <Test()> Public Sub AddしたShapeとName指定と同じインスタンスである()
                Dim line As Shape = workbook.Sheets(0).Shapes.AddLine(10, 20, 30, 40)
                Assert.That(workbook.Sheets(0).Shapes(line.Name), [Is].SameAs(line))
            End Sub

            <Test()> Public Sub Nameとindexとで同じインスタンスである()
                Dim line As Shape = workbook.Sheets(0).Shapes.AddLine(10, 20, 30, 40)
                'Assert.That(workbook.Sheets(sheet.Name), [Is].SameAs(workbook.Sheets(0)))
                Assert.That(workbook.Sheets(0).Shapes(line.Name), [Is].SameAs(workbook.Sheets(0).Shapes(0)))
            End Sub

        End Class

        Public Class ExcelObjectたちTest : Inherits ShapesTest

            <Test()> Public Sub Lineが閉じられること()
                Dim shape As Shape = workbook.Sheets(0).Shapes.AddLine(0, 10, 20, 30)

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

        'Public Class PropertyたちTest : Inherits ShapesTest

        '    <Test()> Public Sub Count_最初は0超_シート数はローカルPCの設定で変わる()
        '        Assert.That(workbook.Sheets.Count, [Is].GreaterThan(0))
        '    End Sub

        'End Class

    End Class
End Namespace