Imports NUnit.Framework

Namespace Core

    Public MustInherit Class ShapeTest

        Private sut As Application
        Private workbook As Workbook
        Private sheet As Worksheet

        <SetUp()> Public Sub SetUp()
            sut = New Application
            workbook = sut.Workbooks.Add
            sheet = workbook.Sheets.Add()
        End Sub

        <TearDown()> Public Sub TearDown()
            If sut Is Nothing Then
                Return
            End If
            sut.Dispose()
        End Sub

        'Public Class CellsTest : Inherits ShapeTest

        '    <Test()> Public Sub Cellsの値をRangeと比較できる()
        '        workbook.Sheets.Item(0).Cells(0, 1).Value = "abc"

        '        Assert.That(workbook.Sheets.Item(0).Range("B1").Value, [Is].EqualTo("abc"))
        '    End Sub

        '    <Test()> Public Sub Cellsの値をRangeと比較できる2()
        '        workbook.Sheets.Item(0).Cells(2, 0).Value = "aiueo"

        '        Assert.That(workbook.Sheets.Item(0).Range("A3").Value, [Is].EqualTo("aiueo"))
        '    End Sub

        '    <Test()> Public Sub Hoge()
        '        Dim start As Range = workbook.Sheets.Item(0).Cells(1, 1)

        '        Dim target As Range = workbook.Sheets.Item(0).Range(start, start)
        '        target.Value = "xyz"
        '        Assert.That(workbook.Sheets.Item(0).Range("B2").Value, [Is].EqualTo("xyz"))
        '    End Sub

        'End Class

        Public Class ExcelObjectたちTest : Inherits ShapeTest

            '<Test()> Public Sub Cellsが閉じられること()
            '    Dim cells As Range = workbook.Sheets(0).Shapes

            '    sut.Dispose()
            '    sut = Nothing

            '    TestUtil.AssertNotExistsExcelPropcess()
            'End Sub

        End Class

        Public Class PropertyたちTest : Inherits ShapeTest

            <Test()> Public Sub Height_()
                Dim shape As Shape = sheet.Shapes.AddLine(10.0F, 20.0F, 30.0F, 40.0F)
                shape.Height = 20.0F
                Assert.That(shape.Height, [Is].EqualTo(20.25F))
            End Sub

            <Test()> Public Sub Width_()
                Dim shape As Shape = sheet.Shapes.AddLine(10.0F, 20.0F, 30.0F, 40.0F)
                shape.Width = 20.0F
                Assert.That(shape.Width, [Is].EqualTo(20.25F))
            End Sub

            <Test()> Public Sub Top_()
                Dim shape As Shape = sheet.Shapes.AddLine(10.0F, 20.0F, 30.0F, 40.0F)
                shape.Top = 20.0F
                Assert.That(shape.Top, [Is].EqualTo(20.25F))
            End Sub

            <Test()> Public Sub Left_()
                Dim shape As Shape = sheet.Shapes.AddLine(10.0F, 20.0F, 30.0F, 40.0F)
                shape.Left = 20.0F
                Assert.That(shape.Left, [Is].EqualTo(20.25F))
            End Sub

        End Class

    End Class
End Namespace