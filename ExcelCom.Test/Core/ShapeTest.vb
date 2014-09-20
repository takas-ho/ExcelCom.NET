Imports NUnit.Framework

Namespace Core

    Public MustInherit Class ShapeTest

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

            '<Test()> Public Sub Visible(<Values(Worksheet.XlSheetVisibility.xlSheetHidden, Worksheet.XlSheetVisibility.xlSheetVeryHidden, Worksheet.XlSheetVisibility.xlSheetVisible)> ByVal value As Worksheet.XlSheetVisibility)
            '    Dim sheet1 As Worksheet = workbook.Sheets.Add()
            '    Dim sheet2 As Worksheet = workbook.Sheets.Add(after:=sheet1)
            '    sheet2.Select()

            '    workbook.Sheets(0).Visible = value
            '    Assert.That(workbook.Sheets(0).Visible, [Is].EqualTo(value))
            'End Sub

        End Class

    End Class
End Namespace