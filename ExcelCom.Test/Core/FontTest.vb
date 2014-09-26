Imports NUnit.Framework

Namespace Core

    Public MustInherit Class FontTest

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

        'Public Class CellsTest : Inherits FontTest

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

        Public Class ExcelObjectたちTest : Inherits FontTest

            '<Test()> Public Sub Cellsが閉じられること()
            '    Dim cells As Range = workbook.Sheets(0).Shapes

            '    sut.Dispose()
            '    sut = Nothing

            '    TestUtil.AssertNotExistsExcelPropcess()
            'End Sub

        End Class

        Public Class PropertyたちTest : Inherits FontTest

            <Test()> Public Sub Name_フォント名を設定する(<Values("MS P GOTHIC", "Arial", "Tahoma")> ByVal fontName As String)
                sheet.Cells(2, 3).Font.Name = fontName
                Assert.That(sheet.Cells(2, 3).Font.Name, [Is].EqualTo(fontName))
            End Sub

            <Test()> Public Sub Size_(<Values(10, 12.5, 14)> ByVal size As Double)
                sheet.Cells(2, 3).Font.Size = size
                Assert.That(sheet.Cells(2, 3).Font.Size, [Is].EqualTo(size))
            End Sub

            <Test()> Public Sub Bold_(<Values(True, False)> ByVal bold As Boolean)
                sheet.Cells(4, 5).Font.Bold = bold
                Assert.That(sheet.Cells(4, 5).Font.Bold, [Is].EqualTo(bold))
            End Sub

            <Test()> Public Sub Italic_(<Values(True, False)> ByVal italic As Boolean)
                sheet.Cells(4, 5).Font.Italic = italic
                Assert.That(sheet.Cells(4, 5).Font.Italic, [Is].EqualTo(italic))
            End Sub

            <Test()> Public Sub Shadow_(<Values(True, False)> ByVal shadow As Boolean)
                sheet.Cells(4, 5).Font.Shadow = shadow
                Assert.That(sheet.Cells(4, 5).Font.Shadow, [Is].EqualTo(shadow))
            End Sub

            <Test()> Public Sub Underline_(<Values(Font.XlUnderlineStyle.xlUnderlineStyleDouble, Font.XlUnderlineStyle.xlUnderlineStyleDoubleAccounting, Font.XlUnderlineStyle.xlUnderlineStyleNon, Font.XlUnderlineStyle.xlUnderlineStyleSingle, Font.XlUnderlineStyle.xlUnderlineStyleSingleAccounting)> ByVal underline As Font.XlUnderlineStyle)
                sheet.Cells(4, 5).Font.Underline = underline
                Assert.That(sheet.Cells(4, 5).Font.Underline, [Is].EqualTo(underline))
            End Sub

            <Test()> Public Sub Color_( _
                    <Values(0, 255, 128)> ByVal r As Integer, _
                    <Values(0, 255, 128)> ByVal g As Integer, _
                    <Values(0, 255, 128)> ByVal b As Integer)
                Dim color1 As Integer = RGB(r, g, b)
                sheet.Cells(4, 5).Value = "aa"
                sheet.Cells(4, 5).Font.Color = color1
                Assert.That(sheet.Cells(4, 5).Font.Color, [Is].EqualTo(color1))
            End Sub

            <Test()> Public Sub Strikethrough_(<Values(True, False)> ByVal strikethrough As Boolean)
                sheet.Cells(4, 5).Font.Strikethrough = strikethrough
                Assert.That(sheet.Cells(4, 5).Font.Strikethrough, [Is].EqualTo(strikethrough))
            End Sub

            <Test()> Public Sub Subscript_(<Values(True, False)> ByVal subscript As Boolean)
                sheet.Cells(4, 5).Font.Subscript = subscript
                Assert.That(sheet.Cells(4, 5).Font.Subscript, [Is].EqualTo(subscript))
            End Sub

            <Test()> Public Sub Superscript(<Values(True, False)> ByVal superscript As Boolean)
                sheet.Cells(4, 5).Font.Superscript = superscript
                Assert.That(sheet.Cells(4, 5).Font.Superscript, [Is].EqualTo(superscript))
            End Sub

        End Class

    End Class
End Namespace