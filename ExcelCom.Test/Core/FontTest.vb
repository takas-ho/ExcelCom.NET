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

            '<Test()> Public Sub Color_( _
            '        <Values(0, 255, 128)> ByVal r As Integer, _
            '        <Values(0, 255, 128)> ByVal g As Integer, _
            '        <Values(0, 255, 128)> ByVal b As Integer)
            '    Dim color1 As Integer = g * 256 * 256 + b * 256 + r
            '    'Assert.That(color1, [Is].EqualTo(g * 256 * 256 + b * 256 + r))
            '    sheet.Cells(4, 5).Value = "aa"
            '    sheet.Cells(4, 5).Font.Color = color1
            '    Assert.That(sheet.Cells(4, 5).Font.Color, [Is].EqualTo(color1))
            'End Sub

            'Public Function ConvHoge(ByVal color As Integer) As Integer
            '    Dim r As Integer = color Mod 256
            '    Dim g As Integer = CInt(Math.Floor(color / 256)) Mod 256
            '    Dim b As Integer = CInt(Math.Floor(Math.Floor(color / 256) / 256)) Mod 256
            '    Dim color2 As Integer = 256 * 256 * r + 256 * g + b
            '    'Return color2.ToString("X6")
            '    Return color2
            'End Function

            <Test()> Public Sub ColorIndex_(<Values(1, 49, 20, 30)> ByVal index As Integer)
                sheet.Cells(4, 5).Font.ColorIndex = index
                Assert.That(sheet.Cells(4, 5).Font.ColorIndex, [Is].EqualTo(index))
            End Sub

            <Test()> Public Sub ColorIndex_ゼロは指定できない_マイナス4105になる()
                sheet.Cells(4, 5).Font.ColorIndex = 0
                Assert.That(sheet.Cells(4, 5).Font.ColorIndex, [Is].EqualTo(-4105))
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