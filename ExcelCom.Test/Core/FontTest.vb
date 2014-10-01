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

            ' 0:000000 1:000000 2:FFFFFF 3:0000FF 4:00FF00 5:FF0000 6:00FFFF 7:FF00FF 8:FFFF00
            '  9:000080 10:008000 11:800000 12:008080 13:800080 14:808000 15:C0C0C0 16:808080
            ' 17:FF9999 18:663399 19:CCFFFF 20:FFFFCC 21:660066 22:8080FF 23:CC6600 24:FFCCCC
            ' 25:800000 26:FF00FF 27:00FFFF 28:FFFF00 29:800080 30:000080 31:808000 32:FF0000
            ' 33:FFCC00 34:FFFFCC 35:CCFFCC 36:99FFFF 37:FFCC99 38:CC99FF 39:FF99CC 40:99CCFF
            ' 41:FF6633 42:CCCC33 43:00CC99 44:00CCFF 45:0099FF 46:0066FF 47:996666 48:969696
            ' 49:663300 50:669933 51:003300 52:003333 53:003399 54:663399 55:993333 56:333333
            <Test()> Public Sub Color_Excel2003以前は56色(<Values(&HFFFFCC, &H99CCFF, &HFFCC99, &H663399)> ByVal color As Integer)
                sheet.Cells(4, 5).Font.Color = color
                Assert.That(sheet.Cells(4, 5).Font.Color, [Is].EqualTo(color))
            End Sub

            <Test()> Public Sub ColorIndex_(<Values(1, 49, 20, 30)> ByVal index As Integer)
                sheet.Cells(4, 5).Font.ColorIndex = index
                Assert.That(sheet.Cells(4, 5).Font.ColorIndex, [Is].EqualTo(index))
            End Sub

            <Test()> Public Sub ColorIndex_ゼロは指定できない_マイナス4105になる()
                For i As Integer = 0 To 56
                    sheet.Cells(4, 5).Font.ColorIndex = i
                    Debug.Print(String.Format("{0}:{1}", i, sheet.Cells(4, 5).Font.Color.ToString("X6")))
                Next
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