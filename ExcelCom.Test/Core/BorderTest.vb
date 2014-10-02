Imports NUnit.Framework
Imports System.Reflection

Namespace Core

    Public MustInherit Class BorderTest

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

        'Public Class CellsTest : Inherits BorderTest

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

        Public Class ExcelObjectたちTest : Inherits BorderTest

            '<Test()> Public Sub Cellsが閉じられること()
            '    Dim cells As Range = workbook.Sheets(0).Shapes

            '    sut.Dispose()
            '    sut = Nothing

            '    TestUtil.AssertNotExistsExcelPropcess()
            'End Sub

        End Class

        Public Class PropertyたちTest : Inherits BorderTest

            <Test()> Public Sub LineStyle_(<Values(Border.XlLineStyle.xlContinuous, Border.XlLineStyle.xlDash, Border.XlLineStyle.xlDashDot, Border.XlLineStyle.xlDashDotDot, Border.XlLineStyle.xlDot, Border.XlLineStyle.xlDouble, Border.XlLineStyle.xlLineStyleNone, Border.XlLineStyle.xlSlantDashDot)> ByVal index As Border.XlLineStyle)
                With sheet.Cells(4, 5).Borders(Border.XlBordersIndex.xlEdgeTop)
                    .LineStyle = index
                    Assert.That(.LineStyle, [Is].EqualTo(index))
                End With
            End Sub

            <Test()> Public Sub Color_Excel2003以前は56色(<Values(&HFFFFCC, &H99CCFF, &HFFCC99, &H663399)> ByVal color As Integer)
                With sheet.Cells(4, 5).Borders(Border.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Border.XlLineStyle.xlDash
                    .Color = color
                    Assert.That(.Color, [Is].EqualTo(color))
                End With
            End Sub

            <Test()> Public Sub ColorIndex_(<Values(1, 56, 20, 30)> ByVal index As Integer)
                With sheet.Cells(4, 5).Borders(Border.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Border.XlLineStyle.xlContinuous
                    .ColorIndex = index
                    Assert.That(.ColorIndex, [Is].EqualTo(index))
                End With
            End Sub

            <Test()> Public Sub ColorIndex_57以上はエラーになる(<Values(57, 100)> ByVal index As Integer)
                Try
                    With sheet.Cells(4, 5).Borders(Border.XlBordersIndex.xlEdgeTop)
                        .LineStyle = Border.XlLineStyle.xlContinuous
                        .ColorIndex = index
                        Assert.That(.ColorIndex, [Is].EqualTo(index))
                    End With
                    Assert.Fail()
                Catch expected As TargetInvocationException
                    Assert.That(expected.InnerException.Message, [Is].EqualTo("Border クラスの ColorIndex プロパティを設定できません。"))
                End Try
            End Sub

        End Class

    End Class
End Namespace