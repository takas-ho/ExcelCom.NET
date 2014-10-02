Imports NUnit.Framework
Imports System.Reflection

Namespace Core

    Public MustInherit Class BordersTest

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

        Public Class ExcelObjectたちTest : Inherits BordersTest

            '<Test()> Public Sub Lineが閉じられること()
            '    Dim Border As Border = workbook.Sheets(0).Borders.AddLine(0, 10, 20, 30)

            '    sut.Dispose()
            '    sut = Nothing

            '    TestUtil.AssertNotExistsExcelPropcess()
            'End Sub

        End Class

        Public Class PropertyたちTest : Inherits BordersTest

            <Test()> Public Sub Count_最初は0超()
                Assert.That(sheet.Cells(3, 4).Borders.Count, [Is].GreaterThan(0))
            End Sub

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

            <Test()> Public Sub Weight_(<Values(Border.XlBorderWeight.xlHairline, Border.XlBorderWeight.xlMedium, Border.XlBorderWeight.xlThick, Border.XlBorderWeight.xlThin)> ByVal weight As Border.XlBorderWeight)
                With sheet.Cells(4, 5).Borders(Border.XlBordersIndex.xlEdgeTop)
                    .LineStyle = Border.XlLineStyle.xlContinuous
                    .Weight = weight
                    Assert.That(.Weight, [Is].EqualTo(weight))
                End With
            End Sub

        End Class

    End Class
End Namespace