Imports NUnit.Framework

Namespace Core

    Public MustInherit Class CommentTest

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

        'Public Class CellsTest : Inherits CommentTest

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

        Public Class TextTest : Inherits CommentTest

            <Test()> Public Sub AddCommentの値を参照できる(<Values("a", "xxxxxxx", "01234", "あいう")> ByVal text As String)
                Dim aComment As Comment = sheet.Cells(4, 5).AddComment(text)
                With aComment
                    Assert.That(.Text, [Is].EqualTo(text))
                End With
            End Sub

            <Test()> Public Sub Textの設定値を参照できる(<Values("a", "xxxxxxx", "01234", "あいう")> ByVal text As String)
                Dim aComment As Comment = sheet.Cells(4, 5).AddComment
                With aComment
                    .Text(aText:=text)
                    Assert.That(.Text, [Is].EqualTo(text))
                End With
            End Sub

        End Class

        Public Class ExcelObjectたちTest : Inherits CommentTest

            <Test()> Public Sub Shapeが閉じられること()
                Dim aComment As Comment = sheet.Cells(5, 6).AddComment
                Dim shape As Shape = aComment.Shape

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

        Public Class PropertyたちTest : Inherits CommentTest

            <Test()> Public Sub Author_()
                Dim author As String = sheet.Cells(5, 6).AddComment.Author
                Assert.That(author, [Is].Not.Empty, "ログインユーザー名")
            End Sub

        End Class

    End Class
End Namespace