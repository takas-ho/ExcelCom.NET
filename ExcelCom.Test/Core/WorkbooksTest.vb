Imports NUnit.Framework

Namespace Core

    Public MustInherit Class WorkbooksTest

        Private sut As Application

        <SetUp()> Public Sub SetUp()
            sut = New Application
        End Sub

        <TearDown()> Public Sub TearDown()
            If sut Is Nothing Then
                Return
            End If
            sut.Dispose()
        End Sub

        Public Class Add_Test : Inherits WorkbooksTest

            <Test()> Public Sub Addすれば1増える()
                Dim baseCount As Integer = sut.Workbooks.Count
                sut.Workbooks.Add()
                Assert.That(sut.Workbooks.Count, [Is].EqualTo(baseCount + 1))
            End Sub

            <Test()> Public Sub Addしたら先頭に追加される()
                Dim addedSheet As Workbook = sut.Workbooks.Add()
                Assert.That(sut.Workbooks(0).Name, [Is].EqualTo(addedSheet.Name))
            End Sub

            <Test()> Public Sub Addを2度したら_2度目は後ろに追加される()
                Dim firstSheet As Workbook = sut.Workbooks.Add()
                Dim secondSheet As Workbook = sut.Workbooks.Add()
                Assert.That(sut.Workbooks(0).Name, [Is].EqualTo(firstSheet.Name))
                Assert.That(sut.Workbooks(1).Name, [Is].EqualTo(secondSheet.Name))
            End Sub

        End Class

        Public Class プロパティたちTest : Inherits WorkbooksTest

            <Test()> Public Sub Count_最初は0()
                Assert.That(sut.Workbooks.Count, [Is].EqualTo(0))
            End Sub

            <Test()> Public Sub Count_Addすれば1になる()
                sut.Workbooks.Add()
                Assert.That(sut.Workbooks.Count, [Is].EqualTo(1))
            End Sub

        End Class

        Public Class ExcelObjectたちTest : Inherits WorkbooksTest

            <Test()> Public Sub AddのWorkbookが閉じられること()
                Dim workbook As Workbook = sut.Workbooks.Add

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

            <Test()> Public Sub ItemのWorkbookが閉じられること()
                sut.Workbooks.Add()
                Dim workbook As Workbook = sut.Workbooks.Item(0)

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

    End Class
End Namespace