Imports System.IO
Imports NUnit.Framework

Namespace Core

    Public MustInherit Class WorkbookTest

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

        Public Class ExcelObjectたちTest : Inherits WorkbookTest

            <Test()> Public Sub Sheetsが閉じられること()
                Dim workbook As Workbook = sut.Workbooks.Add
                Dim sheets As Sheets = workbook.Sheets

                sut.Dispose()
                sut = Nothing

                TestUtil.AssertNotExistsExcelPropcess()
            End Sub

        End Class

        Public Class PropertyたちTest : Inherits WorkbookTest

            <Test()> Public Sub Name_追加した直後はBook1()
                Dim book As Workbook = sut.Workbooks.Add
                Assert.That(book.Name, [Is].EqualTo("Book1"))
            End Sub

            <Test()> Public Sub Name_3度追加すればBook3()
                Dim book As Workbook = sut.Workbooks.Add
                Dim book2 As Workbook = sut.Workbooks.Add
                Dim book3 As Workbook = sut.Workbooks.Add
                Assert.That(book3.Name, [Is].EqualTo("Book3"))
            End Sub

            <Test()> Public Sub MultiUserEditing()
                Dim book As Workbook = sut.Workbooks.Add
                Assert.That(book.MultiUserEditing, [Is].False, "追加したbookはfalse")
            End Sub

            <Test()> Public Sub Path()
                Dim book As Workbook = sut.Workbooks.Add
                Assert.That(book.Path, [Is].EqualTo(""), "追加したbookのパスは空")
            End Sub

            <Test()> Public Sub [ReadOnly]()
                Dim book As Workbook = sut.Workbooks.Add
                Assert.That(book.ReadOnly, [Is].False, "追加したbookのReadOnlyはfalse")
            End Sub

            <Test()> Public Sub HasPassword()
                Dim book As Workbook = sut.Workbooks.Add
                Assert.That(book.HasPassword, [Is].False, "追加したbookのHasPasswordはfalse")
            End Sub

            <Test()> Public Sub ProtectWindows()
                Dim book As Workbook = sut.Workbooks.Add
                Assert.That(book.ProtectWindows, [Is].False, "追加したbookのProtectWindowsはfalse")
            End Sub

            <Test()> Public Sub RevisionNumber()
                Dim book As Workbook = sut.Workbooks.Add
                Assert.That(book.RevisionNumber, [Is].EqualTo(0), "追加したbookのRevisionNumberは0")
            End Sub

            <Test()> Public Sub UpdateRemoteReferences(<Values(False, True)> ByVal value As Boolean)
                Dim book As Workbook = sut.Workbooks.Add
                Assert.That(book.UpdateRemoteReferences, [Is].True, "追加したbookのUpdateRemoteReferencesはtrue")
                book.UpdateRemoteReferences = value
                Assert.That(book.UpdateRemoteReferences, [Is].EqualTo(value))
            End Sub

        End Class

        Public Class その他細かいTest : Inherits WorkbookTest

            <Test()> Public Sub CloseしたらBookが閉じる()
                Dim workbook As Workbook = sut.Workbooks.Add
                workbook.Close()

                Assert.That(sut.Workbooks.Count, [Is].EqualTo(0))
            End Sub

            <Test()> Public Sub SaveAsで指定ファイル名に保存する()
                Const FILE_NAME As String = "a.xls"
                If File.Exists(FILE_NAME) Then
                    File.Delete(FILE_NAME)
                End If
                Dim workbook As Workbook = sut.Workbooks.Add
                workbook.Sheets(0).Cells(4, 4).Value = "ABc"
                workbook.SaveAs(FILE_NAME)
                Try
                    workbook.Close()

                    sut.Workbooks.Open(FILE_NAME)
                    Assert.That(sut.Workbooks(0).Sheets(0).Cells(4, 4).Value, [Is].EqualTo("ABc"))

                Finally
                    File.Delete(FILE_NAME)
                End Try

            End Sub

        End Class

    End Class
End Namespace