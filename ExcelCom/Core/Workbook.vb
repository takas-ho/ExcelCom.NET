Namespace Core

    Public Class Workbook : Inherits AbstractExcelSubObject : Implements IExcelObject

        Private ReadOnly books As Workbooks

        Public Sub New(ByVal parent As Workbooks, ByVal comObject As Object)
            MyBase.New(parent, comObject)
            books = parent
        End Sub

        Public Sub Activate()
            InvokeMethod("Activate")
        End Sub

        ''' <summary>
        ''' 閉じる
        ''' </summary>
        ''' <param name="saveChanges">保存して閉じる場合、true</param>
        ''' <param name="filename">保存する場合のファイル名</param>
        ''' <remarks></remarks>
        Public Sub Close(Optional ByVal saveChanges As Boolean = False, Optional ByVal filename As String = Nothing)
            Dim args As New List(Of Object)
            args.Add(New NamedParameter("SaveChanges", saveChanges))
            If filename IsNot Nothing Then
                args.Add(New NamedParameter("Filename", filename))
            End If
            InvokeMethod("Close", args.ToArray)
            If books.InternalItems.Contains(Me) Then
                books.InternalItems.Remove(Me)
            End If
        End Sub

        Private _sheets As Sheets
        Public Function Sheets() As Sheets
            If _sheets Is Nothing Then
                _sheets = New Sheets(Me, InvokeGetProperty("Sheets"))
            End If
            Return _sheets
        End Function

        ''' <summary>
        ''' 保存する
        ''' </summary>
        ''' <param name="fileName">ファイル名</param>
        ''' <param name="fileFormat">ファイル形式</param>
        ''' <param name="password">設定する読込パスワード</param>
        ''' <param name="writeResPassword">設定する書込みパスワード</param>
        ''' <remarks></remarks>
        Public Sub SaveAs(ByVal fileName As String, Optional ByVal fileFormat As XlFileFormat = Nothing, _
                          Optional ByVal password As String = Nothing, Optional ByVal writeResPassword As String = Nothing)
            Dim args As New List(Of Object)
            args.Add(fileName)
            If fileFormat <> Nothing Then
                args.Add(New NamedParameter("FileFormat", fileFormat))
            End If
            If password IsNot Nothing Then
                args.Add(New NamedParameter("Password", password))
            End If
            If writeResPassword IsNot Nothing Then
                args.Add(New NamedParameter("WriteResPassword", writeResPassword))
            End If
            InvokeMethod("SaveAs", args.ToArray)
        End Sub

    End Class
End Namespace