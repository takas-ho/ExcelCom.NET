Namespace Core

    Public MustInherit Class AbstractExcelSubObject : Inherits AbstractExcelObject : Implements IExcelObject

        Protected ReadOnly Parent As IExcelObject

        Public Sub New(ByVal parent As IExcelObject, ByVal comObject As Object)
            MyBase.New(comObject)
            Me.Parent = parent
            parent.AddToManager(comObject)
        End Sub

        Public Overridable Sub AddToManager(ByVal comObject As Object) Implements IExcelObject.AddToManager
            Parent.AddToManager(comObject)
        End Sub

    End Class
End Namespace