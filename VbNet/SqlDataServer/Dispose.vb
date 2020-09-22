Option Explicit On 
Option Strict On
Public MustInherit Class ReleaseObj
    Implements IDisposable
    Public Overloads Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
    Protected Overridable Overloads Sub Dispose(ByVal Disposing As Boolean)
        If Disposing Then
        End If
    End Sub
    Protected Overrides Sub Finalize()
        Dispose(False)
    End Sub
End Class