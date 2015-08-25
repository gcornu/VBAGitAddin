Imports gcornu.Interop
Imports gcornu.Interop.Extensibility

Public Class EventHandler
    Private _Window As VBAExtensibility.Window

    Friend Sub Close()
        _Window.Visible = False
    End Sub
End Class
