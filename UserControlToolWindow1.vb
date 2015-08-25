Imports gcornu.Interop
Imports gcornu.Interop.VBAExtensibility
Imports System.Windows.Forms

Friend Class UserControlToolWindow1

    Private _VBE As VBE
    Private _Window As VBAExtensibility.Window

    Friend Sub Initialize(ByVal vbe As VBE, ByVal window As VBAExtensibility.Window)

        _VBE = vbe
        _Window = window

    End Sub

    Private Sub ButtonOK_Click(sender As System.Object, e As System.EventArgs) Handles ButtonOK.Click

        'MessageBox.Show("Toolwindow shown in VBA editor version " & _VBE.Version)
        _Window.Visible = False

    End Sub

End Class