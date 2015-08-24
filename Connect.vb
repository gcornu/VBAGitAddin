Imports gcornu.Interop
Imports gcornu.Interop.Extensibility
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports System.Drawing

<ComVisible(True), Guid("01234567-89ab-cdef-0123-456789abcdef"), ProgId("MyVBAAddin.Connect")> _
Public Class Connect
    Implements IDTExtensibility2

    Private _VBE As VBAExtensibility.VBE
    Private _AddIn As VBAExtensibility.AddIn
    Private WithEvents _CommandBarButton1 As CommandBarButton
    Private WithEvents _CommandBarButton2 As CommandBarButton

    Private _toolWindow1 As VBAExtensibility.Window
    Private _toolWindow2 As VBAExtensibility.Window

    Private Sub OnConnection(Application As Object, ConnectMode As ext_ConnectMode, AddInInst As Object, _
       ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection

        Try

            _VBE = DirectCast(Application, VBAExtensibility.VBE)
            _AddIn = DirectCast(AddInInst, VBAExtensibility.AddIn)

            Select Case ConnectMode

                Case ext_ConnectMode.ext_cm_Startup
                    ' OnStartupComplete will be called

                Case ext_ConnectMode.ext_cm_AfterStartup
                    InitializeAddIn()

            End Select

        Catch ex As Exception

            MessageBox.Show(ex.ToString())

        End Try

    End Sub

    Private Sub OnDisconnection(RemoveMode As ext_DisconnectMode, _
       ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection

        If Not _CommandBarButton1 Is Nothing Then

            _CommandBarButton1.Delete()
            _CommandBarButton1 = Nothing

        End If

        If Not _CommandBarButton2 Is Nothing Then

            _CommandBarButton2.Delete()
            _CommandBarButton2 = Nothing

        End If

    End Sub

    Private Sub OnStartupComplete(ByRef custom As System.Array) _
       Implements IDTExtensibility2.OnStartupComplete

        InitializeAddIn()

    End Sub

    Private Sub OnAddInsUpdate(ByRef custom As System.Array) _
       Implements IDTExtensibility2.OnAddInsUpdate

    End Sub

    Private Sub OnBeginShutdown(ByRef custom As System.Array) Implements IDTExtensibility2.OnBeginShutdown

    End Sub

    Private Sub InitializeAddIn()

        Dim standardCommandBar As CommandBar
        Dim commandBarControl As CommandBarControl

        Try

            standardCommandBar = _VBE.CommandBars.Item("Standard")

            commandBarControl = standardCommandBar.Controls.Add(MsoControlType.msoControlButton)
            _CommandBarButton1 = DirectCast(commandBarControl, CommandBarButton)
            _CommandBarButton1.Caption = "Toolwindow 1"
            _CommandBarButton1.FaceId = 59
            _CommandBarButton1.Style = MsoButtonStyle.msoButtonIconAndCaption
            _CommandBarButton1.BeginGroup = True

            commandBarControl = standardCommandBar.Controls.Add(MsoControlType.msoControlButton)
            _CommandBarButton2 = DirectCast(commandBarControl, CommandBarButton)
            _CommandBarButton2.Caption = "Toolwindow 2"
            _CommandBarButton2.FaceId = 59
            _CommandBarButton2.Style = MsoButtonStyle.msoButtonIconAndCaption
            _CommandBarButton2.BeginGroup = True

        Catch ex As Exception

            MessageBox.Show(ex.ToString())

        End Try

    End Sub

    Private Function CreateToolWindow(ByVal toolWindowCaption As String, ByVal toolWindowGuid As String, _
       ByVal toolWindowUserControl As UserControl) As VBAExtensibility.Window

        Dim userControlObject As Object = Nothing
        Dim userControlHost As UserControlHost
        Dim toolWindow As VBAExtensibility.Window
        Dim progId As String

        ' IMPORTANT: ensure that you use the same ProgId value used in the ProgId attribute of the UserControlHost class
        progId = "MyVBAAddin.UserControlHost"

        toolWindow = _VBE.Windows.CreateToolWindow(_AddIn, progId, toolWindowCaption, toolWindowGuid, userControlObject)
        userControlHost = DirectCast(userControlObject, UserControlHost)

        toolWindow.Visible = True

        userControlHost.AddUserControl(toolWindowUserControl)

        Return toolWindow

    End Function

    Private Sub _CommandBarButton1_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, _
       ByRef CancelDefault As Boolean) Handles _CommandBarButton1.Click

        Dim userControlObject As Object = Nothing
        Dim userControlToolWindow1 As UserControlToolWindow1

        Try

            If _toolWindow1 Is Nothing Then

                userControlToolWindow1 = New UserControlToolWindow1()

                _toolWindow1 = CreateToolWindow("My toolwindow 1", "21234567-89ab-cdef-0123-456789abcdef", userControlToolWindow1)

                userControlToolWindow1.Initialize(_VBE)

            Else

                _toolWindow1.Visible = True

            End If

        Catch ex As Exception

            MessageBox.Show(ex.ToString)

        End Try

    End Sub

    Private Sub _CommandBarButton2_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, _
       ByRef CancelDefault As Boolean) Handles _CommandBarButton2.Click

        Dim userControlObject As Object = Nothing
        Dim userControlToolWindow2 As UserControlToolWindow2

        Try

            If _toolWindow2 Is Nothing Then

                userControlToolWindow2 = New UserControlToolWindow2()

                _toolWindow2 = CreateToolWindow("My toolwindow 2", "31234567-89ab-cdef-0123-456789abcdef", userControlToolWindow2)

                userControlToolWindow2.Initialize(_VBE)

            Else

                _toolWindow2.Visible = True

            End If

        Catch ex As Exception

            MessageBox.Show(ex.ToString)

        End Try

    End Sub

End Class