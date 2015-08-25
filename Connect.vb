Imports gcornu.Interop
Imports gcornu.Interop.Extensibility
Imports System.Windows.Forms
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Core
Imports System.Drawing

<ComVisible(True), Guid("2A0D7296-C26B-4EB1-A82F-F97ACC6A3917"), ProgId("MyVBAAddin.Connect")> _
Public Class Connect
    Implements IDTExtensibility2

    Private _VBE As VBAExtensibility.VBE
    Private _AddIn As VBAExtensibility.AddIn
    Private WithEvents _CommandBarButtonCommit As CommandBarButton
    Private WithEvents _CommandBarButtonPush As CommandBarButton
    Private WithEvents _CommandBarPopupButton As CommandBarButton
    Private _CommandBarPopup As CommandBarPopup

    Private _toolWindow1 As VBAExtensibility.Window

    Private _GitManager As GitManager

    Private Sub OnConnection(Application As Object, ConnectMode As ext_ConnectMode, AddInInst As Object, ByRef custom As System.Array) Implements IDTExtensibility2.OnConnection

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

    Private Sub OnDisconnection(RemoveMode As ext_DisconnectMode, ByRef custom As System.Array) Implements IDTExtensibility2.OnDisconnection

        If Not _CommandBarButtonCommit Is Nothing Then

            _CommandBarButtonCommit.Delete()
            _CommandBarButtonCommit = Nothing

        End If

        If Not _CommandBarButtonPush Is Nothing Then

            _CommandBarButtonPush.Delete()
            _CommandBarButtonPush = Nothing

        End If

        If Not _CommandBarPopupButton Is Nothing Then

            _CommandBarPopupButton.Delete()
            _CommandBarPopupButton = Nothing

        End If

        If Not _CommandBarPopup Is Nothing Then

            _CommandBarPopup.Delete()
            _CommandBarPopup = Nothing

        End If

        If Not _toolWindow1 Is Nothing Then

            _toolWindow1.Delete()
            _toolWindow1 = Nothing

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
        Dim toolsCommandBar As CommandBar
        Dim toolsCommandBarControl As CommandBarControl

        Const MY_COMMANDBAR_POPUP1_NAME As String = "CommandBarPopup"
        Const MY_COMMANDBAR_POPUP1_CAPTION As String = "Git For VBA"
        Const TOOLS_COMMANDBAR_NAME As String = "Tools"


        Try

            'New toolab containing addin buttons
            standardCommandBar = _VBE.CommandBars.Item("Standard")

            'Commit button
            _CommandBarButtonCommit = AddCommandBarButton(standardCommandBar, "Commit")

            'Push button
            _CommandBarButtonPush = AddCommandBarButton(standardCommandBar, "Push")


            'Menu item in "Tools" menu
            toolsCommandBar = _VBE.CommandBars.Item(TOOLS_COMMANDBAR_NAME)

            ' Add a new commandbar popup 
            _CommandBarPopup = DirectCast(toolsCommandBar.Controls.Add( _
               MsoControlType.msoControlPopup, System.Type.Missing, System.Type.Missing, _
               toolsCommandBar.Controls.Count + 1, True), CommandBarPopup)

            ' Change some commandbar popup properties
            _CommandBarPopup.CommandBar.Name = MY_COMMANDBAR_POPUP1_NAME
            _CommandBarPopup.Caption = MY_COMMANDBAR_POPUP1_CAPTION

            ' Add a new button on that commandbar popup
            _CommandBarPopupButton = AddCommandBarButton(_CommandBarPopup.CommandBar, "Options")

            ' Make visible the commandbar popup
            _CommandBarPopup.Visible = True

            ' Calculate the position of a new commandbar popup to the right of the "Tools" menu
            toolsCommandBarControl = DirectCast(toolsCommandBar.Parent, CommandBarControl)

            _GitManager = New GitManager

        Catch ex As Exception

            MessageBox.Show(ex.ToString())

        End Try

    End Sub

    Private Function AddCommandBarButton(ByVal commandBar As CommandBar, caption As String) As CommandBarButton

        Dim commandBarButton As CommandBarButton
        Dim commandBarControl As CommandBarControl

        commandBarControl = commandBar.Controls.Add(MsoControlType.msoControlButton)
        commandBarButton = DirectCast(commandBarControl, CommandBarButton)

        commandBarButton.Caption = caption
        commandBarButton.Style = MsoButtonStyle.msoButtonCaption

        Return commandBarButton

    End Function

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

    Private Sub _CommandBarButtonCommit_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles _CommandBarButtonCommit.Click

        _GitManager.PerformGitActions()

    End Sub

    Private Sub _CommandBarButtonPush_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean) Handles _CommandBarButtonPush.Click

        MessageBox.Show("This should git push")

    End Sub

    Private Sub _CommandBarPopupButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, _
       ByRef CancelDefault As Boolean) Handles _CommandBarPopupButton.Click

        Dim userControlObject As Object = Nothing
        Dim userControlToolWindow1 As UserControlToolWindow1

        Try

            If _toolWindow1 Is Nothing Then

                userControlToolWindow1 = New UserControlToolWindow1()

                _toolWindow1 = CreateToolWindow("Options", "5655769A-EB81-436B-8EBB-05183EBDEA79", userControlToolWindow1)

                userControlToolWindow1.Initialize(_VBE, _toolWindow1)

            Else

                _toolWindow1.Visible = True

            End If

        Catch ex As Exception

            MessageBox.Show(ex.ToString)

        End Try

    End Sub



End Class